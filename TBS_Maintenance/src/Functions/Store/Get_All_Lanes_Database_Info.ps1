# ===================================================================================================
#                              FUNCTION: Get_All_Lanes_Database_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Collects per-lane SQL details (machine name, DB server, DB name) for every lane of the
#   store. Uses in-memory results first, then the per-function cache file, and only queries
#   the lanes when both are missing or -RefreshCache is set. With -LaneNumber it returns
#   that lane's info object.
#
# Parameters:
#   - LaneNumber   [string]: return only this lane's data (3-digit, padded)
#   - RefreshCache [switch]: force re-query and cache rewrite
#
# Variables:
#   - $script:FunctionResults['LaneDatabaseInfo']: per-lane info bucket (output)
#   - $script:FunctionResults['LaneNumToMachineName']: lane -> machine map (input)
#
# Returns:
#   The requested lane's info object; nothing for a full-store refresh.
#
# Writes:
#   - Setup_Files\Get_All_Lanes_Database_Info.txt (JSON cache, ordered by lane number)
# ===================================================================================================

function Get_All_Lanes_Database_Info
{
	param (
		[string]$LaneNumber,
		[switch]$RefreshCache
	)

	# Ensure bucket exists (unchanged behavior)
	if (-not $script:FunctionResults.ContainsKey('LaneDatabaseInfo'))
	{
		$script:FunctionResults['LaneDatabaseInfo'] = @{ }
	}
	$LaneDatabaseInfo = $script:FunctionResults['LaneDatabaseInfo']

	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName) { return $null }

	# Guard against null OfficePath causing Join-Path errors (fast fail, no guessing)
	if ([string]::IsNullOrWhiteSpace($OfficePath)) { return $null }

	# ------------------------------------------------------------------------------------------------
	# Per-function cache file (only run queries if memory/file data is missing or -RefreshCache is used)
	# ------------------------------------------------------------------------------------------------
	$CachePath = Join-Path $script:SetupFiles 'Get_All_Lanes_Database_Info.txt'

	# Normalize LaneNumber if provided (keep it simple + safe)
	$RequestedLane = $null
	if ($LaneNumber)
	{
		$digits = ($LaneNumber -replace '[^\d]', '')
		if ($digits) { $RequestedLane = $digits.PadLeft(3, '0') }
	}

	if (-not $RefreshCache)
	{
		# Reuse in-memory data first
		if ($RequestedLane -and $LaneDatabaseInfo.ContainsKey($RequestedLane))
		{
			return $LaneDatabaseInfo[$RequestedLane]
		}
		if ((-not $RequestedLane) -and $LaneDatabaseInfo.Count -gt 0)
		{
			return
		}

		# Reuse file cache if present
		if ($CachePath -and (Test-Path $CachePath))
		{
			try
			{
				$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
				if (-not [string]::IsNullOrWhiteSpace($raw))
				{
					$cacheObj = $raw | ConvertFrom-Json -ErrorAction Stop
					if ($cacheObj -and $cacheObj.PSObject.Properties.Name -contains 'LaneDatabaseInfo' -and $cacheObj.LaneDatabaseInfo)
					{
						foreach ($p in $cacheObj.LaneDatabaseInfo.PSObject.Properties)
						{
							if (-not $LaneDatabaseInfo.ContainsKey($p.Name))
							{
								# Same shape as a fresh lookup (hashtable): callers index it as
								# $info['TcpConnStr'], which returns nothing on a PSCustomObject
								$laneHash = @{ }
								if ($p.Value -and $p.Value.PSObject) { foreach ($q in $p.Value.PSObject.Properties) { $laneHash[$q.Name] = $q.Value } }
								# Caches written before v2.7.0 lack TrustServerCertificate
								foreach ($k in @('ConnectionString', 'NamedPipesConnStr', 'TcpConnStr'))
								{
									if ($laneHash[$k] -and $laneHash[$k] -notmatch '(?i)TrustServerCertificate') { $laneHash[$k] = ([string]$laneHash[$k]).TrimEnd(';') + ';TrustServerCertificate=True;' }
								}
								$LaneDatabaseInfo[$p.Name] = $laneHash
							}
						}
					}
				}
			}
			catch { }

			if ($RequestedLane)
			{
				if ($LaneDatabaseInfo.ContainsKey($RequestedLane)) { return $LaneDatabaseInfo[$RequestedLane] }
			}
			elseif ($LaneDatabaseInfo.Count -gt 0)
			{
				return
			}
		}
	}

	# ------------------------------------------------------------------------------------------------
	# Memory/file cache is missing/unreadable, or caller requested -RefreshCache => query lanes ONCE, then write file
	# ------------------------------------------------------------------------------------------------
	$lanesToProcess = @()

	# Always build a full cache when missing
	$lanesToProcess += $LaneNumToMachineName.Keys

	# Ensure requested lane is included even if it would normally be skipped in "full mode"
	if ($RequestedLane) { $lanesToProcess = @($RequestedLane) + $lanesToProcess }

	# De-dupe + keep only numeric 3-digit lanes
	$lanesToProcess = $lanesToProcess | ForEach-Object {
		$d = ($_.ToString() -replace '[^\d]', '')
		if ($d) { $d.PadLeft(3, '0') }
	} | Where-Object { $_ } | Select-Object -Unique

	$changed = $false

	# Lanes to read: all of them with -RefreshCache (an entry is kept when its lane cannot be
	# read), otherwise only the ones not in memory yet. 8xx (scales) / 9xx (server, backoffices)
	# are not lanes - unless explicitly requested.
	$work = @()
	foreach ($laneNumber in $lanesToProcess)
	{
		if (-not $RefreshCache -and $LaneDatabaseInfo.ContainsKey($laneNumber)) { continue }
		if ($laneNumber -ne $RequestedLane -and $laneNumber -match '^(8|9)') { continue }
		$machineName = $LaneNumToMachineName[$laneNumber]
		if (-not $machineName) { continue }
		$work += [pscustomobject]@{ Lane = $laneNumber; Machine = [string]$machineName; Local = (Join-Path $OfficePath "INI\STARTUP.$laneNumber") }
	}

	# Lanes without a local STARTUP.<lane> copy are read over \\lane\storeman: probe port 445 of
	# all of them at once (1.5 s in total, not 1.5 s per switched-off lane)
	$reachable = @{ }
	$probes = @()
	foreach ($w in $work)
	{
		if (Test-Path -LiteralPath $w.Local) { continue }
		$client = New-Object System.Net.Sockets.TcpClient
		try { $probes += [pscustomobject]@{ Lane = $w.Lane; Client = $client; Task = $client.ConnectAsync($w.Machine, 445) } }
		catch { $client.Dispose() }
	}
	$deadline = (Get-Date).AddMilliseconds(1500)
	foreach ($p in $probes)
	{
		$ok = $false
		try
		{
			$left = [int][Math]::Max(0, ($deadline - (Get-Date)).TotalMilliseconds)
			$ok = ($p.Task.Wait($left) -and $p.Client.Connected)
		}
		catch { $ok = $false }
		$reachable[$p.Lane] = $ok
		try { $p.Client.Dispose() }
		catch { }
	}

	foreach ($w in $work)
	{
		$laneNumber = $w.Lane
		$machineName = $w.Machine

		# ---------- 1. LOCAL Startup file (INI\STARTUP.###), 2. REMOTE \\lane\storeman\Startup.ini ----------
		$content = $null
		$source = $null
		if (Test-Path -LiteralPath $w.Local)
		{
			try
			{
				$content = Get-Content -LiteralPath $w.Local -ErrorAction Stop
				$source = "local"
			}
			catch { $content = $null }
		}
		if (-not $content -and $reachable[$laneNumber])
		{
			try
			{
				$content = Get-Content -LiteralPath "\\$machineName\storeman\Startup.ini" -ErrorAction Stop
				$source = "remote"
			}
			catch { $content = $null }
		}
		if (-not $content) { continue }

		# First line only: a second DBNAME= line would turn the value into an array
		$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' } | Select-Object -First 1
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' } | Select-Object -First 1

		if ($dbNameLine)
		{
			$dbName = ($dbNameLine -replace '^DBNAME=', '').Trim()
		}
		else
		{
			continue
		}

		if ($dbServerLine)
		{
			$dbServerRaw = ($dbServerLine -replace '^DBSERVER=', '').Trim()
		}
		else
		{
			$dbServerRaw = ""
		}

		# -------------- PATCH: Remove trailing backslash unless a named instance is present --------------
		if ($dbServerRaw -match '^[^\\]+\\$')
		{
			$dbServerRaw = $dbServerRaw.TrimEnd('\')
		}
		# PATCH: If (LOCAL) or localhost (any case), replace with actual machine name
		if ($dbServerRaw -match '^(?i)\(LOCAL\)$' -or $dbServerRaw -match '^(?i)localhost$' -or $dbServerRaw -eq "")
		{
			$dbServer = $machineName
		}
		else
		{
			$dbServer = $dbServerRaw
		}
		# END PATCH --------------------------------------------------------------------------------------

		# Parse instance name for Named Pipes/TCP logic
		$serverName = $dbServer
		$instanceName = $null
		if ($dbServer -match '\\')
		{
			$parts = $dbServer -split '\\'
			$serverName = $parts[0]
			$instanceName = $parts[1]
		}
		elseif ($dbServer -match ',')
		{
			$serverName = $dbServer
			$instanceName = $null
		}
		else
		{
			$serverName = $dbServer
			$instanceName = $null
		}

		# Build connection strings
		if ($instanceName -and $instanceName.ToUpper() -ne "MSSQLSERVER")
		{
			$namedPipes = "np:\\$serverName\pipe\MSSQL`$$instanceName\sql\query"
			$tcpServer = "$serverName\$instanceName"
		}
		else
		{
			$namedPipes = "np:\\$serverName\pipe\sql\query"
			$tcpServer = $serverName
		}

		# TrustServerCertificate: SqlServer module 22+ encrypts by default and the lanes use
		# self-signed certificates, so without it every Invoke-Sqlcmd to a lane is refused
		$tcpConnStr = "Server=$tcpServer;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"
		$namedPipesConnStr = "Server=$namedPipes;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"
		$simpleConnStr = "Server=$dbServer;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"

		$laneInfo = @{
			'MachineName'	    = $machineName
			'DBName'		    = $dbName
			'DBServer'		    = $dbServer
			'ServerName'	    = $serverName
			'InstanceName'	    = $instanceName
			'NamedPipes'	    = $namedPipes
			'TcpServer'		    = $tcpServer
			'ConnectionString'  = $simpleConnStr
			'NamedPipesConnStr' = $namedPipesConnStr
			'TcpConnStr'	    = $tcpConnStr
			'Source'		    = $source
		}

		$LaneDatabaseInfo[$laneNumber] = $laneInfo
		$changed = $true
	}

	# ------------------------------------------------------------------------------------------------
	# Write cache ONCE (only if we found anything new)
	#   IMPORTANT: Write LaneDatabaseInfo ordered by lane number for clean file output.
	# ------------------------------------------------------------------------------------------------
	if ($CachePath -and $changed)
	{
		try
		{
			# Build an ORDERED view for file output only (keep in-memory as normal hashtable for .ContainsKey())
			$orderedLaneDb = [ordered]@{ }
			$sortedLaneKeys = @($LaneDatabaseInfo.Keys) | Sort-Object { [int]($_) }
			foreach ($k in $sortedLaneKeys)
			{
				$orderedLaneDb[$k] = $LaneDatabaseInfo[$k]
			}

			$cacheOut = [pscustomobject]@{
				LastUpdated = (Get-Date).ToString('o')
				LaneDatabaseInfo = $orderedLaneDb
			}

			($cacheOut | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
		}
		catch { }
	}

	# Return behavior (same intent as original):
	if ($RequestedLane)
	{
		if ($LaneDatabaseInfo.ContainsKey($RequestedLane)) { return $LaneDatabaseInfo[$RequestedLane] }
		return $null
	}
}
