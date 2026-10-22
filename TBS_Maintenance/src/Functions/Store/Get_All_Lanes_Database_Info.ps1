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
								$LaneDatabaseInfo[$p.Name] = $p.Value
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

	foreach ($laneNumber in $lanesToProcess)
	{
		# If cached already in memory, skip
		if ($LaneDatabaseInfo.ContainsKey($laneNumber)) { continue }

		# Skip unwanted lanes for full mode (same logic as before)
		# IMPORTANT: If the user explicitly requested a lane, do NOT skip it.
		if (-not $RequestedLane -and ($laneNumber -match '^(8|9)' -or $laneNumber -eq '901' -or $laneNumber -eq '999')) { continue }
		if ($RequestedLane -and $laneNumber -ne $RequestedLane -and ($laneNumber -match '^(8|9)' -or $laneNumber -eq '901' -or $laneNumber -eq '999')) { continue }

		$machineName = $LaneNumToMachineName[$laneNumber]
		if (-not $machineName) { continue }

		# ---------- 1. Try LOCAL Startup file (INI\STARTUP.###) ----------
		$startupLocalPath = Join-Path $OfficePath "INI\STARTUP.$($laneNumber.PadLeft(3, '0'))"
		$startupIniPath = $null
		$content = $null
		$source = $null

		if (Test-Path $startupLocalPath)
		{
			try
			{
				$content = Get-Content -Path $startupLocalPath -ErrorAction Stop
				$source = "local"
			}
			catch { $content = $null }
		}

		# ---------- 2. Fallback: Try REMOTE Startup.ini ----------
		if (-not $content)
		{
			$startupIniPath = "\\$machineName\storeman\Startup.ini"
			$remoteShareReachable = $false
			$tcpClient = New-Object System.Net.Sockets.TcpClient
			try
			{
				$asyncResult = $tcpClient.BeginConnect($machineName, 445, $null, $null)
				if ($asyncResult.AsyncWaitHandle.WaitOne(1500))
				{
					$tcpClient.EndConnect($asyncResult)
					$remoteShareReachable = $true
				}
			}
			catch { $remoteShareReachable = $false }
			finally { $tcpClient.Dispose() }

			if ($remoteShareReachable -and (Test-Path $startupIniPath))
			{
				try
				{
					$content = Get-Content -Path $startupIniPath -ErrorAction Stop
					$source = "remote"
				}
				catch { $content = $null }
			}
		}

		if (-not $content) { continue }

		$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' }
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }

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

		$tcpConnStr = "Server=$tcpServer;Database=$dbName;Integrated Security=True;"
		$namedPipesConnStr = "Server=$namedPipes;Database=$dbName;Integrated Security=True;"
		$simpleConnStr = "Server=$dbServer;Database=$dbName;Integrated Security=True;"

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
