# ===================================================================================================
#                              FUNCTION: Start_Lane_Protocol_Jobs
# ---------------------------------------------------------------------------------------------------
# Description:
#   Starts the lane SQL protocol background loop: a timer rescans every lane's SQL protocol
#   (TCP -> Named Pipes -> else File) every minute. Call once from the main script; it
#   returns immediately, and a second call updates the lane map and keeps the existing loop.
#
# Parameters:
#   - LaneNumToMachineName  [hashtable] (required): LaneNumber => MachineName
#   - SqlModuleName         [string] (required): detected SQL module (Invoke-Sqlcmd)
#
# Variables:
#   - $script:LaneProtocols: synchronized hashtable LaneNum => protocol (read by all consumers)
#   - $script:SetupFiles: directory for Protocol_Results.txt (UTF8 no BOM)
#   - $script:LaneProtocolJobs / $script:LaneSqlProtocolEnableJobs: per-lane job and task state
#
# Notes:
#   - Double-start guarded via $script:protocolTimer.
# ===================================================================================================

function Start_Lane_Protocol_Jobs
{
	param (
		[Parameter(Mandatory)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory)]
		[string]$SqlModuleName
	)

	# ---------- assemblies ----------
	try { Add-Type -AssemblyName System.Windows.Forms 2>$null }
	catch { }
	try { Add-Type -AssemblyName System.Data 2>$null }
	catch { }

	# ---------- paths ----------
	$script:ProtocolResultsFile = Join-Path $script:SetupFiles 'Protocol_Results.txt'
	$resultsDir = [System.IO.Path]::GetDirectoryName($script:ProtocolResultsFile)
	if (-not (Test-Path $resultsDir)) { New-Item -Path $resultsDir -ItemType Directory -Force | Out-Null }
	if (-not (Test-Path $script:ProtocolResultsFile)) { New-Item -Path $script:ProtocolResultsFile -ItemType File -Force | Out-Null }

	# ---------- globals ----------
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
	if (-not $script:ProtocolResults) { $script:ProtocolResults = @() }
	if (-not $script:LaneProtocolJobs) { $script:LaneProtocolJobs = @{ } }
	if (-not $script:LaneProtocolJobTimeoutSeconds -or $script:LaneProtocolJobTimeoutSeconds -lt 20) { $script:LaneProtocolJobTimeoutSeconds = 20 }
	if (-not $script:LaneSqlProtocolEnableTaskState) { $script:LaneSqlProtocolEnableTaskState = @{ } }
	if (-not $script:LaneSqlProtocolEnableJobs) { $script:LaneSqlProtocolEnableJobs = @{ } }
	# store lane map for repeated runs
	$script:LaneProtocolLaneMap = $LaneNumToMachineName

	# stop flag
	$script:LaneProtocolStop = $false

	# ---------- prevent double start ----------
	if ($script:protocolTimer)
	{
		return
	}

	# ---------- warm from file (best effort) ----------
	$loadedProtocolCount = 0
	try
	{
		$existing = Get-Content -LiteralPath $script:ProtocolResultsFile -ErrorAction SilentlyContinue
		if ($existing)
		{
			foreach ($line in $existing)
			{
				if ($line -match '^\s*([^,]+),\s*([^,]+)\s*$')
				{
					$laneRaw = $matches[1].Trim()
					$proto = $matches[2].Trim()
					$laneNum = (($laneRaw -replace '[^\d]', '')).PadLeft(3, '0')
					if ([string]::IsNullOrWhiteSpace($laneNum)) { continue }

					$script:LaneProtocols[$laneNum] = $proto
					$script:LaneProtocols[$laneRaw] = $proto
					$loadedProtocolCount++

					$mn = $LaneNumToMachineName[$laneNum]
					if ($mn)
					{
						$script:LaneProtocols[$mn] = $proto
						$script:LaneProtocols[$mn.ToLower()] = $proto
					}

					$script:ProtocolResults = @($script:ProtocolResults | Where-Object { $_.Lane -ne $laneNum })
					$script:ProtocolResults += [pscustomobject]@{ Lane = $laneNum; Protocol = $proto }
				}
			}
		}
	}
	catch { }

	# ---------- runspace pool (kept alive) ----------
	if (-not $script:LaneProtocolPool)
	{
		$minThreads = 1
		$maxThreads = [Math]::Max(8, [Environment]::ProcessorCount * 2)
		$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool($minThreads, $maxThreads, $iss, $Host)
		try { $pool.ApartmentState = 'MTA' }
		catch { }
		$pool.Open()
		$script:LaneProtocolPool = $pool
	}

	# ---------- worker script (NO nested functions) ----------
	$script:LaneProtocolWorker = @'
param([string]$machine,[string]$lane)

Add-Type -AssemblyName System.Data 2>$null
Add-Type -AssemblyName System 2>$null

$protocol = 'File'
$reachable = $true

if (-not [string]::IsNullOrWhiteSpace($machine))
{
  $machineNorm = [string]$machine
  $isLocalMachine = $false
  try {
    if ($machineNorm -match '^(?i)(localhost|\.|\(local\))$' -or $machineNorm -ieq $env:COMPUTERNAME) { $isLocalMachine = $true }
  } catch { }

  if (-not $isLocalMachine)
  {
    $reachable = $false

    try {
      $ping = New-Object System.Net.NetworkInformation.Ping
      $reply = $ping.Send($machineNorm, 700)
      if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) { $reachable = $true }
    } catch { }
  }
}

if (-not $reachable)
{
  [PSCustomObject]@{ Lane = $lane; Protocol = 'Offline' }
  return
}

$cn = $null
try {
  $cs = "Data Source=tcp:$machine;Initial Catalog=master;Integrated Security=True;Connect Timeout=1"
  $cn = New-Object System.Data.SqlClient.SqlConnection $cs
  $cn.Open()
  $protocol = 'TCP'
} catch { }
finally { try { if ($cn) { $cn.Close() } } catch { }; try { if ($cn) { $cn.Dispose() } } catch { } }

if ($protocol -eq 'File') {
  $cn2 = $null
  try {
    $cs2 = "Data Source=np:$machine;Initial Catalog=master;Integrated Security=True;Connect Timeout=1"
    $cn2 = New-Object System.Data.SqlClient.SqlConnection $cs2
    $cn2.Open()
    $protocol = 'Named Pipes'
  } catch { }
  finally { try { if ($cn2) { $cn2.Close() } } catch { }; try { if ($cn2) { $cn2.Dispose() } } catch { } }
}

[PSCustomObject]@{ Lane = $lane; Protocol = $protocol }
'@

	# ---------- cadence ----------
	if (-not $script:LaneProtocolLastRun) { $script:LaneProtocolLastRun = (Get-Date).AddYears(-10) }

	$normalizeLaneProtocolTimestamp = {
		param($Value)

		if ($null -eq $Value) { return $null }

		try
		{
			if ($Value -is [datetime]) { return [datetime]$Value }
		}
		catch { }

		try
		{
			if ($Value -is [datetimeoffset]) { return ([datetimeoffset]$Value).LocalDateTime }
		}
		catch { }

		try
		{
			if ($Value -is [string])
			{
				$text = [string]$Value
				if ([string]::IsNullOrWhiteSpace($text)) { return $null }

				$parsed = [datetime]::MinValue
				if ([datetime]::TryParse($text, [ref]$parsed))
				{
					return $parsed
				}

				return $null
			}
		}
		catch { }

		try
		{
			return [datetime]$Value
		}
		catch { }

		return $null
	}.GetNewClosure()

	$getNextLaneSqlEnableRunDate = {
		$scheduledDate = Get-Date
		try
		{
			$today415 = Get-Date -Hour 4 -Minute 15 -Second 0
			$scheduledDate = if ((Get-Date) -lt $today415) { $today415 }
			else { $today415.AddDays(1) }
		}
		catch
		{
			$scheduledDate = (Get-Date).AddDays(1).Date.AddHours(4).AddMinutes(15)
		}
		return $scheduledDate
	}.GetNewClosure()

	$startLaneSqlEnableBackgroundJob = {
		param (
			[string]$LaneNumber,
			[string]$MachineName
		)

		if ([string]::IsNullOrWhiteSpace($LaneNumber) -or [string]::IsNullOrWhiteSpace($MachineName))
		{
			return
		}

		$laneKey = $LaneNumber.ToString().PadLeft(3, '0')
		$existingJobState = $script:LaneSqlProtocolEnableJobs[$laneKey]
		if ($existingJobState -and $existingJobState.Process -and -not $existingJobState.Process.HasExited)
		{
			return
		}

		$scheduledDate = & $getNextLaneSqlEnableRunDate
		$sourceScriptPath = $null
		try
		{
			if ($script:LaunchEntryPath -and (Test-Path -LiteralPath $script:LaunchEntryPath))
			{
				$sourceScriptPath = $script:LaunchEntryPath
			}
			elseif ($PSCommandPath -and (Test-Path -LiteralPath $PSCommandPath))
			{
				$sourceScriptPath = $PSCommandPath
			}
		}
		catch
		{
			$sourceScriptPath = $null
		}

		if ([string]::IsNullOrWhiteSpace($sourceScriptPath))
		{
			$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
				TaskPresent  = $false
				LastChecked  = Get-Date
				ScheduledFor = $scheduledDate
				MachineName  = $MachineName
			}
			return
		}

		$fnDefinition = $null
		try
		{
			$tokens = $null
			$parseErrors = $null
			$ast = [System.Management.Automation.Language.Parser]::ParseFile($sourceScriptPath, [ref]$tokens, [ref]$parseErrors)
			$fnAst = $ast.Find({
					param($node)
					$node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
					$node.Name -eq 'Ensure_SQL_Enable_Scheduled_Task_For_File_Lane'
				}, $true)
			if ($fnAst)
			{
				$fnDefinition = $fnAst.Extent.Text
			}
		}
		catch
		{
			$fnDefinition = $null
		}

		if ([string]::IsNullOrWhiteSpace($fnDefinition))
		{
			$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
				TaskPresent  = $false
				LastChecked  = Get-Date
				ScheduledFor = $scheduledDate
				MachineName  = $MachineName
			}
			return
		}

		$launcherScriptPath = Join-Path $script:SetupFiles ("LaneSqlEnableLauncher_{0}.ps1" -f $laneKey)
		$launcherScript = @"
param(
	[string]`$LaneNumber,
	[string]`$MachineName,
	[string]`$TcpPort = '1433',
	[string]`$StartTime = '04:15'
)

$fnDefinition

`$ok = `$false
try
{
	`$ok = [bool](Ensure_SQL_Enable_Scheduled_Task_For_File_Lane -LaneNumber `$LaneNumber -MachineName `$MachineName -TcpPort `$TcpPort -StartTime `$StartTime)
}
catch
{
	`$ok = `$false
}

if (`$ok) { exit 0 } else { exit 1 }
"@

		try
		{
			Set-Content -LiteralPath $launcherScriptPath -Value $launcherScript -Encoding UTF8 -Force
			$argList = @(
				'-NoProfile',
				'-ExecutionPolicy', 'Bypass',
				'-File', ('"{0}"' -f $launcherScriptPath),
				'-LaneNumber', $laneKey,
				'-MachineName', $MachineName,
				'-TcpPort', '1433',
				'-StartTime', '04:15'
			)
			$proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -WindowStyle Hidden -PassThru
		}
		catch
		{
			try { Remove-Item -LiteralPath $launcherScriptPath -Force -ErrorAction SilentlyContinue }
			catch { }
			$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
				TaskPresent  = $false
				LastChecked  = Get-Date
				ScheduledFor = $scheduledDate
				MachineName  = $MachineName
			}
			return
		}

		$script:LaneSqlProtocolEnableJobs[$laneKey] = [pscustomobject]@{
			Process      = $proc
			ScriptPath   = $launcherScriptPath
			Started      = Get-Date
			MachineName  = $MachineName
			ScheduledFor = $scheduledDate
		}
		$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
			TaskPresent  = $false
			LastChecked  = Get-Date
			ScheduledFor = $scheduledDate
			MachineName  = $MachineName
			Pending	     = $true
		}
	}.GetNewClosure()

	$harvestLaneSqlEnableJobs = {
		if (-not $script:LaneSqlProtocolEnableJobs -or $script:LaneSqlProtocolEnableJobs.Count -eq 0) { return }

		foreach ($laneKey in @($script:LaneSqlProtocolEnableJobs.Keys))
		{
			$st = $script:LaneSqlProtocolEnableJobs[$laneKey]
			if (-not $st) { continue }

			$proc = $st.Process
			$launcherScriptPath = $st.ScriptPath
			if (-not $proc)
			{
				try { if ($launcherScriptPath) { Remove-Item -LiteralPath $launcherScriptPath -Force -ErrorAction SilentlyContinue } }
				catch { }
				[void]$script:LaneSqlProtocolEnableJobs.Remove($laneKey)
				continue
			}
			try { if (-not $proc.HasExited) { continue } }
			catch { continue }

			$success = $false
			try { $success = ($proc.ExitCode -eq 0) }
			catch { $success = $false }
			try { if ($launcherScriptPath) { Remove-Item -LiteralPath $launcherScriptPath -Force -ErrorAction SilentlyContinue } }
			catch { }
			[void]$script:LaneSqlProtocolEnableJobs.Remove($laneKey)

			$machineName = $null
			$scheduledFor = $null
			try { $machineName = [string]$st.MachineName }
			catch { $machineName = $null }
			try { $scheduledFor = $st.ScheduledFor }
			catch { $scheduledFor = $null }

			$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
				TaskPresent  = $success
				LastChecked  = Get-Date
				ScheduledFor = $scheduledFor
				MachineName  = $machineName
			}
		}
	}.GetNewClosure()

	# ---------- timer (UI thread, safe) ----------
	$script:protocolTimer = New-Object System.Windows.Forms.Timer
	$script:protocolTimer.Interval = 1000

	$script:protocolTimer.add_Tick({

			if ($script:LaneProtocolStop) { return }

			# 0) harvest background SQL-enable scheduling jobs
			try { & $harvestLaneSqlEnableJobs }
			catch { }

			# 1) harvest completed jobs
			$changed = $false
			if ($script:LaneProtocolJobs -and $script:LaneProtocolJobs.Count -gt 0)
			{
				foreach ($laneKey in @($script:LaneProtocolJobs.Keys))
				{
					$st = $script:LaneProtocolJobs[$laneKey]
					if (-not $st) { continue }

					$startedRaw = $st.Started
					$started = & $normalizeLaneProtocolTimestamp $startedRaw
					$jobTimedOut = $false
					if ($started)
					{
						try { $jobTimedOut = (((Get-Date) - $started).TotalSeconds -ge $script:LaneProtocolJobTimeoutSeconds) }
						catch { $jobTimedOut = $true }
					}
					elseif ($null -ne $startedRaw)
					{
						$jobTimedOut = $true
					}

					if ($jobTimedOut)
					{
						$laneNum = (($laneKey -replace '[^\d]', '')).PadLeft(3, '0')
						if ([string]::IsNullOrWhiteSpace($laneNum)) { $laneNum = $laneKey }

						$laneMap = $script:LaneProtocolLaneMap
						$mn = $null
						if ($laneMap) { $mn = $laneMap[$laneNum] }

						$script:LaneProtocols[$laneNum] = 'Offline'
						$script:LaneProtocols[$laneKey] = 'Offline'
						if ($mn)
						{
							$script:LaneProtocols[$mn] = 'Offline'
							$script:LaneProtocols[$mn.ToLower()] = 'Offline'
						}

						$script:ProtocolResults = @($script:ProtocolResults | Where-Object { $_.Lane -ne $laneNum })
						$script:ProtocolResults += [pscustomobject]@{ Lane = $laneNum; Protocol = 'Offline' }

						$changed = $true
						continue
					}

					$handle = $st.Handle
					if ($handle -and $handle.IsCompleted)
					{
						$ps = $st.PS
						$resultList = $null
						try { $resultList = $ps.EndInvoke($handle) }
						catch { $resultList = $null }
						try { if ($ps) { $ps.Dispose() } }
						catch { }

						[void]$script:LaneProtocolJobs.Remove($laneKey)

						$result = $null
						if ($resultList -and $resultList.Count -ge 1) { $result = $resultList[0] }
						if (-not $result) { $result = [pscustomobject]@{ Lane = $laneKey; Protocol = 'Offline' } }

						$rawLane = [string]$result.Lane
						$laneNum = (($rawLane -replace '[^\d]', '')).PadLeft(3, '0')
						if ([string]::IsNullOrWhiteSpace($laneNum)) { $laneNum = $laneKey }

						$proto = [string]$result.Protocol
						if ([string]::IsNullOrWhiteSpace($proto)) { $proto = 'File' }

						$laneMap = $script:LaneProtocolLaneMap
						$mn = $null
						if ($laneMap) { $mn = $laneMap[$laneNum] }

						$script:LaneProtocols[$laneNum] = $proto
						$script:LaneProtocols[$rawLane] = $proto
						if ($mn)
						{
							$script:LaneProtocols[$mn] = $proto
							$script:LaneProtocols[$mn.ToLower()] = $proto
						}

						$script:ProtocolResults = @($script:ProtocolResults | Where-Object { $_.Lane -ne $laneNum })
						$script:ProtocolResults += [pscustomobject]@{ Lane = $laneNum; Protocol = $proto }
						if ($proto -eq 'File' -and -not [string]::IsNullOrWhiteSpace($mn))
						{
							try { & $startLaneSqlEnableBackgroundJob $laneNum $mn }
							catch { }
						}
						elseif ($script:LaneSqlProtocolEnableTaskState -and $script:LaneSqlProtocolEnableTaskState.ContainsKey($laneNum))
						{
							$script:LaneSqlProtocolEnableTaskState[$laneNum] = [pscustomobject]@{
								TaskPresent  = $false
								LastChecked  = Get-Date
								ScheduledFor = $null
								MachineName  = $mn
							}
						}

						$changed = $true
					}
				}
			}

			# 1b) if any protocol jobs went stale, rebuild the pool so stuck probes do not wedge the app
			$staleLaneKeys = @()
			try
			{
				foreach ($laneKey in @($script:LaneProtocolJobs.Keys))
				{
					$st = $script:LaneProtocolJobs[$laneKey]
					if (-not $st) { continue }
					if ($st.Handle -and $st.Handle.IsCompleted) { continue }

					$startedRaw = $st.Started
					$started = & $normalizeLaneProtocolTimestamp $startedRaw
					$jobTimedOut = $false
					if ($started)
					{
						try { $jobTimedOut = (((Get-Date) - $started).TotalSeconds -ge $script:LaneProtocolJobTimeoutSeconds) }
						catch { $jobTimedOut = $true }
					}
					elseif ($null -ne $startedRaw)
					{
						$jobTimedOut = $true
					}

					if ($jobTimedOut)
					{
						$staleLaneKeys += $laneKey
					}
				}
			}
			catch { }

			if ($staleLaneKeys.Count -gt 0)
			{
				foreach ($st in @($script:LaneProtocolJobs.Values))
				{
					try { if ($st -and $st.PS) { $st.PS.Dispose() } }
					catch { }
				}

				$script:LaneProtocolJobs = @{ }

				try
				{
					if ($script:LaneProtocolPool)
					{
						$script:LaneProtocolPool.Close()
						$script:LaneProtocolPool.Dispose()
					}
				}
				catch { }

				$script:LaneProtocolPool = $null

				$minThreads = 1
				$maxThreads = [Math]::Max(8, [Environment]::ProcessorCount * 2)
				$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
				$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool($minThreads, $maxThreads, $iss, $Host)
				try { $pool.ApartmentState = 'MTA' }
				catch { }
				$pool.Open()
				$script:LaneProtocolPool = $pool
				$script:LaneProtocolLastRun = Get-Date
			}

			# 2) write file when something changed
			if ($changed)
			{
				try
				{
					$sorted = $script:ProtocolResults | Sort-Object { ($_.Lane -replace '[^\d]', '') -as [int] }
					$lines = foreach ($r in $sorted) { '{0},{1}' -f $r.Lane, $r.Protocol }
					[System.IO.File]::WriteAllLines($script:ProtocolResultsFile, $lines, [System.Text.Encoding]::UTF8)
				}
				catch { }
			}

			# 3) start a new scan every 60s ONLY if no jobs running
			if ($script:LaneProtocolJobs.Count -eq 0)
			{
				$lastRun = & $normalizeLaneProtocolTimestamp $script:LaneProtocolLastRun
				if (-not $lastRun)
				{
					$lastRun = (Get-Date).AddYears(-10)
					$script:LaneProtocolLastRun = $lastRun
				}
				elseif (-not ($script:LaneProtocolLastRun -is [datetime]))
				{
					$script:LaneProtocolLastRun = $lastRun
				}

				$ageSeconds = [double]::PositiveInfinity
				try { $ageSeconds = ((Get-Date) - $lastRun).TotalSeconds }
				catch
				{
					$script:LaneProtocolLastRun = (Get-Date).AddYears(-10)
					$ageSeconds = [double]::PositiveInfinity
				}

				if ($ageSeconds -ge 60)
				{
					$script:LaneProtocolLastRun = Get-Date

					$laneMap = $script:LaneProtocolLaneMap
					if ($laneMap)
					{
						foreach ($k in $laneMap.Keys)
						{
							$numStr = ($k -replace '[^\d]', '')
							if (-not $numStr) { continue }

							$laneNum = $numStr.PadLeft(3, '0')
							$machine = $laneMap[$k]
							if ([string]::IsNullOrWhiteSpace($machine)) { continue }

							$ps = [System.Management.Automation.PowerShell]::Create()
							$ps.RunspacePool = $script:LaneProtocolPool
							$null = $ps.AddScript($script:LaneProtocolWorker).AddArgument([string]$machine).AddArgument([string]$laneNum)
							$handle = $ps.BeginInvoke()

							$script:LaneProtocolJobs[$laneNum] = @{
								PS	    = $ps
								Handle  = $handle
								Started = Get-Date
							}
						}
					}
				}
			}
		}.GetNewClosure())

	$script:protocolTimer.Start()
}
