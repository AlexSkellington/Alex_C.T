# ===================================================================================================
#                              FUNCTION: Start_Quarterly_System_Health_Maintenance
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs DISM/SFC system health maintenance at most once per interval. Launches a detached
#   hidden PowerShell worker so the scan continues even if the main app closes, and persists
#   run state plus summary/run logs under the configured folders.
#
# Parameters:
#   - IntervalMonths        [int]: minimum months between runs (default 3)
#   - RetentionCount        [int]: number of run logs to keep (default 8)
#   - SummaryRetentionLines [int]: summary log lines to keep (default 250)
#
# Variables:
#   - $script:SystemHealthMaintenance*: state/lock/worker/summary log paths (set at script top)
#
# Returns:
#   @{ Started; Running; Due; Pid; RunLogPath } when a worker started, else $false.
#
# Writes:
#   - state JSON, summary log and run logs under Setup_Files / the script log folder
# ===================================================================================================

function Start_Quarterly_System_Health_Maintenance
{
	[CmdletBinding()]
	param (
		[int]$IntervalMonths = 3,
		[int]$RetentionCount = 8,
		[int]$SummaryRetentionLines = 250
	)

	$statePath = $script:SystemHealthMaintenanceStatePath
	$lockPath = $script:SystemHealthMaintenanceLockPath
	$workerPath = $script:SystemHealthMaintenanceWorkerPath
	$summaryLogPath = $script:SystemHealthMaintenanceSummaryLogPath
	$runLogPattern = $script:SystemHealthMaintenanceRunLogPattern
	$runLogPrefix = $script:SystemHealthMaintenanceRunLogPrefix
	$logFolder = $script:ScriptsLog
	$setupFolder = $script:SetupFiles
	$nowUtc = [DateTime]::UtcNow
	$state = $null

	# DISM and SFC refuse to run without elevation (exit 740 / 1). Without this check a
	# non-elevated launch (e.g. irm | iex from a normal console) recorded a 'completed' run that
	# did nothing and postponed the real one by a full interval.
	$isAdmin = $false
	try { $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }
	catch { }
	if (-not $isAdmin)
	{
		try { Write_Log "Quarterly system health maintenance skipped - not running as administrator." "gray" -FileOnly }
		catch { }
		return @{ Started = $false; Running = $false; Due = $null; Skipped = 'NotElevated' }
	}

	try
	{
		if (-not (Test-Path -LiteralPath $logFolder)) { New-Item -Path $logFolder -ItemType Directory -Force | Out-Null }
		if (-not (Test-Path -LiteralPath $setupFolder)) { New-Item -Path $setupFolder -ItemType Directory -Force | Out-Null }
	}
	catch
	{
		try { Write_Log "Quarterly system health maintenance: failed to ensure log/state folders. $_" "red" }
		catch { }
		return $false
	}

	try
	{
		if (Test-Path -LiteralPath $statePath)
		{
			$state = (Get-Content -LiteralPath $statePath -Raw -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
		}
	}
	catch
	{
		$state = $null
	}

	$workerScriptText = @'
#TBS:EMBED Resources/Workers/Quarterly_System_Health_Worker.ps1
'@

	try
	{
		Set-Content -LiteralPath $workerPath -Value $workerScriptText -Encoding UTF8 -Force
	}
	catch
	{
		try { Write_Log "Quarterly system health maintenance: failed to write worker script. $_" "red" }
		catch { }
		return $false
	}

	$existingWorkerPid = $null
	try
	{
		$escapedWorkerPath = [regex]::Escape($workerPath)
		$existingWorker = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
			Where-Object { $_.CommandLine -and $_.CommandLine -match $escapedWorkerPath } |
			Select-Object -First 1)
		if ($existingWorker.Count -gt 0)
		{
			$existingWorkerPid = [int]$existingWorker[0].ProcessId
		}
	}
	catch { }

	if (-not $existingWorkerPid -and $state -and $state.Pid)
	{
		try
		{
			$null = Get-Process -Id ([int]$state.Pid) -ErrorAction Stop
			$existingWorkerPid = [int]$state.Pid
		}
		catch { }
	}

	if ($existingWorkerPid)
	{
		try
		{
			$runningState = [pscustomobject]@{
				Status    = 'Running'
				Pid	      = $existingWorkerPid
				LastRunId = if ($state -and $state.LastRunId) { [string]$state.LastRunId } else { $null }
				LastStartedTimeUtc = if ($state -and $state.LastStartedTimeUtc) { [string]$state.LastStartedTimeUtc } else { $nowUtc.ToString('o') }
				LastCompletedTimeUtc = if ($state -and $state.LastCompletedTimeUtc) { [string]$state.LastCompletedTimeUtc } else { $null }
				LastOutcome = if ($state -and $state.LastOutcome) { [string]$state.LastOutcome } else { $null }
				LastSummary = if ($state -and $state.LastSummary) { [string]$state.LastSummary } else { $null }
				RestoreHealthExitCode = if ($state -and $null -ne $state.RestoreHealthExitCode) { [int]$state.RestoreHealthExitCode } else { $null }
				StartComponentCleanupExitCode = if ($state -and $null -ne $state.StartComponentCleanupExitCode) { [int]$state.StartComponentCleanupExitCode } else { $null }
				SfcExitCode = if ($state -and $null -ne $state.SfcExitCode) { [int]$state.SfcExitCode } else { $null }
				RunLogPath = if ($state -and $state.RunLogPath) { [string]$state.RunLogPath } else { $null }
			}
			($runningState | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $statePath -Encoding UTF8 -Force
		}
		catch { }

		return @{ Started = $false; Running = $true; Due = $false }
	}

	$lastCompletedUtc = $null
	if ($state -and $state.LastCompletedTimeUtc)
	{
		try { $lastCompletedUtc = [DateTime]::Parse([string]$state.LastCompletedTimeUtc).ToUniversalTime() }
		catch { $lastCompletedUtc = $null }
	}

	if ($state -and $state.Status -eq 'Running')
	{
		try { Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue }
		catch { }
	}

	if ($lastCompletedUtc)
	{
		$nextDueUtc = $lastCompletedUtc.AddMonths($IntervalMonths)
		if ($nextDueUtc -gt $nowUtc)
		{
			return @{ Started = $false; Running = $false; Due = $false; NextDueUtc = $nextDueUtc }
		}
	}

	$runId = [guid]::NewGuid().ToString()
	$runLogName = '{0}{1}.log' -f $runLogPrefix, (Get-Date -Format 'yyyyMMdd_HHmmss')
	$runLogPath = Join-Path $logFolder $runLogName
	# Always Windows PowerShell itself: for a packaged .exe launch the current process path is the
	# TBS exe, which would just start a second (hidden) copy of the app instead of the worker.
	$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
	if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
	$q = { param ([string]$v) '"' + $v + '"' }   # Start-Process does not quote arguments

	$argList = @(
		'-NoProfile',
		'-ExecutionPolicy', 'Bypass',
		'-File', (& $q $workerPath),
		'-StatePath', (& $q $statePath),
		'-LockPath', (& $q $lockPath),
		'-SummaryLogPath', (& $q $summaryLogPath),
		'-RunLogPath', (& $q $runLogPath),
		'-RunLogPattern', (& $q $runLogPattern),
		'-RetentionCount', [string]$RetentionCount,
		'-SummaryRetentionLines', [string]$SummaryRetentionLines,
		'-RunId', $runId
	)

	try
	{
		$proc = Start-Process -FilePath $psExe -ArgumentList $argList -WindowStyle Hidden -PassThru -ErrorAction Stop

		$runningState = [pscustomobject]@{
			Status			     = 'Running'
			Pid				     = $proc.Id
			LastRunId		     = $runId
			LastStartedTimeUtc   = $nowUtc.ToString('o')
			LastCompletedTimeUtc = if ($lastCompletedUtc) { $lastCompletedUtc.ToString('o') } else { $null }
			LastOutcome		     = if ($state -and $state.LastOutcome) { [string]$state.LastOutcome } else { $null }
			LastSummary		     = if ($state -and $state.LastSummary) { [string]$state.LastSummary } else { $null }
			RestoreHealthExitCode = if ($state -and $null -ne $state.RestoreHealthExitCode) { [int]$state.RestoreHealthExitCode } else { $null }
			StartComponentCleanupExitCode = if ($state -and $null -ne $state.StartComponentCleanupExitCode) { [int]$state.StartComponentCleanupExitCode } else { $null }
			SfcExitCode		     = if ($state -and $null -ne $state.SfcExitCode) { [int]$state.SfcExitCode } else { $null }
			RunLogPath		     = $runLogPath
		}
		($runningState | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $statePath -Encoding UTF8 -Force
		return @{ Started = $true; Running = $true; Due = $true; Pid = $proc.Id; RunLogPath = $runLogPath }
	}
	catch
	{
		try { Write_Log "Quarterly system health maintenance: failed to start worker process. $_" "red" }
		catch { }
		return $false
	}
}
