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
param (
	[string]$StatePath,
	[string]$LockPath,
	[string]$SummaryLogPath,
	[string]$RunLogPath,
	[string]$RunLogPattern,
	[int]$RetentionCount = 8,
	[int]$SummaryRetentionLines = 250,
	[string]$RunId
)

$ErrorActionPreference = 'Continue'
$runStartUtc = [DateTime]::UtcNow
$runStartLocal = Get-Date
$runLogFolder = Split-Path -Path $RunLogPath -Parent
$summaryFolder = Split-Path -Path $SummaryLogPath -Parent

if (-not (Test-Path -LiteralPath $runLogFolder)) { New-Item -Path $runLogFolder -ItemType Directory -Force | Out-Null }
if (-not (Test-Path -LiteralPath $summaryFolder)) { New-Item -Path $summaryFolder -ItemType Directory -Force | Out-Null }

$writeState = {
	param (
		[string]$Status,
		[Nullable[int]]$WorkerPid,
		[string]$LastStartedTimeUtc,
		[string]$LastCompletedTimeUtc,
		[string]$LastOutcome,
		[string]$LastSummary,
		[int]$RestoreHealthExitCode,
		[int]$StartComponentCleanupExitCode,
		[int]$SfcExitCode
	)
	
	$stateObj = [pscustomobject]@{
		Status						 = $Status
		Pid							 = $WorkerPid
		LastRunId					 = $RunId
		LastStartedTimeUtc			 = $LastStartedTimeUtc
		LastCompletedTimeUtc		 = $LastCompletedTimeUtc
		LastOutcome					 = $LastOutcome
		LastSummary					 = $LastSummary
		RestoreHealthExitCode		 = $RestoreHealthExitCode
		StartComponentCleanupExitCode = $StartComponentCleanupExitCode
		SfcExitCode					 = $SfcExitCode
		RunLogPath					 = $RunLogPath
	}
	
	($stateObj | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $StatePath -Encoding UTF8 -Force
}

$appendLog = {
	param ([string]$Text)
	Add-Content -LiteralPath $RunLogPath -Value $Text -Encoding UTF8
}

$restoreHealthExit = -1
$startCleanupExit = -1
$sfcExit = -1
$overall = 'Unknown'
$summaryLine = $null

try
{
	$lockObj = [pscustomobject]@{
		Pid		   = $PID
		RunId	   = $RunId
		StartedUtc = $runStartUtc.ToString('o')
		RunLogPath = $RunLogPath
	}
	($lockObj | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $LockPath -Encoding UTF8 -Force
	
	& $writeState 'Running' $PID $runStartUtc.ToString('o') $null $null $null $restoreHealthExit $startCleanupExit $sfcExit
	
	& $appendLog ('=' * 100)
	& $appendLog ('Quarterly system health maintenance started: {0}' -f $runStartLocal.ToString('yyyy-MM-dd HH:mm:ss'))
	& $appendLog ('RunId: {0}' -f $RunId)
	& $appendLog ('PID: {0}' -f $PID)
	& $appendLog ('=' * 100)
	& $appendLog ''
	
	& $appendLog '>>> DISM /Online /Cleanup-Image /RestoreHealth /Source:WindowsUpdate'
	& dism.exe /Online /Cleanup-Image /RestoreHealth /Source:WindowsUpdate *>&1 | Tee-Object -FilePath $RunLogPath -Append | Out-Null
	$restoreHealthExit = $LASTEXITCODE
	& $appendLog ('DISM RestoreHealth exit code: {0}' -f $restoreHealthExit)
	& $appendLog ''
	
	& $appendLog '>>> DISM /Online /Cleanup-Image /StartComponentCleanup'
	& dism.exe /Online /Cleanup-Image /StartComponentCleanup *>&1 | Tee-Object -FilePath $RunLogPath -Append | Out-Null
	$startCleanupExit = $LASTEXITCODE
	& $appendLog ('DISM StartComponentCleanup exit code: {0}' -f $startCleanupExit)
	& $appendLog ''
	
	& $appendLog '>>> SFC /SCANNOW'
	& sfc.exe /SCANNOW *>&1 | Tee-Object -FilePath $RunLogPath -Append | Out-Null
	$sfcExit = $LASTEXITCODE
	& $appendLog ('SFC exit code: {0}' -f $sfcExit)
	& $appendLog ''
	
	if ($restoreHealthExit -eq 0 -and $startCleanupExit -eq 0 -and $sfcExit -eq 0)
	{
		$overall = 'Success'
	}
	elseif ($restoreHealthExit -lt 0 -or $startCleanupExit -lt 0 -or $sfcExit -lt 0)
	{
		$overall = 'Failed'
	}
	else
	{
		$overall = 'CompletedWithIssues'
	}
}
catch
{
	$overall = 'Failed'
	& $appendLog ('Worker exception: {0}' -f $_.Exception.Message)
}
finally
{
	$runEndUtc = [DateTime]::UtcNow
	$duration = $runEndUtc - $runStartUtc
	$summaryLine = '[{0}] Overall={1}; RestoreHealth={2}; StartComponentCleanup={3}; SFC={4}; Duration={5:dd\.hh\:mm\:ss}; Log={6}' -f `
		$runEndUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss'), `
		$overall, `
		$restoreHealthExit, `
		$startCleanupExit, `
		$sfcExit, `
		$duration, `
		$RunLogPath
	
	try
	{
		Add-Content -LiteralPath $SummaryLogPath -Value $summaryLine -Encoding UTF8
	}
	catch { }
	
	try
	{
		if (Test-Path -LiteralPath $SummaryLogPath)
		{
			$summaryLines = @(Get-Content -LiteralPath $SummaryLogPath -ErrorAction SilentlyContinue)
			if ($summaryLines.Count -gt $SummaryRetentionLines)
			{
				$summaryLines | Select-Object -Last $SummaryRetentionLines | Set-Content -LiteralPath $SummaryLogPath -Encoding UTF8 -Force
			}
		}
	}
	catch { }
	
	try
	{
		$oldLogs = @(Get-ChildItem -LiteralPath $runLogFolder -Filter $RunLogPattern -File -ErrorAction SilentlyContinue |
			Sort-Object LastWriteTime -Descending |
			Select-Object -Skip $RetentionCount)
		foreach ($oldLog in $oldLogs)
		{
			try { Remove-Item -LiteralPath $oldLog.FullName -Force -ErrorAction SilentlyContinue }
			catch { }
		}
	}
	catch { }
	
	try
	{
		& $writeState 'Idle' $null $runStartUtc.ToString('o') $runEndUtc.ToString('o') $overall $summaryLine $restoreHealthExit $startCleanupExit $sfcExit
	}
	catch { }
	
	try { Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue }
	catch { }
}
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
	$psExe = $null
	try { $psExe = (Get-Process -Id $PID -ErrorAction Stop).Path }
	catch { $psExe = $null }
	if ([string]::IsNullOrWhiteSpace($psExe)) { $psExe = 'powershell.exe' }
	
	$argList = @(
		'-NoProfile',
		'-ExecutionPolicy', 'Bypass',
		'-File', $workerPath,
		'-StatePath', $statePath,
		'-LockPath', $lockPath,
		'-SummaryLogPath', $summaryLogPath,
		'-RunLogPath', $runLogPath,
		'-RunLogPattern', $runLogPattern,
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
