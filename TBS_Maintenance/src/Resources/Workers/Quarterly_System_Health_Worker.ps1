# ---------------------------------------------------------------------------------------------------
# Quarterly system health worker - embedded at build time into Start_Quarterly_System_Health_Maintenance
#   and written to Scripts\Quarterly_System_Health_Maintenance_Worker.ps1, then run detached as
#   powershell.exe -File (elevated parent required: DISM / SFC need admin).
#   DISM RestoreHealth -> DISM StartComponentCleanup -> SFC /SCANNOW, with a lock file, a state JSON,
#   a per-run log (rotated to RetentionCount) and one summary line per run.
# ---------------------------------------------------------------------------------------------------
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
	# SFC writes UTF-16 to a redirected stdout: drop the NULs so the run log is readable
	& sfc.exe /SCANNOW *>&1 | ForEach-Object { ([string]$_) -replace "`0", '' } | Where-Object { $_.Trim() } | Add-Content -LiteralPath $RunLogPath -Encoding UTF8
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
