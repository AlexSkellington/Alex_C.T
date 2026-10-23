# ===================================================================================================
#                              FUNCTION: Download_NSSM
# ---------------------------------------------------------------------------------------------------
# Description:
#   Ensures nssm.exe exists before any function installs or manages services. Can start a
#   hidden worker and return immediately (Background), wait for completion (Wait), or
#   run synchronously. Always publishes the NSSM paths and writes a status log per run.
#   Returns a hashtable @{ Success; Path; Note; Error; Log } describing the outcome.
#
# Parameters:
#   - ToolsRoot            [string]: tools root, defaults to $script:ToolsDir
#   - ForceDownload        [switch]: redownload even if nssm.exe already exists
#   - Background           [switch]: start a hidden worker and return immediately
#   - Wait                 [switch]: block until nssm.exe appears (or timeout)
#   - WaitTimeoutSeconds   [int]: max wait when -Wait is used (default 120)
#   - WaitPollMilliseconds [int]: poll interval when -Wait is used (default 250)
#   - ThrowOnFail          [switch]: throw instead of returning a failure result
#
# Variables:
#   - $script:NssmDir / $script:NssmExe: published NSSM locations
#   - $script:NssmDownloadLog: Download_NSSM_LOG.txt under the script log folder
#
# Writes:
#   - <ToolsRoot>\NSSM\nssm.exe and the status log file
# ===================================================================================================

function Download_NSSM
{
	[CmdletBinding()]
	param (
		# Where tools live. Defaults to the centralized script path configuration.
		[string]$ToolsRoot = $script:ToolsDir,
		# Force redownload even if nssm.exe already exists
		[switch]$ForceDownload,
		# Start download in background and RETURN immediately (non-blocking)
		[switch]$Background,
		# Wait until NSSM exists (blocks only when requested)
		[switch]$Wait,
		# Max wait time when -Wait is used
		[int]$WaitTimeoutSeconds = 120,
		# Poll rate when -Wait is used
		[int]$WaitPollMilliseconds = 250,
		# If you want failures to throw
		[switch]$ThrowOnFail
	)

	# Single log path (overwrite each run)
	$script:NssmDownloadLog = Join-Path $script:ScriptsLog 'Download_NSSM_LOG.txt'

	# ------------------------------------------------------------------------------------------------
	# Target paths
	# ------------------------------------------------------------------------------------------------
	$script:NssmDir = Join-Path $ToolsRoot 'NSSM'
	$script:NssmExe = Join-Path $script:NssmDir 'nssm.exe'

	# Publish paths for other functions
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	$script:FunctionResults['NssmDir'] = $script:NssmDir
	$script:FunctionResults['NssmExe'] = $script:NssmExe

	# Ensure tools folders exist
	try
	{
		if (-not (Test-Path $ToolsRoot)) { New-Item -Path $ToolsRoot -ItemType Directory -Force | Out-Null }
		if (-not (Test-Path $script:NssmDir)) { New-Item -Path $script:NssmDir -ItemType Directory -Force | Out-Null }
	}
	catch
	{
		if ($ThrowOnFail) { throw "Download_NSSM: Failed creating tools folder(s). $($_.Exception.Message)" }
		return @{ Success = $false; Path = $script:NssmExe; Note = 'Tools folder create failed'; Error = $_.Exception.Message; Log = $script:NssmDownloadLog }
	}

	# Fast return if already present and not forcing
	if (-not $ForceDownload -and (Test-Path $script:NssmExe))
	{
		return @{ Success = $true; Path = $script:NssmExe; Note = 'Already present'; Log = $script:NssmDownloadLog }
	}

	# ------------------------------------------------------------------------------------------------
	# Download worker (source: src\Resources\Workers\Download_NSSM_Worker.ps1) - runs inline or in a
	# background runspace; writes only to the status log (overwritten each run)
	# ------------------------------------------------------------------------------------------------
	$downloadScriptText = @'
#TBS:EMBED Resources/Workers/Download_NSSM_Worker.ps1
'@

	# ------------------------------------------------------------------------------------------------
	# Background job (in-process runspace) - avoids hidden spawned PowerShell that AV hates
	# ------------------------------------------------------------------------------------------------
	if ($Background)
	{
		# If a job is already running, don't start another
		if ($script:NssmDownloadJob -and $script:NssmDownloadJob.Handle -and (-not $script:NssmDownloadJob.Handle.IsCompleted))
		{
			return @{ Success = $true; Path = $script:NssmExe; Note = 'Already downloading in background'; Log = $script:NssmDownloadLog }
		}
		# A finished earlier download nobody waited for: free its runspace before starting again
		if ($script:NssmDownloadJob)
		{
			try { $script:NssmDownloadJob.PS.Dispose(); $script:NssmDownloadJob.Pool.Close(); $script:NssmDownloadJob.Pool.Dispose() }
			catch { }
			$script:NssmDownloadJob = $null
		}

		try
		{
			# Create a 1-thread runspace pool
			$pool = [RunspaceFactory]::CreateRunspacePool(1, 1)
			$pool.ApartmentState = 'MTA'
			$pool.ThreadOptions = 'ReuseThread'
			$pool.Open()

			$ps = [PowerShell]::Create()
			$ps.RunspacePool = $pool

			$null = $ps.AddScript($downloadScriptText)
			$null = $ps.AddArgument($ToolsRoot)
			$null = $ps.AddArgument([bool]$ForceDownload.IsPresent)
			$null = $ps.AddArgument($script:NssmDir)
			$null = $ps.AddArgument($script:NssmExe)
			$null = $ps.AddArgument($script:NssmDownloadLog)

			$handle = $ps.BeginInvoke()

			# Keep a handle so we can wait later if needed
			$script:NssmDownloadJob = @{
				PS	    = $ps
				Pool    = $pool
				Handle  = $handle
				Started = Get-Date
				Log	    = $script:NssmDownloadLog
				Exe	    = $script:NssmExe
			}

			# Return immediately unless -Wait is also requested
			if (-not $Wait)
			{
				return @{ Success = $true; Path = $script:NssmExe; Note = 'Started background download'; Log = $script:NssmDownloadLog }
			}
		}
		catch
		{
			if ($ThrowOnFail) { throw "Download_NSSM: Failed to start background runspace. $($_.Exception.Message)" }
			return @{ Success = $false; Path = $script:NssmExe; Note = 'Background start failed'; Error = $_.Exception.Message; Log = $script:NssmDownloadLog }
		}
	}

	# ------------------------------------------------------------------------------------------------
	# Wait mode (only blocks when asked). Waiting only makes sense while a background download
	# runs - without one, fall through to the synchronous download below instead of waiting
	# for a file nobody is fetching.
	# ------------------------------------------------------------------------------------------------
	$bgRunning = ($script:NssmDownloadJob -and $script:NssmDownloadJob.Handle -and -not $script:NssmDownloadJob.Handle.IsCompleted)
	if ($Wait -and $bgRunning)
	{
		# Wait for the download itself to finish (not just for the file to appear - it may still
		# be copying), keeping the window responsive.
		$job = $script:NssmDownloadJob
		$deadline = (Get-Date).AddSeconds($WaitTimeoutSeconds)
		while ((Get-Date) -lt $deadline -and -not $job.Handle.IsCompleted)
		{
			try { [System.Windows.Forms.Application]::DoEvents() }
			catch { }
			Start-Sleep -Milliseconds $WaitPollMilliseconds
		}

		if (-not $job.Handle.IsCompleted)
		{
			if ($ThrowOnFail) { throw "Download_NSSM: Timed out waiting for NSSM. Check log: $script:NssmDownloadLog" }
			return @{ Success = $false; Path = $script:NssmExe; Note = 'Timed out waiting for NSSM'; Log = $script:NssmDownloadLog }
		}

		$jobResult = $null
		try { $jobResult = @($job.PS.EndInvoke($job.Handle)) | Select-Object -Last 1 }
		catch { }
		try { $job.PS.Dispose(); $job.Pool.Close(); $job.Pool.Dispose() }
		catch { }
		$script:NssmDownloadJob = $null

		if (Test-Path -LiteralPath $script:NssmExe)
		{
			return @{ Success = $true; Path = $script:NssmExe; Note = 'NSSM ready'; Log = $script:NssmDownloadLog }
		}
		$why = if ($jobResult -and $jobResult.Note) { [string]$jobResult.Note } else { 'Download failed' }
		if ($ThrowOnFail) { throw "Download_NSSM: $why. Check log: $script:NssmDownloadLog" }
		return @{ Success = $false; Path = $script:NssmExe; Note = $why; Log = $script:NssmDownloadLog }
	}

	# ------------------------------------------------------------------------------------------------
	# Synchronous (blocking): no -Background, or -Wait with no background download running
	# ------------------------------------------------------------------------------------------------
	try
	{
		$sb = [ScriptBlock]::Create($downloadScriptText)
		$result = & $sb $ToolsRoot ([bool]$ForceDownload.IsPresent) $script:NssmDir $script:NssmExe $script:NssmDownloadLog

		if ($result -and $result.Success)
		{
			return @{ Success = $true; Path = $script:NssmExe; Note = $result.Note; Log = $script:NssmDownloadLog }
		}

		if ($ThrowOnFail) { throw "Download_NSSM: Failed. Check log: $script:NssmDownloadLog" }
		return @{ Success = $false; Path = $script:NssmExe; Note = 'Failed'; Log = $script:NssmDownloadLog }
	}
	catch
	{
		if ($ThrowOnFail) { throw "Download_NSSM: Exception. $($_.Exception.Message). Log: $script:NssmDownloadLog" }
		return @{ Success = $false; Path = $script:NssmExe; Note = 'Exception'; Error = $_.Exception.Message; Log = $script:NssmDownloadLog }
	}
}
