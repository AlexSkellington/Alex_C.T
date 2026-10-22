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
	# Download script (runs in current session OR a background runspace)
	#   - NO nested functions (uses scriptblocks)
	#   - Writes only to $LogFile (overwritten each run)
	# ------------------------------------------------------------------------------------------------
	$downloadScriptText = @'
param(
	[string]$ToolsRoot,
	[bool]$ForceDownload,
	[string]$NssmDir,
	[string]$NssmExe,
	[string]$LogFile
)

$ErrorActionPreference = "Stop"

# Log writer (scriptblock, not a function)
$logLine = {
	param([string]$Line)
	try
	{
		$ts = (Get-Date -Format "M/d/yyyy h:mm:ss tt")
		("[{0}] {1}" -f $ts, $Line) | Add-Content -Path $LogFile -Encoding UTF8
	}
	catch { }
}

# Overwrite log at start (latest results only)
try { "" | Set-Content -Path $LogFile -Encoding UTF8 -Force } catch { }

& $logLine "Worker start"
& $logLine ("ToolsRoot = {0}" -f $ToolsRoot)
& $logLine ("NssmDir   = {0}" -f $NssmDir)
& $logLine ("NssmExe   = {0}" -f $NssmExe)

try
{
	# Ensure dirs exist
	if (-not (Test-Path $ToolsRoot)) { New-Item -Path $ToolsRoot -ItemType Directory -Force | Out-Null }
	if (-not (Test-Path $NssmDir))   { New-Item -Path $NssmDir   -ItemType Directory -Force | Out-Null }

	# If already present and not forcing
	if (-not $ForceDownload -and (Test-Path $NssmExe))
	{
		& $logLine "Already present. Done."
		return @{ Success = $true; Path = $NssmExe; Note = 'Already present' }
	}

	# TLS 1.2
	try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

	$is64 = [Environment]::Is64BitOperatingSystem

	# URLs (multiple sources)
	$urls = @()
	if ($is64) { $urls += "https://github.com/fawno/nssm.cc/releases/download/v2.24.1/nssm-v2.24.1-Win64.zip" }
	else       { $urls += "https://github.com/fawno/nssm.cc/releases/download/v2.24.1/nssm-v2.24.1-Win32.zip" }
	$urls += "https://nssm.cc/release/nssm-2.24.zip"
	$urls += "https://nssm.cc/ci/nssm-2.24-101-g897c7ad.zip"

	# Temp workspace
	$tempRoot = Join-Path $env:TEMP ("NSSM_" + [Guid]::NewGuid().ToString("N"))
	$tempZip  = Join-Path $tempRoot "nssm.zip"
	$tempOut  = Join-Path $tempRoot "out"
	New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
	New-Item -Path $tempOut  -ItemType Directory -Force | Out-Null

	$downloaded = $false
	$lastErr    = ""

	foreach ($u in $urls)
	{
		& $logLine ("Trying URL: {0}" -f $u)

		# wipe old zip
		try { if (Test-Path $tempZip) { Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue } } catch { }

		# Method 1: Invoke-WebRequest (timeout)
		try
		{
			Invoke-WebRequest -Uri $u -OutFile $tempZip -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop

			# IMPORTANT FIX: parentheses so "-and" is an operator
			if ((Test-Path $tempZip) -and ((Get-Item $tempZip).Length -gt 0))
			{
				$downloaded = $true
				& $logLine "Download OK via Invoke-WebRequest"
				break
			}
		}
		catch
		{
			$lastErr = $_.Exception.Message
			& $logLine ("Invoke-WebRequest failed: {0}" -f $lastErr)
		}

		# Method 2: BITS (if available)
		try
		{
			if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue)
			{
				Start-BitsTransfer -Source $u -Destination $tempZip -Priority Foreground -ErrorAction Stop

				# IMPORTANT FIX: parentheses so "-and" is an operator
				if ((Test-Path $tempZip) -and ((Get-Item $tempZip).Length -gt 0))
				{
					$downloaded = $true
					& $logLine "Download OK via BITS"
					break
				}
			}
		}
		catch
		{
			$lastErr = $_.Exception.Message
			& $logLine ("BITS failed: {0}" -f $lastErr)
		}

		# Method 3: WebClient fallback
		try
		{
			$wc = New-Object System.Net.WebClient
			$wc.DownloadFile($u, $tempZip)

			# IMPORTANT FIX: parentheses so "-and" is an operator
			if ((Test-Path $tempZip) -and ((Get-Item $tempZip).Length -gt 0))
			{
				$downloaded = $true
				& $logLine "Download OK via WebClient"
				break
			}
		}
		catch
		{
			$lastErr = $_.Exception.Message
			& $logLine ("WebClient failed: {0}" -f $lastErr)
		}
	}

	if (-not $downloaded)
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine ("RESULT: FAILED - Download failed from all sources. Last error: {0}" -f $lastErr)
		return @{ Success = $false; Path = $NssmExe; Note = 'Download failed'; Error = $lastErr }
	}

	# Unblock zip
	try { Unblock-File -Path $tempZip -ErrorAction SilentlyContinue } catch { }

	# Extract
	$extracted = $false
	try
	{
		Expand-Archive -Path $tempZip -DestinationPath $tempOut -Force -ErrorAction Stop
		$extracted = $true
		& $logLine "Extract OK via Expand-Archive"
	}
	catch
	{
		try
		{
			Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
			[System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $tempOut)
			$extracted = $true
			& $logLine "Extract OK via ZipFile"
		}
		catch
		{
			$extracted = $false
			& $logLine ("Extract failed: {0}" -f $_.Exception.Message)
		}
	}

	if (-not $extracted)
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine "RESULT: FAILED - Extraction failed."
		return @{ Success = $false; Path = $NssmExe; Note = 'Extraction failed' }
	}

	# Find nssm.exe
	$exeCandidates = Get-ChildItem -Path $tempOut -Recurse -Filter "nssm.exe" -File -ErrorAction SilentlyContinue
	if (-not $exeCandidates -or $exeCandidates.Count -eq 0)
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine "RESULT: FAILED - nssm.exe not found in extracted content."
		return @{ Success = $false; Path = $NssmExe; Note = 'nssm.exe not found' }
	}

	# Choose proper binary
	$chosen = $null
	if ($is64)
	{
		$chosen = $exeCandidates | Where-Object { $_.FullName -match "(\\|/)(win64|x64|amd64)(\\|/)" } | Select-Object -First 1
		if (-not $chosen) { $chosen = $exeCandidates | Select-Object -First 1 }
	}
	else
	{
		$chosen = $exeCandidates | Where-Object { $_.FullName -match "(\\|/)(win32|x86|i386)(\\|/)" } | Select-Object -First 1
		if (-not $chosen) { $chosen = $exeCandidates | Select-Object -First 1 }
	}

	if (-not $chosen -or -not (Test-Path $chosen.FullName))
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine "RESULT: FAILED - Could not select a valid nssm.exe candidate."
		return @{ Success = $false; Path = $NssmExe; Note = 'Selection failed' }
	}

	& $logLine ("Selected: {0}" -f $chosen.FullName)

	# Copy to stable location
	Copy-Item -Path $chosen.FullName -Destination $NssmExe -Force -ErrorAction Stop

	# Validate
	if (-not (Test-Path $NssmExe) -or ((Get-Item $NssmExe).Length -lt 50KB))
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine ("RESULT: FAILED - Validation failed after copy to '{0}'." -f $NssmExe)
		return @{ Success = $false; Path = $NssmExe; Note = 'Validation failed' }
	}

	# Cleanup
	try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

	& $logLine ("RESULT: SUCCESS - NSSM ready at {0}" -f $NssmExe)
	return @{ Success = $true; Path = $NssmExe; Note = 'Downloaded/Updated' }
}
catch
{
	& $logLine ("RESULT: FAILED - Unhandled exception: {0}" -f $_.Exception.Message)
	return @{ Success = $false; Path = $NssmExe; Note = 'Exception'; Error = $_.Exception.Message }
}
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
	# Wait mode (only blocks when asked)
	# ------------------------------------------------------------------------------------------------
	if ($Wait)
	{
		$deadline = (Get-Date).AddSeconds($WaitTimeoutSeconds)
		
		while ((Get-Date) -lt $deadline)
		{
			if (Test-Path $script:NssmExe)
			{
				# Cleanup completed job if present
				if ($script:NssmDownloadJob -and $script:NssmDownloadJob.Handle)
				{
					try
					{
						if ($script:NssmDownloadJob.Handle.IsCompleted)
						{
							$null = $script:NssmDownloadJob.PS.EndInvoke($script:NssmDownloadJob.Handle)
							$script:NssmDownloadJob.PS.Dispose()
							$script:NssmDownloadJob.Pool.Close()
							$script:NssmDownloadJob.Pool.Dispose()
							$script:NssmDownloadJob = $null
						}
					}
					catch { }
				}
				
				return @{ Success = $true; Path = $script:NssmExe; Note = 'NSSM ready'; Log = $script:NssmDownloadLog }
			}
			
			# Keep WinForms responsive if applicable
			try { [System.Windows.Forms.Application]::DoEvents() }
			catch { }
			Start-Sleep -Milliseconds $WaitPollMilliseconds
		}
		
		if ($ThrowOnFail) { throw "Download_NSSM: Timed out waiting for NSSM. Check log: $script:NssmDownloadLog" }
		return @{ Success = $false; Path = $script:NssmExe; Note = 'Timed out waiting for NSSM'; Log = $script:NssmDownloadLog }
	}
	
	# ------------------------------------------------------------------------------------------------
	# Synchronous (blocking) if neither -Background nor -Wait
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
