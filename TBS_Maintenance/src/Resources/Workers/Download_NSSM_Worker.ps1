# ---------------------------------------------------------------------------------------------------
# NSSM download worker - embedded at build time into Download_NSSM ($downloadScriptText).
#   Runs in a 1-thread runspace (Download_NSSM -Background) or inline. Tries the fawno GitHub
#   mirror, then nssm.cc (Invoke-WebRequest -> BITS -> WebClient per URL), extracts, picks the
#   win64/win32 binary, copies it to NssmExe and checks it is > 50 KB.
#   Returns @{ Success; Path; Note; [Error] }; every step goes to LogFile (overwritten per run).
# ---------------------------------------------------------------------------------------------------
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
