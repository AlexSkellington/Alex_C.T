# ===================================================================================================
#                              FUNCTION: Get_PsExec
# ---------------------------------------------------------------------------------------------------
# Description:
#   Ensures Sysinternals PsExec.exe exists in the tools directory. If missing, downloads
#   PSTools.zip from Microsoft in a background job, extracts PsExec.exe, shows progress and
#   removes the zip. Reuses a running download job so it never double-downloads.
#
# Parameters:
#   - ToolsDir  [string]: tools folder, defaults to $script:ToolsDir
#
# Writes:
#   - <ToolsDir>\PsExec.exe (PSTools.zip while extracting)
#
# Returns:
#   Full path to PsExec.exe, or $null if it could not be produced.
# ===================================================================================================

function Get_PsExec
{
	param (
		[string]$ToolsDir = $script:ToolsDir
	)

	$psexecPath = Join-Path $ToolsDir "PsExec.exe"
	$pstoolsZip = Join-Path $ToolsDir "PSTools.zip"
	$pstoolsUrl = "https://download.sysinternals.com/files/PSTools.zip"
	$jobName = "Get_PsExec_Download_Job"

	# Check if PsExec.exe already exists
	if (Test-Path $psexecPath)
	{
		Write-Host "PsExec.exe is ready to be used at $psexecPath."
		return $psexecPath
	}

	# Check for existing running job (compatible method)
	$existingJob = Get-Job | Where-Object { $_.Name -eq $jobName -and $_.State -eq 'Running' }
	if ($existingJob)
	{
		Write-Host "Download job already running. Waiting for completion..."
		$job = $existingJob
	}
	else
	{
		Write-Host "PsExec.exe not found. Starting background download and extraction..."
		$job = Start-Job -Name $jobName -ScriptBlock {
			param ($pstoolsUrl,
				$pstoolsZip,
				$ToolsDir,
				$psexecPath)
			try
			{
				if (!(Test-Path $ToolsDir))
				{
					Write-Host "[Job] Creating directory: $ToolsDir"
					New-Item -Path $ToolsDir -ItemType Directory | Out-Null
				}
				Write-Host "[Job] Downloading PSTools.zip..."
				Invoke-WebRequest -Uri $pstoolsUrl -OutFile $pstoolsZip -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop

				Write-Host "[Job] Extracting PsExec.exe..."
				Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
				$zip = [System.IO.Compression.ZipFile]::OpenRead($pstoolsZip)
				$entry = $zip.Entries | Where-Object { $_.Name -ieq "PsExec.exe" }
				if ($entry)
				{
					# Manual extraction (compat with older PowerShell/ISE)
					$fs = $entry.Open()
					$bytes = New-Object byte[] $entry.Length
					[void]$fs.Read($bytes, 0, $entry.Length)
					$fs.Close()
					[System.IO.File]::WriteAllBytes($psexecPath, $bytes)
					Write-Host "[Job] PsExec.exe extracted to $psexecPath"
				}
				else
				{
					Write-Host "[Job] WARNING: PsExec.exe not found in ZIP!"
				}
				$zip.Dispose()

				if (Test-Path $pstoolsZip) { Remove-Item $pstoolsZip -ErrorAction SilentlyContinue }
				if (Test-Path $psexecPath)
				{
					Write-Host "[Job] PsExec is ready at $psexecPath"
				}
				else
				{
					Write-Host "[Job] WARNING: PsExec.exe not found after extraction."
				}
			}
			catch
			{
				Write-Host "[Job] ERROR: $($_.Exception.Message)"
				if (Test-Path $pstoolsZip) { Remove-Item $pstoolsZip -ErrorAction SilentlyContinue }
			}
		} -ArgumentList $pstoolsUrl, $pstoolsZip, $ToolsDir, $psexecPath
	}

	# Simple progress indicator (compatible with ISE)
	Write-Host -NoNewline "Downloading and extracting PsExec.exe"
	while ($job.State -eq "Running")
	{
		Write-Host -NoNewline "."
		Start-Sleep -Seconds 1
		$job = Get-Job | Where-Object { $_.Id -eq $job.Id }
	}
	Write-Host ""

	# Show job output
	Receive-Job -Id $job.Id | ForEach-Object { Write-Host $_ }
	Remove-Job -Id $job.Id -Force

	if (Test-Path $psexecPath)
	{
		Write-Host "All done! PsExec.exe is ready at $psexecPath"
		return $psexecPath
	}
	else
	{
		Write-Host "ERROR: PsExec.exe was not found after extraction. Check above for errors."
		return $null
	}
}
