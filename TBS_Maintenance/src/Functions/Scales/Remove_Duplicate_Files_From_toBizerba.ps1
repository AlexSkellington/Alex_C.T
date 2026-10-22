# ===================================================================================================
#                              FUNCTION: Remove_Duplicate_Files_From_toBizerba
# ---------------------------------------------------------------------------------------------------
# Description:
#   Duplicate File Monitor for C:\Bizerba\RetailConnect\BMS\toBizerba: prompt to run it now
#   (background job, controllable from the GUI) or install it as an always-on Windows service.
#   Removes duplicates by content hash (keeps the oldest by CreationTime) and also removes
#   files whose non-empty lines are fully contained in another file.
#
# Parameters:
#   (none)
#
# Variables:
#   - $script:BmsToBizerbaPath: folder being monitored
#   - $script:DuplicateFileMonitorServiceRoot: service root (ProgramData); watcher script goes in its Scripts subfolder
#   - $script:ScriptsLog: single combined log (service script + NSSM stdout/stderr)
#   - $script:FunctionResults['NssmExe'] / $script:NssmExe / $script:NssmDefaultPath: NSSM candidates
#
# Notes:
#   - Never downloads NSSM; if nssm.exe is missing, install/uninstall errors out and stops.
# ===================================================================================================

function Remove_Duplicate_Files_From_toBizerba
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Remove_Duplicate_Files_From_toBizerba Function ====================`r`n" "blue"
	
	# ==========================================================================================
	# CORE PATHS
	# ==========================================================================================
	$TargetPath = $script:BmsToBizerbaPath
	$serviceName = "Remove_Duplicate_Files_From_toBizerba"
	
	$serviceRoot = $script:DuplicateFileMonitorServiceRoot
	$svcScripts = Join-Path $serviceRoot "Scripts"
	
	$psScriptName = "Remove_Duplicate_Files_From_toBizerba.ps1"
	$psScriptPath = Join-Path $svcScripts $psScriptName
	
	# CHANGE: single combined log for everything (service script + NSSM stdout/stderr)
	$singleLog = Join-Path $script:ScriptsLog "Remove_Duplicate_Files_From_toBizerba.log"
	
	# ==========================================================================================
	# Ensure folders exist
	# ==========================================================================================
	foreach ($p in @($serviceRoot, $svcScripts))
	{
		if (-not (Test-Path -LiteralPath $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
	}
	
	# ==========================================================================================
	# Resolve NSSM executable (NO downloads)
	# ==========================================================================================
	$nssmExe = $null
	$nssmCandidates = @()
	
	try
	{
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('NssmExe') -and $script:FunctionResults['NssmExe'])
		{
			$nssmCandidates += [string]$script:FunctionResults['NssmExe']
		}
	}
	catch { }
	
	try
	{
		if ($script:NssmExe) { $nssmCandidates += [string]$script:NssmExe }
	}
	catch { }
	
	$nssmCandidates += $script:NssmDefaultPath
	
	foreach ($c in $nssmCandidates)
	{
		if ([string]::IsNullOrWhiteSpace($c)) { continue }
		if (Test-Path -LiteralPath $c) { $nssmExe = $c; break }
	}
	
	# ==========================================================================================
	# WRITE WATCHER SCRIPT (NO interpolation bugs; concise logging; never exits)
	# ==========================================================================================
	$psScriptContent = @'
#requires -Version 5.1
param([int]$IntervalSeconds = 5)

$ErrorActionPreference = 'Continue'

$Path    = '__TARGETPATH__'
$LogFile = '__LOGFILE__'

# Ensure log dir exists
try {
	$ld = Split-Path -Parent $LogFile
	if (-not (Test-Path -LiteralPath $ld)) { New-Item -Path $ld -ItemType Directory -Force | Out-Null }
} catch { }

# Concise log helper (scriptblock, not a function)
$log = {
	param([string]$m)
	try { Add-Content -LiteralPath $LogFile -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $m) -Encoding UTF8 } catch { }
}

# Ensure target folder exists (don't exit if missing)
try {
	if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
} catch { }

& $log ("Service script started. Interval={0}s Path={1}" -f $IntervalSeconds, $Path)

while ($true)
{
	try
	{
		# Re-create folder if deleted
		if (-not (Test-Path -LiteralPath $Path))
		{
			try { New-Item -Path $Path -ItemType Directory -Force | Out-Null } catch { }
			Start-Sleep -Seconds $IntervalSeconds
			continue
		}

		$removedHash = 0
		$removedLine = 0
		$removedSamples = New-Object System.Collections.Generic.List[string]

		# --------------------------
		# 1) Remove duplicates by SHA256 hash (keep oldest by CreationTime)
		# --------------------------
		$files = Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue
		$hashTable = @{}

		foreach ($file in $files)
		{
			try { if ($file.Length -eq 0) { continue } } catch { }

			try { $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash }
			catch { continue }

			if (-not $hashTable.ContainsKey($hash)) { $hashTable[$hash] = @() }
			$hashTable[$hash] += $file
		}

		foreach ($entry in $hashTable.GetEnumerator())
		{
			$list = $entry.Value
			if ($list.Count -gt 1)
			{
				$list = $list | Sort-Object CreationTime
				$keep = $list[0]
				$dups = $list[1..($list.Count - 1)]

				foreach ($d in $dups)
				{
					try
					{
						Remove-Item -LiteralPath $d.FullName -Force -ErrorAction Stop
						$removedHash++
						if ($removedSamples.Count -lt 5) { [void]$removedSamples.Add(("HASH:{0}" -f $d.Name)) }
					}
					catch { }
				}
			}
		}

		# --------------------------
		# 2) Remove by line containment
		# --------------------------
		$files = Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue

		for ($i = 0; $i -lt $files.Count; $i++)
		{
			$fileA = $files[$i]

			$linesA = Get-Content -LiteralPath $fileA.FullName -ErrorAction SilentlyContinue |
				Where-Object { $_ -and $_.Trim() -ne "" }

			if (-not $linesA -or $linesA.Count -eq 0) { continue }

			for ($j = 0; $j -lt $files.Count; $j++)
			{
				if ($i -eq $j) { continue }
				$fileB = $files[$j]

				$linesB = Get-Content -LiteralPath $fileB.FullName -ErrorAction SilentlyContinue |
					Where-Object { $_ -and $_.Trim() -ne "" }

				if (-not $linesB) { continue }
				if ($linesB.Count -lt $linesA.Count) { continue }

				$setB = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
				foreach ($lb in $linesB) { [void]$setB.Add($lb.Trim()) }

				$allFound = $true
				foreach ($la in $linesA)
				{
					if (-not $setB.Contains($la.Trim())) { $allFound = $false; break }
				}

				if ($allFound)
				{
					$canDelete = $true
					try
					{
						$s = [System.IO.File]::Open($fileA.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
						$s.Close()
					}
					catch { $canDelete = $false }

					if ($canDelete)
					{
						try
						{
							Remove-Item -LiteralPath $fileA.FullName -Force -ErrorAction Stop
							$removedLine++
							if ($removedSamples.Count -lt 5) { [void]$removedSamples.Add(("LINES:{0}" -f $fileA.Name)) }

							# Refresh and restart to keep indexing safe
							$files = Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue
							$i = -1
							break
						}
						catch { }
					}
				}
			}
		}

		# Concise: only log when something was removed
		if (($removedHash + $removedLine) -gt 0)
		{
			$sampleText = ""
			if ($removedSamples.Count -gt 0) { $sampleText = " Samples: " + ($removedSamples -join ", ") }
			& $log ("Removed: Hash={0}, Lines={1}.{2}" -f $removedHash, $removedLine, $sampleText)
		}
	}
	catch
	{
		& $log ("ERROR: " + $_.Exception.Message)
	}

	Start-Sleep -Seconds $IntervalSeconds
}
'@
	
	# Inject paths safely into the single-quoted script (escape single quotes for PS literals)
	$tpEsc = $TargetPath.Replace("'", "''")
	$lfEsc = $singleLog.Replace("'", "''")
	
	$psScriptContent = $psScriptContent.Replace('__TARGETPATH__', $tpEsc)
	$psScriptContent = $psScriptContent.Replace('__LOGFILE__', $lfEsc)
	
	Set-Content -LiteralPath $psScriptPath -Value $psScriptContent -Encoding UTF8 -Force
	if (-not (Test-Path -LiteralPath $psScriptPath))
	{
		Write_Log "ERROR: Failed to write watcher script: $psScriptPath" "red"
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}
	
	# ==========================================================================================
	# SERVICE EXISTS?
	# ==========================================================================================
	$hasService = [bool](Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
	
	# ==========================================================================================
	# GUI
	# ==========================================================================================
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Duplicate File Monitor"
	$form.Size = New-Object System.Drawing.Size(440, 230)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "How do you want to run the Duplicate File Monitor for:`r`n$TargetPath"
	$label.Location = New-Object System.Drawing.Point(16, 15)
	$label.Size = New-Object System.Drawing.Size(390, 40)
	$label.TextAlign = 'MiddleCenter'
	$form.Controls.Add($label)
	
	$btnRunNow = New-Object System.Windows.Forms.Button
	$btnRunNow.Text = "Run Now (as Job)"
	$btnRunNow.Location = New-Object System.Drawing.Point(10, 70)
	$btnRunNow.Size = New-Object System.Drawing.Size(110, 35)
	$form.Controls.Add($btnRunNow)
	
	$btnService = New-Object System.Windows.Forms.Button
	$btnService.Text = "Install Service"
	$btnService.Location = New-Object System.Drawing.Point(125, 70)
	$btnService.Size = New-Object System.Drawing.Size(140, 35)
	$form.Controls.Add($btnService)
	
	$btnDeleteService = New-Object System.Windows.Forms.Button
	$btnDeleteService.Text = "Uninstall Service"
	$btnDeleteService.Location = New-Object System.Drawing.Point(270, 70)
	$btnDeleteService.Size = New-Object System.Drawing.Size(140, 35)
	$btnDeleteService.Enabled = $hasService
	$form.Controls.Add($btnDeleteService)
	
	$lblSec = New-Object System.Windows.Forms.Label
	$lblSec.Text = "Interval (seconds):"
	$lblSec.Location = New-Object System.Drawing.Point(70, 120)
	$lblSec.Size = New-Object System.Drawing.Size(110, 22)
	$form.Controls.Add($lblSec)
	
	$numSec = New-Object System.Windows.Forms.NumericUpDown
	$numSec.Location = New-Object System.Drawing.Point(185, 118)
	$numSec.Size = New-Object System.Drawing.Size(50, 24)
	$numSec.Minimum = 1
	$numSec.Maximum = 3600
	$numSec.Value = 5
	$form.Controls.Add($numSec)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(10, 148)
	$btnCancel.Size = New-Object System.Drawing.Size(400, 35)
	$form.Controls.Add($btnCancel)
	
	$state = @{ Action = $null; Interval = 5 }
	$monitorJob = $null
	
	$form.Add_FormClosing({
			if ($monitorJob -and $monitorJob.State -eq 'Running')
			{
				try { Stop-Job -Job $monitorJob -Force; Remove-Job -Job $monitorJob -Force; Write_Log "Monitor job stopped on form close." "yellow" }
				catch { }
			}
			
			Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
				$_.CommandLine -and $_.CommandLine -match [Regex]::Escape($psScriptPath)
			} | ForEach-Object {
				try { Stop-Process -Id $_.ProcessId -Force; Write_Log "Stopped powershell.exe PID $($_.ProcessId) on form close" "yellow" }
				catch { }
			}
		})
	
	$btnRunNow.Add_Click({
			if ($btnRunNow.Text -eq "Run Now (as Job)")
			{
				$state.Interval = [int]$numSec.Value
				try { "" | Set-Content -LiteralPath $singleLog -Encoding UTF8 -Force }
				catch { }
				
				$monitorJob = Start-Job -Name 'DuplicateFileMonitorJob' -ScriptBlock {
					param ($scriptPath,
						$interval,
						$logFile)
					& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -IntervalSeconds $interval *>> $logFile
				} -ArgumentList $psScriptPath, $state.Interval, $singleLog
				
				Write_Log "Started monitor job (ID: $($monitorJob.Id)). Log: $singleLog" "green"
				
				$btnRunNow.Text = "Stop Monitor"
				$btnService.Enabled = $false
				$btnDeleteService.Enabled = $false
				$btnCancel.Text = "Close (Monitor Running)"
			}
			else
			{
				if ($monitorJob -and $monitorJob.State -eq 'Running')
				{
					try { Stop-Job -Job $monitorJob -Force; Remove-Job -Job $monitorJob -Force; Write_Log "Monitor job stopped by user." "yellow" }
					catch { }
				}
				
				Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
					$_.CommandLine -and $_.CommandLine -match [Regex]::Escape($psScriptPath)
				} | ForEach-Object {
					try { Stop-Process -Id $_.ProcessId -Force; Write_Log "Stopped powershell.exe PID $($_.ProcessId) by user stop" "yellow" }
					catch { }
				}
				
				$btnRunNow.Text = "Run Now (as Job)"
				$btnService.Enabled = $true
				$btnDeleteService.Enabled = $hasService
				$btnCancel.Text = "Cancel"
			}
		})
	
	$btnService.Add_Click({ $state.Interval = [int]$numSec.Value; $state.Action = "service"; $form.Close() })
	$btnDeleteService.Add_Click({ $state.Action = "uninstall"; $form.Close() })
	$btnCancel.Add_Click({ $form.Close() })
	
	$form.ShowDialog() | Out-Null
	
	if (-not $state.Action)
	{
		Write_Log "User cancelled Duplicate File Monitor dialog." "yellow"
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}
	
	# ==========================================================================================
	# ADMIN CHECK FOR SERVICE ACTIONS
	# ==========================================================================================
	if ($state.Action -in @("service", "uninstall"))
	{
		$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
		).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
		
		if (-not $IsAdmin)
		{
			Write_Log "Service install/uninstall requires Administrator. Re-run as Admin." "red"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}
	}
	
	# ==========================================================================================
	# UNINSTALL (SC.EXE ONLY)
	# ==========================================================================================
	if ($state.Action -eq "uninstall")
	{
		if (-not (Get-Service -Name $serviceName -ErrorAction SilentlyContinue))
		{
			Write_Log "Service '$serviceName' is not installed." "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}
		
		Write_Log "Stopping service '$serviceName' before uninstall..." "yellow"
		try { & $env:ComSpec /c ("sc stop `"{0}`"" -f $serviceName) *> $null }
		catch { }
		try { Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue }
		catch { }
		try { & $env:ComSpec /c ("sc delete `"{0}`"" -f $serviceName) *> $null }
		catch { }
		
		Start-Sleep -Milliseconds 800
		if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
		{
			Write_Log "ERROR: Service still exists after sc delete. Remove manually in Services.msc." "red"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}
		
		Write_Log "Service '$serviceName' uninstalled successfully." "green"
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}
	
	# ==========================================================================================
	# INSTALL SERVICE (requires NSSM present - NO downloads)
	# ==========================================================================================
	if ($state.Action -eq "service")
	{
		if (-not $nssmExe -or -not (Test-Path -LiteralPath $nssmExe))
		{
			Write_Log "ERROR: NSSM is missing. This function will not download it." "red"
			Write_Log "Expected: $script:NssmDefaultPath" "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}
		
		# Fresh log for easy debugging
		try { "" | Set-Content -LiteralPath $singleLog -Encoding UTF8 -Force }
		catch { }
		
		# Remove existing service (clean reinstall)
		if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
		{
			try { & $env:ComSpec /c ("sc stop `"{0}`"" -f $serviceName) *> $null }
			catch { }
			try { Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue }
			catch { }
			try { & $env:ComSpec /c ("sc delete `"{0}`"" -f $serviceName) *> $null }
			catch { }
			Start-Sleep -Seconds 1
		}
		
		$psExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
		$interval = [int]$state.Interval
		
		# Install with args as separate tokens (no quoting issues)
		& $nssmExe install $serviceName $psExe `
				   '-NoProfile' `
				   '-ExecutionPolicy' 'Bypass' `
				   '-File' $psScriptPath `
				   '-IntervalSeconds' "$interval"
		
		if ($LASTEXITCODE -ne 0)
		{
			Write_Log "ERROR: NSSM install failed (ExitCode=$LASTEXITCODE). NSSM: $nssmExe" "red"
			Write_Log "Log: $singleLog" "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}
		
		# Configure service
		& $nssmExe set $serviceName AppDirectory $svcScripts *> $null
		& $nssmExe set $serviceName Start SERVICE_AUTO_START *> $null
		& $nssmExe set $serviceName ObjectName "LocalSystem" *> $null
		
		# ONE log file for NSSM stdout + stderr
		& $nssmExe set $serviceName AppStdout $singleLog *> $null
		& $nssmExe set $serviceName AppStderr $singleLog *> $null
		& $nssmExe set $serviceName AppRotateFiles 0 *> $null
		
		& $nssmExe start $serviceName *> $null
		if ($LASTEXITCODE -ne 0)
		{
			Write_Log "ERROR: Service installed but failed to start (ExitCode=$LASTEXITCODE)." "red"
			Write_Log "Log: $singleLog" "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}
		
		Start-Sleep -Milliseconds 800
		$s = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
		if ($s -and $s.Status -ne 'Running')
		{
			Write_Log "WARNING: Service not Running (Status=$($s.Status)). Check log: $singleLog" "yellow"
		}
		else
		{
			Write_Log "Service '$serviceName' installed and started. Interval: $interval sec." "green"
			Write_Log "Single log: $singleLog" "green"
			Write_Log "NSSM used: $nssmExe" "gray"
		}
		
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}
	
	Write_Log "No valid action selected. Exiting." "yellow"
	Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
}
