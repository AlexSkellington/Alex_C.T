# ===================================================================================================
#                              FUNCTION: Update_Scales_Specials_Interactive
# ---------------------------------------------------------------------------------------------------
# Description:
#   "Update Scales Specials" wizard: run now, or schedule as a daily 5 AM task or a repeating
#   minutes task. Scheduling writes the batch file to the scripts folder and creates the
#   Windows scheduled task (Update_Scales_Specials_Task / ..._Minutes).
#
# Parameters:
#   (none - uses script context)
#
# Variables:
#   - $script:ScriptsFolder: where the batch files are written
#   - $OfficePath: office root holding DEPLOY_CHG.sql
#   - $script:ScaleManagementApp*PathCandidates: FastDEPLOY / regular / UpdateSpecials exe lookup
#   - $script:BmsToBizerbaPath / $script:BmsTerminalsPath: BMS folders used by the batch files
#   - $script:minutesValue: interval for the minutes task
#
# Notes:
#   - Schedule buttons enable/disable based on whether the required files exist.
# ===================================================================================================

function Update_Scales_Specials_Interactive
{
	Write_Log "`r`n==================== Starting Update_Scales_Specials_Interactive Function ====================`r`n" "blue"
	
	$DeployRestored = [ref]$false
	$scriptFolder = $script:ScriptsFolder
	$batchName_Daily = "Update_Scales_Specials.bat"
	$batchPath_Daily = Join-Path $scriptFolder $batchName_Daily
	$batchName_Minutes = "Update_Scales_Specials_Minutes.bat"
	$batchPath_Minutes = Join-Path $scriptFolder $batchName_Minutes
	
	$deployChgFile = Join-Path $OfficePath "DEPLOY_CHG.sql"
	$fastDeployPath = $script:ScaleManagementAppFastDeployPathCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	$regularDeployPath = $script:ScaleManagementAppPathCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	$updateSpecialsPath = $script:ScaleManagementAppUpdateSpecialsPathCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	if (-not $fastDeployPath) { $fastDeployPath = $script:ScaleManagementAppFastDeployPath }
	if (-not $regularDeployPath) { $regularDeployPath = $script:ScaleManagementAppExePath }
	if (-not $updateSpecialsPath) { $updateSpecialsPath = $script:ScaleManagementAppUpdateSpecialsPath }
	$toBizerbaPath = $script:BmsToBizerbaPath
	$terminalsPath = $script:BmsTerminalsPath
	$deployExecRegex = $script:ScaleCommDeployExecRegex
	
	# Check for all executables
	$hasDeployChg = Test-Path $deployChgFile -ErrorAction SilentlyContinue
	$hasFastDeploy = Test-Path -LiteralPath $fastDeployPath -ErrorAction SilentlyContinue
	$hasRegularDeploy = Test-Path -LiteralPath $regularDeployPath -ErrorAction SilentlyContinue
	$hasUpdateSpecials = Test-Path -LiteralPath $updateSpecialsPath -ErrorAction SilentlyContinue
	
	# Enable/disable schedule buttons based on file existence
	$enableDaily = $hasDeployChg -and $hasUpdateSpecials
	$enableMinutes = $hasDeployChg -and ($hasFastDeploy -or $hasRegularDeploy)
	
	if (-not $hasDeployChg)
	{
		Write_Log "DEPLOY_CHG.sql not found in $OfficePath" "yellow"
	}
	if (-not $hasUpdateSpecials)
	{
		Write_Log "ScaleManagementAppUpdateSpecials.exe not found for 5AM task at '$updateSpecialsPath'." "yellow"
	}
	if (-not ($hasFastDeploy -or $hasRegularDeploy))
	{
		Write_Log "Neither FastDEPLOY nor regular ScaleManagementApp.exe found in configured ScaleComm roots for minutes task." "yellow"
	}
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Update Scales Specials"
	$form.Size = New-Object System.Drawing.Size(495, 220)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Schedule Update Scales Specials as a daily (5 AM) or repeating task (in minutes)."
	$label.Location = New-Object System.Drawing.Point(20, 20)
	$label.Size = New-Object System.Drawing.Size(430, 40)
	$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$form.Controls.Add($label)
	
	$btnSchedule = New-Object System.Windows.Forms.Button
	$btnSchedule.Text = "Schedule Daily (5 AM)"
	$btnSchedule.Location = New-Object System.Drawing.Point(5, 90)
	$btnSchedule.Size = New-Object System.Drawing.Size(150, 40)
	$btnSchedule.Enabled = $enableDaily
	$form.Controls.Add($btnSchedule)
	
	$btnRestoreDeployLine = New-Object System.Windows.Forms.Button
	$btnRestoreDeployLine.Text = "Restore DEPLOY_CHG.sql Line"
	$btnRestoreDeployLine.Location = New-Object System.Drawing.Point(155, 90)
	$btnRestoreDeployLine.Size = New-Object System.Drawing.Size(170, 40)
	$btnRestoreDeployLine.Enabled = $false
	$form.Controls.Add($btnRestoreDeployLine)
	
	$btnScheduleMinutes = New-Object System.Windows.Forms.Button
	$btnScheduleMinutes.Text = "Schedule Task (Minutes)"
	$btnScheduleMinutes.Location = New-Object System.Drawing.Point(325, 90)
	$btnScheduleMinutes.Size = New-Object System.Drawing.Size(150, 40)
	$btnScheduleMinutes.Enabled = $enableMinutes
	$form.Controls.Add($btnScheduleMinutes)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(200, 145)
	$btnCancel.Size = New-Object System.Drawing.Size(100, 30)
	$form.Controls.Add($btnCancel)
	
	# --- If missing any necessary files, label displays ---
	if (-not $hasDeployChg)
	{
		$label.Text += "`r`n(DEPLOY_CHG.sql missing - scheduling is disabled)"
	}
	if (-not $hasUpdateSpecials)
	{
		$label.Text += "`r`n(ScaleManagementAppUpdateSpecials.exe missing - 5AM scheduling is disabled)"
	}
	if (-not ($hasFastDeploy -or $hasRegularDeploy))
	{
		$label.Text += "`r`n(Neither FastDEPLOY nor regular ScaleManagementApp.exe present - minutes scheduling is disabled)"
	}
	
	# Determine the correct @EXEC line for deploy file (minutes schedule always prefers FAST if present)
	$correctExeLine = ""
	if ($hasFastDeploy)
	{
		$correctExeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='$fastDeployPath');"
	}
	elseif ($hasRegularDeploy)
	{
		$correctExeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='$regularDeployPath');"
	}
	
	# Enable restore button if the exact correct line is missing (exists but wrong format will enable it)
	$deployLineCorrect = $false
	if (Test-Path $deployChgFile)
	{
		$deployContent = Get-Content $deployChgFile -Raw
		if ($deployContent -match [regex]::Escape($correctExeLine))
		{
			$deployLineCorrect = $true
		}
	}
	if (-not $deployLineCorrect) { $btnRestoreDeployLine.Enabled = $true }
	
	$btnRestoreDeployLine.Add_Click({
			# --- Restore DEPLOY_CHG.sql ---
			if (Test-Path $deployChgFile)
			{
				try
				{
					$content = Get-Content $deployChgFile -Raw
					$newContent = [System.Collections.Generic.List[string]]@(
						($content -split "`r?`n") | Where-Object {
							$_ -notmatch '^\s*/\* Deploy price changes to the scales \*/' -and
							$_ -notmatch $deployExecRegex
						}
					)
					while ($newContent.Count -gt 0 -and ($newContent[-1] -match '^\s*$'))
					{
						$null = $newContent.RemoveAt($newContent.Count - 1)
					}
					$newContent += ""
					$newContent += $correctExeLine
					$newContent -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default
					Write_Log "Restored line to DEPLOY_CHG.sql: $correctExeLine" "green"
					$btnRestoreDeployLine.Enabled = $false
				}
				catch
				{
					Write_Log "Failed to restore line to DEPLOY_CHG.sql: $_" "red"
				}
			}
			else
			{
				Write_Log "DEPLOY_CHG.sql not found for restore." "yellow"
			}
			
			# --- Delete the minutes task if it exists ---
			$minutesTaskName = "Update_Scales_Specials_Task_Minutes"
			$taskExists = schtasks /Query /TN "$minutesTaskName" 2>&1 | Select-String -Quiet -Pattern "$minutesTaskName"
			if ($taskExists)
			{
				$deleteOut = schtasks /Delete /TN "$minutesTaskName" /F 2>&1
				if ($LASTEXITCODE -eq 0)
				{
					Write_Log "Deleted scheduled task '$minutesTaskName' after DEPLOY_CHG.sql restore." "yellow"
				}
				else
				{
					Write_Log "Attempted to delete '$minutesTaskName' after restore but schtasks said: $deleteOut" "yellow"
				}
			}
			
			Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
			$DeployRestored.Value = $true
			$form.Close()
			return
		})
	
	$form.add_FormClosing({
			if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK)
			{
				$script:selectedAction = "cancel"
			}
		})
	
	$selectedAction = $null
	$minutesValue = $null
	
	$btnSchedule.Add_Click({
			$script:selectedAction = "schedule"
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})
	$btnScheduleMinutes.Add_Click({
			$inputForm = New-Object System.Windows.Forms.Form
			$inputForm.Text = "Set Minute Interval"
			$inputForm.Size = New-Object System.Drawing.Size(320, 140)
			$inputForm.StartPosition = "CenterParent"
			
			$lblInput = New-Object System.Windows.Forms.Label
			$lblInput.Text = "How many minutes between runs? (1-1439):"
			$lblInput.Location = New-Object System.Drawing.Point(10, 15)
			$lblInput.Size = New-Object System.Drawing.Size(280, 25)
			$inputForm.Controls.Add($lblInput)
			
			$txtInput = New-Object System.Windows.Forms.TextBox
			$txtInput.Location = New-Object System.Drawing.Point(10, 45)
			$txtInput.Size = New-Object System.Drawing.Size(120, 25)
			$inputForm.Controls.Add($txtInput)
			
			$okBtn = New-Object System.Windows.Forms.Button
			$okBtn.Text = "OK"
			$okBtn.Location = New-Object System.Drawing.Point(150, 42)
			$okBtn.Size = New-Object System.Drawing.Size(60, 25)
			$okBtn.Add_Click({
					$inputForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
					$inputForm.Close()
				})
			$inputForm.Controls.Add($okBtn)
			$inputForm.AcceptButton = $okBtn
			
			if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
			{
				$num = $txtInput.Text
				if ($num -match '^\d+$' -and [int]$num -ge 1 -and [int]$num -le 1439)
				{
					$script:selectedAction = "schedule_minutes"
					$script:minutesValue = [int]$num
					$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
					$form.Close()
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Please enter a valid number between 1 and 1439.", "Invalid Interval", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				}
			}
		})
	$btnCancel.Add_Click({
			$script:selectedAction = "cancel"
			$form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$form.Close()
		})
	
	$form.AcceptButton = $btnSchedule
	$form.CancelButton = $btnCancel
	
	$form.ShowDialog() | Out-Null
	if ($DeployRestored.Value) { return }
	if ($script:selectedAction -eq "cancel" -or -not $script:selectedAction)
	{
		Write_Log "User cancelled Update_Scales_Specials_Interactive." "yellow"
		Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
		return
	}
	
	if (-not (Test-Path $scriptFolder -ErrorAction SilentlyContinue)) { New-Item -Path $scriptFolder -ItemType Directory | Out-Null }
	
	# --- Batch for daily (UpdateSpecials) ---
	if ($script:selectedAction -eq "schedule")
	{
		$batchContent = @"
if "%1" == "" start "" /min "%~f0" MY_FLAG && exit
taskkill /IM ScaleManagementApp.exe /F
taskkill /IM BMSSrv.exe /F
taskkill /IM BMS.exe /F
del /s /q "$toBizerbaPath\*.*"
rmdir /s /q "$terminalsPath" >nul 2>&1
net start BMS /Y
start "" "$updateSpecialsPath"
exit
"@
		Set-Content -Path $batchPath_Daily -Value $batchContent -Encoding ASCII
		
		$taskName = "Update_Scales_Specials_Task"
		$schtasks = "schtasks /create /tn `"$taskName`" /tr `"$batchPath_Daily`" /sc DAILY /st 05:00 /rl HIGHEST /f /ru $env:COMPUTERNAME\Administrator"
		
		Invoke-Expression $schtasks
		
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Scheduled task created successfully for Update_Scales_Specials_Task (daily at 5 AM)." "green"
		}
		else
		{
			Write_Log "Failed to create scheduled task for Update_Scales_Specials_Task." "red"
		}
		Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
		return
	}
	
	# --- Batch for minutes (FastDEPLOY) ---
	if ($script:selectedAction -eq "schedule_minutes" -and $script:minutesValue)
	{
		if (Test-Path $deployChgFile -ErrorAction SilentlyContinue)
		{
			try
			{
				$lines = [System.Collections.Generic.List[string]]@(Get-Content $deployChgFile)
				$toRemoveIdx = @()
				for ($i = 0; $i -lt $lines.Count; $i++)
				{
					if (
						$lines[$i] -match '^\s*/\* Deploy price changes to the scales \*/' -or
						$lines[$i] -match $deployExecRegex
					)
					{
						if ($i -gt 0 -and $lines[$i - 1] -match '^\s*$') { $toRemoveIdx += ($i - 1) }
						$toRemoveIdx += $i
					}
				}
				$toRemoveIdx = $toRemoveIdx | Sort-Object -Unique -Descending
				foreach ($idx in $toRemoveIdx) { $lines.RemoveAt($idx) }
				while ($lines.Count -gt 0 -and $lines[-1] -match '^\s*$') { $null = $lines.RemoveAt($lines.Count - 1) }
				$lines -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default
				if ($toRemoveIdx.Count -gt 0)
				{
					Write_Log "Removed banner, @EXEC line, and any blank line above from $deployChgFile" "green"
				}
				else
				{
					Write_Log "No matching lines found in DEPLOY_CHG.sql for removal." "yellow"
				}
			}
			catch
			{
				Write_Log "Failed to update [$deployChgFile]: $_" "red"
			}
		}
		else
		{
			Write_Log "DEPLOY_CHG.sql not found in $OfficePath" "yellow"
		}
		
		$batchContent = @"
if "%1" == "" start "" /min "%~f0" MY_FLAG && exit
taskkill /IM ScaleManagementApp.exe /F
taskkill /IM BMSSrv.exe /F
taskkill /IM BMS.exe /F
del /s /q "$toBizerbaPath\*.*"
rmdir /s /q "$terminalsPath" >nul 2>&1
net start BMS /Y
if exist "$fastDeployPath" (
    start "" "$fastDeployPath"
) else (
    start "" "$regularDeployPath"
)
exit
"@
		Set-Content -Path $batchPath_Minutes -Value $batchContent -Encoding ASCII
		
		$taskName = "Update_Scales_Specials_Task_Minutes"
		$interval = [int]$script:minutesValue
		$schtasks = "schtasks /create /tn `"$taskName`" /tr `"$batchPath_Minutes`" /sc MINUTE /mo $interval /rl HIGHEST /f /ru $env:COMPUTERNAME\Administrator"
		
		Invoke-Expression $schtasks
		
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Scheduled task created successfully for Update_Scales_Specials_Task_Minutes (every $interval minutes)." "green"
		}
		else
		{
			Write_Log "Failed to create scheduled task for Update_Scales_Specials_Task_Minutes." "red"
		}
		Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
		return
	}
}
