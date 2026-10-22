# ===================================================================================================
#                              FUNCTION: Repair_Windows
# ---------------------------------------------------------------------------------------------------
# Description:
#   Local Windows repair suite with a confirmation prompt and a checkbox form to pick the
#   operations: Defender update + full scan, DISM, SFC, disk cleanup, drive optimization
#   (SSD trim / HDD defrag) and scheduled Check Disk. Logs progress after each step.
#
# Writes:
#   - local system state (Defender, DISM, SFC, cleanup, drive optimization, scheduled chkdsk)
# ===================================================================================================

function Repair_Windows
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Repair_Windows Function ====================`r`n" "blue"
	
	# Import GUI assemblies if needed
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Confirm user intent
	$confirmationResult = [System.Windows.Forms.MessageBox]::Show(
		"The Windows repair process will take a long time and make significant changes to your system. Do you want to proceed?",
		"Confirmation Required",
		[System.Windows.Forms.MessageBoxButtons]::YesNo,
		[System.Windows.Forms.MessageBoxIcon]::Warning
	)
	if ($confirmationResult -ne [System.Windows.Forms.DialogResult]::Yes)
	{
		Write_Log "Windows repair process cancelled by the user." "yellow"
		return
	}
	
	Write_Log "Starting Windows repair process. This might take a while, please wait..." "blue"
	
	# Create the Repair Options Form
	$repairForm = New-Object System.Windows.Forms.Form
	$repairForm.Text = "Select Repair Operations"
	$repairForm.Size = New-Object System.Drawing.Size(400, 380)
	$repairForm.StartPosition = "CenterScreen"
	$repairForm.FormBorderStyle = 'FixedDialog'
	$repairForm.MaximizeBox = $false
	$repairForm.MinimizeBox = $false
	$repairForm.ShowInTaskbar = $false
	
	# Operation Checkboxes (with correct .Location and .Size types)
	$checkboxes = @()
	
	$cb1 = New-Object System.Windows.Forms.CheckBox
	$cb1.Text = "Windows Defender Update and Scan"
	$cb1.Location = [System.Drawing.Point]::new(20, 20)
	$cb1.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb1
	
	$cb2 = New-Object System.Windows.Forms.CheckBox
	$cb2.Text = "Run DISM Commands"
	$cb2.Location = [System.Drawing.Point]::new(20, 60)
	$cb2.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb2
	
	$cb3 = New-Object System.Windows.Forms.CheckBox
	$cb3.Text = "Run System File Checker (SFC)"
	$cb3.Location = [System.Drawing.Point]::new(20, 100)
	$cb3.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb3
	
	$cb4 = New-Object System.Windows.Forms.CheckBox
	$cb4.Text = "Disk Cleanup"
	$cb4.Location = [System.Drawing.Point]::new(20, 140)
	$cb4.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb4
	
	$cb5 = New-Object System.Windows.Forms.CheckBox
	$cb5.Text = "Optimize Drives"
	$cb5.Location = [System.Drawing.Point]::new(20, 180)
	$cb5.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb5
	
	$cb6 = New-Object System.Windows.Forms.CheckBox
	$cb6.Text = "Schedule Check Disk"
	$cb6.Location = [System.Drawing.Point]::new(20, 220)
	$cb6.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb6
	
	foreach ($cb in $checkboxes) { $repairForm.Controls.Add($cb) }
	
	# "Select All" checkbox
	$cbAll = New-Object System.Windows.Forms.CheckBox
	$cbAll.Text = "Select All"
	$cbAll.Location = [System.Drawing.Point]::new(20, 260)
	$cbAll.Size = [System.Drawing.Size]::new(350, 25)
	$cbAll.Add_CheckedChanged({
			foreach ($cb in $checkboxes) { $cb.Checked = $cbAll.Checked }
		})
	$repairForm.Controls.Add($cbAll)
	
	# Enable/Disable Run Button logic
	$runButton = New-Object System.Windows.Forms.Button
	$runButton.Text = "Run"
	$runButton.Location = [System.Drawing.Point]::new(150, 300)
	$runButton.Size = [System.Drawing.Size]::new(100, 30)
	$runButton.Enabled = $false
	$repairForm.Controls.Add($runButton)
	foreach ($cb in $checkboxes)
	{
		$cb.Add_CheckedChanged({ $runButton.Enabled = ($checkboxes | Where-Object { $_.Checked }).Count -gt 0 })
	}
	
	# Show form and get selections
	$selectedParams = @{ }
	$runButton.Add_Click({
			$selectedParams.Defender = $cb1.Checked
			$selectedParams.DISM = $cb2.Checked
			$selectedParams.SFC = $cb3.Checked
			$selectedParams.DiskCleanup = $cb4.Checked
			$selectedParams.OptimizeDrives = $cb5.Checked
			$selectedParams.CheckDisk = $cb6.Checked
			$repairForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$repairForm.Close()
		})
	
	$dialogResult = $repairForm.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Windows repair process cancelled by the user." "yellow"
		return
	}
	
	# -------------------- Operations Begin --------------------
	
	if ($selectedParams.Defender)
	{
		try
		{
			Write_Log "Updating Windows Defender signatures..." "blue"
			$updateOutput = & "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -SignatureUpdate 2>&1
			Write_Log "Windows Defender signatures update output: $updateOutput" "cyan"
			Write_Log "Windows Defender signatures updated successfully." "green"
			Write_Log "Running Windows Defender full scan..." "blue"
			$scanOutput = & "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -Scan -ScanType 2 2>&1
			Write_Log "Windows Defender full scan output: $scanOutput" "cyan"
			Write_Log "Windows Defender full scan completed." "green"
		}
		catch { Write_Log "Defender update/scan failed: $_" "red" }
	}
	else { Write_Log "Skipping Windows Defender update and scan as per user request." "yellow" }
	
	if ($selectedParams.DISM)
	{
		try
		{
			Write_Log "Running DISM RestoreHealth..." "blue"
			$dismRestoreOutput = DISM /Online /Cleanup-Image /RestoreHealth /Source:WindowsUpdate /NoRestart 2>&1
			Write_Log "DISM RestoreHealth output: $dismRestoreOutput" "cyan"
			Write_Log "DISM RestoreHealth completed." "green"
			Write_Log "Running DISM StartComponentCleanup..." "blue"
			$dismCleanupOutput = DISM /Online /Cleanup-Image /StartComponentCleanup /NoRestart 2>&1
			Write_Log "DISM StartComponentCleanup output: $dismCleanupOutput" "cyan"
			Write_Log "DISM StartComponentCleanup completed." "green"
		}
		catch { Write_Log "DISM failed: $_" "red" }
	}
	else { Write_Log "Skipping DISM operations as per user request." "yellow" }
	
	if ($selectedParams.SFC)
	{
		try
		{
			Write_Log "Running System File Checker (SFC)..." "blue"
			$sfcOutput = SFC /scannow 2>&1
			Write_Log "System File Checker output: $sfcOutput" "cyan"
			Write_Log "System File Checker completed." "green"
		}
		catch { Write_Log "SFC failed: $_" "red" }
	}
	else { Write_Log "Skipping System File Checker as per user request." "yellow" }
	
	if ($selectedParams.DiskCleanup)
	{
		try
		{
			Write_Log "Running Disk Cleanup..." "blue"
			Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait
			Write_Log "Disk Cleanup completed." "green"
		}
		catch { Write_Log "Disk Cleanup failed: $_" "red" }
	}
	else { Write_Log "Skipping Disk Cleanup as per user request." "yellow" }
	
	if ($selectedParams.OptimizeDrives)
	{
		try
		{
			Write_Log "Optimizing all fixed drives..." "blue"
			Get-Volume | Where-Object { $_.DriveType -eq 'Fixed' -and $_.DriveLetter } | ForEach-Object {
				Write_Log "Optimizing drive: $($_.DriveLetter)" "blue"
				$optimizeOutput = Optimize-Volume -DriveLetter $_.DriveLetter -Verbose 2>&1
				Write_Log "Optimize drive output: $optimizeOutput" "cyan"
			}
			Write_Log "Disk optimization completed." "green"
		}
		catch { Write_Log "Drive optimization failed: $_" "red" }
	}
	else { Write_Log "Skipping disk optimization as per user request." "yellow" }
	
	if ($selectedParams.CheckDisk)
	{
		try
		{
			Write_Log "Scheduling Check Disk on C: drive..." "blue"
			$checkDiskOutput = Start-Process "cmd.exe" -ArgumentList "/c echo Y|chkdsk C: /f /r" -Verb RunAs -Wait -PassThru
			Write_Log "Check Disk scheduling output: $checkDiskOutput" "cyan"
			Write_Log "Check Disk scheduled. Restart may be required." "green"
		}
		catch { Write_Log "Check Disk scheduling failed: $_" "red" }
	}
	else { Write_Log "Skipping Check Disk scheduling as per user request." "yellow" }
	
	Write_Log "`r`n==================== Repair_Windows Function Completed ====================" "blue"
}
