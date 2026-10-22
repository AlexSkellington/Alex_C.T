# ===================================================================================================
#                              FUNCTION: Repair_Windows
# ---------------------------------------------------------------------------------------------------
# Description:
#   Server menu "Repair Windows": a checkbox form to run server maintenance now. The work runs
#   detached in the background (Start_Server_Maintenance) - the window stays usable and the run
#   continues if the tool is closed. The first four steps are what the scheduled tasks run
#   (weekly: optimize drives; monthly: DISM, SFC, check disk); the rest are manual only.
#
# Writes:
#   - local system state through the worker; Log\Server_Maintenance.log
# ===================================================================================================

function Repair_Windows
{
	[CmdletBinding()]
	param ()

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$repairForm = New-Object System.Windows.Forms.Form
	$repairForm.Text = 'Repair Windows - Server Maintenance'
	$repairForm.ClientSize = New-Object System.Drawing.Size(460, 420)
	$repairForm.StartPosition = 'CenterScreen'
	$repairForm.FormBorderStyle = 'FixedDialog'
	$repairForm.MaximizeBox = $false
	$repairForm.MinimizeBox = $false
	$repairForm.ShowInTaskbar = $false

	$info = New-Object System.Windows.Forms.Label
	$info.Location = [System.Drawing.Point]::new(16, 12)
	$info.Size = [System.Drawing.Size]::new(430, 62)
	$info.Text = "Runs in the background - you can keep working or close this tool.`r`n" +
	"Scheduled: drive optimize every Sunday $($script:ServerMaintenanceWeeklyTime), system repair the first Sunday of the month $($script:ServerMaintenanceMonthlyTime).`r`n" +
	"Log: $($script:ScriptsLog)Server_Maintenance.log"
	$repairForm.Controls.Add($info)

	# Step name, text, checked by default; $null = section heading
	$items = @(
		@($null, 'Scheduled steps (run now)', $false),
		@('Optimize', 'Optimize drives (SSD: TRIM, hard drive: defrag)', $true),
		@('Dism', 'DISM repair + component cleanup', $true),
		@('Sfc', 'System File Checker (SFC)', $true),
		@('ChkScan', 'Check disk (online scan, no restart)', $true),
		@($null, 'Manual only', $false),
		@('Defender', 'Windows Defender signature update + full scan (hours)', $false),
		@('Cleanup', 'Disk cleanup (C:)', $false),
		@('ChkdskReboot', 'Check disk with repair at next restart (C:, can take hours)', $false)
	)
	$checkboxes = @()
	$y = 80
	foreach ($it in $items)
	{
		if (-not $it[0])
		{
			$h = New-Object System.Windows.Forms.Label
			$h.Text = $it[1]
			$h.Font = New-Object System.Drawing.Font($repairForm.Font, [System.Drawing.FontStyle]::Bold)
			$h.Location = [System.Drawing.Point]::new(16, $y + 4)
			$h.Size = [System.Drawing.Size]::new(420, 20)
			$repairForm.Controls.Add($h)
			$y += 28
			continue
		}
		$cb = New-Object System.Windows.Forms.CheckBox
		$cb.Text = $it[1]
		$cb.Tag = $it[0]
		$cb.Checked = $it[2]
		$cb.Location = [System.Drawing.Point]::new(28, $y)
		$cb.Size = [System.Drawing.Size]::new(420, 24)
		$repairForm.Controls.Add($cb)
		$checkboxes += $cb
		$y += 28
	}

	$runButton = New-Object System.Windows.Forms.Button
	$runButton.Text = 'Run in background'
	$runButton.Location = [System.Drawing.Point]::new(220, $y + 12)
	$runButton.Size = [System.Drawing.Size]::new(130, 30)
	$runButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$repairForm.Controls.Add($runButton)
	$repairForm.AcceptButton = $runButton

	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = 'Cancel'
	$cancelButton.Location = [System.Drawing.Point]::new(360, $y + 12)
	$cancelButton.Size = [System.Drawing.Size]::new(80, 30)
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$repairForm.Controls.Add($cancelButton)
	$repairForm.CancelButton = $cancelButton
	$repairForm.ClientSize = [System.Drawing.Size]::new(460, $y + 56)

	# Run needs at least one step (handlers read the form's controls when they fire - no closure)
	foreach ($cb in $checkboxes)
	{
		$cb.Add_CheckedChanged({
				$f = $this.FindForm()
				$any = @($f.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] -and $_.Checked }).Count -gt 0
				foreach ($b in $f.Controls) { if ($b -is [System.Windows.Forms.Button] -and $b.DialogResult -eq 'OK') { $b.Enabled = $any } }
			})
	}

	$result = $repairForm.ShowDialog()
	$steps = @($checkboxes | Where-Object { $_.Checked } | ForEach-Object { [string]$_.Tag })
	$repairForm.Dispose()
	if ($result -ne [System.Windows.Forms.DialogResult]::OK -or $steps.Count -eq 0)
	{
		Write_Log 'Repair Windows cancelled.' 'yellow'
		return
	}

	$null = Start_Server_Maintenance -Steps $steps
}
