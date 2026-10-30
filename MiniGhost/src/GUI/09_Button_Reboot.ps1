############################################################################
# 11) Summary Button
############################################################################
$rebootButton = New-Object System.Windows.Forms.Button
$rebootButton.Text = "Reboot System"
$rebootButton.Location = New-Object System.Drawing.Point(330, 240)
$rebootButton.Size = New-Object System.Drawing.Size(150, 35)
$rebootButton.Add_Click({
		$rebootResult = [System.Windows.Forms.MessageBox]::Show("Do you want to reboot now?", "Reboot", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			Restart-Computer -Force
			# Clean Temp Folder
			Delete_Files_And_Folders -Path "$TempDir" -SpecifiedFiles "MiniGhost.ps1"
		}
	})
