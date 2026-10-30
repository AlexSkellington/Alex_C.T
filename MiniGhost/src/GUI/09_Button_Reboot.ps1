############################################################################
# 11) Reboot Button
############################################################################
$rebootButton = New-Object System.Windows.Forms.Button
$rebootButton.Text = "Reboot System"
$rebootButton.Location = New-Object System.Drawing.Point(330, 240)
$rebootButton.Size = New-Object System.Drawing.Size(150, 35)
$rebootButton.Add_Click({
		$rebootResult = [System.Windows.Forms.MessageBox]::Show("Do you want to reboot now?", "Reboot", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Clean the temp copies first: nothing after Restart-Computer -Force is guaranteed to run
			[void](Remove_MiniGhost_Temp_Copies -TempPath $TempDir)
			Restart-Computer -Force
		}
	})
