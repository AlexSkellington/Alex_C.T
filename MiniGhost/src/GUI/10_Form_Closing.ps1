############################################################################
# Handle form closing event (X button)
############################################################################
$form.add_FormClosing({
		# Confirmation message box to confirm exit
		$confirmResult = [System.Windows.Forms.MessageBox]::Show(
			"Are you sure you want to exit?",
			"Confirm Exit",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Question
		)
		
		# If the user clicks No, cancel the form close action
		if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes)
		{
			$_.Cancel = $true
		}
		else
		{
			# Proceed with form closing and perform actions
			# Write-Log "Form is closing. Performing cleanup." "green"
			
			# Clean Temp Folder
			[void](Remove_MiniGhost_Temp_Copies -TempPath $TempDir)
		}
	})
