############################################################################
# 6) Registry Cleanup Button
############################################################################
$registryCleanupButton = New-Object System.Windows.Forms.Button
$registryCleanupButton.Text = "Registry Cleanup"
$registryCleanupButton.Location = New-Object System.Drawing.Point(330, 160)
$registryCleanupButton.Size = New-Object System.Drawing.Size(150, 35)
$registryCleanupButton.Add_Click({
		$result = [System.Windows.Forms.MessageBox]::Show("Do you want to delete all registry values starting with 'GT'?", "Registry Cleanup", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			$gtRegistryCleanupResult = Remove_GT_Registry_Values
			
			$registryStatus = $gtRegistryCleanupResult.Status
			
			if ($registryStatus -eq 'Successful')
			{
				$operationStatus["RegistryCleanup"].Status = "Successful"
				$operationStatus["RegistryCleanup"].Message = "GT registry values removed successfully."
				$operationStatus["RegistryCleanup"].Details = "$($gtRegistryCleanupResult.DeletedCount) 'GT' registry keys were deleted."
				[System.Windows.Forms.MessageBox]::Show("GT registry values removed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
			elseif ($registryStatus -eq 'Skipped')
			{
				$operationStatus["RegistryCleanup"].Status = "Skipped"
				$operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
				$operationStatus["RegistryCleanup"].Details = ""
			}
			else
			{
				$operationStatus["RegistryCleanup"].Status = "Failed"
				$operationStatus["RegistryCleanup"].Message = "Failed to remove GT registry values."
				$operationStatus["RegistryCleanup"].Details = $gtRegistryCleanupResult.Message
				[System.Windows.Forms.MessageBox]::Show("Failed to remove GT registry values.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		}
		else
		{
			$operationStatus["RegistryCleanup"].Status = "Skipped"
			$operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
		}
	})
