############################################################################
# 3) Update Store Number Button (UPDATED: always runs Update_INIs using latest intended MachineName)
############################################################################
$updateStoreNumberButton = New-Object System.Windows.Forms.Button
$updateStoreNumberButton.Text = "Update Store Number"
$updateStoreNumberButton.Location = New-Object System.Drawing.Point(330, 120)
$updateStoreNumberButton.Size = New-Object System.Drawing.Size(150, 35)
$updateStoreNumberButton.Add_Click({
		
		# Get old store number
		$oldStoreNumber = Get_Store_Number_From_INI
		
		if ($oldStoreNumber -eq $null)
		{
			Show_App_Message -Text "Store number not found in startup.ini." -Title "Error" -Icon Error | Out-Null
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Failed" -Message "Store number not found." -Details "startup.ini not found or store number not defined."
			return
		}
		
		$newStoreNumberInput = Get_NEW_Store_Number
		if ($newStoreNumberInput -eq $null) { return }
		$newStoreNumberInput = ([string]$newStoreNumberInput).Trim()
		if ($newStoreNumberInput -notmatch '^\d{3,4}$')
		{
			Show_App_Message -Text "The new store number input was invalid. Please enter exactly 3 or 4 digits." -Title "Invalid Store Number" -Icon Error | Out-Null
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Failed" -Message "Invalid store number input." -Details "Received value: '$newStoreNumberInput'."
			return
		}
		
		$warningResult = Show_App_Message -Text ("You are about to change the store number from '{0}' to '{1}'. Do you want to proceed?" -f $oldStoreNumber, $newStoreNumberInput) -Title "Warning" -Buttons YesNo -Icon Warning
		
		if ($warningResult -ne [System.Windows.Forms.DialogResult]::Yes)
		{
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Cancelled" -Message "Store number change was cancelled by the user." -Details "Old store number remains '$oldStoreNumber'."
			return
		}
		
		# Prefer the intended machine name if one is already pending from the rename workflow.
		$effectiveMachineName = Get_Effective_Machine_Name
		
		# Always run Update_INIs (best-effort) so INIs are consistent even if other steps later fail
		$iniOk = $false
		try
		{
			$iniOk = Update_INIs -newStoreNumber $newStoreNumberInput -MachineName $effectiveMachineName -OldStoreNumber $oldStoreNumber -OldMachineName $env:COMPUTERNAME -CreateSmsHttpsIfMissing
		}
		catch
		{
			$iniOk = $false
		}
		
		if ($iniOk)
		{
			$script:newStoreNumber = $newStoreNumberInput
			Set_Label_Text -Label $storeNumberLabel -Text "$oldStoreNumber -> $script:newStoreNumber"
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Successful" -Message "Store number updated in INIs." -Details "Store number changed to '$script:newStoreNumber'."
			Show_App_Message -Text "Store number successfully changed to '$script:newStoreNumber'." -Title "Store Number Updated" -Icon Information | Out-Null
		}
		else
		{
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Failed" -Message "Update_INIs failed." -Details "Best-effort update encountered an error."
			Show_App_Message -Text "Update_INIs failed (best-effort). Check console/log output." -Title "Error" -Icon Error | Out-Null
		}
		
		# SQL update (keep your existing logic)
		$sqlUpdateResult = Update_SQL_Tables_For_Store_Number_Change -storeNumber $newStoreNumberInput
		
		if ($sqlUpdateResult.Success)
		{
			$operationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$operationStatus["SQLDatabaseUpdate"].Message = "STD_TAB updated successfully after store number change."
			$operationStatus["SQLDatabaseUpdate"].Details = "STD_TAB updated with new store number."
		}
		else
		{
			$operationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$operationStatus["SQLDatabaseUpdate"].Message = "Failed to update STD_TAB after store number change."
			$operationStatus["SQLDatabaseUpdate"].Details = "Failed commands: $($sqlUpdateResult.FailedCommands -join ', ')"
			
			# The INI result was already shown; without this the SQL failure was only visible in Summary
			$sqlFailureLog = (@($sqlUpdateResult.FailedCommands) -join "`r`n`r`n")
			if ([string]::IsNullOrWhiteSpace($sqlFailureLog)) { $sqlFailureLog = "SQL update result indicated failure." }
			[void](Show_Action_Prompt_With_Log -Text "The store number SQL update (STD_TAB / TER_TAB 901) failed.`r`n`r`nUse View Log to see the failed commands, then run Update SQL Database or fix the database connection." -Title "SQL Update Failed" -PrimaryButtonText "OK" -LogText $sqlFailureLog -LogTitle "Failed SQL Commands")
		}
		
		<# Run Update_INIs again at the end so INIs always reflect the latest state
		try
		{
			$null = Update_INIs -newStoreNumber $newStoreNumberInput -MachineName $effectiveMachineName -OldStoreNumber $oldStoreNumber -OldMachineName $env:COMPUTERNAME -CreateSmsHttpsIfMissing
		}
		catch { }#>
	})
