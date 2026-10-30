# ===================================================================================================
#									SECTION: GUI Buttons
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and configures buttons on the GUI form for various operations.
# ===================================================================================================

############################################################################
# 1) Change Machine Name Button (UPDATED: supports same-name full sync workflow)
############################################################################
$changeMachineNameButton = New-Object System.Windows.Forms.Button
$changeMachineNameButton.Text = "Change Machine Name"
$changeMachineNameButton.Location = New-Object System.Drawing.Point(10, 120)
$changeMachineNameButton.Size = New-Object System.Drawing.Size(150, 35)
$changeMachineNameButton.Add_Click({
		if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
		{
			[System.Windows.Forms.MessageBox]::Show(
				"This script must be run as an administrator.",
				"Insufficient Privileges",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			return
		}
		
		$newMachineNameInput = Get_NEW_Machine_Name
		if ($newMachineNameInput -eq $null)
		{
			Set_Operation_Status -OperationStatus $operationStatus -Key "MachineNameChange" -Status "Cancelled" -Message "Machine name change was cancelled by the user." -Details ""
			Show_App_Message -Text "Machine name change was cancelled." -Title "Cancelled" -Icon Information | Out-Null
			return
		}
		
		Invoke_Machine_Name_Change_Workflow -RequestedMachineName $newMachineNameInput -OperationStatus $operationStatus
	})
