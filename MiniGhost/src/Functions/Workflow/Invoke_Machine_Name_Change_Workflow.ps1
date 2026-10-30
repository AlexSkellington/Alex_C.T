# ===================================================================================================
#                         FUNCTION: Invoke_Machine_Name_Change_Workflow
# ---------------------------------------------------------------------------------------------------
# Description:
# Coordinates the machine-name action from the GUI. If the requested name matches the current name,
# the workflow skips Rename-Computer and still performs the full INI + SQL sync.
# ===================================================================================================

function Invoke_Machine_Name_Change_Workflow
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$RequestedMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus
	)
	
	$oldMachineName = $env:COMPUTERNAME
	$machineContext = Get_Machine_Name_Context -MachineName $RequestedMachineName -CurrentMachineName $oldMachineName
	if (-not $machineContext.Success)
	{
		Show_App_Message -Text $machineContext.ErrorMessage -Title "Error" -Icon Error | Out-Null
		Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Invalid machine name." -Details $machineContext.ErrorMessage
		return
	}
	
	$confirmMessage = $null
	$confirmTitle = $null
	if ($machineContext.IsSameMachineName)
	{
		$confirmMessage = "The requested machine name matches the current machine name '$($machineContext.MachineName)'.`n`nRun a full INI + SQL sync anyway?"
		$confirmTitle = "Confirm Machine Sync"
	}
	else
	{
		$confirmMessage = "Are you sure you want to change the machine name to '$($machineContext.MachineName)' and run the full INI + SQL sync?"
		$confirmTitle = "Confirm Machine Name Change"
	}
	
	$result = Show_App_Message -Text $confirmMessage -Title $confirmTitle -Buttons YesNo -Icon Question
	if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
	
	$currentStoreNumber = Get_Store_Number_From_INI
	if (-not $currentStoreNumber)
	{
		Show_App_Message -Text "Store number not found in startup.ini." -Title "Error" -Icon Error | Out-Null
		Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Store number not found in startup.ini." -Details ""
		return
	}
	
	$startupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	$renamePerformed = $false
	$script:newMachineName = $machineContext.MachineName
	
	try
	{
		if (-not $machineContext.IsSameMachineName)
		{
			Rename-Computer -NewName $machineContext.MachineName -Force -ErrorAction Stop
			$renamePerformed = $true
			
			Set_Label_Text -Label $script:machineNameLabel -Text "The machine name will change from: $env:COMPUTERNAME to $script:newMachineName" -RefreshParent -ProcessEvents
			Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Successful" -Message "Machine name changed successfully." -Details "Machine name changed to '$script:newMachineName'."
			
			Remove_Old_XZ_Folders -MachineName $script:newMachineName -StoreNumber $currentStoreNumber
		}
		
		$syncResult = Invoke_Machine_Name_Sync -StoreNumber $currentStoreNumber -MachineName $machineContext.MachineName -OldMachineName $oldMachineName -MachineContext $machineContext -StartupIniPath $startupIniPath -OperationStatus $OperationStatus
		$syncFailureLogText = $null
		if (-not $syncResult.Success)
		{
			$syncFailureLogParts = @(
				"Machine Name:`r`n$($machineContext.MachineName)",
				"Store Number:`r`n$currentStoreNumber"
			)
			
			if (-not $syncResult.IniSyncSuccess -or -not $syncResult.StartupIniSuccess)
			{
				$iniSectionParts = @()
				
				if ($syncResult.IniSyncError)
				{
					$iniSectionParts += "Update_INIs Error:`r`n$($syncResult.IniSyncError)"
				}
				
				if ($syncResult.StartupIniError)
				{
					$iniSectionParts += "Startup.ini Error:`r`n$($syncResult.StartupIniError)"
				}
				
				if ($syncResult.StartupIniDetails)
				{
					$iniSectionParts += "Additional Details:`r`n$($syncResult.StartupIniDetails)"
				}
				elseif ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate") -and $OperationStatus["StartupIniUpdate"].Details)
				{
					$iniSectionParts += "Operation Status Details:`r`n$($OperationStatus["StartupIniUpdate"].Details)"
				}
				
				if ($iniSectionParts.Count -gt 0)
				{
					$syncFailureLogParts += "INI Sync:`r`n$($iniSectionParts -join "`r`n`r`n")"
				}
			}
			
			if (-not $syncResult.SqlSyncSuccess)
			{
				$sqlSectionParts = @()
				
				if ($syncResult.SqlSyncError)
				{
					$sqlSectionParts += "SQL Sync Error:`r`n$($syncResult.SqlSyncError)"
				}
				
				if ($syncResult.SqlResult)
				{
					if ($syncResult.SqlResult.ContainsKey('DatabaseContext') -and $syncResult.SqlResult.DatabaseContext)
					{
						$sqlSectionParts += "Database Context:`r`n$($syncResult.SqlResult.DatabaseContext)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('DBNAME') -and $syncResult.SqlResult.DBNAME)
					{
						$sqlSectionParts += "DBNAME:`r`n$($syncResult.SqlResult.DBNAME)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('ConfiguredDBSERVER') -and $syncResult.SqlResult.ConfiguredDBSERVER)
					{
						$sqlSectionParts += "Configured DBSERVER:`r`n$($syncResult.SqlResult.ConfiguredDBSERVER)"
					}
					elseif ($syncResult.ConfiguredDbServer)
					{
						$sqlSectionParts += "Configured DBSERVER:`r`n$($syncResult.ConfiguredDbServer)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('RuntimeDBSERVER') -and $syncResult.SqlResult.RuntimeDBSERVER)
					{
						$sqlSectionParts += "Runtime SQL Target:`r`n$($syncResult.SqlResult.RuntimeDBSERVER)"
					}
					elseif ($syncResult.RuntimeDbServer)
					{
						$sqlSectionParts += "Runtime SQL Target:`r`n$($syncResult.RuntimeDbServer)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('FailedCommands') -and $syncResult.SqlResult.FailedCommands)
					{
						$sqlSectionParts += "Failed SQL Commands:`r`n$($syncResult.SqlResult.FailedCommands -join "`r`n`r`n")"
					}
				}
				elseif ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate") -and $OperationStatus["SQLDatabaseUpdate"].Details)
				{
					$sqlSectionParts += "Operation Status Details:`r`n$($OperationStatus["SQLDatabaseUpdate"].Details)"
				}
				
				if ($sqlSectionParts.Count -gt 0)
				{
					$syncFailureLogParts += "SQL Sync:`r`n$($sqlSectionParts -join "`r`n`r`n")"
				}
			}
			
			$syncFailureLogText = ($syncFailureLogParts -join "`r`n`r`n").Trim()
		}
		
		if ($machineContext.IsSameMachineName)
		{
			if ($syncResult.Success)
			{
				Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Successful" -Message "Machine settings synced successfully." -Details "Requested machine name already matched '$script:newMachineName'; full INI + SQL sync was run without renaming."
				Set_Label_Text -Label $script:machineNameLabel -Text "Current Machine Name: $script:newMachineName (sync refreshed)" -RefreshParent
				Show_App_Message -Text "Full INI + SQL sync completed for '$script:newMachineName'. No reboot is required because the machine name did not change." -Title "Sync Complete" -Icon Information | Out-Null
			}
			else
			{
				Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Machine sync completed with errors." -Details "Requested machine name already matched '$script:newMachineName', but one or more INI/SQL updates failed."
				if ($syncFailureLogText)
				{
					[void](Show_Action_Prompt_With_Log -Text "Same-name sync ran, but one or more INI/SQL updates failed.`r`n`r`nUse View Log to see the failure reasons, or use Summary for the full operation status." -Title "Sync Completed With Errors" -PrimaryButtonText "OK" -LogText $syncFailureLogText -LogTitle "Sync Failure Log")
				}
				else
				{
					Show_App_Message -Text "Same-name sync ran, but one or more INI/SQL updates failed. Review the summary for details." -Title "Sync Completed With Errors" -Icon Warning | Out-Null
				}
			}
		}
		elseif ($renamePerformed)
		{
			$rebootMessage = if ($syncResult.Success)
			{
				"Machine name changed successfully to '$script:newMachineName'. The system will need to reboot for changes to take effect. Do you want to reboot now?"
			}
			else
			{
				"Machine name changed to '$script:newMachineName', but one or more INI/SQL updates failed. A reboot is still required for the rename. Do you want to reboot now?"
			}
			
			$rebootIcon = if ($syncResult.Success)
			{
				[System.Windows.Forms.MessageBoxIcon]::Information
			}
			else
			{
				[System.Windows.Forms.MessageBoxIcon]::Warning
			}
			
			$rebootResult = $null
			if (-not $syncResult.Success -and $syncFailureLogText)
			{
				$rebootResult = Show_Action_Prompt_With_Log -Text ($rebootMessage + "`r`n`r`nUse View Log to see the failure reasons.") -Title "Machine Rename Complete" -PrimaryButtonText "Reboot Now" -SecondaryButtonText "Later" -LogText $syncFailureLogText -LogTitle "Sync Failure Log"
			}
			else
			{
				$rebootResult = Show_App_Message -Text $rebootMessage -Title "Machine Rename Complete" -Buttons YesNo -Icon $rebootIcon.ToString()
			}
			
			if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes -or $rebootResult -eq "Primary")
			{
				Restart-Computer -Force
			}
		}
	}
	catch
	{
		$errorMessage = $_.Exception.Message
		Show_App_Message -Text "Error changing machine name: $errorMessage" -Title "Error" -Icon Error | Out-Null
		Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Error changing machine name." -Details "Error: $errorMessage"
	}
}
