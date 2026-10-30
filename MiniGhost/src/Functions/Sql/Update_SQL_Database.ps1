# ===================================================================================================
#                               FUNCTION: Update_SQL_Database
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs the newer protected SQL machine/store sync routine used by rename flows so the standalone
#   "Update SQL Database" button follows the same safer TER/STO/LNK/RUN/STD update path.
#
# Parameters:
#   - CurrentMachineName : Current hostname to use if NewMachineName is empty
#   - NewMachineName     : Optional override hostname (ex: after rename)
#   - OperationStatus    : Shared status hashtable (ex: $OperationStatus["SQLDatabaseUpdate"])
# ===================================================================================================

function Update_SQL_Database
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		[string]$CurrentMachineName,
		[Parameter()]
		[string]$NewMachineName,
		[Parameter(Mandatory)]
		[hashtable]$OperationStatus
	)
	
	try
	{
		$storeNumberFromINI = Get_Store_Number_From_INI
		if ($null -eq $storeNumberFromINI)
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Store number not found in startup.ini.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Store number not found in startup.ini."
			$OperationStatus["SQLDatabaseUpdate"].Details = ""
			return
		}
		
		$storeNumber = ($storeNumberFromINI.ToString()).Trim()
		
		$machineName = if (-not [string]::IsNullOrWhiteSpace($NewMachineName)) { $NewMachineName }
		else { $CurrentMachineName }
		
		$machineContext = Get_Machine_Name_Context -MachineName $machineName -CurrentMachineName $CurrentMachineName
		if (-not $machineContext.Success)
		{
			[System.Windows.Forms.MessageBox]::Show(
				$machineContext.ErrorMessage,
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Invalid machine name."
			$OperationStatus["SQLDatabaseUpdate"].Details = $machineContext.ErrorMessage
			return
		}
		
		$machineName = $machineContext.MachineName
		$machineNumber = $machineContext.MachineNumber
		
		$sqlUpdateResult = Update_SQL_Tables_For_Machine_Name_Change `
															 -storeNumber $storeNumber `
															 -machineName $machineName `
															 -machineNumber $machineNumber `
															 -OldMachineName $CurrentMachineName
		
		if ($sqlUpdateResult -and $sqlUpdateResult.Success)
		{
			$details = "Protected SQL sync completed."
			if ($sqlUpdateResult.ContainsKey('RunTabUpdated'))
			{
				$details += " RUN updated: $($sqlUpdateResult.RunTabUpdated)."
			}
			if ($sqlUpdateResult.ContainsKey('LocalTerminalDB'))
			{
				$details += " Local terminal DB: $($sqlUpdateResult.LocalTerminalDB)."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL database updated successfully."
			$OperationStatus["SQLDatabaseUpdate"].Details = $details
			[System.Windows.Forms.MessageBox]::Show(
				"SQL database updated successfully.",
				"Success",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			)
		}
		else
		{
			$failedSqlCommands = @()
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('FailedCommands') -and $sqlUpdateResult.FailedCommands)
			{
				$failedSqlCommands = @($sqlUpdateResult.FailedCommands)
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Failed to update SQL database."
			$OperationStatus["SQLDatabaseUpdate"].Details = if ($failedSqlCommands.Count -gt 0) { "Failed SQL commands are listed below." }
			else { "SQL update result indicated failure." }
			[System.Windows.Forms.MessageBox]::Show(
				"Failed to update SQL database.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			
			if ($failedSqlCommands.Count -gt 0)
			{
				Show_Text_Report_Form -Title "Failed SQL Commands" -Text ($failedSqlCommands -join "`r`n`r`n") -Width 640 -Height 420
			}
		}
	}
	catch
	{
		$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
		$OperationStatus["SQLDatabaseUpdate"].Message = "Unhandled error during SQL update."
		$OperationStatus["SQLDatabaseUpdate"].Details = $_.Exception.Message
		
		[System.Windows.Forms.MessageBox]::Show(
			"Unhandled error:`r`n$($_.Exception.Message)",
			"Error",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Error
		)
	}
}
