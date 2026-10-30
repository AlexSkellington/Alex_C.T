# ===================================================================================================
#                            FUNCTION: Invoke_Machine_Name_Sync
# ---------------------------------------------------------------------------------------------------
# Description:
# Runs the best-effort INI + Startup.ini + SQL sync for a machine name target. This is used after an
# actual rename and also when the requested name already matches the current machine name.
# ===================================================================================================

function Invoke_Machine_Name_Sync
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName = $env:COMPUTERNAME,
		[Parameter(Mandatory = $false)]
		[psobject]$MachineContext,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus
	)
	
	if (-not $MachineContext -or -not $MachineContext.Success)
	{
		$MachineContext = Get_Machine_Name_Context -MachineName $MachineName -CurrentMachineName $OldMachineName
	}
	
	if (-not $MachineContext.Success)
	{
		if ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate"))
		{
			$OperationStatus["StartupIniUpdate"].Status = "Failed"
			$OperationStatus["StartupIniUpdate"].Message = "INI sync could not start."
			$OperationStatus["StartupIniUpdate"].Details = $MachineContext.ErrorMessage
		}
		if ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate"))
		{
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL sync could not start."
			$OperationStatus["SQLDatabaseUpdate"].Details = $MachineContext.ErrorMessage
		}
		
		return [PSCustomObject]@{
			Success             = $false
			MachineContext      = $MachineContext
			IniSyncSuccess      = $false
			StartupIniSuccess   = $false
			SqlSyncSuccess      = $false
			StartupIniPath      = $null
			IniSyncError        = $MachineContext.ErrorMessage
			StartupIniError     = $null
			SqlSyncError        = $MachineContext.ErrorMessage
			StartupIniDetails   = $null
			SqlResult           = $null
		}
	}
	
	$resolvedStartupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	$null = Ensure_Database_Connection_Context -StartupIniPath $resolvedStartupIniPath
	
	$configuredDbServer = $null
	if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConfiguredDBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConfiguredDBSERVER']))
	{
		$configuredDbServer = [string]$script:FunctionResults['ConfiguredDBSERVER']
	}
	elseif ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBSERVER']))
	{
		$configuredDbServer = [string]$script:FunctionResults['DBSERVER']
	}
	
	$runtimeDbServer = $null
	if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('RuntimeDBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['RuntimeDBSERVER']))
	{
		$runtimeDbServer = [string]$script:FunctionResults['RuntimeDBSERVER']
	}
	
	$iniSyncSuccess = $false
	$iniSyncError = $null
	try
	{
		$iniSyncSuccess = [bool](Update_INIs -newStoreNumber $StoreNumber -MachineName $MachineContext.MachineName -OldStoreNumber $StoreNumber -OldMachineName $OldMachineName -CreateSmsHttpsIfMissing -StartupIniPath $resolvedStartupIniPath)
		if (-not $iniSyncSuccess)
		{
			$iniSyncError = "Update_INIs reported a failure."
		}
	}
	catch
	{
		$iniSyncError = $_.Exception.Message
	}
	
	$startupIniResult = Update_Startup_Ini_For_Machine -MachineContext $MachineContext -StartupIniPath $resolvedStartupIniPath
	$startupIniSuccess = $startupIniResult.Success
	$startupIniError = $startupIniResult.ErrorMessage
	
	if ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate"))
	{
		if ($iniSyncSuccess -and $startupIniSuccess)
		{
			$OperationStatus["StartupIniUpdate"].Status = "Successful"
			$OperationStatus["StartupIniUpdate"].Message = "INI files updated successfully."
			$OperationStatus["StartupIniUpdate"].Details = $startupIniResult.Details
		}
		else
		{
			$detailParts = @()
			if ($iniSyncError) { $detailParts += $iniSyncError }
			if ($startupIniError) { $detailParts += $startupIniError }
			if ($startupIniResult.Details) { $detailParts += $startupIniResult.Details }
			
			$OperationStatus["StartupIniUpdate"].Status = "Failed"
			$OperationStatus["StartupIniUpdate"].Message = "INI sync completed with errors."
			$OperationStatus["StartupIniUpdate"].Details = ($detailParts -join " ")
		}
	}
	
	$sqlUpdateResult = $null
	$sqlSyncSuccess = $false
	$sqlSyncError = $null
	try
	{
		$sqlUpdateResult = Update_SQL_Tables_For_Machine_Name_Change `
															 -storeNumber $StoreNumber `
															 -machineName $MachineContext.MachineName `
															 -machineNumber $MachineContext.MachineNumber `
															 -OldMachineName $OldMachineName
		
		if ($sqlUpdateResult -and $sqlUpdateResult.Success)
		{
			$sqlSyncSuccess = $true
		}
		else
		{
			$sqlSyncError = "SQL update result indicated failure."
		}
	}
	catch
	{
		$sqlSyncError = $_.Exception.Message
	}
	
	if ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate"))
	{
		if ($sqlSyncSuccess)
		{
			$detailParts = @("Protected SQL sync completed.")
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RunTabUpdated'))
			{
				$detailParts += "RUN updated: $($sqlUpdateResult.RunTabUpdated)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('LocalTerminalDB'))
			{
				$detailParts += "Local terminal DB: $($sqlUpdateResult.LocalTerminalDB)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('ConfiguredDBSERVER') -and $sqlUpdateResult.ConfiguredDBSERVER)
			{
				$detailParts += "Configured DBSERVER: $($sqlUpdateResult.ConfiguredDBSERVER)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RuntimeDBSERVER') -and $sqlUpdateResult.RuntimeDBSERVER)
			{
				$detailParts += "Runtime SQL target: $($sqlUpdateResult.RuntimeDBSERVER)."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL tables updated successfully after machine sync."
			$OperationStatus["SQLDatabaseUpdate"].Details = ($detailParts -join " ")
		}
		else
		{
			$detailParts = @()
			if ($sqlSyncError) { $detailParts += $sqlSyncError }
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('ConfiguredDBSERVER') -and $sqlUpdateResult.ConfiguredDBSERVER)
			{
				$detailParts += "Configured DBSERVER: $($sqlUpdateResult.ConfiguredDBSERVER)."
			}
			elseif ($configuredDbServer)
			{
				$detailParts += "Configured DBSERVER: $configuredDbServer."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RuntimeDBSERVER') -and $sqlUpdateResult.RuntimeDBSERVER)
			{
				$detailParts += "Runtime SQL target: $($sqlUpdateResult.RuntimeDBSERVER)."
			}
			elseif ($runtimeDbServer)
			{
				$detailParts += "Runtime SQL target: $runtimeDbServer."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Failed to update SQL tables after machine sync."
			$OperationStatus["SQLDatabaseUpdate"].Details = ($detailParts -join " ")
		}
	}
	
	return [PSCustomObject]@{
		Success             = ($iniSyncSuccess -and $startupIniSuccess -and $sqlSyncSuccess)
		MachineContext      = $MachineContext
		IniSyncSuccess      = $iniSyncSuccess
		StartupIniSuccess   = $startupIniSuccess
		SqlSyncSuccess      = $sqlSyncSuccess
		StartupIniPath      = $resolvedStartupIniPath
		IniSyncError        = $iniSyncError
		StartupIniError     = $startupIniError
		SqlSyncError        = $sqlSyncError
		StartupIniDetails   = $startupIniResult.Details
		SqlResult           = $sqlUpdateResult
		ConfiguredDbServer  = $configuredDbServer
		RuntimeDbServer     = $runtimeDbServer
	}
}
