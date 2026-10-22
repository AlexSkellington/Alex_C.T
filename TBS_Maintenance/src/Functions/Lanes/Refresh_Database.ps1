# ===================================================================================================
#                              FUNCTION: Refresh_Database
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys a database refresh command to the selected registers: for each selected lane it
#   writes Refresh_Database.sqi (ANSI, CRLF) containing @WIZSET(TARGET=@TER) and
#   @EXEC(SQM=exe_activate_accept_sys) into the lane XF folder, then sends a restart
#   command to the machine.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName (input)
#   - $OfficePath: Storeman Office root (XF lane folders)
#
# Writes:
#   - Refresh_Database.sqi in each selected lane XF folder
# ===================================================================================================

function Refresh_Database
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	Write_Log "`r`n==================== Starting Refresh_Database ====================`r`n" "blue"

	# --------------------------------------------------
	# STEP 1: Use Show_Node_Selection_Form to select registers/lanes
	# --------------------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No registers selected or selection cancelled." "yellow"
		Write_Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
		return
	}

	# Normalize to 3-digit lane/register numbers (always as strings)
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$registersToProcess = $selection.Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($LaneNumToMachineName.ContainsKey($_)) { $_ }
		else { $_ }
	}
	$registersToProcess = $registersToProcess | Where-Object { $LaneNumToMachineName.ContainsKey($_) }
	if (-not $registersToProcess -or $registersToProcess.Count -eq 0)
	{
		Write_Log "No valid registers selected for processing." "yellow"
		return
	}

	# --------------------------------------------------
	# STEP 2: Define the SQI content to refresh the database
	# --------------------------------------------------
	$SQIContent = @"
@WIZSET(TARGET=@TER);
@EXEC(SQM=exe_activate_accept_sys);
"@
	# Ensure the SQI content uses CRLF line endings (ANSI PC format)
	$SQIContent = $SQIContent -replace "`r?`n", "`r`n"

	# --------------------------------------------------
	# STEP 3: For each selected register, deploy the SQI file
	# --------------------------------------------------
	foreach ($register in $registersToProcess)
	{
		# Construct the register directory path (assumes folder naming: XF<StoreNumber><Register>)
		$RegisterDirectory = "$OfficePath\XF${StoreNumber}${register}"
		if (-not (Test-Path $RegisterDirectory))
		{
			Write_Log "Register directory $RegisterDirectory not found. Skipping register $register." "yellow"
			continue
		}

		# Define the full path to the SQI file (named "Refresh_Database.sqi")
		$SQIFilePath = Join-Path -Path $RegisterDirectory -ChildPath "Refresh_Database.sqi"

		# Write the SQI file in ANSI (PC) format (using ASCII encoding)
		Set-Content -Path $SQIFilePath -Value $SQIContent -Encoding ASCII

		# Remove the Archive attribute (set file attributes to Normal)
		Set-ItemProperty -Path $SQIFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Deployed Refresh_Database.sqi command to register $register in directory $RegisterDirectory." "green"
	}
	Write_Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
}
