# ===================================================================================================
#                              FUNCTION: Refresh_PIN_Pad_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Refreshes the EMV PIN pad configuration files (PreferredAIDs.ini, EMVCAPKey.ini,
#   EMVAID.ini) on selected lanes: each target file is touched (LastWriteTime updated) so
#   the lane sees it as modified and reloads it.
#
# Parameters:
#   - StoreNumber  [string]: store to process, 3 digits (required)
#
# Variables:
#   - $script:FunctionResults['LaneNumToMachineName']: lane map (input)
#   - $OfficePath: Storeman Office root (XF lane folders)
#
# Notes:
#   - Retrieve_Nodes must have been run first.
# ===================================================================================================

function Refresh_PIN_Pad_Files
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	Write_Log "`r`n==================== Starting Refresh_PIN_Pad_Files Function ====================`r`n" "blue"

	# ============================= Validate Environment =============================
	# Validate OfficePath base directory
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}

	# Ensure required functions are loaded
	foreach ($func in @('Show_Node_Selection_Form', 'Write_Log', 'Retrieve_Nodes'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}

	# Ensure lane mapping is available from Retrieve_Nodes
	if (-not ($script:FunctionResults.ContainsKey('LaneNumToMachineName')))
	{
		Write_Log "No lane information found. Please ensure Retrieve_Nodes has been executed." "red"
		return
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']

	# ============================= Lane Selection UI =============================
	# Use the node selector form to allow the user to choose which lanes to refresh
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($null -eq $selection)
	{
		Write_Log "Operation canceled by user." "yellow"
		return
	}
	$Lanes = $selection.Lanes
	if (-not $Lanes -or $Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for processing." "yellow"
		return
	}

	# ============================= File Types =============================
	# List of PIN pad config files that should be "touched" (refreshed)
	$fileExtensions = @("PreferredAIDs.ini", "EMVCAPKey.ini", "EMVAID.ini")
	Write_Log "Starting refresh of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "green"

	# ============================= Main Processing Loop =============================
	$totalRefreshed = 0
	$totalFailed = 0
	$refreshedLanes = New-Object System.Collections.Generic.List[string]

	foreach ($laneObj in $Lanes)
	{
		# Accept both plain lane number string or PSCustomObject with .LaneNumber property
		$laneNum = if ($laneObj -is [pscustomobject] -and $laneObj.LaneNumber) { $laneObj.LaneNumber }
		else { $laneObj }
		$laneNum = $laneNum.PadLeft(3, '0')

		# Ensure this lane exists in LaneNumToMachineName
		if ($LaneNumToMachineName.ContainsKey($laneNum))
		{
			$machineName = $LaneNumToMachineName[$laneNum]

			if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
			{
				Write_Log "Lane #${laneNum}: Machine name is invalid or unknown. Skipping refresh." "yellow"
				continue
			}

			# Build UNC path to lane's EMV config directory
			$targetPath = "\\$machineName\Storeman\XchDev\EMVConfig\"

			if (-not (Test-Path -Path $targetPath))
			{
				Write_Log "Lane #${laneNum}: Target path '$targetPath' does not exist. Skipping." "yellow"
				continue
			}

			Write_Log "Processing Lane #$laneNum ($machineName) at '$targetPath'." "blue"

			foreach ($file in $fileExtensions)
			{
				$filePath = Join-Path -Path $targetPath -ChildPath $file
				if (Test-Path -Path $filePath)
				{
					try
					{
						# "Touch" the file (set LastWriteTime to current time)
						(Get-Item -Path $filePath).LastWriteTime = Get-Date
						Write_Log "Lane #${laneNum}: Refreshed file '$filePath'." "green"
						$totalRefreshed++
						if (-not $refreshedLanes.Contains($laneNum)) { $refreshedLanes.Add($laneNum) }
					}
					catch
					{
						Write_Log "Lane #${laneNum}: Failed to refresh file '$filePath'. Error: $_" "red"
						$totalFailed++
					}
				}
				else
				{
					Write_Log "Lane #${laneNum}: File '$filePath' does not exist. Skipping." "yellow"
				}
			}
		}
		else
		{
			Write_Log "Lane #${laneNum}: Machine information not found in LaneNumToMachineName. Skipping." "yellow"
			continue
		}
	}
	Write_Log "Refresh Summary for Store Number: $StoreNumber - Total Files Refreshed: $totalRefreshed, Total Failures: $totalFailed." "green"
	# ==== Prompt to Restart All Programs on successful lanes ====
	if ($totalRefreshed -gt 0)
	{
		# ($refreshedLanes is collected in the loop above - lanes where at least one file was touched)
		if ($refreshedLanes.Count -gt 0)
		{
			Add-Type -AssemblyName System.Windows.Forms
			$laneListStr = ($refreshedLanes | Sort-Object) -join ", "
			$msg = "Do you want to send 'Restart All Programs' to the following lanes?`nStore $StoreNumber Lanes: $laneListStr"
			$result = [System.Windows.Forms.MessageBox]::Show($msg, "Restart All Programs?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
			if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User chose to send 'Restart All Programs' to lanes: $laneListStr" "cyan"
				Send_Restart_All_Programs -StoreNumber $StoreNumber -LaneNumbers $refreshedLanes
			}
			else
			{
				Write_Log "User cancelled 'Restart All Programs' action." "yellow"
			}
		}
	}
	Write_Log "`r`n==================== Refresh_PIN_Pad_Files Function Completed ====================" "blue"
}
