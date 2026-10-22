# ===================================================================================================
#                              FUNCTION: Delete_DBS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes specific file types (.txt and .dwr) from the Dbs folders of selected lanes in a
#   store, with a prompt to also include .sus files. Lane machine paths come from the
#   Retrieve_Nodes results; the actual deletion is done by Delete_Files per lane.
#
# Parameters:
#   - StoreNumber  [string]: store whose lanes are processed (required)
#
# Variables:
#   - $script:FunctionResults: LaneMachineNames and LaneNumToMachineName (input)
#
# Notes:
#   - Retrieve_Nodes must have been run first.
# ===================================================================================================

function Delete_DBS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Delete_DBS Function ====================`r`n" "blue"
	
	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneMachineNames') -or
		-not $script:FunctionResults.ContainsKey('LaneNumToMachineName'))
	{
		Write_Log "Lane information is not available. Please run Retrieve_Nodes first." "Red"
		return
	}
	
	# Retrieve lane information
	$LaneMachineNames = $script:FunctionResults['LaneMachineNames']
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	
	if ($LaneMachineNames.Count -eq 0)
	{
		Write_Log "No lanes found for Store Number: $StoreNumber." "Yellow"
		return
	}
	
	# Prompt user to include .sus files in deletion
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Delete_DBS Confirmation"
	$form.Size = New-Object System.Drawing.Size(400, 200)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Do you want to include .sus files in the deletion?"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(20, 30)
	$form.Controls.Add($label)
	
	$checkboxSus = New-Object System.Windows.Forms.CheckBox
	$checkboxSus.Text = "Include .sus files"
	$checkboxSus.AutoSize = $true
	$checkboxSus.Location = New-Object System.Drawing.Point(20, 60)
	$form.Controls.Add($checkboxSus)
	
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(100, 120)
	$buttonOK.Size = New-Object System.Drawing.Size(80, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(200, 120)
	$buttonCancel.Size = New-Object System.Drawing.Size(80, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($buttonCancel)
	
	$form.AcceptButton = $buttonOK
	$form.CancelButton = $buttonCancel
	
	$dialogResult = $form.ShowDialog()
	
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "User canceled the deletion process." "Yellow"
		return
	}
	
	$includeSus = $checkboxSus.Checked
	
	# Define file types to delete
	$fileExtensions = @("*.txt", "*.dwr")
	if ($includeSus)
	{
		$fileExtensions += "*.sus"
	}
	
	Write_Log "Starting deletion of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "Green"
	
	# Show the selection dialog to choose lanes
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	
	if (-not $selection)
	{
		# User canceled the dialog
		Write_Log "User canceled the lane selection." "Yellow"
		return
	}
	
	# Accept all supported forms (lane number, machine name, POSxxx, etc)
	$selectedLanes = $selection.Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($LaneMachineNames.ContainsKey($_)) { $_ }
		elseif ($LaneNumToMachineName.ContainsKey($_)) { $_ }
		else { $_ }
	}
	# Make sure only valid 3-digit lane numbers that are in the mapping
	$selectedLanes = $selectedLanes | Where-Object { $LaneNumToMachineName.ContainsKey($_) }
	
	if (-not $selectedLanes -or $selectedLanes.Count -eq 0)
	{
		Write_Log "No valid lanes selected for processing." "Yellow"
		return
	}
	
	# Initialize counters
	$totalDeleted = 0
	$totalFailed = 0
	
	foreach ($lane in $selectedLanes)
	{
		if ($LaneNumToMachineName.ContainsKey($lane))
		{
			$machineName = $LaneNumToMachineName[$lane]
			
			if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
			{
				Write_Log "Lane #{$lane}: Machine name is invalid or unknown. Skipping deletion." "Yellow"
				continue
			}
			
			# Construct the target path (modify this path as per your environment)
			$targetPath = "\\$machineName\Storeman\Office\DBS\"
			
			if (-not (Test-Path -Path $targetPath))
			{
				Write_Log "Lane #${lane}: Target path '$targetPath' does not exist. Skipping." "Yellow"
				continue
			}
			
			Write_Log "Processing Lane #$lane at '$targetPath', please wait..." "Blue"
			
			try
			{
				# Delete_Files function is now expected to return an integer count
				$deletionCount = Delete_Files -Path $targetPath -SpecifiedFiles $fileExtensions -Exclusions @() -AsJob:$false
				
				if ($deletionCount -is [int])
				{
					$totalDeleted += $deletionCount
				}
				else
				{
					Write_Log "Lane #${lane}: Unexpected response from Delete_Files." "Red"
					$totalFailed++
				}
			}
			catch
			{
				Write_Log "Lane #${lane}: An error occurred while deleting files. Error: $_" "Red"
				$totalFailed++
			}
		}
		else
		{
			Write_Log "Lane #${lane}: Machine information not found. Skipping." "Yellow"
			continue
		}
	}
	
	# Summary of deletion results
	Write_Log "Deletion Summary for Store Number: $StoreNumber - Total Files Deleted: $totalDeleted, Total Failures: $totalFailed." "Blue"
	Write_Log "`r`n==================== Delete_DBS Function Completed ====================" "blue"
}
