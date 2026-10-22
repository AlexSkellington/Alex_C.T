# ===================================================================================================
#                              FUNCTION: Drawer_Control
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys a drawer control command to the selected lanes: pick Enable/Disable in a dialog,
#   select lanes, then for each lane write DrawerControl.sqi (ANSI, CRLF) with the chosen
#   state into the lane XF folder; the SQI activates the change itself (exe_activate_accept_sys).
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName (input)
#   - $OfficePath: Storeman Office root (XF lane folders)
#
# Writes:
#   - DrawerControl.sqi in each selected lane XF folder
# ===================================================================================================

function Drawer_Control
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	Write_Log "`r`n==================== Starting Drawer_Control ====================`r`n" "blue"

	# --------------------------------------------------
	# STEP 1: Prompt for Drawer State using Enable/Disable radio buttons
	# --------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$stateForm = New-Object System.Windows.Forms.Form
	$stateForm.Text = "Select Drawer State"
	$stateForm.Size = New-Object System.Drawing.Size(400, 200)
	$stateForm.StartPosition = "CenterScreen"

	$stateLabel = New-Object System.Windows.Forms.Label
	$stateLabel.Text = "Select Drawer State:"
	$stateLabel.Location = New-Object System.Drawing.Point(10, 20)
	$stateLabel.AutoSize = $true
	$stateForm.Controls.Add($stateLabel)

	# Radio button for Enable (value = 1)
	$radioEnable = New-Object System.Windows.Forms.RadioButton
	$radioEnable.Text = "Enable"
	$radioEnable.Location = New-Object System.Drawing.Point(10, 50)
	$radioEnable.AutoSize = $true
	$radioEnable.Checked = $true # default selection
	$stateForm.Controls.Add($radioEnable)

	# Radio button for Disable (value = 0)
	$radioDisable = New-Object System.Windows.Forms.RadioButton
	$radioDisable.Text = "Disable"
	$radioDisable.Location = New-Object System.Drawing.Point(10, 80)
	$radioDisable.AutoSize = $true
	$stateForm.Controls.Add($radioDisable)

	# OK Button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(80, 120)
	$okButton.Add_Click({
			if ($radioEnable.Checked)
			{
				$stateForm.Tag = "1"
			}
			elseif ($radioDisable.Checked)
			{
				$stateForm.Tag = "0"
			}
			$stateForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$stateForm.Close()
		})
	$stateForm.Controls.Add($okButton)

	# Cancel Button
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(180, 120)
	$cancelButton.Add_Click({
			$stateForm.Tag = "Cancelled"
			$stateForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$stateForm.Close()
		})
	$stateForm.Controls.Add($cancelButton)

	$stateForm.AcceptButton = $okButton
	$stateForm.CancelButton = $cancelButton

	$resultState = $stateForm.ShowDialog()
	if ($stateForm.Tag -eq "Cancelled" -or $resultState -eq [System.Windows.Forms.DialogResult]::Cancel)
	{
		Write_Log "User cancelled the operation at drawer state selection." "yellow"
		Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	$DrawerState = $stateForm.Tag
	Write_Log "Drawer state selected: $DrawerState" "green"

	# --------------------------------------------------
	# STEP 2: Use Show_Lane_Selection_Form to select lanes (Store mode)
	# --------------------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected or selection cancelled." "yellow"
		Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	# Normalize to 3-digit lane numbers (always as strings)
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$lanesToProcess = $selection.Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($LaneNumToMachineName.ContainsKey($_)) { $_ }
		else { $_ }
	}
	$lanesToProcess = $lanesToProcess | Where-Object { $LaneNumToMachineName.ContainsKey($_) }
	if (-not $lanesToProcess -or $lanesToProcess.Count -eq 0)
	{
		Write_Log "No valid lanes selected for processing." "yellow"
		return
	}

	# --------------------------------------------------
	# STEP 3: For each selected lane, deploy the SQI command file in ANSI (PC) format and send restart command.
	# --------------------------------------------------
	foreach ($lane in $lanesToProcess)
	{
		# Construct the lane directory path (assumes folder naming: XF<StoreNumber><Lane>)
		$LaneDirectory = "$OfficePath\XF${StoreNumber}${lane}"
		if (-not (Test-Path $LaneDirectory))
		{
			Write_Log "Lane directory $LaneDirectory not found. Skipping lane $lane." "yellow"
			continue
		}

		# Define the SQI content with the chosen drawer state
		$SQIContent = @"
CREATE VIEW Fct_Load AS SELECT F1063,F1000,F81,F85,F96,F97,F98,F99,F100,F101,F102,F125,F172,F239,F240,F241,F242,F1042,F1043,F1044,F1045,F1046,F1047,F1050,F1051,F1052,F1053,F1054,F1055,F1064,F1081,F1082,F1083,F1084,F1085,F1086,F1088,F1089,F1090,F1091,F1092,F1147,F1817,F1818,F1895,F1897,F1965,F1966 FROM FCT_TAB;

INSERT INTO Fct_Load VALUES
(10010,'PAL',,,,,,,,,,,,,,,,1,1,1,1,1,9,'F3','TRS','Log',1,1,1,'Login operator','DRAWEROPEN=$DrawerState',,,,,,,,,,,,,,'',,,),

@UPDATE_BATCH(JOB=ADDRPL,TAR=FCT_TAB,
KEY=F1063=:F1063 AND F1000=:F1000,
SRC=SELECT * FROM Fct_Load);

DROP TABLE Fct_Load;

@WIZSET(TARGET=@TER);
@EXEC(SQM=exe_activate_accept_sys);
"@

		# Ensure the SQI content uses CRLF line endings (ANSI PC format)
		$SQIContent = $SQIContent -replace "`r?`n", "`r`n"

		# Define the full path to the SQI file (named "DrawerControl.sqi")
		$SQIFilePath = Join-Path -Path $LaneDirectory -ChildPath "DrawerControl.sqi"

		# Write the SQI file using ASCII encoding (ANSI PC)
		Set-Content -Path $SQIFilePath -Value $SQIContent -Encoding ASCII

		# Remove the Archive attribute (set file attributes to Normal)
		Set-ItemProperty -Path $SQIFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Deployed DrawerControl.sqi command to lane $lane with state '$DrawerState' in directory $LaneDirectory." "green"
	}
	Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
}
