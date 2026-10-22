# ===================================================================================================
#                              FUNCTION: Schedule_Lane_DB_Maintenance
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys the lane DB maintenance script to the selected lanes: writes the Lane SQL script
#   to each lane's Office folder and a scheduler macro in the local XF folder so the repair
#   runs on the chosen repeat interval. Files are written as ANSI (Windows-1252), CRLF, no BOM,
#   and the archive bit is cleared after writing.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $LaneSQLScript: lane maintenance SQL text (must be generated first)
#   - $script:FunctionResults['LaneNumToMachineName']: lane map (input, auto-refreshed)
#
# Writes:
#   - Lane maintenance script + scheduler macro on the selected lanes and in the local XF folder
# ===================================================================================================

function Schedule_Lane_DB_Maintenance
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	Write_Log "`r`n==================== Starting Schedule_Lane_DB_Maintenance ====================`r`n" "blue"

	if (-not $LaneSQLScript)
	{
		Write_Log "Lane SQL script content variable (`\$LaneSQLScript`) is empty or not defined." "red"
		return
	}

	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes)
	{
		Write_Log "No lanes selected. Cancelling operation." "yellow"
		return
	}
	# --- ENSURE NODE MAPPING IS PRESENT ---
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName)
	{
		$null = Retrieve_Nodes -StoreNumber $StoreNumber
		$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
		if (-not $LaneNumToMachineName)
		{
			Write_Log "Failed to retrieve lane-to-machine mapping for store $StoreNumber." "red"
			return
		}
	}
	$ansiEncoding = [System.Text.Encoding]::GetEncoding(1252)
	$LaneSQLScriptContent = $LaneSQLScript -replace "`n", "`r`n"

	# Prompt user for the repeat interval in days
	Add-Type -AssemblyName System.Windows.Forms
	$daysPromptForm = New-Object System.Windows.Forms.Form
	$daysPromptForm.Text = "Lane DB Maintenance - Schedule Interval"
	$daysPromptForm.Width = 350
	$daysPromptForm.Height = 160
	$daysPromptForm.StartPosition = "CenterScreen"

	$label = New-Object System.Windows.Forms.Label
	$label.Text = "How many days between each run (minimum 1):"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(15, 20)
	$daysPromptForm.Controls.Add($label)

	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Location = New-Object System.Drawing.Point(20, 50)
	$textBox.Width = 60
	$textBox.Text = "7"
	$daysPromptForm.Controls.Add($textBox)

	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(90, 90)
	$okButton.Add_Click({ $daysPromptForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })
	$daysPromptForm.Controls.Add($okButton)

	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(170, 90)
	$cancelButton.Add_Click({ $daysPromptForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
	$daysPromptForm.Controls.Add($cancelButton)

	$daysPromptForm.AcceptButton = $okButton
	$daysPromptForm.CancelButton = $cancelButton

	if ($daysPromptForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Operation cancelled by user in interval prompt." "yellow"
		return
	}

	[int]$UserDays = 0
	if ([int]::TryParse($textBox.Text, [ref]$UserDays) -and $UserDays -ge 1)
	{
		$RepeatDays = $UserDays
	}
	else
	{
		Write_Log "Invalid or no interval provided, using 7 days." "yellow"
		$RepeatDays = 7
	}

	foreach ($LaneNumber in $selection.Lanes)
	{
		# ----------- USE MAPPING ----------
		$LaneMachineName = $LaneNumToMachineName[$LaneNumber]
		if (-not $LaneMachineName)
		{
			Write_Log "Could not resolve machine name for lane $LaneNumber. Skipping." "red"
			continue
		}
		$LaneOfficeFolder = "\\$LaneMachineName\storeman\office"
		$DestScriptPath = Join-Path $LaneOfficeFolder "LANE_DB_MAINTENANCE.SQI"
		$LocalXFPath = Join-Path $OfficePath "XF$StoreNumber$LaneNumber"
		$SchedulerMacroPath = Join-Path $LocalXFPath "Add_LaneDBMaintenance_to_RUN_TAB.sqi"

		# Prepare scheduler macro content (unique task number per lane if needed)
		$TaskNumber = 750
		$HostTarget = "{0:D3}" -f [int]$LaneNumber
		$CommandToRun = 'sqi=LANE_DB_MAINTENANCE'
		$ExecTarget = $HostTarget
		$TaskName = 'Lane DB Maintenance'
		$ManualAllowed = 1
		$CatchupMissed = 1
		$WeeklyDays = $RepeatDays
		$Months = 0
		$Minutes = 0
		$LastRanDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
		$NextRunDate = (Get-Date).AddDays($RepeatDays).ToString("yyyy-MM-dd HH:mm:ss.fff")


		$SchedulerMacroContent = @"
 /* First delete the scheduled maintenance if it exists */
 DELETE FROM RUN_TAB WHERE F1103 = '$CommandToRun' AND F1000 = '$HostTarget';

 /* Insert the scheduled weekly maintenance */
 INSERT INTO RUN_TAB (F1102, F1000, F1103, F1104, F1105, F1107, F1108, F1109, F1111, F1114, F1115, F1117)
 VALUES ($TaskNumber, '$HostTarget', '$CommandToRun', '$ExecTarget', '$LastRanDate', '$NextRunDate', $ManualAllowed, '$TaskName', $CatchupMissed, $WeeklyDays, $Months, $Minutes);

 /* Activate the new task right away */
 @EXEC(SQL=ACTIVATE_ACCEPT_SYS);
"@ -replace "`n", "`r`n"

		try
		{
			if (-not (Test-Path $LaneOfficeFolder))
			{
				Write_Log "Remote office folder not found: $LaneOfficeFolder (lane may be offline or not shared)." "red"
				continue
			}
			if (-not (Test-Path $LocalXFPath))
			{
				Write_Log "Local XF folder not found: $LocalXFPath (scheduler macro not dropped for this lane)." "red"
				continue
			}
		}
		catch
		{
			Write_Log "Failed to create required directories: $_" "red"
			continue
		}

		# Write the lane repair script
		try
		{
			[System.IO.File]::WriteAllText($DestScriptPath, $LaneSQLScriptContent, $ansiEncoding)
			Set-ItemProperty -Path $DestScriptPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			Write_Log "Wrote lane DB maintenance script to $DestScriptPath" "green"
		}
		catch
		{
			Write_Log "Failed to write script: $_" "red"
			continue
		}

		# Write the scheduler macro
		try
		{
			[System.IO.File]::WriteAllText($SchedulerMacroPath, $SchedulerMacroContent, $ansiEncoding)
			Set-ItemProperty -Path $SchedulerMacroPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			Write_Log "Scheduler SQL macro created at $SchedulerMacroPath" "green"
		}
		catch
		{
			Write_Log "Failed to write scheduler macro: $_" "red"
			continue
		}

		# Remove archive bit if set (optional)
		try
		{
			if (Test-Path $DestScriptPath)
			{
				$file = Get-Item -Path $DestScriptPath
				if ($file.Attributes -band [System.IO.FileAttributes]::Archive)
				{
					$file.Attributes = $file.Attributes -bxor [System.IO.FileAttributes]::Archive
					Write_Log "Removed the archive bit from '$DestScriptPath'." "green"
				}
			}
		}
		catch
		{
			Write_Log "Failed to remove the archive bit from '$DestScriptPath'. Error: $_" "red"
		}
	}

	Write_Log "`r`n==================== Schedule_Lane_DB_Maintenance Function Completed ====================" "blue"
}
