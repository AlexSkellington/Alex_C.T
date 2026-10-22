# ===================================================================================================
#                              FUNCTION: Send_SERVER_time_to_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Time-sync wizard: shows the lane selection form, then hands the per-lane work to
#   Send_TIME_Sync_Lanes_Worker, running it as a background job via
#   Start_Background_Function_Job (inline fallback if the job cannot start).
#
# Parameters:
#   - StoreNumber  [string]: store to process, 3 digits (required)
#   - Schedule     [switch]: prompt for an interval (minutes) for repeated syncs
#
# Variables:
#   - $script:PrimaryStoremanRoot: storeman root (seeded to the worker job)
#   - $OfficePath: office path (seeded to the worker job)
#
# Notes:
#   - Per-lane work lives in Send_TIME_Sync_Lanes_Worker (v2.6.9 split for background jobs).
# ===================================================================================================

function Send_SERVER_time_to_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[switch]$Schedule # Optional: enables scheduling mode (prompt for interval)
	)
	
	Write_Log "`r`n==================== Starting Time Sync ====================`r`n" "blue"
	
	# 1. Unified Lane Picker
	$laneSelection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $laneSelection -or -not $laneSelection.Lanes -or $laneSelection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for time sync." "yellow"
		Write_Log "`r`n==================== Time Sync Aborted ====================`r`n" "blue"
		return
	}
	$selectedLanes = $laneSelection.Lanes
	
	# 2. Scheduling Mode Prompt (Optional)
	$isScheduling = $false
	$intervalMinutes = 0
	if ($Schedule)
	{
		Add-Type -AssemblyName System.Windows.Forms
		$formInterval = New-Object System.Windows.Forms.Form
		$formInterval.Text = "Set Sync Interval"
		$formInterval.Size = New-Object System.Drawing.Size(300, 150)
		$formInterval.StartPosition = "CenterScreen"
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "Enter interval in minutes:"
		$label.Location = New-Object System.Drawing.Point(20, 20)
		$formInterval.Controls.Add($label)
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(20, 50)
		$textBox.Size = New-Object System.Drawing.Size(240, 20)
		$formInterval.Controls.Add($textBox)
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(100, 80)
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$formInterval.Controls.Add($okButton)
		$result = $formInterval.ShowDialog()
		if ($result -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "Scheduling canceled." "yellow"
			return
		}
		if (-not [int]::TryParse($textBox.Text, [ref]$intervalMinutes) -or $intervalMinutes -le 0)
		{
			Write_Log "Invalid interval. Must be a positive integer." "red"
			return
		}
		$isScheduling = $true
	}
	
	# 3. (v2.6.9) the per-lane work runs as a BACKGROUND job so the
	#    UI thread never blocks. Per-lane details go to the file log as
	#    the job progresses; the shared job reaper logs the one-line
	#    completion notice to the log window.
	$bgJob = Start_Background_Function_Job `
		-FunctionName 'Send_TIME_Sync_Lanes_Worker' `
		-Label ('Time sync to ' + @($selectedLanes).Count + ' lane(s)') `
		-CallArgs @{ StoreNumber = $StoreNumber; SelectedLanes = $selectedLanes; IsScheduling = $isScheduling; IntervalMinutes = $intervalMinutes } `
		-SeedVars @{ OfficePath = $OfficePath; PrimaryStoremanRoot = $script:PrimaryStoremanRoot } `
		-DependentBodies @{ Ensure_MailslotSender_Type = (Get-Command Ensure_MailslotSender_Type -CommandType Function).ScriptBlock.ToString() } `
		-NotifyOnComplete
	if ($bgJob)
	{
		Write_Log ("Time sync to {0} lane(s) is running in the background - per-lane details go to the file log." -f @($selectedLanes).Count) "green"
	}
	else
	{
		Write_Log "Could not start the background time-sync job - running inline." "yellow"
		Send_TIME_Sync_Lanes_Worker -StoreNumber $StoreNumber -SelectedLanes $selectedLanes -IsScheduling:$isScheduling -IntervalMinutes $intervalMinutes
	}
}
