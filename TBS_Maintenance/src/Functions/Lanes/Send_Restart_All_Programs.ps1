# ===================================================================================================
#                              FUNCTION: Send_Restart_All_Programs
# ---------------------------------------------------------------------------------------------------
# Description:
#   Sends a restart-all-programs command to the selected lanes of a store: uses the cached
#   lane-to-machine map (Retrieve_Nodes when it is missing), picks lanes (or uses the passed
#   list) and sends "@exec(RESTART_ALL=PROGRAMS)." to each lane's SMSStart mailslot.
#
# Parameters:
#   - StoreNumber  [string]: store to process, 3 digits (required)
#   - LaneNumbers  [array]: lanes to target (default: selection form)
#   - Silent       [switch]: suppress Write_Log output
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName (input via Retrieve_Nodes)
# ===================================================================================================

function Send_Restart_All_Programs
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[array]$LaneNumbers,
		[Parameter(Mandatory = $false)]
		[switch]$Silent
	)

	if (-not $Silent)
	{
		Write_Log "`r`n==================== Starting Send_Restart_All_Programs Function ====================`r`n" "blue"
	}

	# Lane map from the startup Retrieve_Nodes run; only re-run it when the map is missing
	$laneMap = $script:FunctionResults['LaneNumToMachineName']
	if (-not $laneMap -or $laneMap.Count -eq 0)
	{
		$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
		if ($nodes) { $laneMap = $nodes.LaneNumToMachineName }
	}
	if (-not $laneMap)
	{
		Write_Log "Failed to retrieve node information for store $StoreNumber." "red"
		return
	}

	# Lane selection: use passed-in list or picker
	$lanes =
	if ($LaneNumbers -and $LaneNumbers.Count -gt 0)
	{
		$LaneNumbers | ForEach-Object {
			if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
			elseif ($_ -is [int]) { "{0:D3}" -f $_ }
			elseif ($_.Length -lt 3) { $_.PadLeft(3, '0') }
			else { $_ }
		}
	}
	else
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
		if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
		{
			Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
			return
		}
		$selection.Lanes | ForEach-Object {
			if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
			elseif ($_ -is [int]) { "{0:D3}" -f $_ }
			elseif ($_.Length -lt 3) { $_.PadLeft(3, '0') }
			else { $_ }
		}
	}

	foreach ($lane in $lanes)
	{
		$machineName = $laneMap[$lane]
		if (-not $machineName)
		{
			Write_Log "No machine found for lane $lane. Skipping." "yellow"
			continue
		}

		$mailslotAddress = "\\$machineName\Mailslot\SMSStart_${StoreNumber}${lane}"
		$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."

		try
		{
			$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
			if ($result)
			{
				if (-not $Silent)
				{
					Write_Log "Command sent successfully to Machine $machineName (Store $StoreNumber, Lane $lane)." "green"
				}
			}
			else
			{
				if (-not $Silent)
				{
					Write_Log "Failed to send command to Machine $machineName (Store $StoreNumber, Lane $lane)." "red"
				}
			}
		}
		catch
		{
			if (-not $Silent)
			{
				Write_Log "Exception sending to $machineName (Lane $lane): $_" "red"
			}
		}
	}

	if (-not $Silent)
	{
		Write_Log "`r`n==================== Send_Restart_All_Programs Function Completed ====================" "blue"
	}
}
