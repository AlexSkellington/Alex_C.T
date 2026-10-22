# ===================================================================================================
#                              FUNCTION: Reboot_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Reboots selected nodes through a tabbed dialog (Lanes / Scales / Backoffices). Lanes try
#   the SMSStart mailslot first, then shutdown.exe, then Restart-Computer; scales and
#   backoffices try shutdown.exe, then Restart-Computer.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#   - Selection    [any]: pre-made selection object (default: Show_Node_Selection_Form)
#   - NodeTypes    [string[]]: tabs to show - Lane, Scale, Backoffice
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName, ScaleCodeToIPInfo, BackofficeNumToMachineName
# ===================================================================================================

function Reboot_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		$Selection,
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string[]]$NodeTypes = @("Lane", "Scale", "Backoffice")
	)
	Write_Log "`r`n==================== Starting Reboot_Nodes Function ====================`r`n" "blue"

	# Load global node data
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']

	# --- Use the shared selection dialog ---
	if ($Selection)
	{
		$selection = $Selection
	}
	else
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes $NodeTypes
	}
	if (-not $selection)
	{
		Write_Log "No nodes selected or dialog cancelled." "yellow"
		Write_Log "`r`n==================== Reboot_Nodes Function Completed ====================" "blue"
		return
	}

	# shutdown.exe, then Restart-Computer; $true when one of them took the request
	$rebootRemote = {
		param ([string]$Target, [string]$What)
		$null = & shutdown.exe /r /m "\\$Target" /t 0 /f 2>&1
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "shutdown.exe reboot sent to $What." "Green"
			return $true
		}
		Write_Log "shutdown.exe exit code $LASTEXITCODE for $What. Trying Restart-Computer." "Yellow"
		try
		{
			Restart-Computer -ComputerName $Target -Force -ErrorAction Stop
			Write_Log "Restart-Computer sent to $What." "Green"
			return $true
		}
		catch
		{
			Write_Log "All reboot methods failed for ${What}: $($_.Exception.Message)" "Red"
			return $false
		}
	}

	# ---- Lanes: SMSStart mailslot first ----
	if ($NodeTypes -contains "Lane" -and $selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		foreach ($laneNum in $selection.Lanes)
		{
			if (-not $LaneNumToMachineName -or -not $LaneNumToMachineName.ContainsKey($laneNum))
			{
				Write_Log "Lane mapping not found for $laneNum." "yellow"
				continue
			}
			$machine = $LaneNumToMachineName[$laneNum]
			Write_Log "Lane $laneNum on [$machine]: attempting mailslot reboot" "Yellow"
			$msResult = $false
			try { $msResult = [MailslotSender]::SendMailslotCommand("\\$machine\mailslot\SMSStart_${StoreNumber}${laneNum}", '@exec(REBOOT=1).') }
			catch { $msResult = $false }
			if ($msResult)
			{
				Write_Log "Mailslot reboot sent to $machine (Lane $laneNum)" "Green"
				continue
			}
			Write_Log "Mailslot reboot failed for $machine. Falling back to shutdown.exe" "Yellow"
			[void](& $rebootRemote $machine "$machine (Lane $laneNum)")
		}
	}

	# ---- Scales ----
	if ($NodeTypes -contains "Scale" -and $selection.Scales -and $selection.Scales.Count -gt 0)
	{
		foreach ($code in $selection.Scales)
		{
			if (-not $ScaleCodeToIPInfo -or -not $ScaleCodeToIPInfo.ContainsKey($code))
			{
				Write_Log "Scale mapping not found for $code." "yellow"
				continue
			}
			$scale = $ScaleCodeToIPInfo[$code]
			$ip = $scale.FullIP
			$display = if ($scale.ScaleName) { "$($scale.ScaleName) [$ip]" } else { "$code [$ip]" }
			Write_Log "Attempting to reboot scale: $display" "Yellow"
			[void](& $rebootRemote $ip "scale $display")
		}
	}

	# ---- Backoffices ----
	if ($NodeTypes -contains "Backoffice" -and $selection.Backoffices -and $selection.Backoffices.Count -gt 0)
	{
		foreach ($boNum in $selection.Backoffices)
		{
			if (-not $BackofficeNumToMachineName -or -not $BackofficeNumToMachineName.ContainsKey($boNum))
			{
				Write_Log "Backoffice mapping not found for $boNum." "yellow"
				continue
			}
			$machine = $BackofficeNumToMachineName[$boNum]
			Write_Log "Attempting to reboot backoffice $boNum [$machine]" "Yellow"
			[void](& $rebootRemote $machine "backoffice $boNum ($machine)")
		}
	}
	Write_Log "`r`n==================== Reboot_Nodes Function Completed ====================" "blue"
	[void][System.Windows.Forms.MessageBox]::Show("Reboot commands issued for selected nodes.", "Reboot",
		[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
