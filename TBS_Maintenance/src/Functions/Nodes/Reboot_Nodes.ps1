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

	# ---- Lanes ----
	if ($NodeTypes -contains "Lane" -and $selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		foreach ($lane in $selection.Lanes)
		{
			$laneNum = $lane
			if ($LaneNumToMachineName.ContainsKey($laneNum))
			{
				$machine = $LaneNumToMachineName[$laneNum]
				Write_Log "Lane $laneNum on [$machine]: attempting mailslot reboot" "Yellow"
				$msResult = $false
				try
				{
					$mailslot = "\\$machine\mailslot\SMSStart_${StoreNumber}${laneNum}"
					$msResult = [MailslotSender]::SendMailslotCommand($mailslot, '@exec(REBOOT=1).')
				}
				catch { $msResult = $false }
				if ($msResult)
				{
					Write_Log "Mailslot reboot sent to $machine (Lane $laneNum)" "Green"
					continue
				}
				Write_Log "Mailslot reboot failed for $machine. Falling back to shutdown.exe" "Yellow"
				cmd.exe /c "shutdown /r /m \\$machine /t 0 /f" | Out-Null
				if ($LASTEXITCODE -eq 0)
				{
					Write_Log "shutdown.exe reboot succeeded for $machine" "Green"
					continue
				}
				Write_Log "shutdown.exe exit code $LASTEXITCODE. Now trying Restart-Computer" "Yellow"
				try
				{
					Restart-Computer -ComputerName $machine -Force -ErrorAction SilentlyContinue
					if ($?)
					{
						Write_Log "Restart-Computer succeeded for $machine" "Green"
					}
					else
					{
						Write_Log "All reboot methods failed for $machine (Lane $laneNum)" "Red"
					}
				}
				catch
				{
					Write_Log "All reboot methods failed for $machine (Lane $laneNum)" "Red"
				}
			}
			else
			{
				Write_Log "Lane mapping not found for $laneNum." "yellow"
			}
		}
	}

	# ---- Scales ----
	if ($NodeTypes -contains "Scale" -and $selection.Scales -and $selection.Scales.Count -gt 0)
	{
		foreach ($code in $selection.Scales)
		{
			if ($ScaleCodeToIPInfo.ContainsKey($code))
			{
				$scale = $ScaleCodeToIPInfo[$code]
				$ip = $scale.FullIP
				$name = $scale.ScaleName
				$display = if ($name) { "$name [$ip]" }
				else { "$code [$ip]" }
				Write_Log "Attempting to reboot scale: $display at $ip" "Yellow"
				try
				{
					$shutdownArgs = "/r /m \\$ip /t 0 /f"
					$proc = Start-Process -FilePath "shutdown.exe" -ArgumentList $shutdownArgs -Wait -PassThru -ErrorAction Stop
					if ($proc.ExitCode -ne 0) { throw "Shutdown command exited with code $($proc.ExitCode)" }
					Write_Log "Shutdown command executed successfully for $ip." "Green"
				}
				catch
				{
					Write_Log "Shutdown command failed for $ip. Falling back to Restart-Computer." "Red"
					try
					{
						Restart-Computer -ComputerName $ip -Force -ErrorAction Stop
						Write_Log "Restart-Computer command executed successfully for $ip." "Green"
					}
					catch
					{
						Write_Log "Failed to reboot scale $ip using both methods: $_" "Red"
					}
				}
			}
			else
			{
				Write_Log "Scale mapping not found for $code." "yellow"
			}
		}
	}

	# ---- Backoffices ----
	if ($NodeTypes -contains "Backoffice" -and $selection.Backoffices -and $selection.Backoffices.Count -gt 0)
	{
		foreach ($boNum in $selection.Backoffices)
		{
			if ($BackofficeNumToMachineName.ContainsKey($boNum))
			{
				$machine = $BackofficeNumToMachineName[$boNum]
				Write_Log "Attempting to reboot backoffice $boNum [$machine]" "Yellow"
				try
				{
					$shutdownArgs = "/r /m \\$machine /t 0 /f"
					$proc = Start-Process -FilePath "shutdown.exe" -ArgumentList $shutdownArgs -Wait -PassThru -ErrorAction Stop
					if ($proc.ExitCode -ne 0) { throw "Shutdown command exited with code $($proc.ExitCode)" }
					Write_Log "Shutdown command executed successfully for $machine." "Green"
				}
				catch
				{
					Write_Log "Shutdown command failed for $machine. Falling back to Restart-Computer." "Red"
					try
					{
						Restart-Computer -ComputerName $machine -Force -ErrorAction Stop
						Write_Log "Restart-Computer command executed successfully for $machine." "Green"
					}
					catch
					{
						Write_Log "Failed to reboot backoffice $boNum ($machine) using both methods: $_" "Red"
					}
				}
			}
			else
			{
				Write_Log "Backoffice mapping not found for $boNum." "yellow"
			}
		}
	}

	Write_Log "`r`n==================== Reboot_Nodes Function Completed ====================" "blue"
	[System.Windows.Forms.MessageBox]::Show("Reboot commands issued for selected nodes.", "Reboot",
		[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
