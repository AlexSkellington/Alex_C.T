# ===================================================================================================
#                              FUNCTION: Send_TIME_Sync_Lanes_Worker
# ---------------------------------------------------------------------------------------------------
# Description:
#   The per-lane time-sync work behind Send_SERVER_time_to_Lanes (v2.6.9): builds a fresh
#   WIZRPL date/time string at start and pushes a time sync (remote schtasks net time task)
#   to each selected lane. Split out of the wizard so it can run in a background job via
#   Start_Background_Function_Job; running it inline is still supported as a fallback.
#
# Parameters:
#   - StoreNumber     [string]: store to process (required)
#   - SelectedLanes   [any]: lanes to sync, numbers or lane objects (required)
#   - IsScheduling    [switch]: create a recurring task instead of a one-shot sync
#   - IntervalMinutes [int]: schedule interval when -IsScheduling (default 0)
#
# Variables:
#   - $script:FunctionResults['LaneNumToMachineName']: lane map (input)
#
# Writes:
#   - remote scheduled time-sync tasks on the selected lanes
# ===================================================================================================

function Send_TIME_Sync_Lanes_Worker
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		$SelectedLanes,
		[switch]$IsScheduling,
		[int]$IntervalMinutes = 0
	)

	# The mailslot type does not exist in a job runspace; this is a
	# no-op where it is already loaded.
	try { Ensure_MailslotSender_Type } catch { }

	# 3. Build Time String
	$now = Get-Date
	$dateStr = $now.ToString("MM/dd/yyyy")
	$timeStr = $now.ToString("HHmmss")
	$wizCommand = "@WIZRPL(DATE=$dateStr)@WIZRPL(TIME=$timeStr)"

	# This server's IP for 'net time' (resolved once, not per lane)
	$serverIP = $env:COMPUTERNAME
	try { $serverIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -ErrorAction Stop).IPv4Address.IPAddressToString }
	catch { }
	$netTimeCommand = "net time \\$serverIP /set /yes"

	foreach ($lane in $SelectedLanes)
	{
		# Lane number normalization
		$laneNum = if ($lane -is [pscustomobject] -and $lane.LaneNumber) { $lane.LaneNumber }
		elseif ($lane -is [int]) { "{0:D3}" -f $lane }
		elseif ($lane.Length -lt 3) { $lane.PadLeft(3, '0') }
		else { $lane }

		$machine = $script:FunctionResults['LaneNumToMachineName'][$laneNum]
		if (-not $machine)
		{
			Write_Log "No machine name found for lane $laneNum." "red"
			continue
		}

		$success = $false

		# -------- PRIMARY: REMOTE SCHTASKS (SYSTEM task running net time against this server) --------
		try
		{
			$taskName = if ($IsScheduling) { "ScheduledTimeSync" } else { "SyncTime" }
			$createArgs = @('/create', '/s', $machine, '/tn', $taskName, '/tr', $netTimeCommand, '/ru', 'SYSTEM', '/f', '/rl', 'HIGHEST')
			if ($IsScheduling) { $createArgs += @('/sc', 'minute', '/mo', [string]$IntervalMinutes) }
			else
			{
				# /ST is mandatory for /SC ONCE (without it the create always failed before v2.7.0);
				# the task is started right away with /run, the time only satisfies schtasks
				$createArgs += @('/sc', 'once', '/st', (Get-Date).AddMinutes(1).ToString('HH:mm'))
			}
			$createOutput = & schtasks.exe @createArgs 2>&1
			if ($LASTEXITCODE -eq 0)
			{
				if ($IsScheduling)
				{
					Write_Log "Scheduled time sync every $IntervalMinutes minutes on lane $laneNum." "green"
					$success = $true
				}
				else
				{
					$runOutput = & schtasks.exe /run /s $machine /tn $taskName 2>&1
					if ($LASTEXITCODE -eq 0)
					{
						Write_Log "Executed time sync on lane $laneNum." "green"
						$success = $true
						# one-shot: remove the task once it had time to run
						Start-Sleep -Seconds 5
						$deleteOutput = & schtasks.exe /delete /s $machine /tn $taskName /f 2>&1
						if ($LASTEXITCODE -ne 0) { Write_Log "Delete output for [$machine]: $deleteOutput" "yellow" }
					}
					else { Write_Log "Run output for [$machine]: $runOutput" "yellow" }
				}
			}
			else { Write_Log "Create output for [$machine]: $createOutput" "yellow" }
		}
		catch
		{
			Write_Log "Exception during schtasks create/run: $_" "yellow"
		}
		# -------- SECONDARY: FILE-DROP WITH @EXEC --------
		if (-not $success)
		{
			try
			{
				$TempDir = $env:TEMP
				$xfFolder = Join-Path $OfficePath ("XF{0}{1}" -f $StoreNumber, $laneNum)

				if (-not (Test-Path $xfFolder))
				{
					Write_Log "XF folder for lane $laneNum does not exist: $xfFolder" "red"
					continue
				}

				# Always-available lane path
				$remoteShareRoot = "\\$machine\Storeman\Temp"
				if (-not (Test-Path $remoteShareRoot))
				{
					Write_Log "Lane share not reachable: $remoteShareRoot" "yellow"
					throw "Storeman\Temp not reachable"
				}

				$remoteBatPath = Join-Path $remoteShareRoot "sync_time.bat"

				# Create local BAT (ASCII)
				$batContent = "@echo off`r`nnet time \\$serverIP /set /yes"
				$localBatPath = Join-Path $TempDir ("sync_time_{0}.bat" -f $laneNum)
				Set-Content -Path $localBatPath -Value $batContent -Encoding Ascii

				# Copy to lane
				Copy-Item -Path $localBatPath -Destination $remoteBatPath -Force -ErrorAction Stop
				Remove-Item -Path $localBatPath -Force -ErrorAction SilentlyContinue
				Write_Log "Copied sync_time.bat to \\$machine\Storeman\Temp for lane $laneNum." "green"

				# Drop exec trigger in XF (ASCII)
				$execFilePath = Join-Path $xfFolder "exec_time_sync.txt"
				$execContent = "@EXEC(Run='" + (Join-Path $script:PrimaryStoremanRoot 'Temp\sync_time.bat') + "')"
				Write_Sms_File -Path $execFilePath -Content $execContent

				Write_Log "Created @EXEC trigger in XF folder for lane $laneNum (Storeman\Temp path)." "green"

				# The .bat stays in Storeman\Temp (overwritten next time): deleting it after 10 s
				# removed it before a busy lane had run the trigger, and the sync silently did nothing
				$success = $true
			}
			catch
			{
				Write_Log "File/@EXEC fallback failed for lane ${laneNum}: $_" "red"
			}
		}

		# -------- TERTIARY: MAILSLOT --------
		if (-not $success)
		{
			try
			{
				$mailslot = "\\$machine\mailslot\SMSStart_${StoreNumber}${laneNum}"
				$result = [MailslotSender]::SendMailslotCommand($mailslot, $wizCommand)
				if ($result)
				{
					Write_Log "Time sync sent to $machine (Lane $laneNum) via mailslot (final fallback)." "green"
					$success = $true
				}
				else
				{
					Write_Log "Mailslot failed for $machine (Lane $laneNum) [final fallback]." "yellow"
				}
			}
			catch
			{
				Write_Log "Exception during mailslot send (final fallback): $_" "yellow"
			}
		}

		if (-not $success)
		{
			Write_Log "All methods failed for lane $laneNum!" "red"
		}
	}

	Write_Log "`r`n==================== Ending Time Sync ====================`r`n" "blue"
}
