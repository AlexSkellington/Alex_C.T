# ===================================================================================================
#                              FUNCTION: Start_Lane_Protocol_Jobs
# ---------------------------------------------------------------------------------------------------
# Description:
#   Starts the lane SQL protocol background loop: a UI-thread timer probes every lane's SQL
#   protocol (TCP -> Named Pipes -> else File, Offline when unreachable) once a minute in a
#   runspace pool, and keeps the results in $script:LaneProtocols + Protocol_Results.txt.
#   Lanes that end up on File get a one-shot "enable SQL protocols" task scheduled on the lane
#   (Ensure_SQL_Enable_Scheduled_Task_For_File_Lane, run in a hidden helper process).
#   Call once after node discovery; it returns immediately. A second call only updates the
#   lane map - the running loop picks it up on its next scan.
#
# Parameters:
#   - LaneNumToMachineName  [hashtable]: LaneNumber => MachineName
#   - SqlModuleName         [string]: not used (kept so existing callers keep working)
#
# Variables:
#   - $script:LaneProtocols                  lane number / raw key / machine name => protocol
#   - $script:ProtocolResults                @( [pscustomobject]@{ Lane; Protocol } ) sorted on write
#   - $script:ProtocolResultsFile            Protocol_Results.txt (warm start + persistence)
#   - $script:LaneProtocolJobs               lane => @{ PS; Handle; Started } in-flight probes
#   - $script:LaneProtocolPool               runspace pool (rebuilt when a probe hangs)
#   - $script:LaneProtocolLaneMap            the lane map the loop scans
#   - $script:LaneProtocolStop               set by FormClosing to stop the loop
#   - $script:LaneSqlProtocolEnableJobs      lane => helper process scheduling the SQL enable task
#   - $script:LaneSqlProtocolEnableTaskState lane => last scheduling result
#   - $script:LaneProtocolHelpers            the loop's helper scriptblocks (see Notes)
#
# Notes:
#   - The timer handler and its helpers are plain scriptblocks kept in $script: variables.
#     Do NOT use .GetNewClosure() here: a closure runs in its own module scope, so every
#     $script: variable inside it is a different, empty variable (v2.6.7 - v2.6.9 lost all
#     lane probing that way).
#   - Double-start guarded via $script:protocolTimer.
# ===================================================================================================

function Start_Lane_Protocol_Jobs
{
	param (
		[hashtable]$LaneNumToMachineName,
		[string]$SqlModuleName
	)

	try { Add-Type -AssemblyName System.Windows.Forms 2>$null }
	catch { }
	try { Add-Type -AssemblyName System.Data 2>$null }
	catch { }

	# ---------- paths ----------
	$script:ProtocolResultsFile = Join-Path $script:SetupFiles 'Protocol_Results.txt'
	$resultsDir = [System.IO.Path]::GetDirectoryName($script:ProtocolResultsFile)
	if (-not (Test-Path -LiteralPath $resultsDir)) { $null = New-Item -Path $resultsDir -ItemType Directory -Force }
	if (-not (Test-Path -LiteralPath $script:ProtocolResultsFile)) { $null = New-Item -Path $script:ProtocolResultsFile -ItemType File -Force }

	# ---------- state ----------
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
	if (-not $script:ProtocolResults) { $script:ProtocolResults = @() }
	if (-not $script:LaneProtocolJobs) { $script:LaneProtocolJobs = @{ } }
	if (-not $script:LaneProtocolJobTimeoutSeconds -or $script:LaneProtocolJobTimeoutSeconds -lt 20) { $script:LaneProtocolJobTimeoutSeconds = 20 }
	if (-not $script:LaneSqlProtocolEnableTaskState) { $script:LaneSqlProtocolEnableTaskState = @{ } }
	if (-not $script:LaneSqlProtocolEnableJobs) { $script:LaneSqlProtocolEnableJobs = @{ } }
	if (-not $LaneNumToMachineName) { $LaneNumToMachineName = @{ } }
	$script:LaneProtocolLaneMap = $LaneNumToMachineName
	$script:LaneProtocolStop = $false

	if ($script:protocolTimer) { return }

	# ---------- helpers (plain scriptblocks in script scope - see Notes) ----------
	$script:LaneProtocolHelpers = @{
		# Any value -> [datetime] or $null
		ToDate           = {
			param ($Value)
			if ($null -eq $Value) { return $null }
			if ($Value -is [datetime]) { return [datetime]$Value }
			if ($Value -is [datetimeoffset]) { return ([datetimeoffset]$Value).LocalDateTime }
			if ($Value -is [string])
			{
				$parsed = [datetime]::MinValue
				if (-not [string]::IsNullOrWhiteSpace($Value) -and [datetime]::TryParse($Value, [ref]$parsed)) { return $parsed }
				return $null
			}
			try { return [datetime]$Value }
			catch { return $null }
		}

		# Record one lane's protocol under every key consumers look it up by; returns the machine name
		SetProtocol      = {
			param ([string]$LaneNum, [string]$RawLane, [string]$Protocol)
			$script:LaneProtocols[$LaneNum] = $Protocol
			if ($RawLane) { $script:LaneProtocols[$RawLane] = $Protocol }
			$mn = $null
			if ($script:LaneProtocolLaneMap) { $mn = $script:LaneProtocolLaneMap[$LaneNum] }
			if ($mn) { $script:LaneProtocols[[string]$mn] = $Protocol }
			$script:ProtocolResults = @(@($script:ProtocolResults | Where-Object { $_.Lane -ne $LaneNum }) +
				[pscustomobject]@{ Lane = $LaneNum; Protocol = $Protocol })
			return $mn
		}

		NewPool          = {
			$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
			$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, [Math]::Max(8, [Environment]::ProcessorCount * 2), $iss, $Host)
			try { $pool.ApartmentState = 'MTA' }
			catch { }
			$pool.Open()
			return $pool
		}

		# Next StartTime (default 04:15) occurrence
		NextSqlEnableRun = {
			$next = (Get-Date).Date.AddHours(4).AddMinutes(15)
			if ((Get-Date) -ge $next) { $next = $next.AddDays(1) }
			return $next
		}

		# File-protocol lane: schedule the SQL enable task on it from a hidden helper process
		StartSqlEnable   = {
			param ([string]$LaneNumber, [string]$MachineName)
			if ([string]::IsNullOrWhiteSpace($LaneNumber) -or [string]::IsNullOrWhiteSpace($MachineName)) { return }

			$laneKey = $LaneNumber.PadLeft(3, '0')
			$running = $script:LaneSqlProtocolEnableJobs[$laneKey]
			if ($running -and $running.Process -and -not $running.Process.HasExited) { return }

			# The lane is rescanned every minute - re-check a lane at most every 15 minutes
			$state = $script:LaneSqlProtocolEnableTaskState[$laneKey]
			if ($state -and ($state.LastChecked -is [datetime]) -and ((Get-Date) - $state.LastChecked).TotalMinutes -lt 15) { return }

			$scheduledDate = & $script:LaneProtocolHelpers.NextSqlEnableRun
			$fnBody = ${function:Ensure_SQL_Enable_Scheduled_Task_For_File_Lane}
			if (-not $fnBody)
			{
				$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{ TaskPresent = $false; LastChecked = Get-Date; ScheduledFor = $scheduledDate; MachineName = $MachineName; LastResult = 'NoFunction' }
				return
			}

			# The helper process gets the function from memory (works for .ps1, .exe and irm|iex
			# launches alike) plus a Write_Log shim that prints to stdout, so the reason for a
			# failure comes back through the redirected output file.
			$launcherScriptPath = Join-Path $script:SetupFiles ("LaneSqlEnableLauncher_{0}.ps1" -f $laneKey)
			$launcherOutPath = Join-Path $script:SetupFiles ("LaneSqlEnableLauncher_{0}.out.txt" -f $laneKey)
			$launcherScript = (@(
					'param([string]$LaneNumber, [string]$MachineName, [string]$TcpPort = ''1433'', [string]$StartTime = ''04:15'')'
					'function Write_Log { param([string]$Message, $Color = ''Black'') Write-Host $Message }'
					'function Ensure_SQL_Enable_Scheduled_Task_For_File_Lane'
					'{'
					$fnBody.ToString()
					'}'
					'$ok = $false'
					'try { $ok = [bool](Ensure_SQL_Enable_Scheduled_Task_For_File_Lane -LaneNumber $LaneNumber -MachineName $MachineName -TcpPort $TcpPort -StartTime $StartTime) }'
					'catch { Write-Host ("ERROR: " + $_.Exception.Message) }'
					'if ($ok) { exit 0 } else { exit 1 }'
				) -join "`r`n")

			try
			{
				Set-Content -LiteralPath $launcherScriptPath -Value $launcherScript -Encoding UTF8 -Force -ErrorAction Stop
				$argList = @(
					'-NoProfile', '-ExecutionPolicy', 'Bypass',
					'-File', ('"{0}"' -f $launcherScriptPath),
					'-LaneNumber', $laneKey,
					'-MachineName', $MachineName,
					'-TcpPort', '1433',
					'-StartTime', '04:15'
				)
				$proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -WindowStyle Hidden -PassThru `
					-RedirectStandardOutput $launcherOutPath -ErrorAction Stop
			}
			catch
			{
				try { Remove-Item -LiteralPath $launcherScriptPath -Force -ErrorAction SilentlyContinue }
				catch { }
				$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{ TaskPresent = $false; LastChecked = Get-Date; ScheduledFor = $scheduledDate; MachineName = $MachineName; LastResult = 'LaunchFailed' }
				return
			}

			$script:LaneSqlProtocolEnableJobs[$laneKey] = [pscustomobject]@{
				Process      = $proc
				ScriptPath   = $launcherScriptPath
				OutputPath   = $launcherOutPath
				Started      = Get-Date
				MachineName  = $MachineName
				ScheduledFor = $scheduledDate
			}
		}

		# Collect finished helper processes; log only when a lane's result changes
		HarvestSqlEnable = {
			if (-not $script:LaneSqlProtocolEnableJobs -or $script:LaneSqlProtocolEnableJobs.Count -eq 0) { return }
			foreach ($laneKey in @($script:LaneSqlProtocolEnableJobs.Keys))
			{
				$st = $script:LaneSqlProtocolEnableJobs[$laneKey]
				if (-not $st) { continue }
				$proc = $st.Process
				if ($proc)
				{
					try { if (-not $proc.HasExited) { continue } }
					catch { continue }
				}

				$success = $false
				try { $success = ($proc -and $proc.ExitCode -eq 0) }
				catch { $success = $false }
				$reason = ''
				try
				{
					if ($st.OutputPath -and (Test-Path -LiteralPath $st.OutputPath))
					{
						$reason = [string](@(Get-Content -LiteralPath $st.OutputPath -ErrorAction Stop | Where-Object { $_.Trim() }) | Select-Object -Last 1)
					}
				}
				catch { }
				foreach ($tmp in @($st.ScriptPath, $st.OutputPath))
				{
					try { if ($tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } }
					catch { }
				}
				[void]$script:LaneSqlProtocolEnableJobs.Remove($laneKey)

				$result = if ($success) { 'Scheduled' } else { 'Failed' }
				$previous = $script:LaneSqlProtocolEnableTaskState[$laneKey]
				$changedResult = (-not $previous) -or ([string]$previous.LastResult -ne $result)
				$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
					TaskPresent  = $success
					LastChecked  = Get-Date
					ScheduledFor = $st.ScheduledFor
					MachineName  = [string]$st.MachineName
					LastResult   = $result
				}

				if ($changedResult)
				{
					try
					{
						if ($success)
						{
							Write_Log ("Lane {0} ({1}) is on the SQL File protocol - SQL protocol enable scheduled for {2:MM/dd/yyyy HH:mm}." -f $laneKey, $st.MachineName, $st.ScheduledFor) 'green'
						}
						else
						{
							$why = if ($reason) { $reason } else { 'no details' }
							Write_Log ("Lane {0} ({1}) is on the SQL File protocol - could not schedule the SQL protocol enable: {2}" -f $laneKey, $st.MachineName, $why) 'red'
						}
					}
					catch { }
				}
			}
		}
	}

	# ---------- warm from file (best effort) ----------
	try
	{
		foreach ($line in @(Get-Content -LiteralPath $script:ProtocolResultsFile -ErrorAction SilentlyContinue))
		{
			if ($line -match '^\s*([^,]+),\s*([^,]+)\s*$')
			{
				$laneRaw = $matches[1].Trim()
				$laneNum = ($laneRaw -replace '[^\d]', '').PadLeft(3, '0')
				$null = & $script:LaneProtocolHelpers.SetProtocol $laneNum $laneRaw $matches[2].Trim()
			}
		}
	}
	catch { }

	if (-not $script:LaneProtocolPool) { $script:LaneProtocolPool = & $script:LaneProtocolHelpers.NewPool }

	# Probe worker (runs in the pool; source: src\Resources\Workers\Lane_Protocol_Probe.ps1)
	$script:LaneProtocolWorker = @'
#TBS:EMBED Resources/Workers/Lane_Protocol_Probe.ps1
'@

	if (-not $script:LaneProtocolLastRun) { $script:LaneProtocolLastRun = (Get-Date).AddYears(-10) }

	# ---------- timer (UI thread) ----------
	$script:protocolTimer = New-Object System.Windows.Forms.Timer
	$script:protocolTimer.Interval = 1000
	$script:protocolTimer.add_Tick({
			if ($script:LaneProtocolStop) { return }
			$h = $script:LaneProtocolHelpers

			# 0) finished SQL-enable helper processes
			try { & $h.HarvestSqlEnable }
			catch { }

			# 1) finished / hung probes
			$changed = $false
			$stale = $false
			foreach ($laneKey in @($script:LaneProtocolJobs.Keys))
			{
				$st = $script:LaneProtocolJobs[$laneKey]
				if (-not $st) { continue }

				$started = & $h.ToDate $st.Started
				$timedOut = ($null -ne $st.Started)   # an unreadable start time counts as hung
				if ($started)
				{
					try { $timedOut = (((Get-Date) - $started).TotalSeconds -ge $script:LaneProtocolJobTimeoutSeconds) }
					catch { $timedOut = $true }
				}

				if ($timedOut)
				{
					# A probe that hangs this long means the lane is unusable right now
					$laneNum = ($laneKey -replace '[^\d]', '').PadLeft(3, '0')
					$null = & $h.SetProtocol $laneNum $laneKey 'Offline'
					$changed = $true
					$stale = $true
					continue
				}

				if ($st.Handle -and $st.Handle.IsCompleted)
				{
					$resultList = $null
					try { $resultList = $st.PS.EndInvoke($st.Handle) }
					catch { $resultList = $null }
					try { if ($st.PS) { $st.PS.Dispose() } }
					catch { }
					[void]$script:LaneProtocolJobs.Remove($laneKey)

					$result = $null
					if ($resultList -and $resultList.Count -ge 1) { $result = $resultList[0] }
					if (-not $result) { $result = [pscustomobject]@{ Lane = $laneKey; Protocol = 'Offline' } }

					$rawLane = [string]$result.Lane
					$laneNum = ($rawLane -replace '[^\d]', '').PadLeft(3, '0')
					$proto = [string]$result.Protocol
					if ([string]::IsNullOrWhiteSpace($proto)) { $proto = 'File' }

					$mn = & $h.SetProtocol $laneNum $rawLane $proto
					if ($proto -eq 'File' -and -not [string]::IsNullOrWhiteSpace($mn))
					{
						try { & $h.StartSqlEnable $laneNum ([string]$mn) }
						catch { }
					}
					elseif ($script:LaneSqlProtocolEnableTaskState.ContainsKey($laneNum))
					{
						# The lane answers on TCP / Named Pipes now - nothing to schedule any more
						[void]$script:LaneSqlProtocolEnableTaskState.Remove($laneNum)
					}
					$changed = $true
				}
			}

			# 1b) a hung probe: drop every in-flight probe and rebuild the pool so it cannot wedge the loop
			if ($stale)
			{
				foreach ($st in @($script:LaneProtocolJobs.Values))
				{
					try { if ($st -and $st.PS) { $st.PS.Dispose() } }
					catch { }
				}
				$script:LaneProtocolJobs = @{ }
				try
				{
					if ($script:LaneProtocolPool)
					{
						$script:LaneProtocolPool.Close()
						$script:LaneProtocolPool.Dispose()
					}
				}
				catch { }
				$script:LaneProtocolPool = & $h.NewPool
				$script:LaneProtocolLastRun = Get-Date
			}

			# 2) persist when something changed
			if ($changed)
			{
				try
				{
					$lines = foreach ($r in ($script:ProtocolResults | Sort-Object { ($_.Lane -replace '[^\d]', '') -as [int] })) { '{0},{1}' -f $r.Lane, $r.Protocol }
					[System.IO.File]::WriteAllLines($script:ProtocolResultsFile, [string[]]@($lines), [System.Text.Encoding]::UTF8)
				}
				catch { }
			}

			# 3) new scan every 60 s, only when no probe is in flight
			if ($script:LaneProtocolJobs.Count -gt 0) { return }
			$lastRun = & $h.ToDate $script:LaneProtocolLastRun
			if ($lastRun -and ((Get-Date) - $lastRun).TotalSeconds -lt 60) { return }
			$script:LaneProtocolLastRun = Get-Date

			$laneMap = $script:LaneProtocolLaneMap
			if (-not $laneMap) { return }
			foreach ($k in @($laneMap.Keys))
			{
				$numStr = ([string]$k -replace '[^\d]', '')
				if (-not $numStr) { continue }
				$machine = $laneMap[$k]
				if ([string]::IsNullOrWhiteSpace($machine)) { continue }
				$laneNum = $numStr.PadLeft(3, '0')

				$ps = [System.Management.Automation.PowerShell]::Create()
				$ps.RunspacePool = $script:LaneProtocolPool
				$null = $ps.AddScript($script:LaneProtocolWorker).AddArgument([string]$machine).AddArgument([string]$laneNum)
				$script:LaneProtocolJobs[$laneNum] = @{ PS = $ps; Handle = $ps.BeginInvoke(); Started = Get-Date }
			}
		})

	$script:protocolTimer.Start()
}
