# ===================================================================================================
#                              FUNCTION: Ensure_SQL_Enable_Scheduled_Task_For_File_Lane
# ---------------------------------------------------------------------------------------------------
# Description:
#   For a lane stuck on the SQL File protocol, makes sure the "enable SQL protocols" task runs
#   on the lane at 4:15 AM, and reads back what the last run did:
#     1. Finds the lane's Storeman\Temp share and the drive it really is on (a marker file
#        written through \\lane\Storeman is looked for on the c$ / d$ / e$ / f$ admin shares).
#     2. Reads and removes <task>.result.txt left by the last run. A run that changed nothing
#        means rerunning cannot help (the cause is SQL not running, the network, ...): the
#        task is NOT scheduled again and the result says so.
#     3. Keeps a current task that has not run yet; recreates one that ran without leaving a
#        result (wrong path / crashed) or was created by an older version.
#     4. Writes the worker (Sql_Protocol_Settings.ps1 + Enable_SQL_Protocols_Worker.ps1) and
#        creates a DAILY 04:15 SYSTEM task - no /SD date, so the lane's date format does not
#        matter; the worker deletes the task after its run.
#   No logging here: the callers log Message / LastRun (it also runs in a helper process).
#
# Parameters:
#   - LaneNumber  [string]: lane to schedule (required)
#   - MachineName [string]: lane machine name (required)
#   - TcpPort     [string]: static TCP port to enable (default '1433')
#   - StartTime   [string]: task start time, HH:mm (default '04:15')
#
# Returns:
#   [pscustomobject] Status       'Scheduled' | 'AlreadyScheduled' | 'NothingToFix' | 'Failed'
#                    Message      one line for the log, Color its log colour
#                    ScheduledFor next run ([datetime]) when scheduled, else $null
#                    LastRun      summary of the run that was read back ('' if none), LastRunColor
#
# Writes:
#   - Enable_SQL_Protocols_Lane_<n>.ps1 in the lane's Storeman\Temp + the task of the same name
# ===================================================================================================

function Ensure_SQL_Enable_Scheduled_Task_For_File_Lane
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$LaneNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[string]$TcpPort = '1433',
		[string]$StartTime = '04:15'
	)

	$laneKey = $LaneNumber.ToString().PadLeft(3, '0')
	$taskName = "Enable_SQL_Protocols_Lane_$laneKey"
	$scriptName = "$taskName.ps1"
	$resultName = "$taskName.result.txt"

	# Next occurrence of StartTime (for the log only - the task itself is DAILY)
	$startSpan = [TimeSpan]::FromMinutes(255)
	$parsedSpan = [TimeSpan]::Zero
	if ([TimeSpan]::TryParse($StartTime, [ref]$parsedSpan)) { $startSpan = $parsedSpan }
	$StartTime = '{0:00}:{1:00}' -f $startSpan.Hours, $startSpan.Minutes
	$nextRun = (Get-Date).Date.Add($startSpan)
	if ((Get-Date) -ge $nextRun) { $nextRun = $nextRun.AddDays(1) }

	$result = [pscustomobject]@{
		Status       = 'Failed'
		Message      = ''
		Color        = 'red'
		ScheduledFor = $null
		LastRun      = ''
		LastRunColor = 'gray'
	}

	# ---- 1) Temp share + the local path of the same folder ----------------------------------------
	$shareTemp = $null
	$localTemp = $null
	$driveNote = ''
	try
	{
		if (Test-Path -LiteralPath "\\$MachineName\Storeman")
		{
			$temp = "\\$MachineName\Storeman\Temp"
			if (-not (Test-Path -LiteralPath $temp)) { $null = New-Item -Path $temp -ItemType Directory -Force -ErrorAction Stop }
			$marker = '_tbs_probe_{0}_{1}.tmp' -f $laneKey, [guid]::NewGuid().ToString('N')
			Set-Content -LiteralPath (Join-Path $temp $marker) -Value 'probe' -Encoding ASCII -Force -ErrorAction Stop
			$shareTemp = $temp
			foreach ($d in @('C', 'D', 'E', 'F'))
			{
				if (Test-Path -LiteralPath "\\$MachineName\$d`$\Storeman\Temp\$marker") { $localTemp = "${d}:\Storeman\Temp"; break }
			}
			Remove-Item -LiteralPath (Join-Path $temp $marker) -Force -ErrorAction SilentlyContinue
			if (-not $localTemp)
			{
				$localTemp = 'C:\Storeman\Temp'
				$driveNote = ' (admin shares not reachable - assumed Storeman is on C:)'
			}
		}
	}
	catch { $shareTemp = $null }
	if (-not $shareTemp)
	{
		# No usable Storeman share: an admin share of the drive that holds Storeman (never create Storeman itself)
		foreach ($d in @('C', 'D', 'E', 'F'))
		{
			$root = "\\$MachineName\$d`$\Storeman"
			try
			{
				if (-not (Test-Path -LiteralPath $root)) { continue }
				$temp = "$root\Temp"
				if (-not (Test-Path -LiteralPath $temp)) { $null = New-Item -Path $temp -ItemType Directory -Force -ErrorAction Stop }
				$probe = Join-Path $temp ('_tbs_probe_{0}.tmp' -f $laneKey)
				Set-Content -LiteralPath $probe -Value 'probe' -Encoding ASCII -Force -ErrorAction Stop
				Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
				$shareTemp = $temp
				$localTemp = "${d}:\Storeman\Temp"
				break
			}
			catch { }
		}
	}
	if (-not $shareTemp)
	{
		$result.Message = "could not schedule the SQL protocol enable: no writable Storeman\Temp share on $MachineName"
		return $result
	}
	$scriptShare = Join-Path $shareTemp $scriptName
	$resultShare = Join-Path $shareTemp $resultName
	$localScript = "$localTemp\$scriptName"

	# ---- 2) What the last run did ---------------------------------------------------------------------
	try
	{
		if (Test-Path -LiteralPath $resultShare)
		{
			$kv = @{ }
			foreach ($line in @(Get-Content -LiteralPath $resultShare -ErrorAction Stop))
			{
				if ($line -eq '') { break }
				if ($line -match '^(\w+)=(.*)$') { $kv[$matches[1]] = $matches[2].Trim() }
			}
			Remove-Item -LiteralPath $resultShare -Force -ErrorAction SilentlyContinue

			$changed = ($kv['Changed'] -eq 'True')
			$summary = "4:15 SQL enable ran {0}: {1}" -f $kv['Time'], $(if ($changed) { 'settings changed' } else { 'nothing to change' })
			if ($kv['Restarted']) { $summary += ", restarted $($kv['Restarted'])" }
			$summary += ", firewall rule TCP $TcpPort $($kv['Firewall'])"
			if ($kv['Error']) { $summary += " - errors: $($kv['Error'])" }
			$result.LastRun = $summary
			$result.LastRunColor = if ($kv['Error']) { 'red' } elseif ($changed) { 'green' } else { 'yellow' }

			if (-not $changed)
			{
				# Everything was already set: another run changes nothing
				$result.Status = 'NothingToFix'
				$result.Message = "still on the SQL File protocol although its SQL protocols, port $TcpPort and firewall rule are already set - check that SQL Server is running on the lane and that the server can reach it. Not scheduling the 4:15 enable again."
				return $result
			}
		}
	}
	catch { }

	# ---- 3) An existing task ------------------------------------------------------------------------
	# CSV columns are the same in every Windows language: 7 = Last Result, 9 = Task To Run
	$taskRow = $null
	$taskLines = @(& schtasks.exe /Query /S $MachineName /TN $taskName /V /FO CSV /NH 2>$null)
	if ($LASTEXITCODE -eq 0 -and $taskLines.Count -gt 0)
	{
		try { $taskRow = @($taskLines | Where-Object { $_ -match '\S' } | ConvertFrom-Csv -Header (1 .. 30))[0] }
		catch { $taskRow = $null }
		$lastResult = if ($taskRow) { [string]$taskRow.'7' } else { '' }
		$taskToRun = if ($taskRow) { [string]$taskRow.'9' } else { '' }
		$neverRan = ($lastResult -in @('', '267011', '267009'))   # 0x41303 not run yet, 0x41301 running
		$current = ($taskToRun.IndexOf($localScript, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -and $taskToRun -notmatch '-TaskName')
		if ($neverRan -and $current -and (Test-Path -LiteralPath $scriptShare))
		{
			$result.Status = 'AlreadyScheduled'
			$result.Message = "SQL protocol enable already scheduled for {0:MM/dd/yyyy HH:mm}." -f $nextRun
			$result.Color = 'green'
			$result.ScheduledFor = $nextRun
			return $result
		}
		if (-not $neverRan -and -not $result.LastRun)
		{
			$result.LastRun = "the 4:15 SQL enable task ran (result $lastResult) but left no result file - recreating it"
			$result.LastRunColor = 'yellow'
		}
		$null = & schtasks.exe /Delete /S $MachineName /TN $taskName /F 2>$null
	}

	# ---- 4) Worker + DAILY task -------------------------------------------------------------------------
	$workerScript = @'
param([string]$TcpPort = '1433')
#TBS:EMBED Resources/Workers/Sql_Protocol_Settings.ps1
#TBS:EMBED Resources/Workers/Enable_SQL_Protocols_Worker.ps1
'@
	try { Set-Content -LiteralPath $scriptShare -Value $workerScript -Encoding ASCII -Force -ErrorAction Stop }
	catch
	{
		$result.Message = "could not schedule the SQL protocol enable: failed to write $scriptShare ($($_.Exception.Message))"
		return $result
	}

	$taskAction = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \`"$localScript\`" -TcpPort $TcpPort"
	$createArgs = @('/Create', '/S', $MachineName, '/RU', 'SYSTEM', '/RL', 'HIGHEST', '/SC', 'DAILY', '/ST', $StartTime, '/TN', $taskName, '/TR', $taskAction, '/F')
	$createOutput = & schtasks.exe @createArgs 2>&1
	if ($LASTEXITCODE -ne 0)
	{
		Remove-Item -LiteralPath $scriptShare -Force -ErrorAction SilentlyContinue
		$result.Message = "could not schedule the SQL protocol enable on ${MachineName}: " + ([string]($createOutput | Out-String)).Trim()
		return $result
	}

	$result.Status = 'Scheduled'
	$result.Message = "SQL protocol enable scheduled for {0:MM/dd/yyyy HH:mm} ({1}){2}." -f $nextRun, $localScript, $driveNote
	$result.Color = 'green'
	$result.ScheduledFor = $nextRun
	return $result
}
