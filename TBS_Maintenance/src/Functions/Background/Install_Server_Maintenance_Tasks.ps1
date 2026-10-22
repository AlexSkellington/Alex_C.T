# ===================================================================================================
#                              FUNCTION: Install_Server_Maintenance_Tasks
# ---------------------------------------------------------------------------------------------------
# Description:
#   Launch-time step (replaces the v2.7.5 quarterly DISM/SFC check). Makes sure the two Windows
#   scheduled tasks exist and run the current worker (Get_Server_Maintenance_Worker), as SYSTEM:
#     TBS_Server_Maintenance_Weekly    Sundays at WeeklyTime            -Job Weekly  (drive optimize)
#     TBS_Server_Maintenance_Monthly   first Sunday of the month        -Job Monthly (DISM, SFC,
#                                      at MonthlyTime                                check disk)
#   No dates in the schedules, so the server's date format does not matter. A task is only
#   recreated when its command or schedule differs (compared through schtasks /XML, which reads
#   the same in every Windows language). Removes the old quarterly state / lock / worker files.
#   Then logs one line with the last result of each job, read from Log\Server_Maintenance.log.
#
# Parameters:
#   - WeeklyTime  [string]: HH:mm (default '02:30')
#   - MonthlyTime [string]: HH:mm (default '03:00')
#
# Variables:
#   - $script:ScriptsLog / $script:SetupFiles / $script:ScriptsFolder
#
# Writes:
#   - the two scheduled tasks, Scripts\Server_Maintenance.ps1
# ===================================================================================================

function Install_Server_Maintenance_Tasks
{
	param (
		[string]$WeeklyTime = '02:30',
		[string]$MonthlyTime = '03:00'
	)

	$logFolder = $script:ScriptsLog.TrimEnd('\')

	# ---- old quarterly maintenance (v2.7.5 and earlier): its state / lock / worker files ------------
	foreach ($old in @(
			(Join-Path $script:SetupFiles 'Quarterly_System_Health_Maintenance_State.json'),
			(Join-Path $script:SetupFiles 'Quarterly_System_Health_Maintenance.lock'),
			(Join-Path $script:ScriptsFolder 'Quarterly_System_Health_Maintenance_Worker.ps1')))
	{
		try { if (Test-Path -LiteralPath $old) { Remove-Item -LiteralPath $old -Force -ErrorAction Stop } }
		catch { }
	}

	$isAdmin = $false
	try { $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }
	catch { }

	if (-not $isAdmin)
	{
		try { Write_Log 'Server maintenance tasks not checked - not running as administrator.' 'gray' -FileOnly }
		catch { }
	}
	else
	{
		$worker = Get_Server_Maintenance_Worker
		if ($worker)
		{
			$tasks = @(
				[pscustomobject]@{ Name = 'TBS_Server_Maintenance_Weekly'; Job = 'Weekly'; Time = $WeeklyTime; Kind = 'ScheduleByWeek'; Schedule = @('/SC', 'WEEKLY', '/D', 'SUN') },
				[pscustomobject]@{ Name = 'TBS_Server_Maintenance_Monthly'; Job = 'Monthly'; Time = $MonthlyTime; Kind = 'ScheduleByMonthDayOfWeek'; Schedule = @('/SC', 'MONTHLY', '/MO', 'FIRST', '/D', 'SUN') }
			)
			$changed = @()
			foreach ($t in $tasks)
			{
				$arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Job {1} -LogFolder "{2}"' -f $worker, $t.Job, $logFolder

				# Current task: same command + schedule -> leave it (keeps its run history)
				$current = $false
				$xmlLines = @(& schtasks.exe /Query /TN $t.Name /XML 2>$null)
				if ($LASTEXITCODE -eq 0 -and $xmlLines.Count -gt 0)
				{
					try
					{
						$x = [xml]($xmlLines -join "`n")
						$exec = $x.Task.Actions.Exec
						$trigger = $x.Task.Triggers.CalendarTrigger
						$current = (([string]$exec.Arguments -replace '[\s"]', '') -eq ($arguments -replace '[\s"]', '')) -and
						([string]$trigger.StartBoundary -match ('T' + [regex]::Escape($t.Time) + ':00')) -and
						($null -ne $trigger.($t.Kind)) -and
						([string]$x.Task.Principals.Principal.UserId -match 'S-1-5-18|SYSTEM')
					}
					catch { $current = $false }
				}
				if ($current) { continue }

				$taskRun = 'powershell.exe ' + ($arguments -replace '"', '\"')
				$createArgs = @('/Create', '/TN', $t.Name, '/RU', 'SYSTEM', '/RL', 'HIGHEST') + $t.Schedule + @('/ST', $t.Time, '/TR', $taskRun, '/F')
				$out = & schtasks.exe @createArgs 2>&1
				if ($LASTEXITCODE -eq 0) { $changed += $t.Name }
				else
				{
					try { Write_Log ("Could not create the scheduled task {0}: {1}" -f $t.Name, ([string]($out | Out-String)).Trim()) 'red' }
					catch { }
				}
			}
			if ($changed.Count -gt 0)
			{
				try { Write_Log ("Server maintenance scheduled: drive optimize Sundays {0}, system repair first Sunday of the month {1} (log: {2}\Server_Maintenance.log)." -f $WeeklyTime, $MonthlyTime, $logFolder) 'green' }
				catch { }
			}
		}
	}

	# ---- last result of each job -----------------------------------------------------------------------
	$logPath = Join-Path $logFolder 'Server_Maintenance.log'
	$lines = @()
	try { if (Test-Path -LiteralPath $logPath) { $lines = @(Get-Content -LiteralPath $logPath -ErrorAction Stop) } }
	catch { }
	$parts = @()
	$worst = 0   # 0 ok, 1 warn, 2 fail
	foreach ($j in @(
			[pscustomobject]@{ Title = 'Weekly drive optimize'; Short = 'drive optimize'; MaxDays = 14 },
			[pscustomobject]@{ Title = 'Monthly system repair'; Short = 'system repair'; MaxDays = 62 }))
	{
		$last = $null
		for ($i = $lines.Count - 1; $i -ge 0; $i--)
		{
			if ($lines[$i] -match ('^(\d{4}-\d\d-\d\d \d\d:\d\d):\d\d\s+\[(OK|WARN|FAIL)\]\s+' + [regex]::Escape($j.Title) + ' finished')) { $last = $matches; break }
		}
		if (-not $last) { $parts += "$($j.Short): no run yet"; continue }
		$when = [datetime]::ParseExact($last[1], 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
		$text = '{0} {1} ({2:MMM d})' -f $j.Short, $last[2], $when
		if ($last[2] -eq 'FAIL') { $worst = 2 }
		elseif ($last[2] -eq 'WARN' -and $worst -lt 1) { $worst = 1 }
		if (((Get-Date) - $when).TotalDays -gt $j.MaxDays)
		{
			$text += ' - overdue, check the task'
			if ($worst -lt 1) { $worst = 1 }
		}
		$parts += $text
	}
	$color = @('green', 'yellow', 'red')[$worst]
	try { Write_Log ('Server maintenance: ' + ($parts -join ', ') + '.') $color }
	catch { }
}
