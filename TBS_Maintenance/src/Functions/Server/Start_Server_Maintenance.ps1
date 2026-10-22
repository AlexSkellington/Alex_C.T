# ===================================================================================================
#                              FUNCTION: Start_Server_Maintenance
# ---------------------------------------------------------------------------------------------------
# Description:
#   Starts the server maintenance worker (Scripts\Server_Maintenance.ps1, -Job Custom) as a
#   detached hidden PowerShell process and returns at once: the window stays usable, and the run
#   continues when the tool is closed. Elevated tool -> plain start; otherwise UAC is asked
#   (DISM / SFC / check disk need admin). While the tool is open, a timer copies the new lines of
#   Log\Server_Maintenance.log into the log box until the run's "finished" line.
#
# Parameters:
#   - Steps [string[]]: Optimize, Dism, Sfc, ChkScan, Defender, Cleanup, ChkdskReboot (required)
#
# Variables:
#   - $script:ServerMaintenanceTail       @{ Path; Offset; Started; Done } of the followed run
#   - $script:ServerMaintenanceTailTimer  the follow timer (one, reused)
#   - $script:ServerMaintenanceTailTick   its tick scriptblock (no closure - see CLAUDE.md)
#
# Returns:
#   $true when the worker was started.
# ===================================================================================================

function Start_Server_Maintenance
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$Steps
	)

	$worker = Get_Server_Maintenance_Worker
	if (-not $worker) { return $false }
	$logFolder = $script:ScriptsLog.TrimEnd('\')
	$logPath = Join-Path $logFolder 'Server_Maintenance.log'

	# Follow from the current end of the log: only this run's lines reach the log box
	$offset = 0
	try { if (Test-Path -LiteralPath $logPath) { $offset = (Get-Item -LiteralPath $logPath).Length } }
	catch { }

	$running = $false
	try
	{
		$m = [System.Threading.Mutex]::OpenExisting('Global\TBS_Server_Maintenance')
		$m.Dispose()
		$running = $true
	}
	catch [System.UnauthorizedAccessException] { $running = $true }   # exists, owned by SYSTEM
	catch { $running = $false }
	if ($running)
	{
		Write_Log "A server maintenance run is already going - not starting another one; showing its progress (log: $logPath)." 'yellow'
	}
	else
	{
		$isAdmin = $false
		try { $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }
		catch { }
		# Always Windows PowerShell itself (for the packaged .exe the current process is the tool)
		$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
		if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
		$argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden',
			'-File', ('"{0}"' -f $worker), '-Job', 'Custom', '-Steps', ($Steps -join ','), '-LogFolder', ('"{0}"' -f $logFolder))
		try
		{
			if ($isAdmin) { $null = Start-Process -FilePath $psExe -ArgumentList $argList -WindowStyle Hidden -ErrorAction Stop }
			else { $null = Start-Process -FilePath $psExe -ArgumentList $argList -WindowStyle Hidden -Verb RunAs -ErrorAction Stop }
		}
		catch
		{
			Write_Log "Could not start the server maintenance: $($_.Exception.Message)" 'red'
			return $false
		}
		Write_Log ("Server maintenance started in the background ({0}). It keeps running if you close this tool; progress shows here and in {1}." -f ($Steps -join ', '), $logPath) 'green'
	}

	$script:ServerMaintenanceTail = @{ Path = $logPath; Offset = $offset; Started = Get-Date; Done = $false }
	$script:ServerMaintenanceTailTick = {
		$t = $script:ServerMaintenanceTail
		if (-not $t) { $script:ServerMaintenanceTailTimer.Stop(); return }
		try
		{
			if (Test-Path -LiteralPath $t.Path)
			{
				$fs = [System.IO.File]::Open($t.Path, 'Open', 'Read', 'ReadWrite')
				try
				{
					if ($fs.Length -lt $t.Offset) { $t.Offset = $fs.Length }
					if ($fs.Length -gt $t.Offset)
					{
						$null = $fs.Seek($t.Offset, 'Begin')
						$buf = New-Object byte[] ($fs.Length - $t.Offset)
						$n = $fs.Read($buf, 0, $buf.Length)
						$text = [System.Text.Encoding]::UTF8.GetString($buf, 0, $n)
						$cut = $text.LastIndexOf("`n")
						if ($cut -ge 0)
						{
							$chunk = $text.Substring(0, $cut + 1)
							$t.Offset += [System.Text.Encoding]::UTF8.GetByteCount($chunk)
							foreach ($line in ($chunk.TrimStart([char]0xFEFF) -split "`r?`n"))
							{
								if ($line -notmatch '^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d\s+\[(\w+)\]\s+(.*)$') { continue }
								$tag = $matches[1]
								$msg = $matches[2]
								# Write_Log adds its own [ OK ] / [WARN] / [FAIL] tag from the colour
								$color = switch ($tag) { 'OK' { 'green' } 'WARN' { 'yellow' } 'FAIL' { 'red' } 'RUN' { 'blue' } default { 'gray' } }
								if ($tag -eq 'SKIP') { $msg = 'skipped: ' + $msg }
								Write_Log $msg $color
								if ($msg -match ' finished|another maintenance run is still going') { $t.Done = $true }
							}
						}
					}
				}
				finally { $fs.Dispose() }
			}
		}
		catch { }
		if ($t.Done -or ((Get-Date) - $t.Started).TotalHours -gt 12)
		{
			$script:ServerMaintenanceTailTimer.Stop()
			$script:ServerMaintenanceTail = $null
		}
	}
	if (-not $script:ServerMaintenanceTailTimer)
	{
		$script:ServerMaintenanceTailTimer = New-Object System.Windows.Forms.Timer
		$script:ServerMaintenanceTailTimer.Interval = 1000
		$script:ServerMaintenanceTailTimer.Add_Tick({ & $script:ServerMaintenanceTailTick })
	}
	$script:ServerMaintenanceTailTimer.Start()
	return $true
}
