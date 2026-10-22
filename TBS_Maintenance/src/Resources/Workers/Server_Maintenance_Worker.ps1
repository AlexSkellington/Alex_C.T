# ---------------------------------------------------------------------------------------------------
# Server maintenance worker - embedded at build time into Get_Server_Maintenance_Worker and written to
#   Scripts\Server_Maintenance.ps1. Run by the two scheduled tasks (as SYSTEM) and by the
#   Repair Windows menu (detached, elevated):
#     -Job Weekly    drive optimize (SSD: TRIM, HDD: defrag when fragmented, unknown: Windows picks)
#     -Job Monthly   DISM repair + component cleanup, SFC, online check disk
#     -Job Custom    -Steps Optimize,Dism,Sfc,ChkScan,Defender,Cleanup,ChkdskReboot
#   Log\Server_Maintenance.log: one line per step, "yyyy-MM-dd HH:mm:ss  [TAG]  text  (duration)",
#   a [RUN] line when a run starts and "<title> finished" as its last line. Tags: OK WARN FAIL SKIP.
#   Raw tool output goes to Server_Maintenance_Detail_<time>.log (newest 6 kept). One run at a
#   time (Global\TBS_Server_Maintenance mutex).
# ---------------------------------------------------------------------------------------------------
param (
	[ValidateSet('Weekly', 'Monthly', 'Custom')]
	[string]$Job = 'Custom',
	[string]$Steps = '',
	[string]$LogFolder = 'C:\Tecnica_Systems\Alex_C.T\Log'
)

$ErrorActionPreference = 'Continue'
$title = switch ($Job) { 'Weekly' { 'Weekly drive optimize' } 'Monthly' { 'Monthly system repair' } default { 'On-demand maintenance' } }
$stepList = switch ($Job)
{
	'Weekly' { @('Optimize') }
	'Monthly' { @('Dism', 'Sfc', 'ChkScan') }
	default { @($Steps -split '[,; ]+' | Where-Object { $_ }) }
}

if (-not (Test-Path -LiteralPath $LogFolder)) { $null = New-Item -Path $LogFolder -ItemType Directory -Force }
$logPath = Join-Path $LogFolder 'Server_Maintenance.log'
$detailName = 'Server_Maintenance_Detail_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date)
$detailPath = Join-Path $LogFolder $detailName
$counts = @{ OK = 0; WARN = 0; FAIL = 0; SKIP = 0 }
$win8 = ([Environment]::OSVersion.Version -ge [version]'6.2')

$fmtDuration = {
	param ([TimeSpan]$t)
	if ($t.TotalSeconds -lt 90) { '{0:0} s' -f $t.TotalSeconds }
	elseif ($t.TotalMinutes -lt 90) { '{0:0} min' -f $t.TotalMinutes }
	else { '{0:0.0} h' -f $t.TotalHours }
}

# One line in the main log; $Since adds the step duration
$log = {
	param ([string]$Tag, [string]$Text, $Since)
	if ($counts.ContainsKey($Tag)) { $counts[$Tag]++ }
	if ($Since -is [datetime]) { $Text = '{0}  ({1})' -f $Text, (& $fmtDuration ((Get-Date) - $Since)) }
	$line = '{0:yyyy-MM-dd HH:mm:ss}  {1,-6} {2}' -f (Get-Date), "[$Tag]", $Text
	try { Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8 }
	catch { }
}

# Run a native tool; its output (NULs from UTF-16 tools dropped) goes to the detail log
$runTool = {
	param ([string]$Exe, [string[]]$ToolArgs)
	$out = @(& $Exe @ToolArgs 2>&1 | ForEach-Object { ([string]$_) -replace "`0", '' } | Where-Object { $_.Trim() })
	$code = $LASTEXITCODE
	try { Add-Content -LiteralPath $detailPath -Value (@('', ">>> $Exe $($ToolArgs -join ' ')  (exit $code)") + $out) -Encoding UTF8 }
	catch { }
	[pscustomobject]@{ Exit = $code; Text = ($out -join ' ') }
}

$hex = { param ([int]$Code) '0x{0:X8}' -f $Code }

# ---- one run at a time ------------------------------------------------------------------------
$mutex = $null
$owned = $true
try
{
	$mutex = [System.Threading.Mutex]::new($false, 'Global\TBS_Server_Maintenance')
	$owned = $mutex.WaitOne(0)
}
catch [System.Threading.AbandonedMutexException] { $owned = $true }
catch { $owned = $true }
if (-not $owned)
{
	& $log 'SKIP' "$title - another maintenance run is still going"
	exit 0
}

# ---- keep the main log short (done before the run, so a reader following it never sees it shrink)
try
{
	if (Test-Path -LiteralPath $logPath)
	{
		$all = @(Get-Content -LiteralPath $logPath -ErrorAction Stop)
		if ($all.Count -gt 2000) { Set-Content -LiteralPath $logPath -Value ($all | Select-Object -Last 1500) -Encoding UTF8 }
	}
}
catch { }

$runStart = Get-Date
& $log 'RUN' ("{0} started: {1}" -f $title, ($stepList -join ', '))

try
{
	foreach ($step in $stepList)
	{
		$t0 = Get-Date
		switch ($step)
		{
			'Optimize'
			{
				if (-not (Get-Command Optimize-Volume -ErrorAction SilentlyContinue))
				{
					& $log 'SKIP' 'Drive optimize - needs Windows 8 / Server 2012 or newer'
					break
				}
				$vols = @(Get-CimInstance Win32_Volume -Filter 'DriveType=3' -ErrorAction SilentlyContinue |
					Where-Object { $_.DriveLetter -and $_.FileSystem -in @('NTFS', 'ReFS') } | Sort-Object DriveLetter)
				foreach ($v in $vols)
				{
					$letter = $v.DriveLetter.TrimEnd(':')
					$ts = Get-Date
					$media = 'drive'
					try
					{
						$diskNumber = (Get-Partition -DriveLetter $letter -ErrorAction Stop).DiskNumber
						$pd = @(Get-PhysicalDisk -ErrorAction Stop | Where-Object { [string]$_.DeviceId -eq [string]$diskNumber })
						if ($pd.Count -gt 0 -and [string]$pd[0].MediaType -in @('SSD', 'HDD')) { $media = [string]$pd[0].MediaType }
					}
					catch { }
					$fragmented = {
						param ($Volume)
						try
						{
							$r = Invoke-CimMethod -InputObject $Volume -MethodName DefragAnalysis -ErrorAction Stop
							if ($r.ReturnValue -eq 0 -and $r.DefragAnalysis) { return [int]$r.DefragAnalysis.FilePercentFragmentation }
						}
						catch { }
						return $null
					}
					try
					{
						if ($media -eq 'SSD')
						{
							$verbose = Optimize-Volume -DriveLetter $letter -ReTrim -Verbose -ErrorAction Stop 4>&1
							Add-Content -LiteralPath $detailPath -Value (@('', ">>> Optimize-Volume ${letter}: -ReTrim") + @($verbose | ForEach-Object { [string]$_ })) -Encoding UTF8
							& $log 'OK' "${letter}: SSD - TRIM done" $ts
						}
						elseif ($media -eq 'HDD')
						{
							$before = & $fragmented $v
							if ($null -ne $before -and $before -lt 5)
							{
								& $log 'OK' "${letter}: HDD - $before% fragmented, no defrag needed" $ts
							}
							else
							{
								$verbose = Optimize-Volume -DriveLetter $letter -Defrag -Verbose -ErrorAction Stop 4>&1
								Add-Content -LiteralPath $detailPath -Value (@('', ">>> Optimize-Volume ${letter}: -Defrag") + @($verbose | ForEach-Object { [string]$_ })) -Encoding UTF8
								$after = & $fragmented $v
								& $log 'OK' ("${letter}: HDD - defrag {0}% -> {1}% fragmented" -f $before, $after) $ts
							}
						}
						else
						{
							# RAID controllers and virtual disks do not report SSD / HDD: Windows picks
							$verbose = Optimize-Volume -DriveLetter $letter -Verbose -ErrorAction Stop 4>&1
							Add-Content -LiteralPath $detailPath -Value (@('', ">>> Optimize-Volume ${letter}:") + @($verbose | ForEach-Object { [string]$_ })) -Encoding UTF8
							& $log 'OK' "${letter}: drive type not reported (RAID / virtual) - optimized, Windows chose TRIM or defrag" $ts
						}
					}
					catch { & $log 'WARN' "${letter}: $media - optimize failed: $($_.Exception.Message)" $ts }
				}
				if ($vols.Count -eq 0) { & $log 'SKIP' 'Drive optimize - no fixed NTFS / ReFS drive found' }
			}

			'Dism'
			{
				if (-not $win8) { & $log 'SKIP' 'DISM repair - not available on Windows 7 / Server 2008 R2'; break }
				$r = & $runTool 'dism.exe' @('/Online', '/Cleanup-Image', '/RestoreHealth', '/NoRestart')
				$why = switch ($r.Exit)
				{
					-2146498529 { 'repair files not found - Windows Update is not reachable from this server' }   # 0x800F081F
					-2146498298 { 'could not download repair files - check internet / WSUS access' }               # 0x800F0906
					-2146498554 { 'changes are pending - restart the server, then run it again' }                  # 0x800F0806
					740 { 'needs administrator rights' }
					87 { 'option not supported on this Windows version' }
					default { 'exit code ' + (& $hex $r.Exit) }
				}
				if ($r.Exit -eq 0) { & $log 'OK' 'DISM repair - Windows image healthy (any damage repaired)' $t0 }
				elseif ($r.Exit -eq 3010) { & $log 'WARN' 'DISM repair - image repaired, restart the server to finish' $t0 }
				else { & $log 'FAIL' "DISM repair - $why" $t0 }

				$t1 = Get-Date
				$r = & $runTool 'dism.exe' @('/Online', '/Cleanup-Image', '/StartComponentCleanup', '/NoRestart')
				if ($r.Exit -eq 0) { & $log 'OK' 'DISM component cleanup - done' $t1 }
				elseif ($r.Exit -eq -2146498554) { & $log 'WARN' 'DISM component cleanup - changes are pending, restart the server' $t1 }
				else { & $log 'WARN' ('DISM component cleanup - exit code ' + (& $hex $r.Exit)) $t1 }
			}

			'Sfc'
			{
				$r = & $runTool 'sfc.exe' @('/SCANNOW')
				# English and Spanish wording (the exit code says nothing useful)
				if ($r.Text -match 'did not find any integrity violations|no encontr.{0,40}infracci') { & $log 'OK' 'SFC - no damaged system files' $t0 }
				elseif ($r.Text -match 'successfully repaired|repar.{0,20}correctamente') { & $log 'OK' 'SFC - damaged system files found and repaired' $t0 }
				elseif ($r.Text -match 'unable to fix|no pudo reparar|no se pudieron reparar') { & $log 'FAIL' 'SFC - damaged system files it could not repair (run DISM repair, then SFC again; details in C:\Windows\Logs\CBS\CBS.log)' $t0 }
				elseif ($r.Text -match 'pending|pendiente') { & $log 'WARN' 'SFC - a system repair is pending - restart the server, then run it again' $t0 }
				elseif ($r.Text -match 'could not perform|no pudo realizar') { & $log 'FAIL' 'SFC - could not run the scan' $t0 }
				elseif ($r.Exit -eq 0) { & $log 'OK' 'SFC - finished' $t0 }
				else { & $log 'WARN' ("SFC - finished with exit code {0}, see $detailName" -f $r.Exit) $t0 }
			}

			'ChkScan'
			{
				$vols = @(Get-CimInstance Win32_Volume -Filter "DriveType=3 AND FileSystem='NTFS'" -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter } | Sort-Object DriveLetter)
				foreach ($v in $vols)
				{
					$ts = Get-Date
					$d = $v.DriveLetter
					$chkArgs = if ($win8) { @($d, '/scan') } else { @($d) }
					$r = & $runTool 'chkdsk.exe' $chkArgs
					switch ($r.Exit)
					{
						0 { & $log 'OK' "${d} check disk - no errors" $ts }
						1 { & $log 'OK' "${d} check disk - errors found and fixed" $ts }
						2 { & $log 'OK' "${d} check disk - cleanup done" $ts }
						3 { & $log 'FAIL' "${d} check disk - problems it can only fix at a restart: Repair Windows > 'Check disk with repair at next restart'" $ts }
						default { & $log 'WARN' "${d} check disk - exit code $($r.Exit), see $detailName" $ts }
					}
				}
			}

			'Defender'
			{
				$mp = Join-Path $env:ProgramFiles 'Windows Defender\MpCmdRun.exe'
				if (-not (Test-Path -LiteralPath $mp)) { & $log 'SKIP' 'Defender - not installed (another antivirus?)'; break }
				$r = & $runTool $mp @('-SignatureUpdate')
				if ($r.Exit -eq 0) { & $log 'OK' 'Defender - signatures updated' $t0 }
				else { & $log 'WARN' ('Defender - signature update failed ({0}) - no internet, or Defender is turned off by another antivirus' -f (& $hex $r.Exit)) $t0 }
				$t1 = Get-Date
				$r = & $runTool $mp @('-Scan', '-ScanType', '2')
				if ($r.Exit -eq 0) { & $log 'OK' 'Defender full scan - no threats' $t1 }
				elseif ($r.Exit -eq 2) { & $log 'WARN' 'Defender full scan - threats found, see Windows Security > Protection history' $t1 }
				else { & $log 'FAIL' ('Defender full scan - failed ({0}) - Defender may be turned off by another antivirus' -f (& $hex $r.Exit)) $t1 }
			}

			'Cleanup'
			{
				$cleanmgr = Join-Path $env:SystemRoot 'System32\cleanmgr.exe'
				if (-not (Test-Path -LiteralPath $cleanmgr)) { & $log 'SKIP' 'Disk cleanup - cleanmgr.exe not installed'; break }
				$caches = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
				$freeBefore = [double](Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace
				# Profile 99 = every cleanup handler (what /sageset would store)
				foreach ($k in @(Get-ChildItem -LiteralPath $caches -ErrorAction SilentlyContinue))
				{
					try { Set-ItemProperty -LiteralPath $k.PSPath -Name 'StateFlags0099' -Value 2 -Type DWord -ErrorAction Stop }
					catch { }
				}
				$p = Start-Process -FilePath $cleanmgr -ArgumentList '/sagerun:99' -WindowStyle Hidden -PassThru
				if (-not $p.WaitForExit(7200000))
				{
					try { $p.Kill() }
					catch { }
					& $log 'WARN' 'Disk cleanup - still running after 2 hours, stopped' $t0
				}
				else
				{
					$freed = ([double](Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace - $freeBefore) / 1GB
					& $log 'OK' ('Disk cleanup - freed {0:N1} GB on C:' -f [Math]::Max([double]0, $freed)) $t0
				}
				foreach ($k in @(Get-ChildItem -LiteralPath $caches -ErrorAction SilentlyContinue))
				{
					try { Remove-ItemProperty -LiteralPath $k.PSPath -Name 'StateFlags0099' -ErrorAction Stop }
					catch { }
				}
			}

			'ChkdskReboot'
			{
				$null = & $runTool 'cmd.exe' @('/c', 'echo Y| chkdsk C: /f /r')
				$boot = [string](@((Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name BootExecute -ErrorAction SilentlyContinue).BootExecute) -join ' ')
				if ($boot -match 'autochk.*/r.*C:') { & $log 'OK' 'C: check disk with repair scheduled for the next restart (can take hours on a large drive)' $t0 }
				else { & $log 'FAIL' "C: check disk could not be scheduled, see $detailName" $t0 }
			}

			default { & $log 'SKIP' "Unknown step '$step'" }
		}
	}
}
catch { & $log 'FAIL' "$title - stopped: $($_.Exception.Message)" }
finally
{
	$overall = if ($counts.FAIL -gt 0) { 'FAIL' } elseif ($counts.WARN -gt 0) { 'WARN' } else { 'OK' }
	$tail = if ($overall -ne 'OK' -and (Test-Path -LiteralPath $detailPath)) { " - details in $detailName" } else { '' }
	& $log $overall "$title finished$tail" $runStart

	try
	{
		@(Get-ChildItem -LiteralPath $LogFolder -Filter 'Server_Maintenance_Detail_*.log' -File -ErrorAction SilentlyContinue |
			Sort-Object LastWriteTime -Descending | Select-Object -Skip 6) | Remove-Item -Force -ErrorAction SilentlyContinue
	}
	catch { }
	if ($mutex)
	{
		try { $mutex.ReleaseMutex() }
		catch { }
		$mutex.Dispose()
	}
}
