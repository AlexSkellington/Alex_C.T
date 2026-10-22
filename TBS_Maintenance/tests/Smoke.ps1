<#
.SYNOPSIS
	Smoke test: runs the whole built script in a sandbox, shows the window off-screen for a few
	seconds, then reports what was built and every runtime error that occurred.

.DESCRIPTION
	A patched COPY of the built script runs in a hidden child Windows PowerShell (STA):
	  - every Tecnica_Systems / ProgramData / Storeman root is redirected to a TEMP sandbox,
	    so nothing outside %TEMP% is created or changed
	  - the NSSM background download is skipped
	  - the launch-time VNC export, hardware export and quarterly maintenance are switched off
	  - the window opens off-screen and closes itself after -Seconds
	Everything else (bootstrap, storeman detection, every function definition, the whole GUI,
	the pre-window store data load and the post-Show warmup steps) runs for real.

	Output: counts (functions, controls, menu items, warmup steps reached), the log box text,
	and the de-duplicated runtime errors ($Error, including ones the script caught itself).
	Use -Baseline / -Compare to diff two runs (e.g. before and after a refactor).

.PARAMETER ScriptPath
	The built script. Default: TBS_Maintenace_Script_V2.ps1 in the repository root.

.PARAMETER Seconds
	How long the window stays open (warmup runs during this time). Default 25.

.PARAMETER Baseline
	Save this run's report to the given file.

.PARAMETER Compare
	Compare this run's report with a saved baseline and list the differences.

.PARAMETER Detail
	Also print, for every runtime error, the source line that raised it (diagnosis only;
	detail lines are left out of -Baseline / -Compare).

.EXAMPLE
	powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\tests\Smoke.ps1 -Baseline $env:TEMP\smoke-before.txt
#>
[CmdletBinding()]
param (
	[string]$ScriptPath,
	[int]$Seconds = 25,
	[string]$Baseline,
	[string]$Compare,
	[switch]$Detail
)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'TBS_Maintenace_Script_V2.ps1' }
if (-not (Test-Path -LiteralPath $ScriptPath)) { throw "script not found: $ScriptPath" }

$sandbox = Join-Path $env:TEMP 'tbs-smoke'
if (Test-Path -LiteralPath $sandbox) { Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $sandbox -Force | Out-Null
$sbAlex = Join-Path $sandbox 'Tecnica_Systems\Alex_C.T'
$sbProgramData = Join-Path $sandbox 'ProgramData\Tecnica_Systems\Alex_C.T'
$sbStoreman = Join-Path $sandbox 'Storeman'

$text = [System.IO.File]::ReadAllText($ScriptPath)

function Patch([string]$Find, [string]$Replace, [switch]$Last)
{
	# Exactly one target must exist unless -Last (then the LAST occurrence is replaced:
	# '[void]$form.ShowDialog()' also appears inside dialog functions).
	$ix = if ($Last) { $script:text.LastIndexOf($Find) } else { $script:text.IndexOf($Find) }
	if ($ix -lt 0) { throw "smoke patch target not found: $Find" }
	if (-not $Last -and $script:text.IndexOf($Find, $ix + 1) -ge 0) { throw "smoke patch target is not unique: $Find" }
	$script:text = $script:text.Substring(0, $ix) + $Replace + $script:text.Substring($ix + $Find.Length)
}

Patch '$script:AlexRoot = "C:\Tecnica_Systems\Alex_C.T"' ('$script:AlexRoot = "' + $sbAlex + '"')
Patch '$script:AlexProgramDataRoot = Join-Path $env:ProgramData "Tecnica_Systems\Alex_C.T"' ('$script:AlexProgramDataRoot = "' + $sbProgramData + '"')
Patch '$script:StoremanRoots = @("C:\Storeman", "D:\Storeman")' ('$script:StoremanRoots = @("' + $sbStoreman + '")')
Patch 'Download_NSSM -Background | Out-Null' '# [smoke] NSSM download skipped'
# Launch-time tasks that touch the machine (DISM/SFC, remote exports): off at the source.
Patch '$script:AutoExportVncOnLaunch = $true' '$script:AutoExportVncOnLaunch = $false'
Patch '$script:AutoExportRemoteMachineInfoOnLaunch = $true' '$script:AutoExportRemoteMachineInfoOnLaunch = $false'
Patch '$script:AutoQuarterlySystemHealthMaintenanceOnLaunch = $true' '$script:AutoQuarterlySystemHealthMaintenanceOnLaunch = $false'

$harness = @'
# ---------------- [smoke] harness replaces ShowDialog ----------------
$script:AutoExportVncOnLaunch = $false
$script:AutoExportRemoteMachineInfoOnLaunch = $false
$script:AutoQuarterlySystemHealthMaintenanceOnLaunch = $false
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.Location = New-Object System.Drawing.Point(-4000, -4000)
$form.ShowInTaskbar = $false
$script:__smokeTimer = New-Object System.Windows.Forms.Timer
$script:__smokeTimer.Interval = __SMOKE_MS__
$script:__smokeTimer.Add_Tick({
		$script:__smokeTimer.Stop()
		# state that FormClosing tears down - capture it while the window is still open
		$script:__smokeProtocolTimer = [bool]$script:protocolTimer
		$form.Close()
	})
$form.Add_Shown({ $script:__smokeTimer.Start() })
$script:__smokeShowError = $null
try { [void]$form.ShowDialog() } catch { $script:__smokeShowError = $_.Exception.Message }

function __SmokeCount($items) { $n = 0; foreach ($i in @($items)) { $n++; if ($i -is [System.Windows.Forms.ToolStripMenuItem]) { $n += (__SmokeCount $i.DropDownItems) } }; $n }
function __SmokeControls($c) { $n = 0; foreach ($x in @($c.Controls)) { $n += 1 + (__SmokeControls $x) }; $n }
Write-Output ('SMOKE|show_error|' + [string]$script:__smokeShowError)
Write-Output ('SMOKE|functions|' + @(Get-ChildItem function:).Count)
Write-Output ('SMOKE|form_controls|' + (__SmokeControls $form))
foreach ($mn in 'ContextMenuServer', 'ContextMenuLane', 'ContextMenuScale', 'ContextMenuGeneral')
{
	$m = Get-Variable -Name $mn -ValueOnly -ErrorAction SilentlyContinue
	if ($m) { Write-Output ('SMOKE|menu_' + $mn + '|' + (__SmokeCount $m.Items)) } else { Write-Output ('SMOKE|menu_' + $mn + '|MISSING') }
}
Write-Output ('SMOKE|menu_renderer|' + [bool]('ModernToolStripRenderer' -as [type]))
Write-Output ('SMOKE|form_icon|' + [bool]$form.Icon)
Write-Output ('SMOKE|protocol_timer|' + [bool]$script:__smokeProtocolTimer)
Write-Output ('SMOKE|lane_map|' + $(if ($null -eq $script:LaneNumToMachineName) { 'null' } else { $script:LaneNumToMachineName.GetType().Name + ':' + @($script:LaneNumToMachineName.Keys).Count }))
foreach ($jn in @(Get-Job -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object -Unique)) { Write-Output ('SMOKE|job|' + $jn) }
Write-Output ('SMOKE|warmup_started|' + [bool]$script:StartupWarmupStarted)
Write-Output ('SMOKE|warmup_steps_left|' + @($script:StartupWarmupSteps).Count)
Write-Output ('SMOKE|test_env|' + [bool]$script:IsTestEnvironment)
Write-Output ('SMOKE|storeman_source|' + [string]$script:StoremanSource)
$logText = ''
try { $logText = [string]$logBox.Text } catch { }
foreach ($ll in ($logText -split "`r?`n")) { if ($ll.Trim()) { Write-Output ('SMOKE|log|' + ($ll -replace '^\d{1,2}:\d{2}:\d{2}\s*(AM|PM)?\s*', '')) } }
$seenErr = @{ }
foreach ($er in @($Error))
{
	$msg = ''
	try { $msg = [string]$er.Exception.Message } catch { $msg = [string]$er }
	$msg = ($msg -replace '\s+', ' ').Trim()
	$msg = $msg -replace [regex]::Escape($env:TEMP), '%TEMP%'
	$msg = $msg -replace '\b\d+(\.\d+)?\s*(ms|s)\b', 'N'
	if ($msg.Length -gt 220) { $msg = $msg.Substring(0, 220) }
	if (-not $seenErr.ContainsKey($msg)) { $seenErr[$msg] = 1 } else { $seenErr[$msg]++ }
	if (__SMOKE_DETAIL__)
	{
		$where = ''
		try { $where = ([string]$er.InvocationInfo.Line).Trim() } catch { }
		if ($where.Length -gt 200) { $where = $where.Substring(0, 200) }
		Write-Output ('SMOKE|detail|' + $msg.Substring(0, [Math]::Min(80, $msg.Length)) + ' @@ ' + $where)
	}
}
foreach ($k in ($seenErr.Keys | Sort-Object)) { Write-Output ('SMOKE|error|' + $k) }
Write-Output 'SMOKE|done|1'
'@
Patch '[void]$form.ShowDialog()' ($harness.Replace('__SMOKE_MS__', [string]($Seconds * 1000)).Replace('__SMOKE_DETAIL__', $(if ($Detail) { '$true' } else { '$false' }))) -Last

$patched = Join-Path $sandbox 'tbs_smoke_patched.ps1'
[System.IO.File]::WriteAllText($patched, $text, [System.Text.Encoding]::ASCII)
$outFile = Join-Path $sandbox 'smoke_out.txt'
$errFile = Join-Path $sandbox 'smoke_err.txt'

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$cmd = "`$ErrorActionPreference='Continue'; Invoke-Expression ([System.IO.File]::ReadAllText('$patched'))"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$proc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-Command', $cmd) `
	-RedirectStandardOutput $outFile -RedirectStandardError $errFile -PassThru -WindowStyle Hidden
$limitMs = ($Seconds + 90) * 1000
if (-not $proc.WaitForExit($limitMs))
{
	try { $proc.Kill() } catch { }
	Write-Host "SMOKE FAILED: child did not exit within $([int]($limitMs / 1000))s (a dialog may be blocking)" -ForegroundColor Red
}
$sw.Stop()

$lines = @()
if (Test-Path -LiteralPath $outFile) { $lines = @([System.IO.File]::ReadAllLines($outFile) | Where-Object { $_.StartsWith('SMOKE|') }) }
$all = @($lines | ForEach-Object { $_.Substring(6) })
$report = @($all | Where-Object { $_ -notlike 'detail|*' })
foreach ($d in @($all | Where-Object { $_ -like 'detail|*' } | Sort-Object -Unique)) { Write-Host "  $d" -ForegroundColor DarkGray }
$stderr = ''
if (Test-Path -LiteralPath $errFile) { $stderr = [System.IO.File]::ReadAllText($errFile).Trim() }

Write-Host ("Smoke run: {0:N1}s, exit code {1}" -f $sw.Elapsed.TotalSeconds, $(if ($proc.HasExited) { $proc.ExitCode } else { 'killed' })) -ForegroundColor Cyan
foreach ($r in $report) { if ($r -notlike 'log|*') { Write-Host "  $r" } }
Write-Host ("  log lines: {0}" -f @($report | Where-Object { $_ -like 'log|*' }).Count)
if ($stderr) { Write-Host "  stderr: $stderr" -ForegroundColor DarkYellow }
$ok = ($report -contains 'done|1') -and -not ($report | Where-Object { $_ -like 'show_error|?*' })
Write-Host $(if ($ok) { 'SMOKE PASSED (ran to completion)' } else { 'SMOKE FAILED' }) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })

if ($Baseline) { [System.IO.File]::WriteAllLines($Baseline, $report) ; Write-Host "baseline saved: $Baseline" }
if ($Compare)
{
	$base = @([System.IO.File]::ReadAllLines($Compare))
	$diff = @(Compare-Object $base $report)
	if ($diff.Count -eq 0) { Write-Host 'COMPARE: identical to baseline' -ForegroundColor Green }
	else
	{
		Write-Host ("COMPARE: {0} difference(s) vs baseline (<= baseline only, => this run only)" -f $diff.Count) -ForegroundColor Yellow
		$diff | ForEach-Object { Write-Host ("  {0} {1}" -f $_.SideIndicator, $_.InputObject) }
	}
}
if ($ok) { exit 0 } else { exit 1 }
