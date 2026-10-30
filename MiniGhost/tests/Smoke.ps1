<#
.SYNOPSIS
	Smoke test: runs the whole built MiniGhost.ps1, shows the window off-screen, renders it and
	its dialogs to PNG (optional), reports runtime errors, and exits without changing anything.

.DESCRIPTION
	A patched COPY of the built script runs in a hidden child Windows PowerShell (STA). Startup is
	read-only (storeman detection, INI reads, IP configuration), so it is safe on any PC; the
	patch only replaces the final ShowDialog with a harness that:
	  - opens the main window off-screen
	  - optionally renders it and the input / report dialogs (each dialog is opened, drawn and
	    cancelled - no button handler ever runs)
	  - prints counts and the de-duplicated runtime errors, then exits the child process (so the
	    "Are you sure you want to exit?" prompt never appears)

.PARAMETER ScriptPath
	The built script. Default: MiniGhost\MiniGhost.ps1.

.PARAMETER Screenshot
	Folder to save the PNG renders into.

.EXAMPLE
	powershell -NoProfile -ExecutionPolicy Bypass -File .\MiniGhost\tests\Smoke.ps1 -Screenshot $env:TEMP\mg-shots
#>
[CmdletBinding()]
param (
	[string]$ScriptPath,
	[string]$Screenshot,
	[int]$Seconds = 4
)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'MiniGhost.ps1' }
if (-not (Test-Path -LiteralPath $ScriptPath)) { throw "script not found: $ScriptPath" }

$sandbox = Join-Path $env:TEMP 'mg-smoke'
if (Test-Path -LiteralPath $sandbox) { Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $sandbox -Force | Out-Null
if ($Screenshot -and -not (Test-Path -LiteralPath $Screenshot)) { New-Item -ItemType Directory -Path $Screenshot -Force | Out-Null }
$shotDir = if ($Screenshot) { (Resolve-Path -LiteralPath $Screenshot).ProviderPath.Replace("'", "''") } else { '' }

$text = [System.IO.File]::ReadAllText($ScriptPath)
$find = '[void]$form.ShowDialog()'
$ix = $text.LastIndexOf($find)
if ($ix -lt 0) { throw "smoke patch target not found: $find" }

$harness = @'
# ---------------- [smoke] harness replaces ShowDialog ----------------
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.Location = New-Object System.Drawing.Point(-4000, -4000)
$form.ShowInTaskbar = $false
$script:__shotDir = '__SMOKE_SHOT__'
# Output written inside event handlers is discarded: collect it and print after ShowDialog.
# Dialog functions name their own window $form, so handlers use $script:__mainForm.
$script:__mainForm = $form
$script:__out = New-Object System.Collections.Generic.List[string]

# Draws the newest open form that is not the main window, then closes it (= Cancel)
$script:__shotName = $null
$script:__dialogTimer = New-Object System.Windows.Forms.Timer
$script:__dialogTimer.Interval = 700
$script:__dialogTimer.Add_Tick({
		$script:__dialogTimer.Stop()
		$open = @([System.Windows.Forms.Application]::OpenForms)
		$dlg = $open[$open.Count - 1]
		if ($dlg -and $dlg -ne $script:__mainForm)
		{
			try
			{
				$b = New-Object System.Drawing.Bitmap($dlg.Width, $dlg.Height)
				$dlg.DrawToBitmap($b, (New-Object System.Drawing.Rectangle(0, 0, $dlg.Width, $dlg.Height)))
				$b.Save((Join-Path $script:__shotDir ('minighost_' + $script:__shotName + '.png')), [System.Drawing.Imaging.ImageFormat]::Png)
				$b.Dispose()
			}
			catch { $script:__out.Add('shot_error|' + $script:__shotName + ': ' + $_.Exception.Message) }
			$dlg.Close()
		}
	})
function __SmokeDialog([string]$Name, [scriptblock]$Open)
{
	$script:__shotName = $Name
	$script:__dialogTimer.Start()
	try { $null = & $Open } catch { $script:__out.Add('shot_error|' + $Name + ': ' + $_.Exception.Message) }
}

$script:__mainTimer = New-Object System.Windows.Forms.Timer
$script:__mainTimer.Interval = __SMOKE_MS__
$script:__mainTimer.Add_Tick({
		$script:__mainTimer.Stop()
		if ($script:__shotDir)
		{
			try
			{
				$mf = $script:__mainForm
				$b = New-Object System.Drawing.Bitmap($mf.Width, $mf.Height)
				$mf.DrawToBitmap($b, (New-Object System.Drawing.Rectangle(0, 0, $mf.Width, $mf.Height)))
				$b.Save((Join-Path $script:__shotDir 'minighost_main.png'), [System.Drawing.Imaging.ImageFormat]::Png)
				$b.Dispose()
			}
			catch { $script:__out.Add('shot_error|main: ' + $_.Exception.Message) }
			__SmokeDialog 'store_number' { Get_NEW_Store_Number }
			__SmokeDialog 'machine_name' { Get_NEW_Machine_Name }
			__SmokeDialog 'prompt_with_log' { Show_Action_Prompt_With_Log -Text "Machine name changed to 'POS006', but one or more INI/SQL updates failed. A reboot is still required for the rename. Do you want to reboot now?" -Title 'Machine Rename Complete' -PrimaryButtonText 'Reboot Now' -SecondaryButtonText 'Later' -LogText 'log' }
			__SmokeDialog 'report' { Show_Text_Report_Form -Title 'Operation Summary' -Text ("StoreNumberChange: Successful`r`n  Message: Store number updated in INIs.`r`n`r`nMachineNameChange: Pending`r`n") -Width 600 -Height 400 }
		}

		# Hiding a modal window ends ShowDialog without FormClosing (no exit confirmation)
		$script:__mainForm.Visible = $false
	})
$form.Add_Shown({ $script:__mainTimer.Start() })
try { [void]$form.ShowDialog() } catch { Write-Output ('SMOKE|show_error|' + $_.Exception.Message) }
foreach ($o in $script:__out) { Write-Output ('SMOKE|' + $o) }
function __SmokeControls($c) { $n = 0; foreach ($x in @($c.Controls)) { $n += 1 + (__SmokeControls $x) }; $n }
Write-Output ('SMOKE|functions|' + @(Get-ChildItem function:).Count)
Write-Output ('SMOKE|form_controls|' + (__SmokeControls $form))
Write-Output ('SMOKE|node_role|' + [string]$script:NodeRole)
$seenErr = @{ }
foreach ($er in @($Error))
{
	$msg = ''
	try { $msg = [string]$er.Exception.Message } catch { $msg = [string]$er }
	$msg = ($msg -replace '\s+', ' ').Trim()
	if ($msg.Length -gt 220) { $msg = $msg.Substring(0, 220) }
	$seenErr[$msg] = 1
}
foreach ($k in ($seenErr.Keys | Sort-Object)) { Write-Output ('SMOKE|error|' + $k) }
Write-Output 'SMOKE|done|1'
[Environment]::Exit(0)
'@
$harness = $harness.Replace('__SMOKE_MS__', [string]($Seconds * 1000)).Replace('__SMOKE_SHOT__', $shotDir)
$text = $text.Substring(0, $ix) + $harness + $text.Substring($ix + $find.Length)

$patched = Join-Path $sandbox 'minighost_smoke_patched.ps1'
[System.IO.File]::WriteAllText($patched, $text, [System.Text.Encoding]::ASCII)
$outFile = Join-Path $sandbox 'smoke_out.txt'
$errFile = Join-Path $sandbox 'smoke_err.txt'

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$proc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', $patched, '-ShowConsole') `
	-RedirectStandardOutput $outFile -RedirectStandardError $errFile -PassThru -WindowStyle Hidden
if (-not $proc.WaitForExit(($Seconds + 60) * 1000))
{
	try { $proc.Kill() } catch { }
	Write-Host 'SMOKE FAILED: child did not exit (a dialog may be blocking)' -ForegroundColor Red
}

$lines = @()
if (Test-Path -LiteralPath $outFile) { $lines = @([System.IO.File]::ReadAllLines($outFile) | Where-Object { $_.StartsWith('SMOKE|') } | ForEach-Object { $_.Substring(6) }) }
foreach ($l in $lines) { Write-Host "  $l" }
$stderr = ''
if (Test-Path -LiteralPath $errFile) { $stderr = [System.IO.File]::ReadAllText($errFile).Trim() }
if ($stderr) { Write-Host "  stderr: $stderr" -ForegroundColor DarkYellow }
if ($Screenshot) { Get-ChildItem -LiteralPath $Screenshot -Filter 'minighost_*.png' | ForEach-Object { Write-Host ("  screenshot: " + $_.Name) } }

$ok = ($lines -contains 'done|1') -and -not ($lines | Where-Object { $_ -like 'show_error|*' -or $_ -like 'shot_error|*' })
Write-Host $(if ($ok) { 'SMOKE PASSED' } else { 'SMOKE FAILED' }) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
if ($ok) { exit 0 } else { exit 1 }
