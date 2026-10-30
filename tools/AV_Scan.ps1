<#
.SYNOPSIS
	Antivirus check for the shipped scripts: Microsoft Defender's AMSI scan (what runs on
	irm | iex content) and a Defender file scan, so a release is never flagged on the stores.

.DESCRIPTION
	Scans:
	  - TBS_Maintenance\dist\TBS_Maintenance_Script.ps1 (run TBS_Maintenance\build\Build.ps1 first)
	  - MiniGhost\MiniGhost.ps1
	  - the two root launchers
	For each file:
	  1. AMSI (AmsiScanString through amsi.dll, the same interface PowerShell uses before it
	     runs downloaded code) - result >= 32768 means "detected".
	  2. MpCmdRun.exe -Scan -ScanType 3 -File (file signatures; no remediation, nothing deleted).
	Prints Defender's status first: when real-time protection / AMSI is off the results say
	nothing, and the script reports that instead of a false "clean".

	Exit code: 0 = nothing detected, 1 = a detection, 2 = Defender not available (inconclusive).
#>
[CmdletBinding()]
param (
	[string[]]$Path
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Path)
{
	$Path = @(
		(Join-Path $repoRoot 'TBS_Maintenance\dist\TBS_Maintenance_Script.ps1'),
		(Join-Path $repoRoot 'MiniGhost\MiniGhost.ps1'),
		(Join-Path $repoRoot 'TBS_Maintenace_Script_V2.ps1'),
		(Join-Path $repoRoot 'MiniGhost_V38.ps1')
	)
}

$status = $null
try { $status = Get-MpComputerStatus -ErrorAction Stop } catch { }
if ($status)
{
	Write-Host ("Defender: AV enabled={0}  real-time={1}  AM service={2}  engine={3}  signatures={4} ({5})" -f `
		$status.AntivirusEnabled, $status.RealTimeProtectionEnabled, $status.AMServiceEnabled, $status.AMEngineVersion, $status.AntivirusSignatureVersion, $status.AntivirusSignatureLastUpdated)
}
else { Write-Host 'Defender: status not available' -ForegroundColor Yellow }

if (-not ('AvScanAmsi' -as [type]))
{
	Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class AvScanAmsi {
    [DllImport("amsi.dll", CharSet = CharSet.Unicode)] public static extern int AmsiInitialize(string appName, out IntPtr context);
    [DllImport("amsi.dll")] public static extern void AmsiUninitialize(IntPtr context);
    [DllImport("amsi.dll", CharSet = CharSet.Unicode)] public static extern int AmsiScanString(IntPtr context, string text, string contentName, IntPtr session, out int result);
}
'@
}

$ctx = [IntPtr]::Zero
$amsiOk = $false
try { $amsiOk = ([AvScanAmsi]::AmsiInitialize('TBS-AV-Check', [ref]$ctx) -eq 0) } catch { }

# Control: Microsoft's harmless AMSI test string MUST be detected, otherwise AMSI results mean nothing
# (built from two parts so this file itself is not the test sample)
if ($amsiOk)
{
	$control = 'AMSI Test Sample: ' + '7e72c3ce-861b-4339-8740-0ac1484c1386'
	$r = 0
	$hr = [AvScanAmsi]::AmsiScanString($ctx, $control, 'control', [IntPtr]::Zero, [ref]$r)
	if ($hr -eq 0 -and $r -ge 32768) { Write-Host '  OK    AMSI control sample detected - AMSI scanning works' -ForegroundColor Green }
	else
	{
		Write-Host ("  ?     AMSI control sample NOT detected (hr 0x{0:X8}, result {1}) - AMSI results skipped" -f $hr, $r) -ForegroundColor Yellow
		$amsiOk = $false
	}
}
$mpCmd = Join-Path $env:ProgramFiles 'Windows Defender\MpCmdRun.exe'

$detected = 0
$checked = 0
foreach ($p in $Path)
{
	if (-not (Test-Path -LiteralPath $p)) { Write-Host "  SKIP  $p (not found)" -ForegroundColor DarkGray; continue }
	$name = Split-Path -Leaf $p
	$text = [System.IO.File]::ReadAllText($p)

	if ($amsiOk)
	{
		$r = 0
		$hr = [AvScanAmsi]::AmsiScanString($ctx, $text, $name, [IntPtr]::Zero, [ref]$r)
		if ($hr -eq 0) { $checked++ }
		if ($hr -ne 0) { Write-Host ("  ?     AMSI {0}: scan call failed (0x{1:X8})" -f $name, $hr) -ForegroundColor Yellow }
		elseif ($r -ge 32768) { $detected++; Write-Host ("  FLAG  AMSI {0}: DETECTED (result {1})" -f $name, $r) -ForegroundColor Red }
		else { Write-Host ("  OK    AMSI {0}: not detected (result {1})" -f $name, $r) -ForegroundColor Green }
	}

	if (Test-Path -LiteralPath $mpCmd)
	{
		$out = & $mpCmd -Scan -ScanType 3 -File $p -DisableRemediation 2>&1
		$code = $LASTEXITCODE
		if ($code -eq 0 -or $code -eq 2) { $checked++ }
		if ($code -eq 2) { $detected++; Write-Host ("  FLAG  Defender file scan {0}: THREAT FOUND`n{1}" -f $name, (($out | Out-String).Trim())) -ForegroundColor Red }
		elseif ($code -eq 0) { Write-Host ("  OK    Defender file scan {0}: no threats" -f $name) -ForegroundColor Green }
		else { Write-Host ("  ?     Defender file scan {0}: exit {1} {2}" -f $name, $code, (($out | Out-String).Trim())) -ForegroundColor Yellow }
	}
}
if ($amsiOk) { [AvScanAmsi]::AmsiUninitialize($ctx) }

if ($detected -gt 0) { Write-Host "AV CHECK FAILED: $detected detection(s)" -ForegroundColor Red; exit 1 }
if ($checked -eq 0) { Write-Host 'AV CHECK INCONCLUSIVE: Defender / AMSI not available here' -ForegroundColor Yellow; exit 2 }
if (-not $amsiOk) { Write-Host 'note: AMSI (the irm | iex path) was not checked - only the file scan ran' -ForegroundColor Yellow }
Write-Host "AV CHECK PASSED ($checked scans, no detections)" -ForegroundColor Green
exit 0
