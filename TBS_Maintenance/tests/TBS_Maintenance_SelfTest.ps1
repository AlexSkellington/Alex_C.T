# ==============================================================================
#  TBS_Maintenance_SelfTest.ps1
# ------------------------------------------------------------------------------
#  Regression test for the TBS Maintenance Script.
#
#  How it works:
#    Instead of testing a copy, it extracts the real code blocks VERBATIM from
#    the script under test (by marker line), runs them in this process and
#    checks the results. Static checks (parse, ASCII, security strings) run
#    first.
#
#  Usage:
#    powershell -ExecutionPolicy Bypass -File TBS_Maintenance_SelfTest.ps1
#    powershell -ExecutionPolicy Bypass -File TBS_Maintenance_SelfTest.ps1 -ScriptPath C:\path\to\script.ps1
#
#  If -ScriptPath is not given, it uses the built TBS_Maintenance\dist\TBS_Maintenance_Script.ps1
#  (when run from TBS_Maintenance\tests - run Build.ps1 first), else it looks for
#  TBS_Maintenance_Script_V*.ps1 in:
#    1. the folder this file is in
#    2. %TEMP%
#    3. C:\Tecnica_Systems\Alex_C.T
#
#  Exit codes:  0 = all passed (skips allowed)   1 = one or more failed   2 = setup error
#
#  Notes:
#    - Read-only on the script under test; it only ever writes to %TEMP%.
#    - The file log check points the real Write_Log at a TEMP log folder, so
#      the real Tecnica_Systems log folder is never touched.
#    - On a machine without a storeman environment (dev PC) a few checks
#      report SKIP instead of FAIL.
# ==============================================================================

param
(
	[string]$ScriptPath = ''
)

$ErrorActionPreference = 'Stop'

$pass = 0
$fail = 0
$skip = 0

function Check
{
	param([string]$Name, $Cond, [string]$Note = '')
	if ([bool]$Cond)
	{
		$script:pass++
		Write-Host ("  PASS  " + $Name) -ForegroundColor Green
	}
	else
	{
		$script:fail++
		Write-Host ("  FAIL  " + $Name + $(if ($Note) { "  -- " + $Note } else { '' })) -ForegroundColor Red
	}
}

function Skip
{
	param([string]$Name, [string]$Note = '')
	$script:skip++
	Write-Host ("  SKIP  " + $Name + $(if ($Note) { "  -- " + $Note } else { '' })) -ForegroundColor DarkGray
}

# --------------------------------------------------------------------------
# 1) Locate the script under test
# --------------------------------------------------------------------------
if (-not $ScriptPath -and $PSScriptRoot)
{
	# Repository layout: TBS_Maintenance\tests\ -> the built script at the repository root
	$repoBuilt = Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\TBS_Maintenance_Script.ps1'
	if (Test-Path -LiteralPath $repoBuilt) { $ScriptPath = $repoBuilt }
}
if (-not $ScriptPath)
{
	$searchDirs = @()
	if ($PSCommandPath) { $searchDirs += (Split-Path -Parent $PSCommandPath) }
	$searchDirs += $env:TEMP
	$searchDirs += 'C:\Tecnica_Systems\Alex_C.T'
	foreach ($d in $searchDirs)
	{
		if ($d -and (Test-Path -LiteralPath $d))
		{
			$candidate = @(Get-ChildItem -Path $d -Filter 'TBS_Maintenance_Script_V*.ps1' -ErrorAction SilentlyContinue |
				Where-Object { $_.Name -notmatch '\.bak' } |
				Sort-Object LastWriteTime -Descending |
				Select-Object -First 1).FullName
			if ($candidate) { $ScriptPath = $candidate; break }
		}
	}
}

if (-not $ScriptPath -or -not (Test-Path -LiteralPath $ScriptPath))
{
	Write-Host 'Could not find the TBS Maintenance Script. Pass -ScriptPath <file>.' -ForegroundColor Red
	exit 2
}

Write-Host ("Testing: " + $ScriptPath) -ForegroundColor Cyan
Write-Host ''

$scriptLines = [System.IO.File]::ReadAllLines($ScriptPath)
$rawText = [System.IO.File]::ReadAllText($ScriptPath)

# --------------------------------------------------------------------------
# Helpers: find marker lines and extract code blocks verbatim
# --------------------------------------------------------------------------
function Find-Marker
{
	param([string]$Marker, [int]$StartAfter = -1)
	for ($i = [Math]::Max(0, $StartAfter + 1); $i -lt $scriptLines.Count; $i++)
	{
		if ($scriptLines[$i].Contains($Marker)) { return $i }
	}
	throw "marker not found: $Marker"
}

function Find-FunctionStart
{
	param([string]$Name)
	for ($i = 0; $i -lt $scriptLines.Count; $i++)
	{
		if ($scriptLines[$i] -match ('^function\s+' + [regex]::Escape($Name) + '\s*$')) { return $i }
	}
	throw ('function not found: ' + $Name)
}

function Get-BlockFromLine
{
	# From a start line, returns the lines up to (and including) the line where
	# the curly braces balance back to zero.
	param([int]$StartLine)
	$depth = 0
	$started = $false
	$e = -1
	for ($i = $StartLine; $i -lt $scriptLines.Count; $i++)
	{
		foreach ($ch in $scriptLines[$i].ToCharArray())
		{
			if ($ch -eq '{') { $depth++; $started = $true }
			elseif ($ch -eq '}') { $depth-- }
		}
		if ($started -and $depth -eq 0) { $e = $i; break }
	}
	if ($e -lt 0) { throw ('braces never balanced from line ' + ($StartLine + 1)) }
	return ($scriptLines[$StartLine..$e] -join "`n")
}

function Get-FunctionBlock
{
	param([string]$Name)
	return (Get-BlockFromLine (Find-FunctionStart $Name))
}

# --------------------------------------------------------------------------
# 2) Static checks (no code executed)
# --------------------------------------------------------------------------
Write-Host 'Static checks:' -ForegroundColor Cyan

$tokens = $null
$parseErrors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
Check 'parse: no syntax errors under Windows PowerShell 5.1' (@($parseErrors).Count -eq 0) ('{0} error(s)' -f @($parseErrors).Count)
if (@($parseErrors).Count -gt 0)
{
	foreach ($er in @($parseErrors))
	{
		Write-Host ("      line {0}: {1}" -f $er.Extent.StartLineNumber, $er.Message) -ForegroundColor DarkRed
	}
}

$nonAscii = [regex]::Matches($rawText, '[^\u0000-\u007F]')
Check 'file is pure ASCII (no special characters)' ($nonAscii.Count -eq 0) ('{0} non-ASCII char(s)' -f $nonAscii.Count)

$ver = [regex]::Match($rawText, '^\s*\$VersionNumber\s*=\s*"([\d.]+)"', 'Multiline')
Check 'version variables present' $ver.Success ''
if ($ver.Success)
{
	$verDate = [regex]::Match($rawText, '^\s*\$VersionDate\s*=\s*"([\d-]+)"', 'Multiline')
	Write-Host ("      version = {0}  date = {1}" -f $ver.Groups[1].Value, $verDate.Groups[1].Value) -ForegroundColor DarkGray
	Check 'changelog has an entry for the current version' ($rawText.Contains('#  v' + $ver.Groups[1].Value + '  -')) ''
}

$codeOnly = ($rawText -split "`r?`n" | Where-Object { $_.Trim() -notmatch '^#' }) -join "`n"
Check 'security: no remote activation URLs in live code' (($codeOnly -notmatch 'get\.activated\.win') -and ($codeOnly -notmatch 'massgrave\.dev')) ''
Check 'security: no hard-coded activation password in file' (-not $rawText.Contains('112922')) ''

Write-Host ''
Write-Host 'Functional checks (extracted code runs for real):' -ForegroundColor Cyan

# Safe sandbox for anything the extracted code may want to write
$testRoot = Join-Path $env:TEMP 'tbs-selftest'
if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

$script:FunctionResults = @{}
$script:IsTestEnvironment = $false
$script:StoremanKeyStale = $false
$script:StoremanSource = 'Production'
$script:StoremanSourceReason = ''
$script:StoremanKeyDate = $null
$script:StoremanRoots = @('C:\Storeman', 'D:\Storeman')
$script:PrimaryStoremanRoot = 'C:\Storeman'

Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------
# A) Storeman detection - the real block, run against THIS machine's state.
#    The only re-derivation is one date comparison read straight from the
#    key file (the same comparison the block itself makes); everything else
#    is an invariant of the block's own output, so the test stays valid on
#    any machine (fresh key, expired key, no share, ...).
# --------------------------------------------------------------------------
$hasShare = $false
$sharePath = ''
$infoFound = $false
$keyStringOk = $false
$keyDate = $null
try
{
	$realShare = @(Get-SmbShare | Where-Object { $_.Name -ieq 'storeman' } | Select-Object -First 1)
	if ($realShare)
	{
		$hasShare = $true
		$sharePath = $realShare.Path
		$infoFile = @(Get-ChildItem -Path (Join-Path $sharePath 'Office\Dbs') -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue | Select-Object -First 1)
		if ($infoFile)
		{
			$infoFound = $true
			$iniText = [System.IO.File]::ReadAllText($infoFile.FullName)
			$mk = [regex]::Match($iniText, '(?im)^\s*KeyString\s*=\s*(.+?)\s*$')
			if ($mk.Success) { $keyStringOk = ($mk.Groups[1].Value.Trim() -ne '') }
			$md = [regex]::Match($iniText, '(?im)^\s*KeyValidDate\s*=\s*(\d{2}/\d{2}/\d{4})\s*$')
			if ($md.Success)
			{
				$dv = [datetime]::MinValue
				# NOTE: 5-argument TryParseExact - the 4-argument overload does
				# not exist in Windows PowerShell 5.1 (.NET Framework)
				if ([datetime]::TryParseExact($md.Groups[1].Value, 'MM/dd/yyyy', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dv))
				{
					$keyDate = $dv
				}
			}
		}
	}
}
catch { }
$shouldBeStale = $hasShare -and $infoFound -and $keyStringOk -and ($null -ne $keyDate) -and ($keyDate.Date -lt (Get-Date).Date.AddDays(-14))
$shouldBeProduction = $hasShare -and $infoFound -and $keyStringOk -and ($null -ne $keyDate) -and (-not $shouldBeStale)

$detStart = Find-Marker '$BasePath = $null'
$detEnd = Find-Marker '-ForegroundColor Magenta' $detStart
$detBlock = ($scriptLines[$detStart..$detEnd]) -join "`n"

try
{
	Invoke-Expression "`n$detBlock`n`$script:TestBasePath = `$BasePath"
	Check 'A: detection block ran without error' $true ''
}
catch
{
	Check 'A: detection block ran without error' $false $_.Exception.Message
}

Check 'A: BasePath resolved non-empty' (-not [string]::IsNullOrWhiteSpace($script:TestBasePath)) ('BasePath: ' + $script:TestBasePath)
Check 'A: KEY EXPIRED flag matches the key file' ($script:StoremanKeyStale -eq [bool]$shouldBeStale) ('expected ' + [bool]$shouldBeStale + ', key date ' + $(if ($keyDate) { $keyDate.ToString('MM/dd/yyyy') } else { 'n/a' }))

if ($script:StoremanKeyStale)
{
	Check 'A: KEY EXPIRED -> test mode flagged' ($script:IsTestEnvironment -eq $true) ''
	Check 'A: KEY EXPIRED -> reason mentions re-key' ($script:StoremanSourceReason -like '*re-key*') ('reason: ' + $script:StoremanSourceReason)
	$gotDate = ''
	if ($script:StoremanKeyDate -is [datetime]) { $gotDate = $script:StoremanKeyDate.ToString('MM/dd/yyyy') }
	Check 'A: KEY EXPIRED -> key date recorded' ($keyDate -and ($gotDate -eq $keyDate.ToString('MM/dd/yyyy'))) ('expected ' + $(if ($keyDate) { $keyDate.ToString('MM/dd/yyyy') } else { 'n/a' }) + ', got ' + $gotDate)
}

if ($shouldBeProduction)
{
	Check 'A: fresh key -> production mode (no test flag)' ($script:IsTestEnvironment -eq $false) ('IsTest=' + $script:IsTestEnvironment + ' Stale=' + $script:StoremanKeyStale)
	Check 'A: fresh key -> BasePath is the share' ($script:TestBasePath -ieq $sharePath) ('BasePath: ' + $script:TestBasePath)
}

if (-not $hasShare)
{
	Check 'A: no share -> test/fallback mode with a reason' ($script:IsTestEnvironment -eq $true -and ($script:StoremanSourceReason -ne '')) ('source: ' + $script:StoremanSource + ' reason: ' + $script:StoremanSourceReason)
}

# --------------------------------------------------------------------------
# B) Drive-scan fallback: no share, no test root -> scan drive top levels.
#    Skipped on machines that have no storeman folder at all.
# --------------------------------------------------------------------------
$script:StoremanRoots = @('C:\NoSuchStoremanRoot')
$script:PrimaryStoremanRoot = 'C:\NoSuchStoremanRoot'
$script:IsTestEnvironment = $false
$script:StoremanSource = 'Production'
$script:StoremanSourceReason = ''
$script:StoremanKeyStale = $false
$script:StoremanKeyDate = $null

$machineStoreman = $null
foreach ($drv in @('C:\', 'D:\', 'E:\', 'F:\'))
{
	if (-not (Test-Path -LiteralPath $drv)) { continue }
	foreach ($c in @(Get-ChildItem -Path $drv -Directory -Filter '*storeman*' -ErrorAction SilentlyContinue))
	{
		if (Test-Path -LiteralPath (Join-Path $c.FullName 'Startup.ini')) { $machineStoreman = $c.FullName; break }
	}
	if ($machineStoreman) { break }
}

if (-not $machineStoreman)
{
	Skip 'B: drive scan finds a local storeman folder' 'no storeman folder on this machine'
}
else
{
	function Get-SmbShare { [CmdletBinding()] param([switch]$All) return $null }
	try
	{
		$detBlock2 = ($scriptLines[$detStart..$detEnd]) -join "`n"
		Invoke-Expression "`n$detBlock2`n`$script:TestBasePath2 = `$BasePath"
		Check 'B: drive scan ran without error' $true ''
		Check 'B: drive scan found the local storeman folder' ($script:TestBasePath2 -ieq $machineStoreman) ('BasePath: ' + $script:TestBasePath2)
		Check 'B: source flagged as drive-scan fallback' ($script:StoremanSource -eq 'Drive-scan fallback') ('source: ' + $script:StoremanSource)
		Check 'B: test mode flagged' ($script:IsTestEnvironment -eq $true) ''
		Check 'B: reason mentions drive scan' ($script:StoremanSourceReason -like '*drive scan*') ('reason: ' + $script:StoremanSourceReason)
	}
	catch
	{
		Check 'B: drive scan ran without error' $false $_.Exception.Message
	}
	finally
	{
		Remove-Item function:Get-SmbShare -ErrorAction SilentlyContinue
	}
}

# --------------------------------------------------------------------------
# C) Store number normalization - the real startup block, 9 cases
# --------------------------------------------------------------------------
$sn = Find-Marker '$script:StoreNumber = [string]$script:FunctionResults[''StoreNumber'']'
$snEnd = Find-Marker '$script:StoreName = $script:FunctionResults[''StoreName'']' $sn
$normBlock = ($scriptLines[$sn..($snEnd - 1)]) -join "`n"
$cases = @(
	@('1', '001'),
	@('01', '001'),
	@('001', '001'),
	@('12', '012'),
	@('123', '123'),
	@('1234', '1234'),
	@('12345', '12345'),
	@('abc', 'abc'),
	@('', '')
)
foreach ($c in $cases)
{
	$script:FunctionResults = @{}
	$script:FunctionResults['StoreNumber'] = $c[0]
	$script:StoreNumber = $null
	Invoke-Expression $normBlock
	Check ("C: store number '{0}' -> '{1}'" -f $c[0], [string]$script:StoreNumber) ([string]$script:StoreNumber -ceq $c[1]) ''
}

# --------------------------------------------------------------------------
# D) Old remote activation: disabled, inert, no prompt, no network
# --------------------------------------------------------------------------
$funcBlock = Get-FunctionBlock 'Invoke_Secure_Script'
$ixDisabled = $funcBlock.IndexOf('Security hardening')
# v2.7.0 removed the unreachable prompt/download code entirely
Check 'D: disabled, with no prompt or download code left' (($ixDisabled -ge 0) -and ($funcBlock -notmatch 'ShowDialog|Invoke-RestMethod|Invoke-WebRequest|DownloadString|storedPassword')) ''

$script:logLines = @()
function Write_Log { [CmdletBinding()] param([Parameter(Mandatory)][string]$Message, $Color = 'Black', [switch]$IncludeTimestamp, [switch]$Raw) $script:logLines += $Message }
Invoke-Expression $funcBlock
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Invoke_Secure_Script
$sw.Stop()
$joined = ($script:logLines -join "`n")
Check 'D: logs the disabled state' ($joined -like '*disabled*') ''
# 'irm' appears in the disabled message itself, so the real invariant is: no
# password prompt and none of the old remote URLs anywhere in the log.
Check 'D: no authentication prompt, no remote URLs in log' (($joined -notlike '*Authentication*') -and ($joined -notlike '*get.activated.win*') -and ($joined -notlike '*massgrave.dev*')) ''
Check 'D: returns quickly (no network round trip)' ($sw.Elapsed.TotalSeconds -lt 2) ('{0:N1}s' -f $sw.Elapsed.TotalSeconds)

# --------------------------------------------------------------------------
# E) TCP-445 reachability probe - the real block (bounded, no SMB timeout stall)
# --------------------------------------------------------------------------
$ps = Find-Marker '$remoteShareReachable = $false'
$pe = Find-Marker 'finally { $tcpClient.Dispose() }' $ps
$probeBlock = ($scriptLines[$ps..$pe]) -join "`n"

$machineName = $env:COMPUTERNAME
Invoke-Expression $probeBlock
if ($remoteShareReachable)
{
	Check 'E: probe of this machine => reachable' $true ''
}
else
{
	Skip 'E: probe of this machine => reachable' 'no local SMB (dev machine)'
}

$machineName = 'no-such-host-tbs-selftest.invalid'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Invoke-Expression $probeBlock
$sw.Stop()
Check 'E: probe of bogus host => unreachable' ($remoteShareReachable -eq $false) ''
Check 'E: bogus-host probe is bounded (<4s)' ($sw.Elapsed.TotalSeconds -lt 4) ('{0:N1}s' -f $sw.Elapsed.TotalSeconds)

# --------------------------------------------------------------------------
# F) VNC job throttle + collection - the real blocks (the v2.6.9 fix)
# --------------------------------------------------------------------------
$throttleText = Get-BlockFromLine (Find-Marker 'while ($jobs.Count -ge $MaxConcurrentJobs)')
$cs = Find-Marker 'if (@($jobs).Count -gt 0)'
$ce = Find-Marker 'Remove-Job $j -Force' $cs
$collectText = ($scriptLines[$cs..($ce + 1)]) -join "`n"

$ftok = $null; $ferr = $null
$null = [System.Management.Automation.Language.Parser]::ParseInput($throttleText, [ref]$ftok, [ref]$ferr)
Check 'F: throttle block extracted cleanly' (@($ferr).Count -eq 0) ''
$ftok2 = $null; $ferr2 = $null
$null = [System.Management.Automation.Language.Parser]::ParseInput($collectText, [ref]$ftok2, [ref]$ferr2)
Check 'F: collect block extracted cleanly' (@($ferr2).Count -eq 0) ''

$AllVNCPasswords = @{}
$MaxConcurrentJobs = 2
$jobs = @()
$fakeMachines = @('vm-a', 'vm-b', 'vm-c', 'vm-d', 'vm-e', 'vm-f')
foreach ($m in $fakeMachines)
{
	$jobs += Start-Job -Name ('tbstest_' + $m) -ScriptBlock {
		param ($mm)
		Start-Sleep -Seconds 2
		return @{ Machine = $mm; Password = ('pw-' + $mm) }
	} -ArgumentList $m
}
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Invoke-Expression ($throttleText + "`n" + $collectText)
$sw.Stop()
Check 'F: all 6 VNC results collected (throttle drops none)' ($AllVNCPasswords.Count -eq 6) ('got ' + $AllVNCPasswords.Count)
$allOk = $true
foreach ($m in $fakeMachines) { if ($AllVNCPasswords[$m] -ne ('pw-' + $m)) { $allOk = $false } }
Check 'F: every password value correct' $allOk ''
Check 'F: finished without hanging (<30s)' ($sw.Elapsed.TotalSeconds -lt 30) ('{0:N1}s' -f $sw.Elapsed.TotalSeconds)

# --------------------------------------------------------------------------
# G) File log - the real Write_Log, pointed at a TEMP log folder
# --------------------------------------------------------------------------
$logDir = Join-Path $testRoot 'Log'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$script:ScriptsLog = $logDir
$script:LogFileLogDate = $null
$script:LogFileLogPath = $null

try
{
	$wlText = Get-FunctionBlock 'Write_Log'
	$wtok = $null; $werr = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($wlText, [ref]$wtok, [ref]$werr)
	Check 'G: Write_Log extracted cleanly' (@($werr).Count -eq 0) ''
	Invoke-Expression $wlText
}
catch
{
	Check 'G: Write_Log extracted cleanly' $false $_.Exception.Message
}

Write_Log -Message 'SelfTest file-log probe' -Color 'green'
Write_Log -Message 'SelfTest file-log probe' -Color 'green'   # identical line again: 120s suppression
$logFile = Join-Path $logDir ('TBS_Maintenance_' + (Get-Date -Format 'yyyy-MM-dd') + '.log')
$probeCount = 0
$probeLine = ''
if (Test-Path -LiteralPath $logFile)
{
	$hits = @(Select-String -Path $logFile -Pattern 'SelfTest file-log probe')
	$probeCount = $hits.Count
	if ($probeCount -gt 0) { $probeLine = $hits[0].Line }
}
Check 'G: log file created in the Log folder' (Test-Path -LiteralPath $logFile) $logFile
Check 'G: probe line written once (repeat suppression honored)' ($probeCount -eq 1) ('count=' + $probeCount)
Check 'G: line format is "HH:mm:ss [ OK ] message"' ($probeLine -match '^\d{2}:\d{2}:\d{2} \[ OK \] SelfTest file-log probe$') ('line: ' + $probeLine)

# v2.6.9 -FileOnly: the line must still reach the FILE log (with its severity
# tag) even though the box and console are skipped for it.
Write_Log -Message 'SelfTest file-only probe' -Color 'blue' -FileOnly
Write_Log -Message 'SelfTest file-only probe' -Color 'blue' -FileOnly   # identical line again: 120s suppression
$hitsFo = @(Select-String -Path $logFile -Pattern 'SelfTest file-only probe')
Check 'G: -FileOnly line written to the file log' (@($hitsFo).Count -eq 1) ('count=' + @($hitsFo).Count)
if (@($hitsFo).Count -gt 0)
{
	Check 'G: -FileOnly line keeps the [INFO] tag in the file' ($hitsFo[0].Line -match '\[INFO\] SelfTest file-only probe$') ('line: ' + $hitsFo[0].Line)
}
else
{
	Check 'G: -FileOnly line keeps the [INFO] tag in the file' $false 'no line found'
}

# Test mode: known expected conditions are reworded, every other line passes unchanged.
$script:IsTestEnvironment = $true
Write_Log -Message 'Startup.ini path not provided (selftest)' -Color 'red'
Write_Log -Message 'SelfTest plain line in test mode' -Color 'green'
$script:IsTestEnvironment = $false
$tmLines = @(Get-Content -LiteralPath $logFile -ErrorAction SilentlyContinue)
Check 'G: test mode rewords an expected condition (orange [WARN])' (@($tmLines | Where-Object { $_ -match '\[WARN\] Store config \(Startup\.ini\) not found on this machine' }).Count -eq 1) ''
Check 'G: test mode leaves other lines unchanged' (@($tmLines | Where-Object { $_ -match '\[ OK \] SelfTest plain line in test mode$' }).Count -eq 1) ('last: ' + ($tmLines | Select-Object -Last 1))

# --------------------------------------------------------------------------
# H) Background job engine - the real helper functions, end-to-end round trip
# --------------------------------------------------------------------------
# Extracts Start_Background_Function_Job + Get_Background_Function_Job_Result
# verbatim and writes them into a TEMP .ps1 TOGETHER with a fake function and
# a small driver. They must share one file scope (PowerShell 5.1 function:
# resolution quirk), so the driver runs in a child powershell process and
# prints machine-readable RT_ lines that this test checks.
#
# Verified invariants:
#   - job starts and registers in BgfJobs
#   - a poll while the job runs returns $null (no early merge)
#   - the job's FunctionResults merge back into the parent script scope
#   - a pre-existing key is not clobbered
#   - named ExtraVars land in the parent script scope
#   - the job is removed and BgfJobs is cleaned up
#   - the failure path reports Success=$false with the error message
try
{
	$bgfStartText = Get-FunctionBlock 'Start_Background_Function_Job'
	$bgfTok = $null; $bgfErr = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($bgfStartText, [ref]$bgfTok, [ref]$bgfErr)
	Check 'H: Start_Background_Function_Job extracted cleanly' (@($bgfErr).Count -eq 0) ''
	$bgfEndText = Get-FunctionBlock 'Get_Background_Function_Job_Result'
	$bgfTok2 = $null; $bgfErr2 = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($bgfEndText, [ref]$bgfTok2, [ref]$bgfErr2)
	Check 'H: Get_Background_Function_Job_Result extracted cleanly' (@($bgfErr2).Count -eq 0) ''
}
catch
{
	Check 'H: job engine extracted cleanly' $false $_.Exception.Message
	$bgfStartText = $null
	$bgfEndText = $null
}

if ($bgfStartText -and $bgfEndText)
{
	$rtFile = Join-Path $testRoot 'tbs_selftest_h_roundtrip.ps1'
	$rtHead = @'
$ErrorActionPreference = 'Stop'
$script:FunctionResults = @{ PreExisting = 'kept' }
$script:FakeOutVar = $null
$script:ScriptsLog = $env:TEMP
$script:IsTestEnvironment = $false
function Write_Log { [CmdletBinding()] param([string]$Message, [string]$Color = 'Black', [switch]$IncludeTimestamp, [switch]$Raw) }
'@
	$rtTail = @'

function Test-FakeBgFunc
{
	[CmdletBinding()]
	param([int]$DelayMs = 0, [switch]$ShouldThrow, [string]$OutValue = '')
	if ($ShouldThrow) { throw 'intentional selftest failure' }
	Start-Sleep -Milliseconds $DelayMs
	$script:FunctionResults['FakeResult'] = ('hello-' + $OutValue)
	$script:FakeOutVar = $OutValue
}

$j = Start_Background_Function_Job -FunctionName 'Test-FakeBgFunc' `
	-Label 'selftest' `
	-CallArgs @{ DelayMs = 800; OutValue = 'xyz' } `
	-SeedVars @{ SomeVar = 'seeded' } `
	-ReturnVarNames @('FakeOutVar')
Write-Output ('RT_S_JOB_STARTED=' + ($j -is [System.Management.Automation.Job]))
$pollNow = $null
try { $pollNow = Get_Background_Function_Job_Result -Job $j } catch { $pollNow = 'THREW' }
Write-Output ('RT_S_POLL_WHILE_RUNNING=' + $(if ($pollNow -eq $null) { 'null' } else { 'early' }))
$result = $null
$deadline = (Get-Date).AddSeconds(60)
while ((Get-Date) -lt $deadline)
{
	Start-Sleep -Milliseconds 300
	$r = Get_Background_Function_Job_Result -Job $j
	if ($null -ne $r) { $result = $r; break }
}
Write-Output ('RT_S_RESULT_RECEIVED=' + ($null -ne $result))
if ($null -ne $result)
{
	Write-Output ('RT_S_SUCCESS=' + $result.Success)
	Write-Output ('RT_S_ERROR=' + [string]$result.Error)
	Write-Output ('RT_S_MERGED=' + [string]$script:FunctionResults['FakeResult'])
	Write-Output ('RT_S_KEPT=' + [string]$script:FunctionResults['PreExisting'])
	Write-Output ('RT_S_EXTRA=' + [string]$script:FakeOutVar)
	Write-Output ('RT_S_BGF_COUNT=' + $script:BgfJobs.Count)
	Write-Output ('RT_S_JOB_REMOVED=' + (-not (Get-Job -Id $j.Id -ErrorAction SilentlyContinue)))
}

$script:FunctionResults = @{}
$script:BgfJobs = @{}
$j2 = Start_Background_Function_Job -FunctionName 'Test-FakeBgFunc' `
	-Label 'selftest-fail' `
	-CallArgs @{ ShouldThrow = $true }
Write-Output ('RT_F_JOB_STARTED=' + ($j2 -is [System.Management.Automation.Job]))
$result2 = $null
$deadline2 = (Get-Date).AddSeconds(60)
while ((Get-Date) -lt $deadline2)
{
	Start-Sleep -Milliseconds 300
	$r2 = Get_Background_Function_Job_Result -Job $j2
	if ($null -ne $r2) { $result2 = $r2; break }
}
Write-Output ('RT_F_RESULT_RECEIVED=' + ($null -ne $result2))
if ($null -ne $result2)
{
	Write-Output ('RT_F_SUCCESS=' + $result2.Success)
	Write-Output ('RT_F_ERROR=' + [string]$result2.Error)
	Write-Output ('RT_F_JOB_REMOVED=' + (-not (Get-Job -Id $j2.Id -ErrorAction SilentlyContinue)))
}
Write-Output 'RT_DONE'
'@
	[System.IO.File]::WriteAllText($rtFile, ($rtHead + "`n" + $bgfStartText + "`n" + $bgfEndText + "`n" + $rtTail), [System.Text.Encoding]::ASCII)
	Check 'H: round-trip script written' (Test-Path -LiteralPath $rtFile) ''
	
	$rtLines = @()
	$rtOutFile = Join-Path $testRoot 'tbs_selftest_h_out.txt'
	try
	{
		$psExe = (Get-Process -Id $PID).Path
		$rtProc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $rtFile) `
			-RedirectStandardOutput $rtOutFile -NoNewWindow -Wait -PassThru
		$rtLines = @([System.IO.File]::ReadAllLines($rtOutFile) | Where-Object { $_.StartsWith('RT_') })
	}
	catch
	{
		Check 'H: round-trip child process ran' $false $_.Exception.Message
	}
	$rt = @{}
	foreach ($ln in $rtLines)
	{
		$ix = $ln.IndexOf('=')
		if ($ix -lt 1) { $rt[$ln] = '1' } else { $rt[$ln.Substring(0, $ix)] = $ln.Substring($ix + 1) }
	}
	
	Check 'H: round-trip completed (RT_DONE)' ($rt.ContainsKey('RT_DONE')) ('child exit: ' + $(if ($rtProc) { $rtProc.ExitCode } else { 'n/a' }))
	Check 'H: job started' ($rt['RT_S_JOB_STARTED'] -eq 'True') ''
	Check 'H: poll while running returns $null' ($rt['RT_S_POLL_WHILE_RUNNING'] -eq 'null') ('got: ' + $rt['RT_S_POLL_WHILE_RUNNING'])
	Check 'H: result received within 60s' ($rt['RT_S_RESULT_RECEIVED'] -eq 'True') ''
	Check 'H: job reports Success' ($rt['RT_S_SUCCESS'] -eq 'True') ('error: ' + $rt['RT_S_ERROR'])
	Check 'H: FunctionResults merged into main session' ($rt['RT_S_MERGED'] -ceq 'hello-xyz') ('got: ' + $rt['RT_S_MERGED'])
	Check 'H: pre-existing key not clobbered' ($rt['RT_S_KEPT'] -ceq 'kept') ('got: ' + $rt['RT_S_KEPT'])
	Check 'H: ExtraVar landed in main script scope' ($rt['RT_S_EXTRA'] -ceq 'xyz') ('got: ' + $rt['RT_S_EXTRA'])
	Check 'H: BgfJobs cleaned up (count 0)' ($rt['RT_S_BGF_COUNT'] -eq '0') ('count: ' + $rt['RT_S_BGF_COUNT'])
	Check 'H: job removed after result' ($rt['RT_S_JOB_REMOVED'] -eq 'True') ''
	Check 'H: failure-path result received' ($rt['RT_F_RESULT_RECEIVED'] -eq 'True') ''
	Check 'H: failure path reports Success=$false' ($rt['RT_F_SUCCESS'] -eq 'False') ''
	Check 'H: failure path carries the error message' ([string]$rt['RT_F_ERROR'] -like '*intentional selftest failure*') ('error: ' + $rt['RT_F_ERROR'])
	Check 'H: failure-path job removed' ($rt['RT_F_JOB_REMOVED'] -eq 'True') ''
	
	try { Remove-Item -LiteralPath $rtFile -Force -ErrorAction Stop } catch { }
	try { if (Test-Path -LiteralPath $rtOutFile) { Remove-Item -LiteralPath $rtOutFile -Force -ErrorAction Stop } } catch { }
}
else
{
	Skip 'H: job engine round trip' 'helper functions not found in script under test'
}

# --------------------------------------------------------------------------
# I) Real time-sync worker through the job engine (fake lane, no real touches)
# --------------------------------------------------------------------------
# Runs the REAL Send_TIME_Sync_Lanes_Worker as a real background job, using
# the extracted real helpers, the real mailslot-type loader (Add-Type must
# work inside the job runspace) and the real result merge. The lane list is
# ONE unmapped lane number, so the worker's loop logs "No machine name found"
# and exits without touching any real machine.
try
{
	$msText = Get-FunctionBlock 'Ensure_MailslotSender_Type'
	$wkText = Get-FunctionBlock 'Send_TIME_Sync_Lanes_Worker'
	$mkTok = $null; $mkErr = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($msText, [ref]$mkTok, [ref]$mkErr)
	$wkTok = $null; $wkErr = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($wkText, [ref]$wkTok, [ref]$wkErr)
	Check 'I: time-sync worker + mailslot loader extracted cleanly' ((@($mkErr).Count -eq 0) -and (@($wkErr).Count -eq 0)) ''
}
catch
{
	Check 'I: time-sync worker extracted cleanly' $false $_.Exception.Message
	$msText = $null
	$wkText = $null
}

if ($bgfStartText -and $bgfEndText -and $msText -and $wkText)
{
	$itFile = Join-Path $testRoot 'tbs_selftest_i_roundtrip.ps1'
	$itHead = @'
$ErrorActionPreference = 'Stop'
$script:FunctionResults = @{ LaneNumToMachineName = @{} }
$script:ScriptsLog = $env:TEMP
$script:IsTestEnvironment = $false
function Write_Log { [CmdletBinding()] param([string]$Message, [string]$Color = 'Black', [switch]$IncludeTimestamp, [switch]$Raw) }
'@
	$itTail = @'

$j = Start_Background_Function_Job -FunctionName 'Send_TIME_Sync_Lanes_Worker' `
	-Label 'selftest-time-sync' `
	-CallArgs @{ StoreNumber = '001'; SelectedLanes = @('999') } `
	-SeedVars @{ OfficePath = $null; PrimaryStoremanRoot = $null } `
	-DependentBodies @{ Ensure_MailslotSender_Type = (Get-Command Ensure_MailslotSender_Type -CommandType Function).ScriptBlock.ToString() }
Write-Output ('RT_I_JOB_STARTED=' + ($j -is [System.Management.Automation.Job]))
$result = $null
$deadline = (Get-Date).AddSeconds(90)
while ((Get-Date) -lt $deadline)
{
	Start-Sleep -Milliseconds 300
	$r = Get_Background_Function_Job_Result -Job $j
	if ($null -ne $r) { $result = $r; break }
}
Write-Output ('RT_I_RESULT_RECEIVED=' + ($null -ne $result))
if ($null -ne $result)
{
	Write-Output ('RT_I_SUCCESS=' + $result.Success)
	Write-Output ('RT_I_ERROR=' + [string]$result.Error)
}
Write-Output ('RT_I_JOB_REMOVED=' + (-not (Get-Job -Id $j.Id -ErrorAction SilentlyContinue)))
Write-Output 'RT_DONE'
'@
	[System.IO.File]::WriteAllText($itFile, ($itHead + "`n" + $bgfStartText + "`n" + $bgfEndText + "`n" + $msText + "`n" + $wkText + "`n" + $itTail), [System.Text.Encoding]::ASCII)
	
	$itOutFile = Join-Path $testRoot 'tbs_selftest_i_out.txt'
	$itLines = @()
	$itProc = $null
	try
	{
		$psExe = (Get-Process -Id $PID).Path
		$itProc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $itFile) `
			-RedirectStandardOutput $itOutFile -NoNewWindow -Wait -PassThru
		$itLines = @([System.IO.File]::ReadAllLines($itOutFile) | Where-Object { $_.StartsWith('RT_') })
	}
	catch
	{
		Check 'I: time-sync job round trip ran' $false $_.Exception.Message
	}
	$it = @{}
	foreach ($ln in $itLines)
	{
		$ix = $ln.IndexOf('=')
		if ($ix -lt 1) { $it[$ln] = '1' } else { $it[$ln.Substring(0, $ix)] = $ln.Substring($ix + 1) }
	}
	Check 'I: round trip completed (RT_DONE)' ($it.ContainsKey('RT_DONE')) ('child exit: ' + $(if ($itProc) { $itProc.ExitCode } else { 'n/a' }))
	Check 'I: real worker job started' ($it['RT_I_JOB_STARTED'] -eq 'True') ''
	Check 'I: result received within 90s' ($it['RT_I_RESULT_RECEIVED'] -eq 'True') ''
	Check 'I: worker reports Success' ($it['RT_I_SUCCESS'] -eq 'True') ('error: ' + $it['RT_I_ERROR'])
	Check 'I: job removed after result' ($it['RT_I_JOB_REMOVED'] -eq 'True') ''
	
	try { Remove-Item -LiteralPath $itFile -Force -ErrorAction Stop } catch { }
	try { if (Test-Path -LiteralPath $itOutFile) { Remove-Item -LiteralPath $itOutFile -Force -ErrorAction Stop } } catch { }
}
else
{
	Skip 'I: time-sync job round trip' 'required functions not found in script under test'
}

# --------------------------------------------------------------------------
# J) Parallel probe + ping function (1 reachable + 2 unreachable)
# --------------------------------------------------------------------------
# Tests the REAL Test-Hosts_Parallel and the REAL Ping_All_Nodes in a child
# process (no job engine - ping runs on the UI thread, so in the real app its
# results go to the log box). Lane data is seeded with THREE fake lanes: the
# local hostname (always reachable) and two RFC-2606 .invalid names (instant
# NXDOMAIN, never real hosts). This functionally tests the parallel probe
# path without touching any real machine.
try
{
	$thpText = Get-FunctionBlock 'Test-Hosts_Parallel'
	$pnText = Get-FunctionBlock 'Ping_All_Nodes'
	$tpTok = $null; $tpErr = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($thpText, [ref]$tpTok, [ref]$tpErr)
	$pnTok = $null; $pnErr = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($pnText, [ref]$pnTok, [ref]$pnErr)
	Check 'J: parallel probe + ping function extracted cleanly' ((@($tpErr).Count -eq 0) -and (@($pnErr).Count -eq 0)) ''
}
catch
{
	Check 'J: parallel probe + ping function extracted cleanly' $false $_.Exception.Message
	$thpText = $null
	$pnText = $null
}

if ($thpText -and $pnText)
{
	$jFile = Join-Path $testRoot 'tbs_selftest_j_direct.ps1'
	$jHead = @'
$ErrorActionPreference = 'Stop'
$script:FunctionResults = @{
	LaneNumToMachineName = @{
		'001' = $env:COMPUTERNAME
		'002' = 'tbsst1.invalid'
		'003' = 'tbsst2.invalid'
	}
}
$script:ScriptsLog = $env:TEMP
$script:IsTestEnvironment = $false
function Write_Log { [CmdletBinding()] param([string]$Message, [string]$Color = 'Black', [switch]$IncludeTimestamp, [switch]$Raw) }
'@
	$jTail = @'

# --- direct probe test: 1 reachable (local) + 2 unreachable (.invalid) ---
$probe = Test-Hosts_Parallel -Targets @($env:COMPUTERNAME, 'tbsst1.invalid', 'tbsst2.invalid')
Write-Output ('RT_J_PROBE_LOCAL=' + $probe[$env:COMPUTERNAME])
Write-Output ('RT_J_PROBE_BAD1=' + $probe['tbsst1.invalid'])
Write-Output ('RT_J_PROBE_BAD2=' + $probe['tbsst2.invalid'])
Write-Output ('RT_J_PROBE_COUNT=' + $probe.Count)

# --- full Ping_All_Nodes run (must not throw) ---
try
{
	Ping_All_Nodes -NodeType 'Lane' -StoreNumber '001'
	Write-Output 'RT_J_PING_RAN=1'
}
catch
{
	Write-Output ('RT_J_PING_RAN=0 ERR=' + $_.Exception.Message)
}
Write-Output 'RT_DONE'
'@
	[System.IO.File]::WriteAllText($jFile, ($jHead + "`n" + $thpText + "`n" + $pnText + "`n" + $jTail), [System.Text.Encoding]::ASCII)
	
	$jOutFile = Join-Path $testRoot 'tbs_selftest_j_out.txt'
	$jLines = @()
	$jProc = $null
	try
	{
		$psExe = (Get-Process -Id $PID).Path
		$jProc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $jFile) `
			-RedirectStandardOutput $jOutFile -NoNewWindow -Wait -PassThru
		$jLines = @([System.IO.File]::ReadAllLines($jOutFile) | Where-Object { $_.StartsWith('RT_') })
	}
	catch
	{
		Check 'J: direct ping round trip ran' $false $_.Exception.Message
	}
	$jv = @{}
	foreach ($ln in $jLines)
	{
		$ix = $ln.IndexOf('=')
		if ($ix -lt 1) { $jv[$ln] = '1' } else { $jv[$ln.Substring(0, $ix)] = $ln.Substring($ix + 1) }
	}
	Check 'J: round trip completed (RT_DONE)' ($jv.ContainsKey('RT_DONE')) ('child exit: ' + $(if ($jProc) { $jProc.ExitCode } else { 'n/a' }))
	Check 'J: probe returned all 3 targets' ($jv['RT_J_PROBE_COUNT'] -eq '3') ('count: ' + $jv['RT_J_PROBE_COUNT'])
	Check 'J: local host probed reachable (OK)' ($jv['RT_J_PROBE_LOCAL'] -eq 'OK') ('state: ' + $jv['RT_J_PROBE_LOCAL'])
	Check 'J: bogus host 1 not reachable' (($jv['RT_J_PROBE_BAD1'] -eq 'FAIL') -or ($jv['RT_J_PROBE_BAD1'] -eq 'ERROR')) ('state: ' + $jv['RT_J_PROBE_BAD1'])
	Check 'J: bogus host 2 not reachable' (($jv['RT_J_PROBE_BAD2'] -eq 'FAIL') -or ($jv['RT_J_PROBE_BAD2'] -eq 'ERROR')) ('state: ' + $jv['RT_J_PROBE_BAD2'])
	Check 'J: Ping_All_Nodes ran without error' ($jv['RT_J_PING_RAN'] -like '1*') ('' + $jv['RT_J_PING_RAN'])
	
	try { Remove-Item -LiteralPath $jFile -Force -ErrorAction Stop } catch { }
	try { if (Test-Path -LiteralPath $jOutFile) { Remove-Item -LiteralPath $jOutFile -Force -ErrorAction Stop } } catch { }
}
else
{
	Skip 'J: parallel probe + ping function' 'required functions not found in script under test'
}

# --------------------------------------------------------------------------
# K) Shared job reaper: success notices file-log-only, failures still in box
# --------------------------------------------------------------------------
# Runs the REAL job engine in a child process WITH a WinForms message pump
# (hidden form + Application.Run) so the REAL shared reaper (a WinForms timer)
# actually ticks - unlike sections H/I, which poll manually. Two jobs are
# started: one succeeds, one throws. A fake Write_Log records every call with
# its -FileOnly flag, and this test verifies the reaper's logging behavior:
#   - the "finished in Xs" success notice is recorded WITH FileOnly (file log
#     only - the box stays quiet while the status bar shows progress)
#   - the "failed" notice is recorded WITHOUT it (box + file, so the user
#     still sees failures)
if ($bgfStartText -and $bgfEndText)
{
	$kFile = Join-Path $testRoot 'tbs_selftest_k_reaper.ps1'
	$kHead = @'
$ErrorActionPreference = 'Stop'
$script:FunctionResults = @{}
$script:ScriptsLog = $env:TEMP
$script:IsTestEnvironment = $false
$script:LogCalls = @()
function Write_Log { [CmdletBinding()] param([Parameter(Mandatory = $true)][string]$Message, [string]$Color = 'Black', [switch]$IncludeTimestamp, [switch]$Raw, [switch]$FileOnly) $script:LogCalls += [pscustomobject]@{ Msg = $Message; Color = [string]$Color; FileOnly = [bool]$FileOnly } }
'@
	$kTail = @'

function Test-KGood
{
	[CmdletBinding()]
	param([int]$DelayMs = 0)
	Start-Sleep -Milliseconds $DelayMs
	$script:FunctionResults['KGood'] = 'ok'
}
function Test-KBad
{
	[CmdletBinding()]
	param()
	throw 'intentional reaper test failure'
}

# Hidden form so the WinForms timer (the reaper) has a message loop to tick on.
try
{
	Add-Type -AssemblyName System.Windows.Forms
	$pumpForm = New-Object System.Windows.Forms.Form
	$pumpForm.ShowInTaskbar = $false
	$pumpForm.Visible = $false
	$pumpTimer = New-Object System.Windows.Forms.Timer
	$pumpTimer.Interval = 12000
	$pumpTimer.Add_Tick({ $pumpForm.Close() })
	$pumpTimer.Start()
}
catch
{
	Write-Output ('RT_K_PUMP_SETUP=0 ERR=' + $_.Exception.Message)
	exit 1
}

# Both jobs start BEFORE the pump runs; the first call creates the reaper.
$j1 = Start_Background_Function_Job -FunctionName 'Test-KGood' -Label 'k-good' -CallArgs @{ DelayMs = 600 }
$j2 = Start_Background_Function_Job -FunctionName 'Test-KBad' -Label 'k-bad'
Write-Output ('RT_K_JOBS_STARTED=' + (($j1 -is [System.Management.Automation.Job]) -and ($j2 -is [System.Management.Automation.Job])))

# Pump for ~12s: the reaper (1s tick) picks up both finished jobs in this window.
[System.Windows.Forms.Application]::Run($pumpForm)
Write-Output 'RT_K_PUMP_DONE=1'

$goodLines = @($script:LogCalls | Where-Object { $_.Msg -like 'Background job *k-good*finished*' })
$badLines  = @($script:LogCalls | Where-Object { $_.Msg -like 'Background job *k-bad*failed*' })
Write-Output ('RT_K_GOOD_LINE_COUNT=' + $goodLines.Count)
if ($goodLines.Count -gt 0) { Write-Output ('RT_K_GOOD_FILEONLY=' + $goodLines[0].FileOnly) }
Write-Output ('RT_K_BAD_LINE_COUNT=' + $badLines.Count)
if ($badLines.Count -gt 0) { Write-Output ('RT_K_BAD_FILEONLY=' + $badLines[0].FileOnly) }
Write-Output 'RT_DONE'
'@
	[System.IO.File]::WriteAllText($kFile, ($kHead + "`n" + $bgfStartText + "`n" + $bgfEndText + "`n" + $kTail), [System.Text.Encoding]::ASCII)
	Check 'K: reaper round-trip script written' (Test-Path -LiteralPath $kFile) ''

	$kOutFile = Join-Path $testRoot 'tbs_selftest_k_out.txt'
	$kLines = @()
	$kProc = $null
	try
	{
		$psExe = (Get-Process -Id $PID).Path
		$kProc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $kFile) `
			-RedirectStandardOutput $kOutFile -NoNewWindow -Wait -PassThru
		$kLines = @([System.IO.File]::ReadAllLines($kOutFile) | Where-Object { $_.StartsWith('RT_') })
	}
	catch
	{
		Check 'K: reaper round trip ran' $false $_.Exception.Message
	}
	$kv = @{}
	foreach ($ln in $kLines)
	{
		$ix = $ln.IndexOf('=')
		if ($ix -lt 1) { $kv[$ln] = '1' } else { $kv[$ln.Substring(0, $ix)] = $ln.Substring($ix + 1) }
	}
	Check 'K: round trip completed (RT_DONE)' ($kv.ContainsKey('RT_DONE')) ('child exit: ' + $(if ($kProc) { $kProc.ExitCode } else { 'n/a' }))
	Check 'K: both reaper jobs started' ($kv['RT_K_JOBS_STARTED'] -eq 'True') ''
	Check 'K: message pump ran' ($kv['RT_K_PUMP_DONE'] -eq '1') ('lines: ' + ($kLines -join ' | '))
	Check 'K: success notice recorded' ($kv['RT_K_GOOD_LINE_COUNT'] -eq '1') ('count: ' + $kv['RT_K_GOOD_LINE_COUNT'])
	Check 'K: success notice is FILE-ONLY (box stays quiet)' ($kv['RT_K_GOOD_FILEONLY'] -eq 'True') ('fileonly: ' + $kv['RT_K_GOOD_FILEONLY'])
	Check 'K: failure notice recorded' ($kv['RT_K_BAD_LINE_COUNT'] -eq '1') ('count: ' + $kv['RT_K_BAD_LINE_COUNT'])
	Check 'K: failure notice still goes to the box' ($kv['RT_K_BAD_FILEONLY'] -eq 'False') ('fileonly: ' + $kv['RT_K_BAD_FILEONLY'])

	try { Remove-Item -LiteralPath $kFile -Force -ErrorAction Stop } catch { }
	try { if (Test-Path -LiteralPath $kOutFile) { Remove-Item -LiteralPath $kOutFile -Force -ErrorAction Stop } } catch { }
}
else
{
	Skip 'K: shared reaper logging behavior' 'job engine functions not found in script under test'
}

# --------------------------------------------------------------------------
# L) Lane SQL protocol loop (v2.7.0 fix): the real Start_Lane_Protocol_Jobs
# --------------------------------------------------------------------------
# v2.6.7 - v2.6.9 ran the protocol timer as a .GetNewClosure() closure, where
# every $script: variable is a separate empty one, so no lane was ever probed.
# Static: no closures anywhere. Functional: the real function in a child
# process with a WinForms message pump and two RFC-2606 .invalid lanes (no
# real machine is touched) must probe both, record them Offline under every
# key, and persist Protocol_Results.txt.
$codeLines = @($scriptLines | Where-Object { $_.Trim() -notmatch '^#' })
Check 'L: no .GetNewClosure() in live code (closures break $script: state)' (-not (($codeLines -join "`n") -match '\.GetNewClosure\(\)')) ''

$lpText = $null
try
{
	$lpText = Get-FunctionBlock 'Start_Lane_Protocol_Jobs'
	$lpTok = $null; $lpErr = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($lpText, [ref]$lpTok, [ref]$lpErr)
	Check 'L: Start_Lane_Protocol_Jobs extracted cleanly' (@($lpErr).Count -eq 0) ''
}
catch
{
	Check 'L: Start_Lane_Protocol_Jobs extracted cleanly' $false $_.Exception.Message
	$lpText = $null
}

if ($lpText)
{
	$lDir = Join-Path $testRoot 'L_SetupFiles'
	New-Item -ItemType Directory -Path $lDir -Force | Out-Null
	$lFile = Join-Path $testRoot 'tbs_selftest_l_protocol.ps1'
	$lHead = @'
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
function Write_Log { [CmdletBinding()] param([string]$Message, [string]$Color = 'Black', [switch]$IncludeTimestamp, [switch]$Raw, [switch]$FileOnly) }
'@
	$lTail = @'

$script:SetupFiles = $args[0]
Start_Lane_Protocol_Jobs -LaneNumToMachineName @{ '001' = 'tbsst-l1.invalid'; '002' = 'tbsst-l2.invalid' } -SqlModuleName ''
Write-Output ('RT_L_TIMER=' + [bool]$script:protocolTimer)
$pumpForm = New-Object System.Windows.Forms.Form
$pumpForm.ShowInTaskbar = $false
$pumpForm.WindowState = 'Minimized'
$pumpTimer = New-Object System.Windows.Forms.Timer
$pumpTimer.Interval = 12000
$pumpTimer.Add_Tick({ $pumpTimer.Stop(); $pumpForm.Close() })
$pumpTimer.Start()
[System.Windows.Forms.Application]::Run($pumpForm)
Write-Output ('RT_L_001=' + [string]$script:LaneProtocols['001'])
Write-Output ('RT_L_002=' + [string]$script:LaneProtocols['002'])
Write-Output ('RT_L_BYNAME=' + [string]$script:LaneProtocols['tbsst-l1.invalid'])
$pf = Join-Path $script:SetupFiles 'Protocol_Results.txt'
Write-Output ('RT_L_FILE=' + ((@(Get-Content -LiteralPath $pf -ErrorAction SilentlyContinue) | Where-Object { $_.Trim() }) -join ';'))
Write-Output 'RT_DONE'
'@
	[System.IO.File]::WriteAllText($lFile, ($lHead + "`n" + $lpText + "`n" + $lTail), [System.Text.Encoding]::ASCII)

	$lOutFile = Join-Path $testRoot 'tbs_selftest_l_out.txt'
	$lv = @{}
	$lProc = $null
	try
	{
		$psExe = (Get-Process -Id $PID).Path
		$lProc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', $lFile, $lDir) `
			-RedirectStandardOutput $lOutFile -NoNewWindow -Wait -PassThru
		foreach ($ln in @([System.IO.File]::ReadAllLines($lOutFile) | Where-Object { $_.StartsWith('RT_') }))
		{
			$ix = $ln.IndexOf('=')
			if ($ix -lt 1) { $lv[$ln] = '1' } else { $lv[$ln.Substring(0, $ix)] = $ln.Substring($ix + 1) }
		}
	}
	catch
	{
		Check 'L: protocol loop child process ran' $false $_.Exception.Message
	}
	Check 'L: round trip completed (RT_DONE)' ($lv.ContainsKey('RT_DONE')) ('child exit: ' + $(if ($lProc) { $lProc.ExitCode } else { 'n/a' }))
	Check 'L: protocol timer started' ($lv['RT_L_TIMER'] -eq 'True') ''
	Check 'L: lane 001 probed (Offline)' ($lv['RT_L_001'] -eq 'Offline') ('got: ' + $lv['RT_L_001'])
	Check 'L: lane 002 probed (Offline)' ($lv['RT_L_002'] -eq 'Offline') ('got: ' + $lv['RT_L_002'])
	Check 'L: result also stored by machine name' ($lv['RT_L_BYNAME'] -eq 'Offline') ('got: ' + $lv['RT_L_BYNAME'])
	Check 'L: Protocol_Results.txt written' ($lv['RT_L_FILE'] -eq '001,Offline;002,Offline') ('got: ' + $lv['RT_L_FILE'])

	try { Remove-Item -LiteralPath $lFile -Force -ErrorAction Stop } catch { }
	try { if (Test-Path -LiteralPath $lOutFile) { Remove-Item -LiteralPath $lOutFile -Force -ErrorAction Stop } } catch { }
}
else
{
	Skip 'L: lane protocol loop' 'Start_Lane_Protocol_Jobs not found in script under test'
}

# --------------------------------------------------------------------------
# 4) Summary
# --------------------------------------------------------------------------
Write-Host ''
Write-Host ("RESULT: {0} passed, {1} failed, {2} skipped" -f $pass, $fail, $skip)
if ($fail -gt 0)
{
	Write-Host 'SelfTest FAILED' -ForegroundColor Red
	exit 1
}
else
{
	Write-Host 'SelfTest PASSED' -ForegroundColor Green
	exit 0
}
