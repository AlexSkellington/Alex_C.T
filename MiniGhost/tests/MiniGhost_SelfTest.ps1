# ==============================================================================
#  MiniGhost_SelfTest.ps1
# ------------------------------------------------------------------------------
#  Regression test for MiniGhost.
#
#  How it works:
#    Static checks (parse, ASCII, version vs changelog) run on the built
#    MiniGhost\MiniGhost.ps1. Then every function is loaded VERBATIM from that
#    file (by AST - the script's own top-level code never runs) and the ones that
#    do not need a store (name rules, INI rewrites, folder / temp cleanup, the
#    database lookup) are run against files in %TEMP%.
#
#  Usage (from the repository root, after MiniGhost\build\Build.ps1):
#    powershell -NoProfile -ExecutionPolicy Bypass -File .\MiniGhost\tests\MiniGhost_SelfTest.ps1
#    ... -ScriptPath C:\path\to\MiniGhost.ps1
#
#  Exit codes:  0 = all passed (skips allowed)   1 = one or more failed   2 = setup error
#
#  Notes:
#    - Never renames the computer and never touches storeman, SQL tables or the
#      registry: everything it writes is under %TEMP%\mg-selftest-*.
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

function Write-TestFile([string]$Path, [string[]]$Lines)
{
	[System.IO.File]::WriteAllText($Path, (($Lines -join "`r`n") + "`r`n"), [System.Text.Encoding]::ASCII)
}

function Read-IniValue([string]$Path, [string]$Key)
{
	foreach ($l in [System.IO.File]::ReadAllLines($Path))
	{
		if ($l -match ('^\s*' + [regex]::Escape($Key) + '\s*=\s*(.*)$')) { return $Matches[1].Trim() }
	}
	return $null
}

# --------------------------------------------------------------------------
# 1) Locate the script under test
# --------------------------------------------------------------------------
if (-not $ScriptPath -and $PSScriptRoot)
{
	$ScriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'MiniGhost.ps1'
}
if (-not $ScriptPath -or -not (Test-Path -LiteralPath $ScriptPath))
{
	Write-Host 'Could not find MiniGhost.ps1. Run MiniGhost\build\Build.ps1 or pass -ScriptPath <file>.' -ForegroundColor Red
	exit 2
}
$ScriptPath = (Resolve-Path -LiteralPath $ScriptPath).ProviderPath
Write-Host ("Testing: " + $ScriptPath) -ForegroundColor Cyan
Write-Host ''

$rawText = [System.IO.File]::ReadAllText($ScriptPath)

# --------------------------------------------------------------------------
# 2) Static checks
# --------------------------------------------------------------------------
Write-Host 'Static checks' -ForegroundColor Cyan

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($rawText, [ref]$tokens, [ref]$parseErrors)
Check 'script parses without errors' (@($parseErrors).Count -eq 0) ((@($parseErrors) | Select-Object -First 3 | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Message)" }) -join '; ')

$nonAscii = [regex]::Matches($rawText, '[^\u0000-\u007F]').Count
Check 'script is pure ASCII (PowerShell 5.1 reads BOM-less files as Windows-1252)' ($nonAscii -eq 0) "$nonAscii non-ASCII character(s)"

Check 'no .GetNewClosure() (closures get their own empty $script: scope)' ($rawText -notmatch '\.GetNewClosure\(')

$version = [regex]::Match($rawText, '(?m)^\$VersionNumber\s*=\s*"([\d.]+)"').Groups[1].Value
$changelogTop = [regex]::Match($rawText, '(?m)^#\s+v([\d.]+)\s+-\s+\d{4}-\d{2}-\d{2}').Groups[1].Value
Check "version $version matches the newest changelog entry" ($version -and $version -eq $changelogTop) "changelog top: $changelogTop"

$functionAsts = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })
$names = @($functionAsts | ForEach-Object { $_.Name })
$expected = @(
	'Invoke_Elevation_Bootstrap', 'Get_Database_Connection_String', 'Ensure_Database_Connection_Context', 'Execute_SQL_Commands',
	'Get_Machine_Name_Context', 'Update_INIs', 'Update_Startup_Ini_For_Machine', 'Invoke_Machine_Name_Sync',
	'Update_SQL_Tables_For_Machine_Name_Change', 'Update_SQL_Tables_For_Store_Number_Change', 'Remove_Old_XZ_Folders',
	'Remove_MiniGhost_Temp_Copies', 'Show_Action_Prompt_With_Log', 'Get_Active_IP_Config'
)
$missing = @($expected | Where-Object { $names -notcontains $_ })
Check 'all core functions are present' ($missing.Count -eq 0) ("missing: " + ($missing -join ', '))
Check 'no function defined twice' (@($names | Group-Object | Where-Object { $_.Count -gt 1 }).Count -eq 0)

# --------------------------------------------------------------------------
# 3) Load every function verbatim (top-level code is NOT run)
# --------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
foreach ($fa in $functionAsts)
{
	. ([scriptblock]::Create($fa.Extent.Text))
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('mg-selftest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try
{
	# ----------------------------------------------------------------------
	Write-Host ''
	Write-Host 'Subnet helpers' -ForegroundColor Cyan
	Check 'mask 255.255.255.0 -> /24' ((Convert_Subnet_Mask_To_Prefix_Length -SubnetMask '255.255.255.0') -eq 24)
	Check 'mask 255.255.0.0 -> /16' ((Convert_Subnet_Mask_To_Prefix_Length -SubnetMask '255.255.0.0') -eq 16)
	Check 'mask 255.255.255.128 -> /25' ((Convert_Subnet_Mask_To_Prefix_Length -SubnetMask '255.255.255.128') -eq 25)
	Check 'non-contiguous mask -> $null' ($null -eq (Convert_Subnet_Mask_To_Prefix_Length -SubnetMask '255.0.255.0'))
	Check '/24 -> 255.255.255.0' ((Convert_Prefix_Length_To_Subnet_Mask -PrefixLength 24) -eq '255.255.255.0')
	Check '/25 -> 255.255.255.128' ((Convert_Prefix_Length_To_Subnet_Mask -PrefixLength 25) -eq '255.255.255.128')
	Check '/33 -> $null' ($null -eq (Convert_Prefix_Length_To_Subnet_Mask -PrefixLength 33))

	# ----------------------------------------------------------------------
	Write-Host ''
	Write-Host 'Machine name rules (Get_Machine_Name_Context)' -ForegroundColor Cyan
	$script:FunctionResults = @{ StoreNumber = '231' }
	$script:NodeRole = 'Lane'
	$c = Get_Machine_Name_Context -MachineName 'pos006' -CurrentMachineName 'POS005'
	Check 'lane POS006 -> terminal 006' ($c.Success -and $c.MachineName -eq 'POS006' -and $c.MachineNumber -eq '006' -and -not $c.IsServerMoniker)
	$c = Get_Machine_Name_Context -MachineName '007' -CurrentMachineName '0231POS-005'
	Check 'terminal-only 007 keeps the current style (0231POS-005 -> 0231POS-007)' ($c.Success -and $c.MachineName -eq '0231POS-007' -and $c.InputWasTerminalOnly)
	$c = Get_Machine_Name_Context -MachineName 'POS000' -CurrentMachineName 'POS005'
	Check 'terminal 000 is rejected' (-not $c.Success)
	$c = Get_Machine_Name_Context -MachineName 'ABCDEFGHIJKLMNO001' -CurrentMachineName 'POS005'
	Check 'names over 15 characters are rejected' (-not $c.Success)
	$script:NodeRole = 'StoreServer'
	$c = Get_Machine_Name_Context -MachineName '0231SERVER001' -CurrentMachineName 'SERVER001'
	Check 'server 0231SERVER001 -> terminal 901, store prefix 0231' ($c.Success -and $c.IsServerMoniker -and $c.MachineNumber -eq '901' -and $c.StorePrefix -eq '0231')
	$c = Get_Machine_Name_Context -MachineName '001' -CurrentMachineName 'SERVER001'
	Check 'servers reject terminal-only input' (-not $c.Success)

	# ----------------------------------------------------------------------
	Write-Host ''
	Write-Host 'INI rewrite on a store number change (Update_INIs)' -ForegroundColor Cyan
	$script:NodeRole = 'Lane'
	$storeman = Join-Path $tempRoot 'storeman'
	New-Item -ItemType Directory -Path $storeman -Force | Out-Null
	$startupIni = Join-Path $storeman 'Startup.ini'
	$smsStartIni = Join-Path $storeman 'SMSStart.ini'
	Write-TestFile $startupIni @('[STARTUP]', 'STORE=231', 'TER=005', 'REDIRMAIL=231901', 'REDIRMSG=231901', 'SERVERNAME=231server', 'SERVERIP=10.0.231.5', 'DBSERVER=POS005', 'DBNAME=LANESQL')
	Write-TestFile $smsStartIni @('[SMSSTART]', 'STORE=231', 'TER=005', '[TIMESYNC]', 'SERVER=192.168.231.10', 'ENABLE=0')

	$ok = Update_INIs -newStoreNumber '242' -MachineName 'POS006' -OldStoreNumber '231' -StartupIniPath $startupIni -GlobalSmsStartIniPath $smsStartIni
	Check 'Update_INIs reports success' ($ok -eq $true)
	Check 'Startup.ini STORE=242' ((Read-IniValue $startupIni 'STORE') -eq '242')
	Check 'Startup.ini TER=006 (from the machine name)' ((Read-IniValue $startupIni 'TER') -eq '006')
	Check 'Startup.ini REDIRMAIL=242901' ((Read-IniValue $startupIni 'REDIRMAIL') -eq '242901')
	Check 'Startup.ini SERVERNAME store prefix replaced and uppercased' ((Read-IniValue $startupIni 'SERVERNAME') -eq '242SERVER')
	Check 'an IP octet equal to the old store number is NOT rewritten' ((Read-IniValue $startupIni 'SERVERIP') -eq '10.0.231.5') ("got " + (Read-IniValue $startupIni 'SERVERIP'))
	Check 'Startup.ini keeps CRLF line endings' ([System.IO.File]::ReadAllText($startupIni) -match "STORE=242`r`n")
	Check 'SMSStart.ini [SMSSTART] STORE=242 / TER=006' ((Read-IniValue $smsStartIni 'STORE') -eq '242' -and (Read-IniValue $smsStartIni 'TER') -eq '006')
	Check 'SMSStart.ini [TIMESYNC] SERVER= IP address is left alone' ((Read-IniValue $smsStartIni 'SERVER') -eq '192.168.231.10') ("got " + (Read-IniValue $smsStartIni 'SERVER'))
	Check 'SMSStart.ini [TIMESYNC] ENABLE=1' ((Read-IniValue $smsStartIni 'ENABLE') -eq '1')

	Write-TestFile $smsStartIni @('[TIMESYNC]', 'SERVER=0231SERVER', 'ENABLE=1')
	$null = Update_INIs -newStoreNumber '242' -MachineName 'POS006' -OldStoreNumber '231' -StartupIniPath $startupIni -GlobalSmsStartIniPath $smsStartIni
	Check 'SMSStart.ini [TIMESYNC] SERVER=0231SERVER -> 0242SERVER' ((Read-IniValue $smsStartIni 'SERVER') -eq '0242SERVER') ("got " + (Read-IniValue $smsStartIni 'SERVER'))

	# ----------------------------------------------------------------------
	Write-Host ''
	Write-Host 'Database lookup without an explicit Startup.ini path' -ForegroundColor Cyan
	$dbIni = Join-Path $tempRoot 'Startup_db.ini'
	Write-TestFile $dbIni @('DBSERVER=localhost', 'DBNAME=STORESQL')
	$StartupIniPath = $dbIni
	$script:FunctionResults = @{ }
	$cs = Ensure_Database_Connection_Context
	Check 'reads DBNAME from the script-level Startup.ini' ($script:FunctionResults['DBNAME'] -eq 'STORESQL')
	Check 'builds a connection string (was empty: "Database connection string is not available")' ($cs -match 'Database=STORESQL')

	# ----------------------------------------------------------------------
	Write-Host ''
	Write-Host 'Old X*/Z* folder cleanup (Remove_Old_XZ_Folders)' -ForegroundColor Cyan
	$office = Join-Path $tempRoot 'office'
	foreach ($d in @('XF0242006', 'XF0242901', 'ZW0242900', 'XF0231006', 'XF0242005', 'XW0231901', 'XNOTAFOLDER'))
	{
		New-Item -ItemType Directory -Path (Join-Path $office $d) -Force | Out-Null
	}
	Remove_Old_XZ_Folders -StoreNumber '242' -MachineName 'POS006' -OfficePath $office 6>$null | Out-Null
	$left = @(Get-ChildItem -LiteralPath $office -Directory | ForEach-Object { $_.Name } | Sort-Object)
	Check 'keeps this lane, 900 and 901 of the new store (and unrelated names)' (($left -join ',') -eq 'XF0242006,XF0242901,XNOTAFOLDER,ZW0242900') ("left: " + ($left -join ','))
	New-Item -ItemType Directory -Path (Join-Path $office 'XF0231005') -Force | Out-Null
	Remove_Old_XZ_Folders -StoreNumber '242' -MachineName 'SERVER901' -OfficePath $office 6>$null | Out-Null
	Check 'refuses to run for terminal 901 (backoffice safeguard)' (Test-Path -LiteralPath (Join-Path $office 'XF0231005'))

	# ----------------------------------------------------------------------
	Write-Host ''
	Write-Host 'Temp cleanup (Remove_MiniGhost_Temp_Copies)' -ForegroundColor Cyan
	$tmp = Join-Path $tempRoot 'tmp'
	New-Item -ItemType Directory -Path (Join-Path $tmp 'sub') -Force | Out-Null
	foreach ($f in @('MiniGhost.ps1', 'MiniGhost_1.3.9_Elevated.ps1', 'MiniGhostNotes.ps1', 'sub\MiniGhost.ps1')) { Write-TestFile (Join-Path $tmp $f) @('#') }
	$deleted = Remove_MiniGhost_Temp_Copies -TempPath $tmp
	Check 'deletes MiniGhost.ps1 and the _Elevated copy' ($deleted -eq 2 -and -not (Test-Path (Join-Path $tmp 'MiniGhost.ps1')) -and -not (Test-Path (Join-Path $tmp 'MiniGhost_1.3.9_Elevated.ps1')))
	Check 'leaves other files and subfolders alone' ((Test-Path (Join-Path $tmp 'MiniGhostNotes.ps1')) -and (Test-Path (Join-Path $tmp 'sub\MiniGhost.ps1')))

	# ----------------------------------------------------------------------
	Write-Host ''
	Write-Host 'Dialog result (Show_Action_Prompt_With_Log)' -ForegroundColor Cyan
	if (-not [Environment]::UserInteractive)
	{
		Skip 'returns the clicked button' 'not an interactive session'
	}
	else
	{
		# A timer clicks "Later" once the dialog is up; a second safety net closes it.
		$script:mgTicks = 0
		$timer = New-Object System.Windows.Forms.Timer
		$timer.Interval = 300
		$timer.Add_Tick({
				$script:mgTicks++
				foreach ($f in @([System.Windows.Forms.Application]::OpenForms))
				{
					if ($f.Text -ne 'MG-SelfTest') { continue }
					foreach ($c in @($f.Controls))
					{
						foreach ($b in @($c.Controls)) { if ($b.Text -eq 'Later') { $b.PerformClick(); return } }
					}
					if ($script:mgTicks -gt 20) { $f.Close() }
				}
			})
		$timer.Start()
		try
		{
			$answer = Show_Action_Prompt_With_Log -Text 'self-test' -Title 'MG-SelfTest' -PrimaryButtonText 'Reboot Now' -SecondaryButtonText 'Later' -LogText 'log'
			Check 'returns the clicked button (was always "Closed")' ($answer -eq 'Secondary') "got '$answer'"
		}
		catch
		{
			Skip 'returns the clicked button' ("cannot show a dialog here: " + $_.Exception.Message)
		}
		finally
		{
			$timer.Stop()
			$timer.Dispose()
		}
	}
}
finally
{
	Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host ("Result: {0} passed, {1} failed, {2} skipped" -f $pass, $fail, $skip) -ForegroundColor $(if ($fail) { 'Red' } else { 'Green' })
if ($fail -gt 0) { exit 1 }
exit 0
