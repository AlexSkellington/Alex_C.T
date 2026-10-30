# ==============================================================================
#  Pump_Upsert_Test.ps1
# ------------------------------------------------------------------------------
#  Runs the REAL lane upsert worker of Pump_Tables (extracted verbatim from the
#  built script) against two SQL Server LocalDB databases - "server" and "lane" -
#  and checks that a pump only updates and adds rows:
#    - a changed row is updated, a new row is added, an unchanged row stays
#    - rows that only the lane has are kept, lane-only columns are kept
#    - a table missing on the lane is created (with the server's primary key)
#    - a table without any primary key is refused and left unchanged
#  Skips (exit 0) when SqlLocalDB is not installed.
#
#  Usage: powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\tests\Pump_Upsert_Test.ps1
#  Exit codes: 0 = passed or skipped, 1 = failed, 2 = setup error
# ==============================================================================
param ([string]$ScriptPath = '')

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\TBS_Maintenance_Script.ps1' }
if (-not (Test-Path -LiteralPath $ScriptPath)) { Write-Host "Built script not found: $ScriptPath (run Build.ps1)" -ForegroundColor Red; exit 2 }

$localDb = Get-Command SqlLocalDB.exe -ErrorAction SilentlyContinue
if (-not $localDb)
{
	foreach ($p in @(Get-ChildItem "$env:ProgramFiles\Microsoft SQL Server\*\Tools\Binn\SqlLocalDB.exe" -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)) { $localDb = $p; break }
}
if (-not $localDb) { Write-Host 'SKIP: SqlLocalDB is not installed here' -ForegroundColor DarkGray; exit 0 }
$localDbExe = if ($localDb -is [System.Management.Automation.CommandInfo]) { $localDb.Source } else { $localDb.FullName }
& $localDbExe start MSSQLLocalDB | Out-Null

$pass = 0; $fail = 0
function Check([string]$Name, $Cond, [string]$Note = '')
{
	if ([bool]$Cond) { $script:pass++; Write-Host "  PASS  $Name" -ForegroundColor Green }
	else { $script:fail++; Write-Host ("  FAIL  $Name" + $(if ($Note) { "  -- $Note" })) -ForegroundColor Red }
}

# The worker, verbatim from the built script
$text = [System.IO.File]::ReadAllText($ScriptPath)
$m = [regex]::Match($text, "(?s)\`$laneUpsertWorker = @'\r?\n(?<body>.*?)\r?\n'@")
if (-not $m.Success) { Write-Host 'Could not find $laneUpsertWorker in the built script' -ForegroundColor Red; exit 2 }
$worker = $m.Groups['body'].Value

$master = 'Server=(localdb)\MSSQLLocalDB;Integrated Security=True;Connect Timeout=30'
function Exec([string]$ConnStr, [string]$Sql)
{
	$cn = New-Object System.Data.SqlClient.SqlConnection $ConnStr
	$cn.Open()
	try { $c = $cn.CreateCommand(); $c.CommandText = $Sql; $c.CommandTimeout = 120; return $c.ExecuteScalar() }
	finally { $cn.Dispose() }
}
$suffix = [guid]::NewGuid().ToString('N').Substring(0, 6)
$srcDb = "PumpSrc_$suffix"; $laneDb = "PumpLane_$suffix"
Exec $master "CREATE DATABASE [$srcDb]; CREATE DATABASE [$laneDb];" | Out-Null
$src = "$master;Database=$srcDb"
$lane = "$master;Database=$laneDb"
try
{
	Exec $src @"
CREATE TABLE dbo.ITEM_TAB (F01 varchar(13) NOT NULL PRIMARY KEY, F02 varchar(40) NULL, F30 decimal(9,2) NULL);
INSERT dbo.ITEM_TAB VALUES ('0001','SAME',1.00), ('0002','CHANGED NAME',2.50), ('0003','NEW ON SERVER',3.00), ('0005',NULL,5.00);
CREATE TABLE dbo.NEW_TAB (ID int NOT NULL PRIMARY KEY, NAME varchar(20) NULL);
INSERT dbo.NEW_TAB VALUES (1,'one'), (2,'two');
CREATE TABLE dbo.NOPK_TAB (A int NULL, B varchar(10) NULL);
INSERT dbo.NOPK_TAB VALUES (1,'x');
"@ | Out-Null
	Exec $lane @"
CREATE TABLE dbo.ITEM_TAB (F01 varchar(13) NOT NULL PRIMARY KEY, F02 varchar(40) NULL, F30 decimal(9,2) NULL, LANE_ONLY varchar(10) NULL);
INSERT dbo.ITEM_TAB VALUES ('0001','SAME',1.00,'keep1'), ('0002','OLD NAME',2.50,'keep2'), ('0004','LANE ONLY ROW',4.00,'keep4'), ('0005','HAD NAME',5.00,NULL);
CREATE TABLE dbo.NOPK_TAB (A int NULL, B varchar(10) NULL);
INSERT dbo.NOPK_TAB VALUES (9,'lane');
"@ | Out-Null

	$plans = @(
		[pscustomobject]@{ Table = 'ITEM_TAB'; Columns = [string[]]@('F01', 'F02', 'F30'); SourcePk = [string[]]@('F01'); CreateSql = $null; Rows = 4 },
		[pscustomobject]@{ Table = 'NEW_TAB'; Columns = [string[]]@('ID', 'NAME'); SourcePk = [string[]]@('ID'); CreateSql = 'CREATE TABLE [dbo].[NEW_TAB] ([ID] int NOT NULL, [NAME] varchar(20) NULL, CONSTRAINT [PK_NEW_TAB] PRIMARY KEY ([ID]))'; Rows = 2 },
		[pscustomobject]@{ Table = 'NOPK_TAB'; Columns = [string[]]@('A', 'B'); SourcePk = [string[]]@(); CreateSql = $null; Rows = 1 }
	)
	$queue = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
	$ps = [PowerShell]::Create()
	$null = $ps.AddScript($worker).AddArgument($src).AddArgument($lane).AddArgument('001').AddArgument('TESTLANE').AddArgument('TCP').AddArgument($plans).AddArgument($queue)
	$res = @($ps.Invoke()) | Select-Object -Last 1
	foreach ($e in $ps.Streams.Error) { Write-Host "  worker error: $e" -ForegroundColor DarkYellow }
	$ps.Dispose()
	$msg = $null
	while ($queue.TryDequeue([ref]$msg)) { Write-Host ("    [{0}] {1}" -f $msg.Color, $msg.Text) -ForegroundColor DarkGray }

	Check 'worker connected' ($res -and $res.Connected)
	Check 'only the table without a primary key failed' ((@($res.Failed) -join ',') -eq 'NOPK_TAB') ("failed: " + (@($res.Failed) -join ','))
	Check 'changed row updated' ((Exec $lane "SELECT F02 FROM dbo.ITEM_TAB WHERE F01='0002'") -eq 'CHANGED NAME')
	Check 'NULL from the server is written over a value' ((Exec $lane "SELECT COUNT(*) FROM dbo.ITEM_TAB WHERE F01='0005' AND F02 IS NULL") -eq 1)
	Check 'new server row added' ((Exec $lane "SELECT F02 FROM dbo.ITEM_TAB WHERE F01='0003'") -eq 'NEW ON SERVER')
	Check 'lane-only row kept' ((Exec $lane "SELECT F02 FROM dbo.ITEM_TAB WHERE F01='0004'") -eq 'LANE ONLY ROW')
	Check 'lane-only column kept on updated rows' ((Exec $lane "SELECT LANE_ONLY FROM dbo.ITEM_TAB WHERE F01='0002'") -eq 'keep2')
	Check 'row count = lane rows + new server rows' ((Exec $lane "SELECT COUNT(*) FROM dbo.ITEM_TAB") -eq 5)
	Check 'missing table created and filled' ((Exec $lane "SELECT COUNT(*) FROM dbo.NEW_TAB") -eq 2)
	Check 'created table has the primary key' ((Exec $lane "SELECT COUNT(*) FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.NEW_TAB') AND is_primary_key = 1") -eq 1)
	Check 'table without a key left unchanged' ((Exec $lane "SELECT COUNT(*) FROM dbo.NOPK_TAB WHERE A = 9") -eq 1 -and (Exec $lane "SELECT COUNT(*) FROM dbo.NOPK_TAB") -eq 1)

	# Second pump: nothing changed -> 0 updated, 0 added
	$queue2 = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
	$ps = [PowerShell]::Create()
	$null = $ps.AddScript($worker).AddArgument($src).AddArgument($lane).AddArgument('001').AddArgument('TESTLANE').AddArgument('TCP').AddArgument(@($plans[0])).AddArgument($queue2)
	$null = $ps.Invoke(); $ps.Dispose()
	$texts = @(); while ($queue2.TryDequeue([ref]$msg)) { $texts += $msg.Text }
	Check 'repeat pump changes nothing (0 updated, 0 added)' (@($texts | Where-Object { $_ -match 'ITEM_TAB - 0 updated, 0 added' }).Count -eq 1) ($texts -join ' | ')
}
finally
{
	try { Exec $master "ALTER DATABASE [$srcDb] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [$srcDb]; ALTER DATABASE [$laneDb] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [$laneDb];" | Out-Null } catch { }
}

Write-Host ("Result: {0} passed, {1} failed" -f $pass, $fail) -ForegroundColor $(if ($fail) { 'Red' } else { 'Green' })
if ($fail -gt 0) { exit 1 }
exit 0
