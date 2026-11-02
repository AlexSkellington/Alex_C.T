# ==============================================================================
#  Backup_Copy_Test.ps1
# ------------------------------------------------------------------------------
#  Tests the "Also copy to" part of the backup schedulers with the REAL functions
#  from the built script (loaded by AST, nothing else runs):
#    - Resolve_Backup_Copy_Paths creates / checks folders and drops duplicates
#      and the main folder
#    - the lines New_Backup_Copy_Snippet generates parse, copy the new backup into
#      <folder>\<machine> for every folder, keep only the newest N there, and
#      log a failed copy without stopping the others
#    - a folder name with ' and $ cannot break the generated script
#  Everything happens under %TEMP%.
#
#  Usage: powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\tests\Backup_Copy_Test.ps1
# ==============================================================================
param ([string]$ScriptPath = '')

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\TBS_Maintenance_Script.ps1' }
if (-not (Test-Path -LiteralPath $ScriptPath)) { Write-Host "Built script not found: $ScriptPath (run Build.ps1)" -ForegroundColor Red; exit 2 }

$pass = 0; $fail = 0
function Check([string]$Name, $Cond, [string]$Note = '')
{
	if ([bool]$Cond) { $script:pass++; Write-Host "  PASS  $Name" -ForegroundColor Green }
	else { $script:fail++; Write-Host ("  FAIL  $Name" + $(if ($Note) { "  -- $Note" })) -ForegroundColor Red }
}

$tokens = $null; $errs = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errs)
foreach ($name in @('Resolve_Backup_Folder', 'Resolve_Backup_Copy_Paths', 'New_Backup_Copy_Snippet', 'Select-FolderOrDefault'))
{
	$fn = $ast.Find({ param ($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
	if (-not $fn) { Write-Host "Function $name not found in the built script" -ForegroundColor Red; exit 2 }
	. ([scriptblock]::Create($fn.Extent.Text))
}

# The Windows folder picker (Select Folder button, typed / network paths) is reached through .NET
# internals - make sure they are there on this Windows / .NET (otherwise the old tree is shown)
Check 'the Explorer folder picker can be built (FOS_PICKFOLDERS)' ((Select-FolderOrDefault -ProbeOnly) -eq $true)

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('tbs-copytest-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try
{
	$main = Join-Path $root 'main'
	$copyA = Join-Path $root 'copy A'
	$copyB = Join-Path $root "odd 'name' `$x"
	New-Item -ItemType Directory -Path $main -Force | Out-Null

	$r = Resolve_Backup_Copy_Paths -Paths @($copyA, $copyB, $copyA, $main) -Primary $main
	Check 'extra folders are created and accepted' ($r.Ok -and (Test-Path -LiteralPath $copyA) -and (Test-Path -LiteralPath $copyB))
	Check 'duplicates and the main folder are dropped' ($r.Paths.Count -eq 2) ("got: " + ($r.Paths -join ' | '))
	$bad = Resolve_Backup_Copy_Paths -Paths @('relative\folder') -Primary $main
	Check 'a relative path is refused with a message' ((-not $bad.Ok) -and $bad.Error)

	# a destination that cannot be written: a FILE where the folder should be
	$blocked = Join-Path $root 'blocked'
	Set-Content -LiteralPath $blocked -Value 'not a folder'
	$paths = @($r.Paths) + @($blocked)

	$snippet = New_Backup_Copy_Snippet -Paths $paths -SubFolder 'HOST01' -Filter 'DB_*.bak' -Keep 2 -SourceExpression '$backupFile' -LogExpression '(Join-Path $BackupRoot ''backup.log'')'
	$t2 = $null; $e2 = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput($snippet, [ref]$t2, [ref]$e2)
	Check 'generated copy lines parse' (@($e2).Count -eq 0) ((@($e2) | ForEach-Object { $_.Message }) -join '; ')
	Check 'no extra folders = no lines' ((New_Backup_Copy_Snippet -Paths @() -SubFolder 'X' -Filter 'x' -Keep 1 -SourceExpression '$a' -LogExpression '$b') -eq '')

	# Run it three times like three nights of backups (keep 2)
	foreach ($n in 1 .. 3)
	{
		$BackupRoot = $main
		$backupFile = Join-Path $main ("DB_2026010{0}_010000.bak" -f $n)
		Set-Content -LiteralPath $backupFile -Value "backup $n"
		(Get-Item -LiteralPath $backupFile).LastWriteTime = (Get-Date).AddDays($n - 10)
		& ([scriptblock]::Create($snippet))
	}
	$inA = @(Get-ChildItem -LiteralPath (Join-Path $copyA 'HOST01') -Filter 'DB_*.bak' | ForEach-Object { $_.Name } | Sort-Object)
	$inB = @(Get-ChildItem -LiteralPath (Join-Path $copyB 'HOST01') -Filter 'DB_*.bak' | ForEach-Object { $_.Name } | Sort-Object)
	Check 'copies land in <folder>\<machine>' ($inA.Count -gt 0 -and $inB.Count -gt 0)
	Check 'only the newest 2 are kept in each extra folder' (($inA -join ',') -eq 'DB_20260102_010000.bak,DB_20260103_010000.bak' -and ($inB -join ',') -eq ($inA -join ',')) ("A: " + ($inA -join ',') + " B: " + ($inB -join ','))
	$log = Get-Content -LiteralPath (Join-Path $main 'backup.log') -Raw
	Check 'each copy is logged' (([regex]::Matches($log, 'Copy saved to')).Count -eq 6)
	Check 'a folder that cannot be written is logged as FAILED (others still copied)' (([regex]::Matches($log, 'FAILED')).Count -eq 3)
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ("Result: {0} passed, {1} failed" -f $pass, $fail) -ForegroundColor $(if ($fail) { 'Red' } else { 'Green' })
if ($fail -gt 0) { exit 1 }
exit 0
