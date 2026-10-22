<#
.SYNOPSIS
	Builds the single-file TBS Maintenance Script from the modular sources in src\.

.DESCRIPTION
	Concatenates the files listed in Manifest.psd1 (in order) into one .ps1, expands
	'#TBS:EMBED <path>' lines with the referenced resource file, appends the changelog,
	and validates the result:
	  - parses cleanly under the running PowerShell (use Windows PowerShell 5.1)
	  - pure ASCII (PowerShell 5.1 reads BOM-less files as Windows-1252)
	  - no function defined twice, no source file left out of (or listed twice in) the manifest
	The output is written as ASCII with CRLF line endings.

	The output is what ships:
	  - 'irm https://bit.ly/TBSMaintenanceScript | iex' downloads it from GitHub
	  - PowerShell Studio packages it into TBS_Maintenance_Script_V<version>.exe

.PARAMETER OutFile
	Where to write the built script. Default: the OutputFile from Manifest.psd1, relative to
	the repository root (TBS_Maintenace_Script_V2.ps1).

.PARAMETER Check
	Build in memory and compare with the existing output file instead of writing it.
	Exit code 0 = up to date, 1 = the committed output is stale (run Build.ps1).

.PARAMETER PsBuild
	Also write the PowerShell Studio package settings (<OutFile>.psbuild) from
	TBS_Maintenance.psbuild.template, with the product name set to the current version.

.EXAMPLE
	powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\build\Build.ps1

.EXAMPLE
	powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\build\Build.ps1 -Check
#>
[CmdletBinding()]
param (
	[string]$OutFile,
	[switch]$Check,
	[switch]$PsBuild
)

$ErrorActionPreference = 'Stop'

$buildDir = $PSScriptRoot
$projectRoot = Split-Path -Parent $buildDir          # ...\TBS_Maintenance
$repoRoot = Split-Path -Parent $projectRoot          # repository root
$manifest = Import-PowerShellDataFile -Path (Join-Path $buildDir 'Manifest.psd1')

if (-not $OutFile) { $OutFile = Join-Path $repoRoot $manifest.OutputFile }

function Fail([string]$Message)
{
	Write-Host "BUILD FAILED: $Message" -ForegroundColor Red
	exit 1
}

function Read-SourceLines([string]$Path)
{
	# Normalise line endings; keep every line exactly as written otherwise.
	$text = [System.IO.File]::ReadAllText($Path)
	$text = $text -replace "`r`n", "`n" -replace "`r", "`n"
	$arr = $text.Split("`n")
	if ($arr.Count -gt 0 -and $arr[$arr.Count - 1] -eq '') { $arr = $arr[0 .. ($arr.Count - 2)] }
	return , $arr
}

function Get-RelativePath([string]$Path)
{
	return $Path.Substring($projectRoot.Length + 1).Replace('\', '/')
}

# ---------------------------------------------------------------------------------------------
# 1) Resolve the manifest to an ordered file list (and catch orphans / duplicates)
# ---------------------------------------------------------------------------------------------
$ordered = New-Object System.Collections.Generic.List[string]
$seen = @{ }
foreach ($pattern in $manifest.Sources)
{
	$full = Join-Path $projectRoot ($pattern -replace '/', '\')
	$matched = @(Get-ChildItem -Path $full -File -ErrorAction SilentlyContinue | Sort-Object FullName)
	if ($matched.Count -eq 0) { Fail "manifest pattern matched no files: $pattern" }
	foreach ($f in $matched)
	{
		if ($seen.ContainsKey($f.FullName)) { Fail ("file matched by two manifest patterns: " + (Get-RelativePath $f.FullName)) }
		$seen[$f.FullName] = $true
		$ordered.Add($f.FullName)
	}
}

$srcRoot = Join-Path $projectRoot 'src'
$resourceRoot = Join-Path $srcRoot 'Resources'
$orphans = @(Get-ChildItem -Path $srcRoot -Recurse -File -Filter '*.ps1' |
	Where-Object { -not $_.FullName.StartsWith($resourceRoot + '\', [System.StringComparison]::OrdinalIgnoreCase) } |
	Where-Object { -not $seen.ContainsKey($_.FullName) })
if ($orphans.Count -gt 0) { Fail ("source file(s) not in the manifest: " + (($orphans | ForEach-Object { Get-RelativePath $_.FullName }) -join ', ')) }

# ---------------------------------------------------------------------------------------------
# 2) Concatenate (one blank line between files) and expand #TBS:EMBED resources
# ---------------------------------------------------------------------------------------------
$out = New-Object System.Collections.Generic.List[string]
$embedRegex = '^\s*#TBS:EMBED\s+(?<path>\S+)\s*$'
$embedded = 0

foreach ($file in $ordered)
{
	$src = Read-SourceLines $file
	# trim leading / trailing blank lines so the joins are uniform
	$a = 0; $b = $src.Count - 1
	while ($a -le $b -and $src[$a].Trim() -eq '') { $a++ }
	while ($b -ge $a -and $src[$b].Trim() -eq '') { $b-- }
	if ($a -gt $b) { continue }

	if ($out.Count -gt 0) { $out.Add('') }
	for ($i = $a; $i -le $b; $i++)
	{
		$line = $src[$i]
		$m = [regex]::Match($line, $embedRegex)
		if ($m.Success)
		{
			$resPath = Join-Path $srcRoot ($m.Groups['path'].Value -replace '/', '\')
			if (-not (Test-Path -LiteralPath $resPath)) { Fail ("embedded resource not found: " + $m.Groups['path'].Value + " (in " + (Get-RelativePath $file) + ")") }
			foreach ($rl in (Read-SourceLines $resPath))
			{
				# a line starting with '@ or "@ would terminate the surrounding here-string
				if ($rl -match '^[''"]@') { Fail ("embedded resource contains a here-string terminator: " + $m.Groups['path'].Value) }
				$out.Add($rl)
			}
			$embedded++
			continue
		}
		$out.Add($line)
	}
}

# Changelog: comment block after the final 'exit'
if ($manifest.Changelog)
{
	$clPath = Join-Path $projectRoot $manifest.Changelog
	if (-not (Test-Path -LiteralPath $clPath)) { Fail "changelog not found: $($manifest.Changelog)" }
	$out.Add('')
	foreach ($cl in (Read-SourceLines $clPath)) { $out.Add($cl) }
}

$text = ($out -join "`r`n") + "`r`n"

# ---------------------------------------------------------------------------------------------
# 3) Validate
# ---------------------------------------------------------------------------------------------
$nonAscii = [regex]::Matches($text, '[^\u0000-\u007F]')
if ($nonAscii.Count -gt 0)
{
	$firstIx = $nonAscii[0].Index
	$lineNo = ([regex]::Matches($text.Substring(0, $firstIx), "`n")).Count + 1
	Fail ("{0} non-ASCII character(s); first on output line {1}" -f $nonAscii.Count, $lineNo)
}

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0)
{
	foreach ($pe in @($parseErrors) | Select-Object -First 10)
	{
		Write-Host ("  line {0}: {1}" -f $pe.Extent.StartLineNumber, $pe.Message) -ForegroundColor DarkRed
	}
	Fail ("{0} parse error(s) in the built script" -f @($parseErrors).Count)
}

$functionNames = @($ast.EndBlock.Statements |
	Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] } |
	ForEach-Object { $_.Name })
$dupes = @($functionNames | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
if ($dupes.Count -gt 0) { Fail ("function defined more than once: " + ($dupes -join ', ')) }

$version = [regex]::Match($text, '(?m)^\s*\$VersionNumber\s*=\s*"([\d.]+)"').Groups[1].Value
if (-not $version) { Fail 'could not find $VersionNumber in the sources' }

# ---------------------------------------------------------------------------------------------
# 4) Write (or compare)
# ---------------------------------------------------------------------------------------------
$bytes = [System.Text.Encoding]::ASCII.GetBytes($text)
$sha = [System.Security.Cryptography.SHA256]::Create()
$hash = ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '')

if ($Check)
{
	if (-not (Test-Path -LiteralPath $OutFile)) { Write-Host "STALE: $OutFile does not exist" -ForegroundColor Yellow; exit 1 }
	$existing = [System.IO.File]::ReadAllText($OutFile) -replace "`r`n", "`n"
	if ($existing -ceq ($text -replace "`r`n", "`n"))
	{
		Write-Host "UP TO DATE: $OutFile (v$version)" -ForegroundColor Green
		exit 0
	}
	Write-Host "STALE: $OutFile differs from the sources - run Build.ps1" -ForegroundColor Yellow
	exit 1
}

[System.IO.File]::WriteAllBytes($OutFile, $bytes)

if ($PsBuild)
{
	$tpl = Join-Path $buildDir 'TBS_Maintenance.psbuild.template'
	if (Test-Path -LiteralPath $tpl)
	{
		$psb = [System.IO.File]::ReadAllText($tpl).Replace('{VERSION}', $version)
		[System.IO.File]::WriteAllText($OutFile + '.psbuild', $psb, [System.Text.Encoding]::ASCII)
	}
}

Write-Host ("BUILT  v{0}  ->  {1}" -f $version, $OutFile) -ForegroundColor Green
Write-Host ("       {0} source files, {1} embedded resource(s), {2} functions, {3} lines" -f $ordered.Count, $embedded, $functionNames.Count, $out.Count)
Write-Host ("       SHA256 {0}" -f $hash)
exit 0
