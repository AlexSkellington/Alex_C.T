<#
.SYNOPSIS
	Normalises whitespace in the TBS Maintenance sources without changing any code.

.DESCRIPTION
	For every .ps1 / .psd1 under src\ (or the files given with -Path):
	  - removes trailing spaces / tabs (PowerShell Studio leaves indentation on blank lines)
	  - collapses 3+ consecutive blank lines to 2
	  - writes CRLF line endings, ASCII, one final newline
	Lines inside multi-line strings (here-strings, quoted strings that span lines) are never
	touched, so string contents - SQL, C#, generated worker scripts - stay byte-identical.

.PARAMETER Path
	Specific files to format. Default: everything under ..\src.

.PARAMETER Check
	Report files that are not formatted (exit 1) instead of rewriting them.
#>
[CmdletBinding()]
param (
	[string[]]$Path,
	[switch]$Check
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not $Path)
{
	$Path = @(Get-ChildItem -Path (Join-Path $projectRoot 'src') -Recurse -File |
		Where-Object { $_.Extension -in '.ps1', '.psd1' } | ForEach-Object { $_.FullName })
}

$changed = 0
foreach ($file in $Path)
{
	$original = [System.IO.File]::ReadAllText($file)
	$text = $original -replace "`r`n", "`n" -replace "`r", "`n"
	$lines = $text.Split("`n")
	if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines = $lines[0 .. ($lines.Count - 2)] }

	# Protect every line that is (partly) inside a multi-line string token.
	$protect = New-Object 'System.Collections.Generic.HashSet[int]'
	$tokens = $null
	$parseErrors = $null
	$null = [System.Management.Automation.Language.Parser]::ParseInput(($lines -join "`n"), [ref]$tokens, [ref]$parseErrors)
	if ($file.EndsWith('.ps1') -and @($parseErrors).Count -gt 0)
	{
		# Resource fragments (embedded worker bodies etc.) may not parse on their own; leave them alone.
		Write-Verbose "skip (does not parse on its own): $file"
		continue
	}
	foreach ($t in @($tokens))
	{
		if ($t.Kind -in 'HereStringLiteral', 'HereStringExpandable', 'StringLiteral', 'StringExpandable' -and
			$t.Extent.EndLineNumber -gt $t.Extent.StartLineNumber)
		{
			for ($n = $t.Extent.StartLineNumber; $n -lt $t.Extent.EndLineNumber; $n++) { [void]$protect.Add($n) }
		}
	}

	$out = New-Object System.Collections.Generic.List[string]
	$blankRun = 0
	for ($i = 0; $i -lt $lines.Count; $i++)
	{
		$line = $lines[$i]
		if (-not $protect.Contains($i + 1)) { $line = $line.TrimEnd(" `t") }
		if ($line -eq '' -and -not $protect.Contains($i + 1))
		{
			$blankRun++
			if ($blankRun -gt 2) { continue }
		}
		else { $blankRun = 0 }
		$out.Add($line)
	}
	while ($out.Count -gt 0 -and $out[$out.Count - 1] -eq '') { $out.RemoveAt($out.Count - 1) }
	$newText = ($out -join "`r`n") + "`r`n"

	if ($newText -cne $original)
	{
		$changed++
		if ($Check) { Write-Host ("not formatted: " + $file.Substring($projectRoot.Length + 1)) -ForegroundColor Yellow }
		else { [System.IO.File]::WriteAllText($file, $newText, [System.Text.Encoding]::ASCII) }
	}
}

if ($Check)
{
	if ($changed -gt 0) { Write-Host "$changed file(s) need formatting - run Format-Sources.ps1" -ForegroundColor Yellow; exit 1 }
	Write-Host 'all sources formatted' -ForegroundColor Green
	exit 0
}
Write-Host "formatted $changed file(s)" -ForegroundColor Green
exit 0
