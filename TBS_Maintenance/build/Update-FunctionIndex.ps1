<#
.SYNOPSIS
	Regenerates docs\FUNCTIONS.md - the map of every function: category, file, size, what it
	does (first line of its header description) and which menu item / startup step calls it.

.DESCRIPTION
	Reads src\Functions\<Category>\*.ps1 (one function per file) and the GUI / Main files.
	"Called from" lists the menu items (ToolStripMenuItem text) whose click handler calls the
	function, the warmup steps that run it, and the other functions that call it.
	Run by Build.ps1 after every successful build; can also be run on its own.
#>
[CmdletBinding()]
param ()

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$srcRoot = Join-Path $projectRoot 'src'
$outFile = Join-Path $projectRoot 'docs\FUNCTIONS.md'
$categoryOrder = @('Core', 'Store', 'Background', 'Dialogs', 'Server', 'Lanes', 'Scales', 'Nodes', 'General')
$categoryText = @{
	Core       = 'Infrastructure every other part uses: logging, background jobs, SQL connection strings, tools.'
	Store      = 'Store / node / database discovery and the startup data the GUI shows.'
	Background = 'Launch-time background tasks (exports, XE cleanup, quarterly maintenance).'
	Dialogs    = 'Reusable selection dialogs (nodes, tables, sections).'
	Server     = 'Server Tools menu: the store server and its database.'
	Lanes      = 'Lane Tools menu: lane databases, files, protocols, time, keys.'
	Scales     = 'Scale Tools menu: Bizerba / Ishida scales, BMS, ScaleCommApp.'
	Nodes      = 'Tasks that work on any node type (ping, reboot, C$ shares, hosts, copy).'
	General    = 'General Tools menu.'
}

function Get-Description([string[]]$Lines)
{
	# First paragraph after '# Description:' in the header block, cut to its first sentence
	# (or ~180 characters)
	$inDesc = $false
	$parts = New-Object System.Collections.Generic.List[string]
	foreach ($l in $Lines)
	{
		if ($l -match '^\s*function\s') { break }
		if ($l -match '^#\s*Description:') { $inDesc = $true; continue }
		if (-not $inDesc) { continue }
		$t = ($l -replace '^#\s?', '').Trim()
		if (-not $t -or $t -match '^[=-]{5,}$' -or $t -match '^\w[\w ]*:$') { if ($parts.Count) { break } else { continue } }
		$parts.Add($t)
	}
	$text = ($parts -join ' ')
	$m = [regex]::Match($text, '^.+?[.!](?=\s|$)')
	if ($m.Success -and $m.Length -ge 40) { $text = $m.Value }
	if ($text.Length -gt 180) { $text = $text.Substring(0, 177).TrimEnd() + '...' }
	return $text
}

$fnFiles = @(Get-ChildItem -Path (Join-Path $srcRoot 'Functions') -Recurse -File -Filter '*.ps1')
$functions = @{ }
foreach ($f in $fnFiles)
{
	$lines = [System.IO.File]::ReadAllLines($f.FullName)
	$functions[$f.BaseName] = [pscustomobject]@{
		Name        = $f.BaseName
		Category    = $f.Directory.Name
		File        = ('src/Functions/{0}/{1}' -f $f.Directory.Name, $f.Name)
		Lines       = $lines.Count
		Description = (Get-Description $lines)
		Text        = [string]::Join("`n", $lines)
		CalledFrom  = New-Object System.Collections.Generic.List[string]
	}
}
$names = @($functions.Keys | Sort-Object { $_.Length } -Descending)

function Find-Calls([string]$Text, [string]$Exclude)
{
	$hits = New-Object System.Collections.Generic.List[string]
	foreach ($n in $names)
	{
		if ($n -eq $Exclude) { continue }
		if ($Text -match ('(?<![\w-])' + [regex]::Escape($n) + '(?![\w-])')) { $hits.Add($n) }
	}
	return $hits
}

# Menu items: '$X = New-Object ...ToolStripMenuItem("Text")' followed by its Add_Click block
foreach ($g in @(Get-ChildItem -Path (Join-Path $srcRoot 'GUI') -File -Filter '*.ps1'))
{
	$text = [System.IO.File]::ReadAllText($g.FullName)
	$items = [regex]::Matches($text, '(?m)^\$(?<var>\w+)\s*=\s*New-Object\s+System\.Windows\.Forms\.ToolStripMenuItem\("(?<text>[^"]+)"\)')
	foreach ($m in $items)
	{
		$var = $m.Groups['var'].Value
		$clickIx = $text.IndexOf('$' + $var + '.Add_Click(', $m.Index)
		if ($clickIx -lt 0) { continue }
		$nextItem = $text.IndexOf('ToolStripMenuItem(', $clickIx)
		if ($nextItem -lt 0) { $nextItem = $text.Length }
		$handler = $text.Substring($clickIx, $nextItem - $clickIx)
		$menu = ($g.BaseName -replace '^\d+_Menu_', '' -replace '^\d+_', '')
		foreach ($called in (Find-Calls $handler ''))
		{
			# '&&' is how a menu text shows a literal '&' (a single '&' marks the access key)
			$functions[$called].CalledFrom.Add(('menu: {0} > "{1}"' -f $menu, ($m.Groups['text'].Value -replace '&&', '&')))
		}
	}
	# non-menu GUI uses (info labels, popups, timers)
	if ($items.Count -eq 0)
	{
		foreach ($called in (Find-Calls $text '')) { $functions[$called].CalledFrom.Add(('GUI: {0}' -f $g.BaseName)) }
	}
}

# Startup files
foreach ($s in @(Get-ChildItem -Path (Join-Path $srcRoot 'Main'), (Join-Path $srcRoot 'Init') -File -Filter '*.ps1'))
{
	$text = [System.IO.File]::ReadAllText($s.FullName)
	foreach ($called in (Find-Calls $text '')) { $functions[$called].CalledFrom.Add(('startup: {0}/{1}' -f $s.Directory.Name, $s.BaseName)) }
}

# Function -> function
foreach ($fn in $functions.Values)
{
	foreach ($called in (Find-Calls $fn.Text $fn.Name)) { $functions[$called].CalledFrom.Add(('function: {0}' -f $fn.Name)) }
}

# ---------------------------------------------------------------------------------------------
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('# Function index')
[void]$sb.AppendLine('')
[void]$sb.AppendLine('Generated by `build/Update-FunctionIndex.ps1` (runs with every build) - do not edit by hand.')
[void]$sb.AppendLine('')
[void]$sb.AppendLine(('{0} functions, {1} lines in src/Functions.' -f $functions.Count, (($functions.Values | Measure-Object Lines -Sum).Sum)))
[void]$sb.AppendLine('')
foreach ($cat in $categoryOrder)
{
	$inCat = @($functions.Values | Where-Object { $_.Category -eq $cat } | Sort-Object Name)
	if ($inCat.Count -eq 0) { continue }
	[void]$sb.AppendLine(('## {0}' -f $cat))
	[void]$sb.AppendLine('')
	[void]$sb.AppendLine($categoryText[$cat])
	[void]$sb.AppendLine('')
	[void]$sb.AppendLine('| Function | Lines | What it does | Called from |')
	[void]$sb.AppendLine('|---|---:|---|---|')
	foreach ($fn in $inCat)
	{
		$callers = @($fn.CalledFrom | Sort-Object -Unique)
		$fnCallers = @($callers | Where-Object { $_ -like 'function: *' })
		if ($fnCallers.Count -gt 6)
		{
			# widely used helpers: summarise the function callers instead of listing them all
			$callers = @($callers | Where-Object { $_ -notlike 'function: *' }) + @(('{0} functions' -f $fnCallers.Count))
		}
		$callerText = if ($callers.Count) { ($callers -join '<br>') } else { '**not called anywhere**' }
		$desc = $fn.Description -replace '\|', '\|'
		[void]$sb.AppendLine(('| [{0}](../{1}) | {2} | {3} | {4} |' -f $fn.Name, $fn.File, $fn.Lines, $desc, ($callerText -replace '\|', '\|')))
	}
	[void]$sb.AppendLine('')
}

$dir = Split-Path -Parent $outFile
if (-not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
[System.IO.File]::WriteAllText($outFile, ($sb.ToString() -replace "`r?`n", "`r`n"), [System.Text.Encoding]::ASCII)
Write-Host ("index: {0} functions -> {1}" -f $functions.Count, $outFile)
