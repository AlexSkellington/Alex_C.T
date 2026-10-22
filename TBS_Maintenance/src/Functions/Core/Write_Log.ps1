# ===================================================================================================
#                              FUNCTION: Write_Log
# ---------------------------------------------------------------------------------------------------
# Description:
#   The single logging entry point of this script. Writes a line to the GUI log box (colored),
#   optionally to the console, and to the daily file log. Handles UI rebuilds, cross-thread
#   updates, test-mode rewording and repeat suppression.
#
# Parameters:
#   - Message          [string]: the line to log (required)
#   - Color            [any]: color name or System.Drawing.Color (default Black)
#   - IncludeTimestamp [switch]: prefix the line with a timestamp (off by default)
#   - Raw              [switch]: skip the presentation layer (severity tags, rewording)
#   - FileOnly         [switch]: file log only - skip the box and console (v2.6.9)
#
# Variables:
#   - $script:LogBox / $logBox / $global:LogBox: the active log control (found robustly)
#
# Notes:
#   - Background job success lines use -FileOnly; only failures reach the box.
# ===================================================================================================

function Write_Log
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Message,
		[Parameter(Mandatory = $false)]
		$Color = "Black",
		# accepts string names OR [System.Drawing.Color]
		# Default OFF (user preference): log lines are clean, no timestamp prefix.
		# Pass -IncludeTimestamp if a caller ever wants it back.
		[Parameter(Mandatory = $false)]
		[switch]$IncludeTimestamp = $false,
		# Skip the presentation layer (severity tags, test-mode rewording,
		# exception squashing, repeat suppression). Used by the startup
		# summary block which formats its own lines.
		[Parameter(Mandatory = $false)]
		[switch]$Raw,
		# v2.6.9: skip the log BOX and the console entirely - the line goes to
		# the file log only. Used for background job lifecycle lines: the status
		# bar already shows what the app is doing in the background, and the box
		# only sees failures (user preference).
		[Parameter(Mandatory = $false)]
		[switch]$FileOnly
	)
	
	# ----------------------------------------------------------------------------------------------
	# Resolve timestamp (OFF by default - user preference; opt in per call with -IncludeTimestamp)
	# ----------------------------------------------------------------------------------------------
	$timestamp = $null
	if ($IncludeTimestamp)
	{
		try { $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss' }
		catch { $timestamp = $null }
	}
	
	# ----------------------------------------------------------------------------------------------
	# Resolve the active log control robustly (handles refactors / recreated controls)
	# Priority:
	#   1) local $logBox (if valid)
	#   2) $script:LogBox / $script:logBox
	#   3) $global:LogBox / $global:logBox
	#   4) $script:FunctionResults['LogBox'] (if used)
	# ----------------------------------------------------------------------------------------------
	$lb = $null
	
	# Helper inline checks (no nested functions required)
	try
	{
		if ($logBox -and ($logBox -is [System.Windows.Forms.RichTextBox]) -and -not $logBox.IsDisposed) { $lb = $logBox }
	}
	catch { }
	
	if (-not $lb)
	{
		try
		{
			if ($script:LogBox -and ($script:LogBox -is [System.Windows.Forms.RichTextBox]) -and -not $script:LogBox.IsDisposed) { $lb = $script:LogBox }
		}
		catch { }
	}
	if (-not $lb)
	{
		try
		{
			if ($script:logBox -and ($script:logBox -is [System.Windows.Forms.RichTextBox]) -and -not $script:logBox.IsDisposed) { $lb = $script:logBox }
		}
		catch { }
	}
	if (-not $lb)
	{
		try
		{
			if ($global:LogBox -and ($global:LogBox -is [System.Windows.Forms.RichTextBox]) -and -not $global:LogBox.IsDisposed) { $lb = $global:LogBox }
		}
		catch { }
	}
	if (-not $lb)
	{
		try
		{
			if ($global:logBox -and ($global:logBox -is [System.Windows.Forms.RichTextBox]) -and -not $global:logBox.IsDisposed) { $lb = $global:logBox }
		}
		catch { }
	}
	if (-not $lb)
	{
		try
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('LogBox'))
			{
				$tmp = $script:FunctionResults['LogBox']
				if ($tmp -and ($tmp -is [System.Windows.Forms.RichTextBox]) -and -not $tmp.IsDisposed) { $lb = $tmp }
			}
		}
		catch { }
	}
	
	# ----------------------------------------------------------------------------------------------
	# Normalize color (string name OR actual System.Drawing.Color)
	# ----------------------------------------------------------------------------------------------
	$selColor = [System.Drawing.Color]::Black
	try
	{
		if ($Color -is [System.Drawing.Color])
		{
			$selColor = $Color
		}
		else
		{
			$colorName = ([string]$Color).Trim().ToLower()
			$selColor = switch ($colorName)
			{
				"green"   { [System.Drawing.Color]::Green }
				"red"     { [System.Drawing.Color]::Red }
				"yellow"  { [System.Drawing.Color]::Goldenrod }
				"blue"    { [System.Drawing.Color]::DodgerBlue } # nicer than default Blue
				"magenta" { [System.Drawing.Color]::Magenta }
				"gray"    { [System.Drawing.Color]::Gray }
				"grey"    { [System.Drawing.Color]::Gray }
				"cyan"    { [System.Drawing.Color]::Cyan }
				"white"   { [System.Drawing.Color]::White }
				"orange"  { [System.Drawing.Color]::Orange }
				default   { [System.Drawing.Color]::Black }
			}
		}
	}
	catch { $selColor = [System.Drawing.Color]::Black }
	
	# ----------------------------------------------------------------------------------------------
	# Presentation layer (v2.6.8) - makes the log output calmer and easier to read:
	#   1) Test mode: known expected conditions are reworded into friendly explanations
	#   2) Raw .NET exception dumps are squashed to a short first sentence
	#   3) A severity tag is prefixed by color: [ OK ] / [WARN] / [FAIL] / [INFO]
	#   4) Identical lines are suppressed for 120s (stops per-lane repeat spam)
	# Skipped entirely when -Raw is passed (the startup summary block formats itself).
	# ----------------------------------------------------------------------------------------------
	if (-not $Raw)
	{
		# --- 1) Test-mode friendly rewording -------------------------------------------
		# These conditions are EXPECTED when the storeman folder / store database is
		# not available. Reword them so the log reads as a calm explanation instead
		# of a wall of red errors. In production the original text is kept - there
		# these conditions are real problems.
		if ($script:IsTestEnvironment)
		{
			$friendlyMap = @(
				@{ P = '^Startup\.ini path not provided';                 T = 'Store config (Startup.ini) not found on this machine. Continuing without store/terminal details - expected in test mode.'; C = 'orange' }
				@{ P = '^WIN\.INI is unresolved';                         T = 'Store info file (WIN.INI) not found. Continuing without store details - expected in test mode.'; C = 'orange' }
				@{ P = '^(No INFO_\*_WIN\.INI|No store info file)';        T = 'Store info file (INFO_*_WIN.INI) not found - expected in test mode.'; C = 'orange' }
				@{ P = '^Failed to retrieve counts from the database';    T = 'Store database is not reachable - expected in test mode. Lane/scale info will come from local files where available.'; C = 'orange' }
				@{ P = '\(Urgent Messages\) does not exist\.';            T = 'Urgent messages folder not found - expected in test mode (no storeman folder).'; C = 'orange' }
			)
			foreach ($entry in $friendlyMap)
			{
				if ($Message -match $entry.P)
				{
					$Message = $entry.T
					$selColor = switch ($entry.C)
					{
						'orange' { [System.Drawing.Color]::Orange }
						'blue'   { [System.Drawing.Color]::DodgerBlue }
						default  { [System.Drawing.Color]::Gray }
					}
					break
				}
			}
		}
		
		# --- 2) Squash raw .NET exception dumps ----------------------------------------
		# Turns:
		#   "...: Exception calling ""Open"" with ""0"" argument(s): ""A network-related or...""
		# into:
		#   "...: A network-related or..."  (first ~200 chars, on a word boundary)
		if ($Message -match 'Exception calling "[^"]*" with "\d+" argument\(s\): "(.*)"')
		{
			$inner = $matches[1]
			if ($inner.Length -gt 200)
			{
				$cut = $inner.Substring(0, 200)
				$sp = $cut.LastIndexOf(' ')
				if ($sp -gt 120) { $cut = $cut.Substring(0, $sp) }
				$inner = $cut + '...'
			}
			$pre = $Message.Substring(0, $Message.IndexOf('Exception calling', [System.StringComparison]::Ordinal)).TrimEnd().TrimEnd(':').TrimEnd()
			if ($pre) { $Message = "${pre}: $inner" } else { $Message = $inner }
		}
		elseif ($Message.Length -gt 300)
		{
			$cut = $Message.Substring(0, 300)
			$sp = $cut.LastIndexOf(' ')
			if ($sp -gt 200) { $cut = $cut.Substring(0, $sp) }
			$Message = $cut + '...'
		}
		
		# --- 3) Severity tag by color --------------------------------------------------
		# NOTE: explicit -eq tests, not a switch() - PowerShell switch() does not
		# match System.Drawing.Color values (falls to default).
		$tag = ''
		if ($selColor -eq [System.Drawing.Color]::Green)        { $tag = '[ OK ] ' }
		elseif ($selColor -eq [System.Drawing.Color]::Orange)   { $tag = '[WARN] ' }
		elseif ($selColor -eq [System.Drawing.Color]::Goldenrod){ $tag = '[WARN] ' }
		elseif ($selColor -eq [System.Drawing.Color]::Red)      { $tag = '[FAIL] ' }
		elseif ($selColor -eq [System.Drawing.Color]::DodgerBlue){ $tag = '[INFO] ' }
		elseif ($selColor -eq [System.Drawing.Color]::Cyan)     { $tag = '[INFO] ' }
		
		# --- 4) Repeat suppression -------------------------------------------------------
		# The same line (tag + text) is only written once per 120 seconds. This stops
		# per-lane startup jobs from filling the log with the identical error.
		if (-not $script:LogSeen) { $script:LogSeen = @{} }
		if ($script:LogSeen.Count -gt 400) { $script:LogSeen.Clear() }
		$lineKey = $tag + $Message
		$nowT = [datetime]::Now
		if ($script:LogSeen.ContainsKey($lineKey) -and ($nowT - $script:LogSeen[$lineKey]).TotalSeconds -lt 120)
		{
			return
		}
		$script:LogSeen[$lineKey] = $nowT
		
		$Message = $tag + $Message
	}
	
	# Compose message
	$fullMessage = if ($timestamp) { "[$timestamp] $Message" }
	else { $Message }
	
	# ----------------------------------------------------------------------------------------------
	# File log (v2.6.9): every line is appended to
	#   C:\Tecnica_Systems\Alex_C.T\Log\TBS_Maintenance_<date>.log   (one file per day)
	# - lines the UI shows (box writes are the presentation of the same line)
	# - -FileOnly lines, which the UI deliberately does NOT show (job lifecycle)
	# Lines suppressed by the 120-second repeat rule above are skipped here too,
	# so the file stays as readable as the window.
	# Opened with ReadWrite sharing because background jobs may append at the same time.
	# Any failure here is silent - logging must never break maintenance work.
	# ----------------------------------------------------------------------------------------------
	try
	{
		$logToday = (Get-Date -Format 'yyyy-MM-dd')
		if (-not $script:LogFileLogDate -or ($script:LogFileLogDate -ne $logToday))
		{
			$script:LogFileLogDate = $logToday
			$script:LogFileLogPath = Join-Path $script:ScriptsLog ("TBS_Maintenance_" + $logToday + ".log")
		}
		$logLine = (Get-Date -Format 'HH:mm:ss') + " " + $fullMessage
		try
		{
			$logStream = [System.IO.File]::Open($script:LogFileLogPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
			$logWriter = New-Object System.IO.StreamWriter($logStream)
			$logWriter.WriteLine($logLine)
			$logWriter.Close()
		}
		catch { }
	}
	catch
	{
		# Silent - a missing log folder or disk problem must not surface in the UI
	}
	
	# ----------------------------------------------------------------------------------------------
	# If we have a UI log, write to it safely (cross-thread aware).
	# -FileOnly lines never reach the box (file log only - see param docs).
	# ----------------------------------------------------------------------------------------------
	if (-not $FileOnly -and $lb -and -not $lb.IsDisposed)
	{
		# Define the UI write action as a scriptblock (NO nested function)
		$uiWrite = {
			param ($rtb,
				$text,
				$color)
			
			try
			{
				if (-not $rtb -or $rtb.IsDisposed) { return }
				
				# Append colored line
				$rtb.SelectionStart = $rtb.TextLength
				$rtb.SelectionLength = 0
				$rtb.SelectionColor = $color
				$rtb.AppendText($text + "`r`n")
				
				# Reset to default color for future appends
				$rtb.SelectionColor = $rtb.ForeColor
				
				# Autoscroll
				$rtb.SelectionStart = $rtb.TextLength
				$rtb.ScrollToCaret()
			}
			catch { }
		}
		
		try
		{
			# Marshal to UI thread if needed
			if ($lb.InvokeRequired)
			{
				# BeginInvoke avoids deadlocks if called during closing
				[void]$lb.BeginInvoke($uiWrite, @($lb, $fullMessage, $selColor))
			}
			else
			{
				& $uiWrite $lb $fullMessage $selColor
			}
			return
		}
		catch
		{
			# If anything about UI writing fails, fall back to console below
		}
	}
	
	# ----------------------------------------------------------------------------------------------
	# Console fallback (until logBox is initialized OR if UI write failed).
	# Skipped for -FileOnly lines: the file log is their only destination.
	# ----------------------------------------------------------------------------------------------
	if ($FileOnly)
	{
		return
	}
	try
	{
		# Write-Host expects a ConsoleColor, so map common values
		$cc = ([string]$Color).Trim().ToLower()
		$consoleColor = switch ($cc)
		{
			"green"   { "Green" }
			"red"     { "Red" }
			"yellow"  { "Yellow" }
			"blue"    { "Cyan" } # closer visibility on dark consoles
			"magenta" { "Magenta" }
			"gray"    { "Gray" }
			"grey"    { "Gray" }
			"cyan"    { "Cyan" }
			"white"   { "White" }
			"orange"  { "Yellow" }
			default   { "White" }
		}
		Write-Host $fullMessage -ForegroundColor $consoleColor
	}
	catch
	{
		# last-resort: silent failure
	}
}
