# ===================================================================================================
#                              FUNCTION: Write_Log
# ---------------------------------------------------------------------------------------------------
# Description:
#   The single logging entry point of this script. Writes a line to the GUI log box (colored),
#   to the console while the box does not exist yet, and to the daily file log. Handles
#   cross-thread updates, test-mode rewording and repeat suppression.
#
# Parameters:
#   - Message          [string]: the line to log (required)
#   - Color            [any]: color name or System.Drawing.Color (default Black)
#   - IncludeTimestamp [switch]: prefix the line with a timestamp (off by default)
#   - Raw              [switch]: skip the presentation layer (severity tags, rewording,
#                                repeat suppression) - the startup summary formats itself
#   - FileOnly         [switch]: file log only - skip the box and the console (v2.6.9)
#
# Variables:
#   - $logBox / $script:LogBox / $global:LogBox / FunctionResults['LogBox']: the log control,
#     first valid one wins (GUI\05_Log_Area publishes all of them)
#   - $script:ScriptsLog: folder of the daily file log TBS_Maintenance_<yyyy-MM-dd>.log
#   - $script:IsTestEnvironment: enables the friendly test-mode rewording
#   - $script:LogSeen: line -> last time written (120 s repeat suppression)
#   - $script:WriteLogColors / $script:WriteLogFriendly: lookup tables, built on first use
#
# Notes:
#   - Shipped into background jobs as text (Start_Background_Function_Job), so it must stay
#     self-contained: everything it needs is created lazily inside it.
#   - File-log line format "HH:mm:ss [TAG] message" is checked by the self-test (section G).
# ===================================================================================================

function Write_Log
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Message,
		[Parameter(Mandatory = $false)]
		$Color = "Black",
		[Parameter(Mandatory = $false)]
		[switch]$IncludeTimestamp = $false,
		[Parameter(Mandatory = $false)]
		[switch]$Raw,
		[Parameter(Mandatory = $false)]
		[switch]$FileOnly
	)

	# ----------------------------------------------------------------------------------------------
	# Lookup tables (built once per session / job)
	# ----------------------------------------------------------------------------------------------
	if (-not $script:WriteLogColors)
	{
		$script:WriteLogColors = @{
			green   = @([System.Drawing.Color]::Green, 'Green', '[ OK ] ')
			red     = @([System.Drawing.Color]::Red, 'Red', '[FAIL] ')
			yellow  = @([System.Drawing.Color]::Goldenrod, 'Yellow', '[WARN] ')
			orange  = @([System.Drawing.Color]::Orange, 'Yellow', '[WARN] ')
			blue    = @([System.Drawing.Color]::DodgerBlue, 'Cyan', '[INFO] ')
			cyan    = @([System.Drawing.Color]::Cyan, 'Cyan', '[INFO] ')
			magenta = @([System.Drawing.Color]::Magenta, 'Magenta', '')
			gray    = @([System.Drawing.Color]::Gray, 'Gray', '')
			grey    = @([System.Drawing.Color]::Gray, 'Gray', '')
			white   = @([System.Drawing.Color]::White, 'White', '')
		}
		# Expected conditions in test mode (no storeman folder / store database): reworded
		# into calm explanations. In production the original text is kept - there these
		# conditions are real problems.
		# (hashtables, not nested arrays: @( @(a,b) <newline> @(c,d) ) flattens into one list)
		$script:WriteLogFriendly = @(
			@{ P = '^Startup\.ini path not provided'; T = 'Store config (Startup.ini) not found on this machine. Continuing without store/terminal details - expected in test mode.' }
			@{ P = '^WIN\.INI is unresolved'; T = 'Store info file (WIN.INI) not found. Continuing without store details - expected in test mode.' }
			@{ P = '^(No INFO_\*_WIN\.INI|No store info file)'; T = 'Store info file (INFO_*_WIN.INI) not found - expected in test mode.' }
			@{ P = '^Failed to retrieve counts from the database'; T = 'Store database is not reachable - expected in test mode. Lane/scale info will come from local files where available.' }
			@{ P = '\(Urgent Messages\) does not exist\.'; T = 'Urgent messages folder not found - expected in test mode (no storeman folder).' }
		)
	}

	# ----------------------------------------------------------------------------------------------
	# Colour: a System.Drawing.Color as given, or one of the names above (anything else = black)
	# ----------------------------------------------------------------------------------------------
	$colorName = ''
	if ($Color -is [System.Drawing.Color])
	{
		$selColor = $Color
	}
	else
	{
		$colorName = ([string]$Color).Trim().ToLower()
		$entry = $script:WriteLogColors[$colorName]
		$selColor = if ($entry) { $entry[0] } else { [System.Drawing.Color]::Black }
	}

	# ----------------------------------------------------------------------------------------------
	# Presentation layer (v2.6.8) - calmer, easier-to-read output (skipped with -Raw):
	#   1) test mode: known expected conditions are reworded (shown orange)
	#   2) raw .NET exception dumps are squashed to their message (~200 chars)
	#   3) a severity tag by colour: [ OK ] / [WARN] / [FAIL] / [INFO]
	#   4) identical lines are suppressed for 120 s (stops per-lane repeat spam)
	# ----------------------------------------------------------------------------------------------
	if (-not $Raw)
	{
		if ($script:IsTestEnvironment)
		{
			foreach ($fm in $script:WriteLogFriendly)
			{
				if ($Message -match $fm.P)
				{
					$Message = $fm.T
					$selColor = [System.Drawing.Color]::Orange
					break
				}
			}
		}

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
			$Message = if ($pre) { "${pre}: $inner" } else { $inner }
		}
		elseif ($Message.Length -gt 300)
		{
			$cut = $Message.Substring(0, 300)
			$sp = $cut.LastIndexOf(' ')
			if ($sp -gt 200) { $cut = $cut.Substring(0, $sp) }
			$Message = $cut + '...'
		}

		# Tag from the final colour (explicit -eq tests: switch() does not match Color values)
		$tag = ''
		if ($selColor -eq [System.Drawing.Color]::Green) { $tag = '[ OK ] ' }
		elseif ($selColor -eq [System.Drawing.Color]::Orange -or $selColor -eq [System.Drawing.Color]::Goldenrod) { $tag = '[WARN] ' }
		elseif ($selColor -eq [System.Drawing.Color]::Red) { $tag = '[FAIL] ' }
		elseif ($selColor -eq [System.Drawing.Color]::DodgerBlue -or $selColor -eq [System.Drawing.Color]::Cyan) { $tag = '[INFO] ' }

		if (-not $script:LogSeen) { $script:LogSeen = @{ } }
		if ($script:LogSeen.Count -gt 400) { $script:LogSeen.Clear() }
		$lineKey = $tag + $Message
		$nowT = [datetime]::Now
		if ($script:LogSeen.ContainsKey($lineKey) -and ($nowT - $script:LogSeen[$lineKey]).TotalSeconds -lt 120) { return }
		$script:LogSeen[$lineKey] = $nowT

		$Message = $tag + $Message
	}

	$fullMessage = if ($IncludeTimestamp) { "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f [datetime]::Now, $Message } else { $Message }

	# ----------------------------------------------------------------------------------------------
	# File log (v2.6.9): every line (shown or -FileOnly) is appended to
	#   <ScriptsLog>\TBS_Maintenance_<yyyy-MM-dd>.log
	# ReadWrite sharing because background jobs append at the same time. Silent on failure -
	# logging must never break maintenance work.
	# ----------------------------------------------------------------------------------------------
	try
	{
		$now = [datetime]::Now
		$logToday = $now.ToString('yyyy-MM-dd')
		if ($script:LogFileLogDate -ne $logToday)
		{
			$script:LogFileLogDate = $logToday
			$script:LogFileLogPath = Join-Path $script:ScriptsLog ("TBS_Maintenance_" + $logToday + ".log")
		}
		$logStream = [System.IO.File]::Open($script:LogFileLogPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
		$logWriter = New-Object System.IO.StreamWriter($logStream)
		$logWriter.WriteLine($now.ToString('HH:mm:ss') + " " + $fullMessage)
		$logWriter.Close()
	}
	catch { }

	if ($FileOnly) { return }

	# ----------------------------------------------------------------------------------------------
	# Log box (cross-thread aware)
	# ----------------------------------------------------------------------------------------------
	$lb = $null
	foreach ($candidate in @($logBox, $script:LogBox, $global:LogBox))
	{
		if ($candidate -is [System.Windows.Forms.RichTextBox] -and -not $candidate.IsDisposed) { $lb = $candidate; break }
	}
	if (-not $lb -and $script:FunctionResults -and $script:FunctionResults.ContainsKey('LogBox'))
	{
		$tmp = $script:FunctionResults['LogBox']
		if ($tmp -is [System.Windows.Forms.RichTextBox] -and -not $tmp.IsDisposed) { $lb = $tmp }
	}

	if ($lb)
	{
		$uiWrite = {
			param ($rtb, $text, $color)
			try
			{
				if (-not $rtb -or $rtb.IsDisposed) { return }
				$rtb.SelectionStart = $rtb.TextLength
				$rtb.SelectionLength = 0
				$rtb.SelectionColor = $color
				$rtb.AppendText($text + "`r`n")
				$rtb.SelectionColor = $rtb.ForeColor
				$rtb.SelectionStart = $rtb.TextLength
				$rtb.ScrollToCaret()
			}
			catch { }
		}
		try
		{
			# BeginInvoke from other threads (never blocks, safe while closing)
			if ($lb.InvokeRequired) { [void]$lb.BeginInvoke($uiWrite, @($lb, $fullMessage, $selColor)) }
			else { & $uiWrite $lb $fullMessage $selColor }
			return
		}
		catch { }
	}

	# ----------------------------------------------------------------------------------------------
	# Console (before the log box exists, or if the box write failed)
	# ----------------------------------------------------------------------------------------------
	try
	{
		$entry = $script:WriteLogColors[$colorName]
		$consoleColor = if ($entry) { $entry[1] } else { 'White' }
		Write-Host $fullMessage -ForegroundColor $consoleColor
	}
	catch { }
}
