# ===================================================================================================
#                                   SECTION: Main - Show the Window
# ---------------------------------------------------------------------------------------------------
# Description:
#   Shows the main window (modal - the script waits here until it closes), then exits.
#   Unless -ShowConsole was given, the console window is hidden 400 ms after the main window
#   appears (so the "Script started" line is visible for a moment while it loads).
#
# Notes:
#   - 'exit' also ends the PowerShell session for irm | iex launches - intentional, the console
#     is hidden by then and would otherwise stay open invisibly.
#   - tests\Smoke.ps1 replaces the LAST '[void]$form.ShowDialog()' in the built file - keep it.
# ===================================================================================================

Write-Host "Script started" -ForegroundColor Green

if (-not $script:ShowConsole -and [Environment]::UserInteractive)
{
	try
	{
		$script:PostOpenConsoleHideTimer = New-Object System.Windows.Forms.Timer
		$script:PostOpenConsoleHideTimer.Interval = 400
		$script:PostOpenConsoleHideTimer.Add_Tick({
				try
				{
					$script:PostOpenConsoleHideTimer.Stop()
					$script:PostOpenConsoleHideTimer.Dispose()
					$script:PostOpenConsoleHideTimer = $null
				}
				catch { }
				try
				{
					$__consoleHandle = [ConsoleWindowHelper]::GetConsoleWindow()
					if ($__consoleHandle -ne [IntPtr]::Zero) { [void][ConsoleWindowHelper]::ShowWindow($__consoleHandle, 0) }
				}
				catch { }
			})
		$form.Add_Shown({
				try { if ($script:PostOpenConsoleHideTimer) { $script:PostOpenConsoleHideTimer.Start() } }
				catch { }
			})
	}
	catch { }
}

[void]$form.ShowDialog()

if ($script:ShowConsole) { Write-Host "Script closing..." -ForegroundColor Yellow }

exit
