# ===================================================================================================
#                                       SECTION: Show the Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays the main GUI form to the user.
# ===================================================================================================

if (-not $script:ShowConsole -and [Environment]::UserInteractive)
{
	try
	{
		$consoleHandle = [ConsoleWindowHelper]::GetConsoleWindow()
		if ($consoleHandle -ne [IntPtr]::Zero)
		{
			[void][ConsoleWindowHelper]::ShowWindow($consoleHandle, 0)
			$script:ConsoleHidden = $true
		}
	}
	catch { }
}

[void]$form.ShowDialog()

# Indicate the script is closing
Write-Host "Script closing..." -ForegroundColor Yellow

# Close the console to aviod duplicate logging to the richbox
exit
