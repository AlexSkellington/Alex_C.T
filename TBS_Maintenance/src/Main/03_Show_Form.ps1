# Clear %Temp% folder on start
# $ClearTempAtLaunch = Delete_Files -Path "$TempDir" -Exclusions "Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1" -AsJob
# $ClearWinTempAtLaunch = Delete_Files -Path "$env:SystemRoot\Temp" -AsJob

# Indicate the script has started
Write-Host "Script started" -ForegroundColor Green

if (-not $script:ShowConsole -and [Environment]::UserInteractive)
{
	try
	{
		if (-not $script:PostOpenConsoleHideTimer)
		{
			$script:PostOpenConsoleHideTimer = New-Object System.Windows.Forms.Timer
			$script:PostOpenConsoleHideTimer.Interval = 400
			$script:PostOpenConsoleHideTimer.Add_Tick({
				try
				{
					if ($script:PostOpenConsoleHideTimer)
					{
						$script:PostOpenConsoleHideTimer.Stop()
					}
				}
				catch { }
				try
				{
					$__consoleHandle = [ConsoleWindowHelper]::GetConsoleWindow()
					if ($__consoleHandle -ne [IntPtr]::Zero)
					{
						[void][ConsoleWindowHelper]::ShowWindow($__consoleHandle, 0)
						$script:ConsoleHidden = $true
						$script:RestoreConsoleOnExit = $false
					}
				}
				catch { }
				try
				{
					if ($script:PostOpenConsoleHideTimer)
					{
						$script:PostOpenConsoleHideTimer.Dispose()
						$script:PostOpenConsoleHideTimer = $null
					}
				}
				catch { }
			})
			$form.Add_Shown({
				try
				{
					if ($script:PostOpenConsoleHideTimer)
					{
						$script:PostOpenConsoleHideTimer.Start()
					}
				}
				catch { }
			})
		}
	}
	catch { }
}

# ===================================================================================================
#                                       SECTION: Show the GUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays the main form to the user and manages the script's execution flow based on user interactions.
# ===================================================================================================

[void]$form.ShowDialog()

# Indicate the script is closing
if ($script:RestoreConsoleOnExit -and $script:ShowConsole)
{
	try
	{
		$__consoleHandle = [ConsoleWindowHelper]::GetConsoleWindow()
		if ($__consoleHandle -ne [IntPtr]::Zero)
		{
			[void][ConsoleWindowHelper]::ShowWindow($__consoleHandle, 5)
		}
	}
	catch { }
}

if ($script:ShowConsole) { Write-Host "Script closing..." -ForegroundColor Yellow }

# Close the console to avoid duplicate logging to the richbox
exit
