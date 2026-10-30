# ===================================================================================================
#                           SECTION: Import Necessary Assemblies and Modules
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports required .NET assemblies for creating and managing Windows Forms and graphical components
#   and imports necessary PowerShell modules required for the script's operation.
# ===================================================================================================

# Add the WinForms assembly early so the elevation gate can show a warning dialog if needed.
Add-Type -AssemblyName System.Windows.Forms

# No hidden console: the console is no longer hidden through user32 ShowWindow (that needed C# compiled
# at run time, and hiding a window is a classic malware trait antivirus scores). The launcher and
# the elevated relaunch start the console minimized instead - visible in the taskbar.

# Prefer the standard WinForms startup order, but don't fail if the current PowerShell
# session has already created a window handle before MiniGhost starts.
[System.Windows.Forms.Application]::EnableVisualStyles()
try
{
	[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
}
catch
{
	$innerException = $_.Exception.InnerException
	$alreadyInitialized = ($innerException -is [System.InvalidOperationException]) -or ($_.Exception.Message -like '*SetCompatibleTextRenderingDefault must be called before the first IWin32Window object is created*')
	if (-not $alreadyInitialized)
	{
		throw
	}
}
