# ===================================================================================================
#                           SECTION: Import Necessary Assemblies and Modules
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports required .NET assemblies for creating and managing Windows Forms and graphical components
#   and imports necessary PowerShell modules required for the script's operation.
# ===================================================================================================

# Add the WinForms assembly early so the elevation gate can show a warning dialog if needed.
Add-Type -AssemblyName System.Windows.Forms

if (-not ([System.Management.Automation.PSTypeName]'ConsoleWindowHelper').Type)
{
	Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class ConsoleWindowHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
}

$script:ShowConsole = [bool](@($args | Where-Object { $_ -in @('-ShowConsole', '/ShowConsole', '-KeepConsole', '/KeepConsole') }).Count -gt 0)
$script:ConsoleHidden = $false

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
