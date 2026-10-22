#######################################################################################################
#                                                                                                     #
#                                     TBS MAINTENANCE SCRIPT                                          #
#                                                                                                     #
#                                        Author: Alex_C.T                                             #
#                                                                                                     #
#  > Edit only in consultation with Alex_C.T                                                          #
#  > This script performs advanced maintenance and diagnostics on TBS systems                         #
#                                                                                                     #
#  > GENERATED FILE - edit the sources in TBS_Maintenance\src and run                                 #
#    TBS_Maintenance\build\Build.ps1 (see TBS_Maintenance\README.md)                                  #
#                                                                                                     #
#######################################################################################################

# ===================================================================================================
#                                     SECTION: Launch Detection
# ---------------------------------------------------------------------------------------------------
# Description:
#   Works out HOW the script was started, before anything else runs:
#     Ps1    - a saved .ps1 file (powershell -File ..., "Run with PowerShell")
#     Exe    - the PowerShell Studio package (TBS_Maintenance_Script_V*.exe)
#     Inline - irm https://bit.ly/TBSMaintenanceScript | iex, a paste, or an ISE selection
#   and reads the optional command-line switches:
#     -ShowConsole / -KeepConsole      keep the console window visible (troubleshooting)
#     -LaneSqlEnableWorker             headless worker mode, with -LaneSqlEnableWorkerLane,
#                                      -LaneSqlEnableWorkerMachine, -LaneSqlEnableWorkerTcpPort and
#                                      -LaneSqlEnableWorkerStartTime (used by the scheduled task
#                                      that Ensure_SQL_Enable_Scheduled_Task_For_File_Lane creates)
#
# Variables:
#   - $script:LaunchSourceType       'Ps1' | 'Exe' | 'Inline'
#   - $script:LaunchEntryPath        the .ps1 / .exe that was started ($null for Inline)
#   - $script:IsInlineLaunch         $true for Inline launches
#   - $script:ShowConsole            -ShowConsole / -KeepConsole was given
#   - $script:LaneSqlEnableWorker*   worker-mode switch + its four values
#   - $__scriptPath, $__entryPath, $__launchArgs, $__isSavedPs1Launch, $__isExeLaunch
#                                    read by the elevation step (03_Elevation)
# ===================================================================================================

# Path of the running .ps1 (empty for Inline launches)
$__scriptPath = $null
try
{
	$__scriptPath = $PSCommandPath
	if ([string]::IsNullOrWhiteSpace($__scriptPath))
	{
		try { $__scriptPath = $MyInvocation.MyCommand.Path }
		catch { $__scriptPath = $null }
	}
}
catch { $__scriptPath = $null }

# Process entry point (powershell.exe / the packaged .exe) and its arguments
$__entryPath = $null
$__launchArgs = @()
try
{
	$__cmdLineArgs = [Environment]::GetCommandLineArgs()
	if ($__cmdLineArgs -and $__cmdLineArgs.Count -gt 0)
	{
		$__entryPath = $__cmdLineArgs[0]
		if ($__cmdLineArgs.Count -gt 1) { $__launchArgs = @($__cmdLineArgs[1 .. ($__cmdLineArgs.Count - 1)]) }
	}
}
catch
{
	$__entryPath = $null
	$__launchArgs = @()
}

# Saved .ps1?
$__isSavedPs1Launch = $false
try
{
	$__isSavedPs1Launch = (-not [string]::IsNullOrWhiteSpace($__scriptPath) -and
		(Test-Path -LiteralPath $__scriptPath) -and
		([string]::Equals([System.IO.Path]::GetExtension($__scriptPath), '.ps1', [System.StringComparison]::OrdinalIgnoreCase)))
}
catch { $__isSavedPs1Launch = $false }

# Packaged .exe? (a PowerShell host .exe running us inline does not count)
$__isExeLaunch = $false
if (-not $__isSavedPs1Launch)
{
	try
	{
		if (-not [string]::IsNullOrWhiteSpace($__entryPath) -and
			(Test-Path -LiteralPath $__entryPath) -and
			([string]::Equals([System.IO.Path]::GetExtension($__entryPath), '.exe', [System.StringComparison]::OrdinalIgnoreCase)))
		{
			$__entryName = [System.IO.Path]::GetFileNameWithoutExtension($__entryPath)
			$__isExeLaunch = ($__entryName -notmatch '^(powershell|pwsh|powershell_ise|powershellstudio)$')
		}
	}
	catch { $__isExeLaunch = $false }
}

if ($__isSavedPs1Launch)
{
	$script:LaunchSourceType = 'Ps1'
	$script:LaunchEntryPath = $__scriptPath
}
elseif ($__isExeLaunch)
{
	$script:LaunchSourceType = 'Exe'
	$script:LaunchEntryPath = $__entryPath
}
else
{
	$script:LaunchSourceType = 'Inline'
	$script:LaunchEntryPath = $null
}
$script:IsInlineLaunch = ($script:LaunchSourceType -eq 'Inline')

# ---------------------------------------------------------------------------------------------------
# Command-line switches
# ---------------------------------------------------------------------------------------------------
$script:ShowConsole = [bool](@($__launchArgs | Where-Object { $_ -in @('-ShowConsole', '/ShowConsole', '-KeepConsole', '/KeepConsole') }).Count -gt 0)
$script:LaneSqlEnableWorkerMode = $false
$script:LaneSqlEnableWorkerLane = $null
$script:LaneSqlEnableWorkerMachine = $null
$script:LaneSqlEnableWorkerTcpPort = '1433'
$script:LaneSqlEnableWorkerStartTime = '04:15'

for ($__argIndex = 0; $__argIndex -lt $__launchArgs.Count; $__argIndex++)
{
	$__argValue = [string]$__launchArgs[$__argIndex]
	$__hasValue = (($__argIndex + 1) -lt $__launchArgs.Count)
	switch -Regex ($__argValue)
	{
		'^-LaneSqlEnableWorker$' { $script:LaneSqlEnableWorkerMode = $true }
		'^-LaneSqlEnableWorkerLane$' { if ($__hasValue) { $__argIndex++; $script:LaneSqlEnableWorkerLane = [string]$__launchArgs[$__argIndex] } }
		'^-LaneSqlEnableWorkerMachine$' { if ($__hasValue) { $__argIndex++; $script:LaneSqlEnableWorkerMachine = [string]$__launchArgs[$__argIndex] } }
		'^-LaneSqlEnableWorkerTcpPort$' { if ($__hasValue) { $__argIndex++; $script:LaneSqlEnableWorkerTcpPort = [string]$__launchArgs[$__argIndex] } }
		'^-LaneSqlEnableWorkerStartTime$' { if ($__hasValue) { $__argIndex++; $script:LaneSqlEnableWorkerStartTime = [string]$__launchArgs[$__argIndex] } }
	}
}

# ---------------------------------------------------------------------------------------------------
# Console window helper (the console is hidden once the main window is up - see Main\03_Show_Form)
# ---------------------------------------------------------------------------------------------------
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

Write-Host "Script starting, pls wait..." -ForegroundColor Yellow
