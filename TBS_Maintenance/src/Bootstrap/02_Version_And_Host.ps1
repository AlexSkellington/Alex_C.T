# ===================================================================================================
#                                       SECTION: Parameters
# ---------------------------------------------------------------------------------------------------
# Description:
#   Defines the script parameters, allowing users to run the script in silent mode.
# ===================================================================================================

# Script build version (cunsult with Alex_C.T before changing this)
$VersionNumber = "2.6.9"
$VersionDate = "2026-09-01"

# Startup milestone logging (v2.6.9): time zero for the "window shown in Xs"
# and warmup-step durations that go to the Tecnica_Systems log folder.
$script:StartupT0 = Get-Date

# Retrieve Major, Minor, Build, and Revision version numbers of PowerShell
$major = $PSVersionTable.PSVersion.Major
$minor = $PSVersionTable.PSVersion.Minor
$build = $PSVersionTable.PSVersion.Build
$revision = $PSVersionTable.PSVersion.Revision

# Idle timeout for the whole script
$script:IdleMinutesAllowed = 15 # <<< adjust as needed
$script:SuppressClosePrompt = $true
if (-not $script:LaunchSourceType) { $script:LaunchSourceType = 'Unknown' }
if ($null -eq $script:LaunchEntryPath) { $script:LaunchEntryPath = $null }
if ($null -eq $script:IsInlineLaunch) { $script:IsInlineLaunch = $false }
if ($null -eq $script:ShowConsole) { $script:ShowConsole = $false }
if ($null -eq $script:ConsoleHidden) { $script:ConsoleHidden = $false }
if ($null -eq $script:RestoreConsoleOnExit) { $script:RestoreConsoleOnExit = $false }

# Combine them into a single version string
$PowerShellVersion = "$major.$minor.$build.$revision"

# Set Execution Policy to Bypass for the current process
Set-ExecutionPolicy Bypass -Scope Process -Force
