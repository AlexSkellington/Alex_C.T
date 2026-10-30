# ===================================================================================================
#                                    SECTION: Version and Host
# ---------------------------------------------------------------------------------------------------
# Description:
#   Script version (shown in the status bar and the file log), the startup clock, the
#   PowerShell version string, the idle-close timeout and the process execution policy.
#
# Variables:
#   - $VersionNumber / $VersionDate   build version (consult with Alex_C.T before changing)
#   - $script:StartupT0               time zero for the "window shown in Xs" startup log lines
#   - $PowerShellVersion              "major.minor.build.revision" of the running host
#   - $script:IdleMinutesAllowed      idle minutes before the window closes itself (0 = never)
# ===================================================================================================

# Script build version (consult with Alex_C.T before changing this)
$VersionNumber = "2.7.8"
$VersionDate = "2026-10-30"

$script:StartupT0 = Get-Date

$PowerShellVersion = "{0}.{1}.{2}.{3}" -f $PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor, $PSVersionTable.PSVersion.Build, $PSVersionTable.PSVersion.Revision

# Idle timeout for the whole script
$script:IdleMinutesAllowed = 15 # <<< adjust as needed

# Bypass for this process only. A Group Policy that pins the policy makes this throw even
# though the script already runs - nothing to do about it, so stay quiet.
try { Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop }
catch { }
