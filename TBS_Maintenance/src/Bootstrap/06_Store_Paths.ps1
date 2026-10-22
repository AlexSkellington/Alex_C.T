# ---------------------------------------------------------------------------------------------------
# Build All Core Paths and File Locations
# ---------------------------------------------------------------------------------------------------
$script:BasePath = $BasePath
$script:FunctionResults['StoremanPath'] = $BasePath
$OfficePath = Join-Path $BasePath "Office"
$script:OfficePath = $OfficePath
$script:StoremanScriptsFolder = Join-Path $BasePath "Scripts_by_Alex_C.T"
$LoadPath = Join-Path $OfficePath "Load"
$StartupIniPath = Join-Path $BasePath "Startup.ini"
$SystemIniPath = Join-Path $OfficePath "system.ini"
$GasInboxPath = Join-Path $OfficePath "XchGAS\INBOX"
$DbsPath = Join-Path $OfficePath "Dbs"
$TempDir = [System.IO.Path]::GetTempPath()

# Find first INFO_*901_WIN.INI and INFO_*901_SMSStart.ini in Office\Dbs
$WinIniPath = $null
$SmsStartIniPath = $null

$WinIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($WinIniMatch) { $WinIniPath = $WinIniMatch.FullName }
$SmsStartIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_SMSStart.ini' -ErrorAction SilentlyContinue | Select-Object -First 1
if ($SmsStartIniMatch) { $SmsStartIniPath = $SmsStartIniMatch.FullName }

# SQI temporary output file paths (used by maintenance routines)
$LanesqlFilePath = Join-Path $TempDir "Lane_Database_Maintenance.sqi"
$StoresqlFilePath = Join-Path $TempDir "Server_Database_Maintenance.sqi"

# Local machine name
$script:LocalHost = $env:COMPUTERNAME

# ---------------------------------------------------------------------------------------------------
# (Optional) Script Name Extraction
# ---------------------------------------------------------------------------------------------------
# $scriptName = Split-Path -Leaf $PSCommandPath

$script:SqlcmdSupportsConnectionString = $null
