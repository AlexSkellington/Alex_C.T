# ===================================================================================================
#                                    SECTION: Store Paths
# ---------------------------------------------------------------------------------------------------
# Description:
#   Derives the storeman sub-paths and INI file locations from the resolved $BasePath.
#
# Variables:
#   - $script:BasePath / $script:OfficePath   (+ unscoped copies: older code reads both forms)
#   - $LoadPath, $StartupIniPath, $SystemIniPath, $DbsPath, $TempDir
#   - $WinIniPath       first Office\Dbs\INFO_*901_WIN.INI       ($null if none)
#   - $SmsStartIniPath  first Office\Dbs\INFO_*901_SMSStart.ini  ($null if none)
#   - $StoresqlFilePath %TEMP%\Server_Database_Maintenance.sqi (Process_Server / scheduling)
#   - $script:LocalHost this machine's name
# ===================================================================================================

$script:BasePath = $BasePath
$script:FunctionResults['StoremanPath'] = $BasePath
$OfficePath = Join-Path $BasePath "Office"
$script:OfficePath = $OfficePath
$LoadPath = Join-Path $OfficePath "Load"
$StartupIniPath = Join-Path $BasePath "Startup.ini"
$SystemIniPath = Join-Path $OfficePath "system.ini"
$DbsPath = Join-Path $OfficePath "Dbs"
$TempDir = [System.IO.Path]::GetTempPath()

$WinIniPath = $null
$SmsStartIniPath = $null
if (Test-Path -LiteralPath $DbsPath)
{
	$WinIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue | Select-Object -First 1
	if ($WinIniMatch) { $WinIniPath = $WinIniMatch.FullName }
	$SmsStartIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_SMSStart.ini' -ErrorAction SilentlyContinue | Select-Object -First 1
	if ($SmsStartIniMatch) { $SmsStartIniPath = $SmsStartIniMatch.FullName }
}

$StoresqlFilePath = Join-Path $TempDir "Server_Database_Maintenance.sqi"

$script:LocalHost = $env:COMPUTERNAME
