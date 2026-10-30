# ===================================================================================================
#                                   SECTION: Initialize Variables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Initializes all necessary variables and paths required for the script's operation, including
#   dynamic detection of the main Storeman folder, Office subpaths, INI files, encoding settings,
#   counters, and interop types for MailSlot messaging.
# ===================================================================================================

# ---------------------------------------------------------------------------------------------------
# Script-Scoped Result and Process Tracking Structures
# ---------------------------------------------------------------------------------------------------
$script:FunctionResults = @{ }

# Get the current machine name
$currentMachineName = $env:COMPUTERNAME

# Initialize script-scoped variables for new store number and new machine name
$script:newStoreNumber = $null
$script:newMachineName = $null
$script:NodeRole = "Unknown"

# ---------------------------------------------------------------------------------------------------
# Encoding Settings
# ---------------------------------------------------------------------------------------------------
$script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252) # Windows-1252 legacy files
$script:utf8NoBOM = New-Object System.Text.UTF8Encoding($false) # UTF-8 no BOM (for output)
$script:utf8NoBOM = $utf8NoBOM
$script:ansiPcEncoding = $ansiPcEncoding

# ---------------------------------------------------------------------------------------------------
# Locate Base Path: Storeman Folder Detection (case-insensitive)
# ---------------------------------------------------------------------------------------------------
$BasePath = $null
$targetSubPathPattern = 'Office\Dbs\INFO_*901_WIN.INI'
$storemanDirs = @()
$fixedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 -and $_.Root -match '^[A-Z]:\\$' }

foreach ($drive in $fixedDrives)
{
	# Case-insensitive match for any *storeman* variation in directory name
	$dirs = Get-ChildItem -Path "$($drive.Root)" -Directory -ErrorAction SilentlyContinue |
	Where-Object { $_.Name -imatch 'storeman' } |
	ForEach-Object {
		$candidatePath = Join-Path $_.FullName '\'
		$files = Get-ChildItem -Path $candidatePath -Filter 'Startup.ini' -ErrorAction SilentlyContinue
		if ($files) { $_ }
	}
	if ($dirs) { $storemanDirs += $dirs }
}

if ($storemanDirs.Count -gt 1)
{
	# Prefer a path that is actually shared as a Windows share
	$shares = Get-SmbShare -ErrorAction SilentlyContinue
	foreach ($dir in $storemanDirs)
	{
		if ($shares.Path -contains $dir.FullName)
		{
			$BasePath = $dir.FullName
			break
		}
	}
	if (-not $BasePath) { $BasePath = $storemanDirs[0].FullName }
}
elseif ($storemanDirs.Count -eq 1)
{
	$BasePath = $storemanDirs[0].FullName
}

# Final fallback: Default to C:\storeman if none found
if (-not $BasePath)
{
	$fallback = "C:\storeman"
	$candidatePath = Join-Path $fallback '\'
	$files = Get-ChildItem -Path $candidatePath -Filter 'Startup.ini' -ErrorAction SilentlyContinue
	if ($files) { $BasePath = $fallback }
	else
	{
		Write-Warning "Could not locate any storeman folder containing 'storeman\Startup.ini'.`nRunning with limited functionality."
		$BasePath = $fallback
	}
}

Write-Host "Selected (storeman) folder: '$BasePath'" -ForegroundColor Magenta

# ---------------------------------------------------------------------------------------------------
# Build All Core Paths and File Locations
# ---------------------------------------------------------------------------------------------------

# Storeman root paths
$OfficePath = Join-Path $BasePath "Office"
$LoadPath = Join-Path $OfficePath "Load"
$StartupIniPath = Join-Path $BasePath "Startup.ini"
$GlobalSmsStartIniPath = Join-Path $BasePath "SMSStart.ini"
$SystemIniPath = Join-Path $OfficePath "system.ini"
$GasInboxPath = Join-Path $OfficePath "XchGAS\INBOX"
$DbsPath = Join-Path $OfficePath "Dbs"
# SmsHttps.ini (auto)
$SmsHttpsIniPath = Join-Path $BasePath "SmsHttps64\SmsHttps.INI"
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = Join-Path $BasePath "SmsHttps64\SmsHttps.ini" }
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = Join-Path $BasePath "SmsHttps32\SmsHttps.ini" }
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = Join-Path $BasePath "SmsHttps\SmsHttps.ini" }
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = $null }

$TempDir = [System.IO.Path]::GetTempPath()

# Initialize variables for the INFO_*901 files
$WinIniPath = $null
$SmsStartIniPath = $null # INFO_*901_SMSStart.ini inside Dbs

# ---------------------------------------------------------------------------------------------------
# Find INFO_*901_WIN.INI
# ---------------------------------------------------------------------------------------------------
try
{
	$WinIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_WIN.INI' -ErrorAction Stop |
	Select-Object -First 1
	if ($WinIniMatch)
	{
		$WinIniPath = $WinIniMatch.FullName
	}
}
catch { }

# ---------------------------------------------------------------------------------------------------
# Find INFO_*901_SMSStart.ini
# ---------------------------------------------------------------------------------------------------
try
{
	$SmsStartIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_SMSStart.ini' -ErrorAction Stop |
	Select-Object -First 1
	
	if ($SmsStartIniMatch)
	{
		$SmsStartIniPath = $SmsStartIniMatch.FullName
	}
}
catch { }

# Initialize a hashtable to track the status of each operation
$operationStatus = @{
	"StoreNumberChange" = @{ Status = "Pending"; Message = ""; Details = "" }
	"MachineNameChange" = @{ Status = "Pending"; Message = ""; Details = "" }
	"OldXFoldersDeletion" = @{ Status = "Pending"; Message = ""; Details = "" }
	"StartupIniUpdate"  = @{ Status = "Pending"; Message = ""; Details = "" }
	"IPConfiguration"   = @{ Status = "Pending"; Message = ""; Details = "" }
	"TableTruncation"   = @{ Status = "Pending"; Message = ""; Details = "" }
	"RegistryCleanup"   = @{ Status = "Pending"; Message = ""; Details = "" }
	"SQLDatabaseUpdate" = @{ Status = "Pending"; Message = ""; Details = "" }
}
