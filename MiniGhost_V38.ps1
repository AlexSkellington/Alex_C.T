#######################################################################################################
#                                                                                                     #
#                                        MiniGhost SCRIPT                                             #
#                                                                                                     #
#                                        Author: Alex_C.T                                             #
#                                                                                                     #
#  > Edit only in consultation with Alex_C.T                                                          #
#  > This script performs advanced maintenance and diagnostics on TBS systems                         #
#                                                                                                     #
#######################################################################################################

Write-Host "Script starting, pls wait..." -ForegroundColor Yellow

# ===================================================================================================
#                                       SECTION: Parameters
# ---------------------------------------------------------------------------------------------------
# Description:
#   Defines the script parameters, allowing users to run the script in silent mode.
# ===================================================================================================

# Script build version (cunsult with Alex_C.T before changing this)
$VersionNumber = "1.3.0"
$VersionDate = "2026-01-07"

# Retrieve Major, Minor, Build, and Revision version numbers of PowerShell
$major = $PSVersionTable.PSVersion.Major
$minor = $PSVersionTable.PSVersion.Minor
$build = $PSVersionTable.PSVersion.Build
$revision = $PSVersionTable.PSVersion.Revision

# Combine them into a single version string
$PowerShellVersion = "$major.$minor.$build.$revision"

# Set Execution Policy to Bypass for the current process
Set-ExecutionPolicy Bypass -Scope Process -Force

# ===================================================================================================
#                           SECTION: Import Necessary Assemblies and Modules
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports required .NET assemblies for creating and managing Windows Forms and graphical components
#   and imports necessary PowerShell modules required for the script's operation.
# ===================================================================================================

# Import necessary modules
Import-Module -Name Microsoft.PowerShell.Utility

# Add necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Enable modern visual styles for WinForms
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

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

# ===================================================================================================
#                               FUNCTION: Get_Database_Connection_String
# ---------------------------------------------------------------------------------------------------
# Description:
#   Searches for the Startup.ini file in specified locations, extracts the DBNAME value,
#   constructs the connection string, and stores it in a script-level hashtable.
# ===================================================================================================

function Get_Database_Connection_String
{		
	if (-not $startupIniPath)
	{
		return
	}
	
	# Read the Startup.ini file
	try
	{
		$content = Get-Content -Path $startupIniPath -ErrorAction Stop
		
		# Extract DBSERVER
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }
		if ($dbServerLine)
		{
			$dbServer = $dbServerLine -replace '^DBSERVER=', ''
			$dbServer = $dbServer.Trim()
			if (-not $dbServer)
			{
				$dbServer = "localhost"
			}
		}
		else
		{
			$dbServer = "localhost"
		}
		
		# Extract DBNAME
		$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' }
		if ($dbNameLine)
		{
			$dbName = $dbNameLine -replace '^DBNAME=', ''
			$dbName = $dbName.Trim()
			if (-not $dbName)
			{
				return
			}
		}
		else
		{
			return
		}
	}
	catch
	{
		return
	}
	
	# Store DBSERVER and DBNAME in the FunctionResults hashtable
	$script:FunctionResults['DBSERVER'] = $dbServer
	$script:FunctionResults['DBNAME'] = $dbName
	
	# Build the connection string
	$connectionString = "Server=$dbServer;Database=$dbName;Integrated Security=True;"
	
	# Store the connection string in the FunctionResults hashtable
	$script:FunctionResults['ConnectionString'] = $connectionString
}

# ===================================================================================================
#                                       FUNCTION: Get_Store_Number_From_INI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store number from the startup.ini file.
# ===================================================================================================

function Get_Store_Number_From_INI
{
	[CmdletBinding()]
	param (
		[switch]$UpdateLabel
	)
	
	# Ensure FunctionResults exists (PS 5.1 safe)
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	
	# Default/fallback
	$script:FunctionResults['StoreNumber'] = "N/A"
	$foundStore = $null
	
	if (Test-Path $startupIniPath)
	{
		try
		{
			$iniContent = Get-Content $startupIniPath -ErrorAction Stop
			
			foreach ($line in $iniContent)
			{
				if ($line -match '^\s*STORE\s*=\s*(\d{3,4})\s*$')
				{
					$foundStore = $matches[1]
					$script:FunctionResults['StoreNumber'] = $foundStore
					break
				}
			}
		}
		catch
		{
			Write-Warning ("Failed to read startup.ini: {0}" -f $_)
		}
	}
	
	# Update label ONLY if requested
	if ($UpdateLabel -and (-not $SilentMode) -and $script:storeNumberLabel)
	{
		$script:storeNumberLabel.Text = "Store Number: $($script:FunctionResults['StoreNumber'])"
		$script:storeNumberLabel.Refresh()
		
		if ($script:storeNumberLabel.Parent)
		{
			$script:storeNumberLabel.Parent.PerformLayout()
			$script:storeNumberLabel.Parent.Refresh()
		}
		
		[System.Windows.Forms.Application]::DoEvents()
	}
	
	if ($foundStore) { return $foundStore }
	return $null
}

# ===================================================================================================
#                                      FUNCTION: Get_Store_Name_From_INI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store name from the system.ini file.
#   Stores the result in $script:FunctionResults['StoreName'].
# ===================================================================================================

function Get_Store_Name_From_INI
{
	param (
		[string]$INIPath = $SystemIniPath
	)
	
	# Initialize StoreName
	$script:FunctionResults['StoreName'] = "N/A"
	
	if (Test-Path $INIPath)
	{
		$storeName = Select-String -Path $INIPath -Pattern "^NAME=" | ForEach-Object {
			$_.Line.Split('=')[1].Trim()
		}
		if ($storeName)
		{
			$script:FunctionResults['StoreName'] = $storeName
			# Write-Log "Store name found in system.ini: $storeName" "green"
		}
		else
		{
			# Write-Log "Store name not found in system.ini." "yellow"
		}
	}
	else
	{
		# Write-Log "INI file not found: $INIPath" "yellow"
	}
	
	# Update the storeNameLabel in the GUI
	if (-not $SilentMode -and $storeNameLabel -ne $null)
	{
		$storeNameLabel.Text = "Store Name: $($script:FunctionResults['StoreName'])"
		$form.Refresh()
		[System.Windows.Forms.Application]::DoEvents()
	}
}

# ===================================================================================================
#                                       FUNCTION: Get_Active_IP_Config
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the active IP configuration for network adapters that are up and have valid IPv4 addresses.
#   Optimized for performance by prefiltering and minimizing pipeline overhead.
#   Attempts WMI first for compatibility, falls back to NetAdapter methods if WMI fails.
# ===================================================================================================

function Get_Active_IP_Config
{
	try
	{
		# WMI attempt: Get default route from Win32_IP4RouteTable
		$defaultRoute = Get-WmiObject -Class Win32_IP4RouteTable -ErrorAction Stop |
		Where-Object { $_.Destination -eq '0.0.0.0' -and $_.Mask -eq '0.0.0.0' -and $_.NextHop -ne '0.0.0.0' } |
		Sort-Object -Property Metric1 |
		Select-Object -First 1
		
		if ($defaultRoute -and $defaultRoute.InterfaceIndex)
		{
			$idx = [int]$defaultRoute.InterfaceIndex
			
			# Check adapter status via Win32_NetworkAdapter (NetConnectionStatus 2 = Connected)
			$ad = Get-WmiObject -Class Win32_NetworkAdapter -Filter "DeviceID = $idx" -ErrorAction SilentlyContinue
			if ($ad -and $ad.NetConnectionStatus -eq 2)
			{
				# Get config via Win32_NetworkAdapterConfiguration
				$wmiCfg = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "Index = $idx" -ErrorAction SilentlyContinue
				if ($wmiCfg -and $wmiCfg.IPEnabled -and $wmiCfg.DefaultIPGateway)
				{
					# Filter for valid IPv4 (non-APIPA)
					$validIPs = $wmiCfg.IPAddress | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' -and $_ -notmatch '^169\.254\.' }
					if ($validIPs)
					{
						# Construct a custom object similar to Get-NetIPConfiguration for consistency
						$ipv4Address = [PSCustomObject]@{
							IPAddress    = $validIPs[0]
							PrefixLength = ($wmiCfg.IPSubnet | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' })[0] | ForEach-Object {
								[int]([math]::Log([uint32]::MaxValue -bxor ([ipaddress]$_).Address, 2)) + 1
							}
						}
						$cfg = [PSCustomObject]@{
							InterfaceAlias	   = $ad.NetConnectionID
							InterfaceIndex	   = $idx
							IPv4Address	       = $ipv4Address
							IPv4DefaultGateway = [PSCustomObject]@{ NextHop = $wmiCfg.DefaultIPGateway[0] }
							DNSServer		   = $wmiCfg.DNSServerSearchOrder
							# Add more properties if needed for parity
						}
						return @($cfg) # Return as array for consistent .Count / [0]
					}
				}
			}
		}
	}
	catch
	{
		# Ignore and fall back to original methods
	}
	
	# =========================================================================================
	# FALLBACK: Original approach (NetRoute/NetAdapter/NetIPConfiguration)
	# =========================================================================================
	
	try
	{
		$defaultRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 -ErrorAction Stop |
		Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } |
		Sort-Object -Property RouteMetric, InterfaceMetric |
		Select-Object -First 1
		
		if ($defaultRoute -and $defaultRoute.InterfaceIndex)
		{
			$idx = [int]$defaultRoute.InterfaceIndex
			
			# Ensure adapter is Up (do NOT restrict to -Physical here; active interface might be Wi-Fi/VPN/etc.)
			$ad = Get-NetAdapter -InterfaceIndex $idx -ErrorAction SilentlyContinue
			if ($ad -and $ad.Status -eq 'Up')
			{
				# Make sure there is at least one valid IPv4 (non-APIPA) on this interface
				$ipRow = Get-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -ErrorAction SilentlyContinue |
				Where-Object { $_.IPAddress -and $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '0.0.0.0' } |
				Select-Object -First 1
				
				if ($ipRow)
				{
					$cfg = Get-NetIPConfiguration -InterfaceIndex $idx -ErrorAction SilentlyContinue
					if ($cfg -and $cfg.IPv4Address)
					{
						# Validate IPv4Address (can be array). If any non-APIPA exists, return it.
						$hasValid = $false
						foreach ($addr in @($cfg.IPv4Address))
						{
							if ($addr -and $addr.IPAddress -and ($addr.IPAddress -notlike '169.254.*'))
							{
								$hasValid = $true
								break
							}
						}
						
						if ($hasValid)
						{
							return @($cfg) # return as array for consistent .Count / [0]
						}
					}
				}
			}
		}
	}
	catch
	{
		# ignore and fall back
	}
	
	# =========================================================================================
	# SECONDARY FALLBACK: Physical + Up adapters -> Get-NetIPConfiguration
	# =========================================================================================
	
	# Prefilter by adapter status to reduce objects early
	$adapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object Status -eq 'Up'
	if (-not $adapters) { return $null }
	
	# Get only configs for up adapters
	$adapterNames = $adapters | Select-Object -ExpandProperty Name
	$ipConfigs = Get-NetIPConfiguration -InterfaceAlias $adapterNames -ErrorAction SilentlyContinue
	
	# Filter for valid IPv4 (not APIPA/169.254.x.x, not null)
	$validConfigs = $ipConfigs | Where-Object {
		$_.IPv4Address -and
		$_.IPv4Address.IPAddress -and
		($_.IPv4Address.IPAddress -notlike '169.254*')
	}
	
	if ($validConfigs) { return @($validConfigs) }
	return $null
}

# ===================================================================================================
# FUNCTION: Remove_Old_XF/XW_Folders
# ---------------------------------------------------------------------------------------------------
# Description:
# Removes old XF and XW folders based on the provided store number and machine name.
# Supports both:
#   - 6-digit format: XF123456   (3 digit store + 3 digit terminal)
#   - 7-digit format: XF1234006  (4 digit store + 3 digit terminal)
# The machine number is always extracted from the last 3 characters of $MachineName.
# ===================================================================================================

function Remove_Old_XF/XW_Folders
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		# add this because your code writes to it
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus
	)
	
	# Define prefixes to process
	$folderPrefixes = @("XF", "XW")
	
	# Initialize results
	$deletedFolders = @()
	$failedToDeleteFolders = @()
	
	# Status tracking
	$anyFoldersSeen = $false
	$anyCandidatesMatched = $false
	
	# Validate store number (3 or 4 digits, not all zeros)
	$storeNumberTrim = $StoreNumber
	if ($null -eq $storeNumberTrim) { $storeNumberTrim = "" }
	$storeNumberTrim = $storeNumberTrim.Trim()
	
	if ($storeNumberTrim -notmatch '^(?!0{3,4})\d{3,4}$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Invalid StoreNumber '$StoreNumber' (must be 3 or 4 digits, not all zeros)."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: Invalid StoreNumber '$StoreNumber'." -ForegroundColor Red
		return
	}
	
	# Normalize store to numeric + both 3/4 digit strings for matching folders
	$storeInt = [int]$storeNumberTrim
	$store3 = $storeInt.ToString("D3")
	$store4 = $storeInt.ToString("D4")
	
	# Define possible base paths in order of priority
	$possibleBasePaths = @("\\localhost\storeman\office", "C:\storeman\office", "D:\storeman\office")
	
	# Find the first existing base path
	$basePath = $null
	foreach ($p in $possibleBasePaths)
	{
		if (Test-Path $p)
		{
			$basePath = $p
			break
		}
	}
	
	if (-not $basePath)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "None of the base paths exist: $($possibleBasePaths -join ', ')"
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: None of the base paths exist: $($possibleBasePaths -join ', ')" -ForegroundColor Red
		return
	}
	
	# Normalize machine name to host only (handles UNC/path/FQDN), then extract last 1-3 digits and pad to 3
	$mn = $MachineName
	if ($null -eq $mn) { $mn = "" }
	$mn = $mn.Trim()
	$mn = $mn -replace '^[\\\/]+', '' # strip leading \\ or /
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] } # keep host portion
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] } # drop domain
	$mn = $mn.Trim().ToUpper()
	
	if ($mn -notmatch '(\d{1,3})$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "MachineName '$MachineName' does not end with 1-3 digits; cannot extract terminal number."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: MachineName '$MachineName' cannot provide terminal digits." -ForegroundColor Red
		return
	}
	
	$machineNumber = ([int]$Matches[1]).ToString("D3")
	
	# Safety: prevent wiping all lane folders when machineNumber is 901
	if ($machineNumber -eq "901")
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Refusing to run because extracted terminal number is 901 (backoffice)."
			$OperationStatus["OldXFoldersDeletion"].Details = "This would delete all non-901 XF/XW folders for the store."
		}
		Write-Host "Failed: Refusing to run on terminal 901 (backoffice safeguard)." -ForegroundColor Red
		return
	}
	
	foreach ($prefix in $folderPrefixes)
	{
		# enumerate only XF* or XW*
		$folders = Get-ChildItem -Path $basePath -Directory -Filter ($prefix + "*") -ErrorAction SilentlyContinue
		if (-not $folders) { continue }
		
		$anyFoldersSeen = $true
		
		foreach ($folder in $folders)
		{
			$folderName = $folder.Name
			
			# Match expected pattern
			if ($folderName -match "^(?<prefix>XF|XW)(?<store>\d{3,4})(?<terminal>\d{3})$")
			{
				$folderStore = $matches['store']
				$folderTerminal = $matches['terminal']
				
				# Store match: accept either 3 or 4 digit representation of the SAME store
				$storeMatches = ($folderStore -eq $store3) -or ($folderStore -eq $store4)
				
				if ($storeMatches -and $folderTerminal -ne "901" -and $folderTerminal -ne $machineNumber)
				{
					$anyCandidatesMatched = $true
					
					$maxRetries = 3
					$retryCount = 0
					$deleted = $false
					
					while ($retryCount -lt $maxRetries -and -not $deleted)
					{
						try
						{
							Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
							$deletedFolders += $folderName
							$deleted = $true
						}
						catch
						{
							$retryCount++
							Start-Sleep -Seconds 2
							
							if ($retryCount -ge $maxRetries)
							{
								$failedToDeleteFolders += $folderName
								Write-Host "Failed to delete folder: $folderName. Error: $_" -ForegroundColor Red
							}
						}
					}
				}
			}
			else
			{
				Write-Host "Skipped: Folder '$folderName' does not match expected pattern (XF/XW + 3-4 digits + 3 digits)" -ForegroundColor Yellow
			}
		}
	}
	
	# Build result message
	$resultMessage = ""
	if ($deletedFolders.Count -gt 0)
	{
		$resultMessage += "Deleted folders:`n$($deletedFolders -join "`n")`n"
	}
	if ($failedToDeleteFolders.Count -gt 0)
	{
		$resultMessage += "Failed to delete folders:`n$($failedToDeleteFolders -join "`n")`n"
	}
	
	# Decide outcome
	if (-not $anyFoldersSeen)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Folders Found"
			$OperationStatus["OldXFoldersDeletion"].Message = "No XF/XW folders were found under '$basePath'."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: No XF/XW folders found under '$basePath'." -ForegroundColor Cyan
		return
	}
	
	if (-not $anyCandidatesMatched)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Matching Folders"
			$OperationStatus["OldXFoldersDeletion"].Message = "No old XF/XW folders matched store $StoreNumber (excluding $machineNumber and 901)."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: No matching old XF/XW folders to delete for store $StoreNumber." -ForegroundColor Cyan
		return
	}
	
	if ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Successful"
			$OperationStatus["OldXFoldersDeletion"].Message = "Old XF and XW folders deleted successfully."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Success: Old XF and XW folders deleted successfully." -ForegroundColor Green
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Green }
	}
	elseif ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -gt 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Partial Failure"
			$OperationStatus["OldXFoldersDeletion"].Message = "Some old XF and XW folders could not be deleted."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Warning: Some old XF and XW folders could not be deleted." -ForegroundColor Yellow
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Yellow }
	}
	else
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Failed to delete any old XF and XW folders."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Error: Failed to delete any old XF and XW folders." -ForegroundColor Red
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Red }
	}
}

# ===================================================================================================
#                                       FUNCTION: Execute_SQL_Commands
# ---------------------------------------------------------------------------------------------------
# Description:
#   Executes a given SQL command using the provided connection string.
# ===================================================================================================

function Execute_SQL_Commands
{
	param (
		[string]$commandText
	)
	
	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$sqlConnection.ConnectionString = $connectionString
	$sqlCommand = $sqlConnection.CreateCommand()
	$sqlCommand.CommandText = $commandText
	
	try
	{
		$sqlConnection.Open()
		$sqlCommand.ExecuteNonQuery() | Out-Null
		return $true # Indicate success
	}
	catch
	{
		return $false # Indicate failure
	}
	finally
	{
		$sqlConnection.Close()
	}
}

# ===================================================================================================
# 									FUNCTION: Get_NEW_Store_Number
# ---------------------------------------------------------------------------------------------------
# Description:
# Prompts the user via a GUI to enter a valid store number.
# Accepts exactly 3 or 4 digits (no leading zeros required, not all zeros).
# Returns the entered value as-is (no padding) or $null if cancelled.
# ===================================================================================================

function Get_NEW_Store_Number
{
	while ($true)
	{
		# Decide required store length based on CURRENT store number length (3 vs 4)
		$requiredLen = 4
		
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$cur = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
			if ($cur -match '^\d{3,4}$') { $requiredLen = $cur.Length }
		}
		else
		{
			$iniStore = Get_Store_Number_From_INI # returns as-is (3 or 4) if your function is set that way
			if ($iniStore)
			{
				$iniStore = ($iniStore.ToString()).Trim()
				if ($iniStore -match '^\d{3,4}$') { $requiredLen = $iniStore.Length }
			}
		}
		
		$storeNumberForm = New-Object System.Windows.Forms.Form
		$storeNumberForm.Text = "Enter New Store Number"
		$storeNumberForm.Size = New-Object System.Drawing.Size(350, 180)
		$storeNumberForm.StartPosition = "CenterParent"
		
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "New Store Number (exactly $requiredLen digits, e.g. " + ($(if ($requiredLen -eq 3) { "123" }
				else { "4123" })) + "):"
		$label.Location = New-Object System.Drawing.Point(10, 20)
		$label.Size = New-Object System.Drawing.Size(315, 40)
		$label.AutoSize = $false
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(10, 65)
		$textBox.Size = New-Object System.Drawing.Size(320, 20)
		
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(85, 100)
		$okButton.Size = New-Object System.Drawing.Size(75, 23)
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		
		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Text = "Cancel"
		$cancelButton.Location = New-Object System.Drawing.Point(175, 100)
		$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		
		$storeNumberForm.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))
		$storeNumberForm.AcceptButton = $okButton
		$storeNumberForm.CancelButton = $cancelButton
		
		$dialogResult = $storeNumberForm.ShowDialog()
		
		# IMPORTANT: read before Dispose (PS WinForms reliability)
		$userInput = $textBox.Text
		$storeNumberForm.Dispose()
		
		if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
		{
			return $null
		}
		
		if ($null -eq $userInput) { $userInput = "" }
		$userInput = $userInput.Trim()
		
		if ([string]::IsNullOrEmpty($userInput))
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Store number cannot be empty.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			continue
		}
		
		# Build regex based on required length; reject all zeros for that length
		$re = '^(?!0{' + $requiredLen + '})\d{' + $requiredLen + '}$'
		
		if ($userInput -match $re)
		{
			if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
			$script:FunctionResults['StoreNumber'] = $userInput
			return $userInput # return as-is, no padding
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Invalid store number.`n`nMust be exactly $requiredLen digits (numeric).`nNot allowed: all zeros, too short, too long.`nExample: " + ($(if ($requiredLen -eq 3) { "123" }
						else { "4123" })),
				"Invalid Input",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
		}
	}
}

# ===================================================================================================
# 									FUNCTION: Update_Store_Number_In_INI
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates ALL required INI files with the new store number.
# Supports both 3-digit and 4-digit store numbers.
# NOTE: No backups are created.
# ===================================================================================================

function Update_Store_Number_In_INI
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0{3,4})\d{3,4}$')]
		[string]$newStoreNumber
	)

	$success = $true

	# ===============================================================================================
	# 1) Update Startup.ini (STORE + REDIRs)
	# ===============================================================================================
	if (-not (Test-Path $StartupIniPath))
	{
		Write-Host "startup.ini not found at $StartupIniPath" -ForegroundColor Red
		return $false
	}

	try
	{
		$startupLines = Get-Content -Path $StartupIniPath -ErrorAction Stop

		# STORE= (3 or 4 digits)
		$startupLines = $startupLines -replace '^[ \t]*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreNumber)

		# REDIRMAIL/REDIRMSG=<store>901  -> keep the 901 suffix
		$startupLines = $startupLines -replace '^[ \t]*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$',
			('`$1=' + $newStoreNumber + '`$2')

		# IMPORTANT: Set-Content default encoding differs by PS version; keep UTF8 if that’s what you want.
		Set-Content -Path $StartupIniPath -Value $startupLines -Encoding UTF8 -Force -ErrorAction Stop

		Write-Host "Updated startup.ini" -ForegroundColor Green
	}
	catch
	{
		$success = $false
		Write-Host "Failed updating startup.ini: $($_.Exception.Message)" -ForegroundColor Red
	}

	# ===============================================================================================
	# 2) Update Global SMSStart.ini → Only STORE= inside [SMSSTART]
	# ===============================================================================================
	if ($GlobalSmsStartIniPath -and (Test-Path $GlobalSmsStartIniPath))
	{
		try
		{
			$globalLines = Get-Content -Path $GlobalSmsStartIniPath -ErrorAction Stop
			$inSmsStartSection = $false

			for ($i = 0; $i -lt $globalLines.Count; $i++)
			{
				$line = $globalLines[$i]

				if ($line -match '^\s*\[(.+?)\]\s*$')
				{
					$sectionName = $matches[1].Trim()
					$inSmsStartSection = ($sectionName -ieq 'SMSSTART')
					continue
				}

				if ($inSmsStartSection -and $line -match '^\s*STORE\s*=\s*\d{1,4}\s*$')
				{
					$globalLines[$i] = "STORE=$newStoreNumber"
				}
			}

			Set-Content -Path $GlobalSmsStartIniPath -Value $globalLines -Encoding UTF8 -Force -ErrorAction Stop
			Write-Host "Updated SMSStart.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SMSStart.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}

	# ===============================================================================================
	# 3) Update INFO_*901_WIN.ini → STORE anywhere
	# ===============================================================================================
	if ($WinIniPath -and (Test-Path $WinIniPath))
	{
		try
		{
			$winLines = Get-Content -Path $WinIniPath -ErrorAction Stop

			for ($j = 0; $j -lt $winLines.Count; $j++)
			{
				if ($winLines[$j] -match '^\s*STORE\s*=\s*\d{1,4}\s*$')
				{
					$winLines[$j] = "STORE=$newStoreNumber"
				}
			}

			Set-Content -Path $WinIniPath -Value $winLines -Encoding UTF8 -Force -ErrorAction Stop
			Write-Host "Updated INFO_*901_WIN.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating INFO_*901_WIN.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}

	return $success
}

# ===================================================================================================
#                             FUNCTION: Update_SQL_Tables_For_Store_Number_Change
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates STD_TAB in the SQL database after store number change.
# ===================================================================================================

function Update_SQL_Tables_For_Store_Number_Change
{
	param (
		[string]$storeNumber # New store number to apply
	)
	
	# =============================
	# TABLE NAMES
	# =============================
	$stdTableName = "STD_TAB"
	$terTableName = "TER_TAB"
	
	# =============================
	# SQL FOR STD_TAB
	# =============================
	
	# Create view (required by Storeman load process)
	$createViewCommandStd = @"
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
"@
	
	# Update STD_TAB.F1056 to the new store number
	$updateStdTabCommand = @"
UPDATE $stdTableName
SET F1056 = '$storeNumber';
"@
	
	# Drop the view
	$dropViewCommandStd = "DROP VIEW Std_Load;"
	
	
	# =============================
	# SQL FOR TER_TAB
	# =============================
	# Only update rows where F1057 = '901'
	# This ensures ONLY the back-office record gets changed
	# and does NOT break lane entries.
	# =============================
	$updateTerTabCommand = @"
UPDATE $terTableName
SET F1056 = '$storeNumber'
WHERE F1057 = '901';
"@
	
	
	# =============================
	# EXECUTION PIPELINE
	# =============================
	$sqlCommands = @(
		$createViewCommandStd,
		$updateStdTabCommand,
		$dropViewCommandStd,
		$updateTerTabCommand # <-- NEW COMMAND
	)
	
	$allSqlSuccessful = $true
	$failedSqlCommands = @()
	
	foreach ($command in $sqlCommands)
	{
		# Execute_SQL_Commands is part of your existing framework
		# so no changes here.
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$allSqlSuccessful = $false
			$failedSqlCommands += $command
		}
	}
	
	# =============================
	# RETURN STRUCTURED RESULT
	# =============================
	return @{
		Success	       = $allSqlSuccessful
		FailedCommands = $failedSqlCommands
	}
}

# ===================================================================================================
# 									FUNCTION: Get_NEW_Machine_Name
# ---------------------------------------------------------------------------------------------------
# Description:
# Prompts user for a new machine name via GUI.
# Supports: 
#   - Classic: PREFIX + 3 digits (e.g. LANE003, SCO012)
#   - Prefixed: optional 1-4 digit store + PREFIX + 3 digits (e.g. 0231LANE006)
# If prefixed format used → offers to sync store number
# Returns uppercase validated name or $null if cancelled
# ===================================================================================================

function Get_NEW_Machine_Name
{
	while ($true)
	{
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Enter New Machine Name"
		$form.Size = New-Object System.Drawing.Size(420, 190)
		$form.StartPosition = "CenterParent"
		$form.FormBorderStyle = 'FixedDialog'
		$form.MaximizeBox = $false
		$form.MinimizeBox = $false
		$form.ShowInTaskbar = $false
		
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "Machine Name examples:`nPOS003  -  SCO012  -  LANE005 - 0231LANE006  -  1234POS999"
		$label.Location = New-Object System.Drawing.Point(10, 15)
		$label.Size = New-Object System.Drawing.Size(395, 45)
		$label.AutoSize = $false
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(10, 70)
		$textBox.Size = New-Object System.Drawing.Size(395, 22)
		
		$okBtn = New-Object System.Windows.Forms.Button
		$okBtn.Text = "OK"
		$okBtn.Location = New-Object System.Drawing.Point(130, 110)
		$okBtn.Size = New-Object System.Drawing.Size(70, 28)
		$okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
		
		$cancelBtn = New-Object System.Windows.Forms.Button
		$cancelBtn.Text = "Cancel"
		$cancelBtn.Location = New-Object System.Drawing.Point(215, 110)
		$cancelBtn.Size = New-Object System.Drawing.Size(70, 28)
		$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		
		$form.Controls.AddRange(@($label, $textBox, $okBtn, $cancelBtn))
		$form.AcceptButton = $okBtn
		$form.CancelButton = $cancelBtn
		
		$dialogResult = $form.ShowDialog()
		
		# PS 5.1-safe: read textbox BEFORE disposing
		$userInputRaw = $textBox.Text
		$form.Dispose()
		
		if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
		{
			return $null
		}
		
		# PS 5.1-safe null handling (no ??)
		$userInput = $userInputRaw
		if ($null -eq $userInput) { $userInput = "" }
		$userInput = $userInput.Trim().ToUpper()
		
		if ([string]::IsNullOrEmpty($userInput))
		{
			[System.Windows.Forms.MessageBox]::Show("Machine name cannot be empty.", "Error")
			continue
		}
		
		# Regex: optional 1-4 digits + 2-8 letters + exactly 3 digits
		if ($userInput -match '^(\d{1,4})?([A-Z]{2,8})(\d{3})$')
		{
			$storePrefixRaw = $matches[1] # may be empty
			$namePrefix = $matches[2]
			$terminal = $matches[3]
			
			# Ensure FunctionResults exists
			if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
			
			# Pull current store (as-is) from cache or INI
			$currentStore = $null
			if ($script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
			{
				$currentStore = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
			}
			else
			{
				$currentStore = Get_Store_Number_From_INI
				if ($currentStore) { $currentStore = ($currentStore.ToString()).Trim() }
			}
			
			# Decide store width:
			# - If current store is 3 or 4 digits -> use THAT length
			# - Else if user typed store prefix (3 or 4 digits) -> use THAT length
			# - Else default 4
			$storeWidth = 4
			if ($currentStore -match '^\d{3,4}$')
			{
				$storeWidth = $currentStore.Length
			}
			elseif ($storePrefixRaw -match '^\d{3,4}$')
			{
				$storeWidth = $storePrefixRaw.Length
			}
			
			# Normalize store prefix to chosen width (INLINE)
			$storePrefixNorm = $null
			if ($storePrefixRaw)
			{
				$val = [int]$storePrefixRaw
				$storePrefixNorm = $val.ToString(("D{0}" -f $storeWidth))
			}
			
			# Normalize current store to same width (INLINE)
			$currentStoreNorm = $null
			if ($currentStore -match '^\d{3,4}$')
			{
				$val2 = [int]$currentStore
				$currentStoreNorm = $val2.ToString(("D{0}" -f $storeWidth))
			}
			
			# Build normalized final machine name
			$normalizedName = if ($storePrefixNorm) { "$storePrefixNorm$namePrefix$terminal" }
			else { "$namePrefix$terminal" }
			
			# If user included store prefix, optionally sync INI when mismatch/unknown
			if ($storePrefixNorm)
			{
				if (-not $currentStoreNorm -or $storePrefixNorm -ne $currentStoreNorm)
				{
					$displayStore = $currentStoreNorm
					if ([string]::IsNullOrEmpty($displayStore)) { $displayStore = "UNKNOWN" }
					
					$sync = [System.Windows.Forms.MessageBox]::Show(
						"Machine name has store prefix '$storePrefixNorm',`nbut current store is '$displayStore'.`n`nUpdate store number to '$storePrefixNorm'?",
						"Store Number Mismatch",
						[System.Windows.Forms.MessageBoxButtons]::YesNo,
						[System.Windows.Forms.MessageBoxIcon]::Question
					)
					
					if ($sync -eq [System.Windows.Forms.DialogResult]::Yes)
					{
						$success = Update_Store_Number_In_INI -newStoreNumber $storePrefixNorm
						if ($success)
						{
							$script:newStoreNumber = $storePrefixNorm
							$script:FunctionResults['StoreNumber'] = $storePrefixNorm
							
							if ($storeNumberLabel)
							{
								$storeNumberLabel.Text = "Store Number: $storePrefixNorm (updated)"
								$storeNumberLabel.Refresh()
							}
						}
						else
						{
							[System.Windows.Forms.MessageBox]::Show("Failed to update store number in INI.", "Error")
							continue
						}
					}
				}
			}
			
			# Prevent duplicate name (compare against current computername)
			if ($normalizedName -eq $env:COMPUTERNAME.ToUpper())
			{
				[System.Windows.Forms.MessageBox]::Show(
					"New name is the same as current computer name ($env:COMPUTERNAME).`nChoose a different name.",
					"Invalid Name"
				)
				continue
			}
			
			return $normalizedName
		}
		
		[System.Windows.Forms.MessageBox]::Show(
			"Invalid format.`n`nValid examples:`n  LANE003`n  SCO012`n  0231LANE006`n  1234POS999`n`nRule: [optional 1-4 digits] + [2-8 letters] + [3 digits]",
			"Invalid Machine Name",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Error
		)
	}
}

# ===================================================================================================
#                                       FUNCTION: Update_SQL_Tables_For_Machine_Name_Change
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates STO_TAB, TER_TAB, LNK_TAB, and RUN_TAB in the SQL database after machine name change.
# ===================================================================================================

function Update_SQL_Tables_For_Machine_Name_Change
{
	param (
		[string]$storeNumber,
		[string]$machineName,
		[string]$machineNumber
	)
	
	# Variables
	$terTableName = "TER_TAB"
	$runTableName = "RUN_TAB"
	$stoTableName = "STO_TAB"
	$lnkTableName = "LNK_TAB"
	
	# Prepare SQL commands
	
	# TER_TAB commands
	$createViewCommandTer = @"
CREATE VIEW Ter_Load AS
SELECT F1056, F1057, F1058, F1125, F1169
FROM $terTableName;
"@
	
	$deleteOldRecordCommand = @"
DELETE FROM $terTableName 
WHERE F1057 NOT IN ('$machineNumber', '901');
"@
	
	$insertOrUpdateCommand = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$machineNumber')
BEGIN
    UPDATE $terTableName
    SET F1058='Terminal $machineNumber', 
        F1125='\\$machineName\storeman\office\XF$storeNumber$machineNumber\', 
        F1169='\\$machineName\storeman\office\XF${storeNumber}901\' 
    WHERE F1056='$storeNumber' AND F1057='$machineNumber';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$machineNumber', 
     'Terminal $machineNumber', 
     '\\$machineName\storeman\office\XF$storeNumber$machineNumber\', 
     '\\$machineName\storeman\office\XF${storeNumber}901\');
END
"@
	
	$dropViewCommandTer = "DROP VIEW Ter_Load;"
	
	# RUN_TAB commands
	$createViewCommandRun = @"
CREATE VIEW Run_Load AS
SELECT F1000, F1104
FROM $runTableName;
"@
	
	$updateRunTabCommand = @"
UPDATE $runTableName 
SET F1000 = '$machineNumber'
WHERE F1000 <> 'SMS';

UPDATE $runTableName 
SET F1104 = '$machineNumber'
WHERE F1104 <> '901';
"@
	
	$dropViewCommandRun = "DROP VIEW Run_Load;"
	
	# STO_TAB commands
	$createViewCommandSto = @"
CREATE VIEW Sto_Load AS
SELECT F1000, F1018, F1180, F1181, F1182
FROM $stoTableName;
"@
	
	$insertOrUpdateStoCommand = @"
MERGE INTO $stoTableName AS target
USING (VALUES 
    ('$machineNumber', 'Terminal $machineNumber', 1, 1, 1)
) AS source (F1000, F1018, F1180, F1181, F1182)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET 
        F1018 = source.F1018,
        F1180 = source.F1180,
        F1181 = source.F1181,
        F1182 = source.F1182
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182);
"@
	
	$deleteOldStoTabEntries = @"
DELETE FROM $stoTableName 
WHERE F1000 <> '$machineNumber'
AND F1000 NOT LIKE 'DSM%' 
AND F1000 NOT LIKE 'PAL%' 
AND F1000 NOT LIKE 'RAL%' 
AND F1000 NOT LIKE 'XAL%';
"@
	
	$dropViewCommandSto = "DROP VIEW Sto_Load;"
	
	# LNK_TAB commands
	$createViewCommandLnk = @"
CREATE VIEW Lnk_Load AS
SELECT F1000, F1056, F1057
FROM $lnkTableName;
"@
	
	$insertOrUpdateLnkCommand = @"
MERGE INTO $lnkTableName AS target
USING (VALUES 
    ('$machineNumber', '$storeNumber', '$machineNumber'),
    ('DSM', '$storeNumber', '$machineNumber'),
    ('PAL', '$storeNumber', '$machineNumber'),
    ('RAL', '$storeNumber', '$machineNumber'),
    ('XAL', '$storeNumber', '$machineNumber')
) AS source (F1000, F1056, F1057)
ON target.F1000 = source.F1000 AND target.F1056 = source.F1056 AND target.F1057 = source.F1057
WHEN NOT MATCHED THEN
    INSERT (F1000, F1056, F1057) VALUES (source.F1000, source.F1056, source.F1057);
"@
	
	$deleteOldLnkTabEntries = @"
DELETE FROM $lnkTableName 
WHERE F1057 <> '$machineNumber';
"@
	
	$dropViewCommandLnk = "DROP VIEW Lnk_Load;"
	
	# Execute the SQL commands
	$sqlCommands = @(
		# TER_TAB commands
		$createViewCommandTer,
		$deleteOldRecordCommand,
		$insertOrUpdateCommand,
		$dropViewCommandTer,
		
		# RUN_TAB commands
		$createViewCommandRun,
		$updateRunTabCommand,
		$dropViewCommandRun,
		
		# STO_TAB commands
		$createViewCommandSto,
		$insertOrUpdateStoCommand,
		$deleteOldStoTabEntries,
		$dropViewCommandSto,
		
		# LNK_TAB commands
		$createViewCommandLnk,
		$insertOrUpdateLnkCommand,
		$deleteOldLnkTabEntries,
		$dropViewCommandLnk
	)
	
	$allSqlSuccessful = $true
	$failedSqlCommands = @()
	
	foreach ($command in $sqlCommands)
	{
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$allSqlSuccessful = $false
			$failedSqlCommands += $command
		}
	}
	
	# Return the result
	return @{
		Success	       = $allSqlSuccessful
		FailedCommands = $failedSqlCommands
	}
}

# ===================================================================================================
#                                       FUNCTION: RPT_Blank
# ---------------------------------------------------------------------------------------------------
# Description:
#   Truncates the specified list of tables in the SQL database.
# ===================================================================================================

function RPT_Blank
{
	param (
		[string[]]$tables
	)
	
	# Initialize an array to store failed truncate commands
	$failedTruncateTables = @()
	
	# Prepare and execute Truncate Commands
	foreach ($table in $tables)
	{
		$command = "TRUNCATE TABLE $table;"
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$failedTruncateTables += $table # Add failed table to the array
		}
	}
	
	# Return the list of failed tables
	return $failedTruncateTables
}

# ===================================================================================================
#                               FUNCTION: Update_SQL_Database
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates the local SQL database configuration for the current store/terminal.
#
#   - Reads Store Number from: startup.ini (via Get_Store_Number_From_INI)
#   - Derives Terminal Number from the Machine Name (last 3 digits)
#   - Updates/cleans key tables to match the current terminal:
#       * TER_TAB  (Terminal records + XF paths)
#       * RUN_TAB  (Runtime terminal references)
#       * STO_TAB  (Terminal definition entry)
#       * LNK_TAB  (Terminal links: DSM/PAL/RAL/XAL + terminal)
#       * STD_TAB  (Store number)
#
#   Executes each SQL block via Execute_SQL_Commands, tracks results in $OperationStatus, and
#   displays success/failure dialogs (with optional failed-command viewer).
#
# Requirements:
#   - Get_Store_Number_From_INI
#   - Execute_SQL_Commands
#
# Parameters:
#   - CurrentMachineName : Current hostname to use if NewMachineName is empty
#   - NewMachineName     : Optional override hostname (ex: after rename)
#   - OperationStatus    : Shared status hashtable (ex: $OperationStatus["SQLDatabaseUpdate"])
# ===================================================================================================

function Update_SQL_Database
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		[string]$CurrentMachineName,
		[Parameter()]
		[string]$NewMachineName,
		[Parameter(Mandatory)]
		[hashtable]$OperationStatus
	)
	
	try
	{
		# Read the store number directly from startup.ini
		$storeNumberFromINI = Get_Store_Number_From_INI
		if ($null -eq $storeNumberFromINI)
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Store number not found in startup.ini.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Store number not found in startup.ini."
			$OperationStatus["SQLDatabaseUpdate"].Details = ""
			return
		}
		
		$storeNumber = $storeNumberFromINI
		
		# Determine the machine name to use
		$machineName = if (-not [string]::IsNullOrWhiteSpace($NewMachineName)) { $NewMachineName }
		else { $CurrentMachineName }
		
		# -------------------------------
		# Extract machine number (robust)
		# Supports: POS006, POS6, 0231LANE006, IFB001901, \\POS006\share, POS006.domain.local
		# Always normalizes to 3 digits (e.g. POS6 -> 006)
		# -------------------------------
		
		# Normalize to a clean host name:
		# - trim
		# - strip leading slashes
		# - if UNC/path is pasted, keep only the host part
		# - strip domain suffix if present
		$machineName = ($machineName -as [string])
		if ($null -eq $machineName) { $machineName = "" }
		$machineName = $machineName.Trim()
		$machineName = $machineName -replace '^[\\\/]+', '' # remove leading \\ or /
		if ($machineName -match '[\\\/]') { $machineName = ($machineName -split '[\\\/]')[0] } # keep host only
		if ($machineName -match '\.') { $machineName = ($machineName -split '\.')[0] } # remove domain
		$machineName = $machineName.Trim().ToUpper()
		
		# Pull trailing digits (allow 1-3) and normalize to 3 digits
		if ($machineName -notmatch '(\d{1,3})$')
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Invalid terminal number in machine name '$machineName'. Must end with 1-3 digits (ex: POS6, POS006, ...901).",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Invalid machine name."
			$OperationStatus["SQLDatabaseUpdate"].Details = "Cannot extract machine number."
			return
		}
		
		# Normalize to exactly 3 digits (POS6 -> 006, POS06 -> 006)
		$machineNumber = ([int]$Matches[1]).ToString("D3")
		
		# Reject 000 (almost always invalid)
		if ($machineNumber -eq '000')
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Invalid terminal number extracted from '$machineName' (000 is not allowed).",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Invalid machine name."
			$OperationStatus["SQLDatabaseUpdate"].Details = "Extracted terminal number 000."
			return
		}
		
		# Variables
		$terTableName = "TER_TAB"
		$runTableName = "RUN_TAB"
		$stoTableName = "STO_TAB"
		$lnkTableName = "LNK_TAB"
		$stdTableName = "STD_TAB"
		
		# ---------------------------
		# TER_TAB commands
		# ---------------------------
		$createViewCommandTer = @"
IF OBJECT_ID('Ter_Load','V') IS NOT NULL DROP VIEW Ter_Load;
CREATE VIEW Ter_Load AS
SELECT F1056, F1057, F1058, F1125, F1169
FROM $terTableName;
"@
		
		$deleteOldRecordCommand = @"
DELETE FROM $terTableName
WHERE F1057 NOT IN ('$machineNumber', '901');
"@
		
		$insertOrUpdateCommand = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$machineNumber')
BEGIN
    UPDATE $terTableName
    SET F1058='Terminal $machineNumber',
        F1125='\\$machineName\storeman\office\XF$storeNumber$machineNumber\',
        F1169='\\$machineName\storeman\office\XF${storeNumber}901\'
    WHERE F1056='$storeNumber' AND F1057='$machineNumber';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$machineNumber',
     'Terminal $machineNumber',
     '\\$machineName\storeman\office\XF$storeNumber$machineNumber\',
     '\\$machineName\storeman\office\XF${storeNumber}901\');
END
"@
		
		$dropViewCommandTer = "IF OBJECT_ID('Ter_Load','V') IS NOT NULL DROP VIEW Ter_Load;"
		
		# ---------------------------
		# RUN_TAB commands
		# ---------------------------
		$createViewCommandRun = @"
IF OBJECT_ID('Run_Load','V') IS NOT NULL DROP VIEW Run_Load;
CREATE VIEW Run_Load AS
SELECT F1000, F1104
FROM $runTableName;
"@
		
		$updateRunTabCommand = @"
UPDATE $runTableName
SET F1000 = '$machineNumber'
WHERE F1000 <> 'SMS';

UPDATE $runTableName
SET F1104 = '$machineNumber'
WHERE F1104 <> '901';
"@
		
		$dropViewCommandRun = "IF OBJECT_ID('Run_Load','V') IS NOT NULL DROP VIEW Run_Load;"
		
		# ---------------------------
		# STO_TAB commands
		# ---------------------------
		$createViewCommandSto = @"
IF OBJECT_ID('Sto_Load','V') IS NOT NULL DROP VIEW Sto_Load;
CREATE VIEW Sto_Load AS
SELECT F1000, F1018, F1180, F1181, F1182
FROM $stoTableName;
"@
		
		$insertOrUpdateStoCommand = @"
MERGE INTO $stoTableName AS target
USING (VALUES
    ('$machineNumber', 'Terminal $machineNumber', 1, 1, 1)
) AS source (F1000, F1018, F1180, F1181, F1182)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET
        F1018 = source.F1018,
        F1180 = source.F1180,
        F1181 = source.F1181,
        F1182 = source.F1182
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182);
"@
		
		$deleteOldStoTabEntries = @"
DELETE FROM $stoTableName
WHERE F1000 <> '$machineNumber'
AND F1000 NOT LIKE 'DSM%'
AND F1000 NOT LIKE 'PAL%'
AND F1000 NOT LIKE 'RAL%'
AND F1000 NOT LIKE 'XAL%';
"@
		
		$dropViewCommandSto = "IF OBJECT_ID('Sto_Load','V') IS NOT NULL DROP VIEW Sto_Load;"
		
		# ---------------------------
		# LNK_TAB commands
		# ---------------------------
		$createViewCommandLnk = @"
IF OBJECT_ID('Lnk_Load','V') IS NOT NULL DROP VIEW Lnk_Load;
CREATE VIEW Lnk_Load AS
SELECT F1000, F1056, F1057
FROM $lnkTableName;
"@
		
		$insertOrUpdateLnkCommand = @"
MERGE INTO $lnkTableName AS target
USING (VALUES
    ('$machineNumber', '$storeNumber', '$machineNumber'),
    ('DSM', '$storeNumber', '$machineNumber'),
    ('PAL', '$storeNumber', '$machineNumber'),
    ('RAL', '$storeNumber', '$machineNumber'),
    ('XAL', '$storeNumber', '$machineNumber')
) AS source (F1000, F1056, F1057)
ON target.F1000 = source.F1000 AND target.F1056 = source.F1056 AND target.F1057 = source.F1057
WHEN NOT MATCHED THEN
    INSERT (F1000, F1056, F1057) VALUES (source.F1000, source.F1056, source.F1057);
"@
		
		$deleteOldLnkTabEntries = @"
DELETE FROM $lnkTableName
WHERE F1057 <> '$machineNumber';
"@
		
		$dropViewCommandLnk = "IF OBJECT_ID('Lnk_Load','V') IS NOT NULL DROP VIEW Lnk_Load;"
		
		# ---------------------------
		# STD_TAB commands
		# ---------------------------
		$createViewCommandStd = @"
IF OBJECT_ID('Std_Load','V') IS NOT NULL DROP VIEW Std_Load;
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
"@
		
		$updateStdTabCommand = @"
UPDATE $stdTableName
SET F1056 = '$storeNumber';
"@
		
		$dropViewCommandStd = "IF OBJECT_ID('Std_Load','V') IS NOT NULL DROP VIEW Std_Load;"
		
		# Execute SQL commands
		$allSqlSuccessful = $true
		$failedSqlCommands = @()
		
		$sqlCommands = @(
			# TER_TAB
			$createViewCommandTer,
			$deleteOldRecordCommand,
			$insertOrUpdateCommand,
			$dropViewCommandTer,
			
			# RUN_TAB
			$createViewCommandRun,
			$updateRunTabCommand,
			$dropViewCommandRun,
			
			# STO_TAB
			$createViewCommandSto,
			$insertOrUpdateStoCommand,
			$deleteOldStoTabEntries,
			$dropViewCommandSto,
			
			# LNK_TAB
			$createViewCommandLnk,
			$insertOrUpdateLnkCommand,
			$deleteOldLnkTabEntries,
			$dropViewCommandLnk,
			
			# STD_TAB
			$createViewCommandStd,
			$updateStdTabCommand,
			$dropViewCommandStd
		)
		
		foreach ($command in $sqlCommands)
		{
			if (-not (Execute_SQL_Commands -commandText $command))
			{
				$allSqlSuccessful = $false
				$failedSqlCommands += $command
			}
		}
		
		if ($allSqlSuccessful)
		{
			$OperationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL database updated successfully."
			$OperationStatus["SQLDatabaseUpdate"].Details = "All SQL commands executed successfully."
			[System.Windows.Forms.MessageBox]::Show(
				"SQL database updated successfully.",
				"Success",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			)
		}
		else
		{
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Failed to execute some SQL commands."
			$OperationStatus["SQLDatabaseUpdate"].Details = "Failed SQL commands are listed below."
			[System.Windows.Forms.MessageBox]::Show(
				"Failed to execute some SQL commands.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			
			$failedCommandsForm = New-Object System.Windows.Forms.Form
			$failedCommandsForm.Text = "Failed SQL Commands"
			$failedCommandsForm.Size = New-Object System.Drawing.Size(600, 400)
			$failedCommandsForm.StartPosition = "CenterParent"
			
			$textBox = New-Object System.Windows.Forms.TextBox
			$textBox.Multiline = $true
			$textBox.ReadOnly = $true
			$textBox.ScrollBars = "Vertical"
			$textBox.Dock = "Fill"
			$textBox.Text = $failedSqlCommands -join "`r`n`r`n"
			
			$failedCommandsForm.Controls.Add($textBox)
			[void]$failedCommandsForm.ShowDialog()
		}
	}
	catch
	{
		$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
		$OperationStatus["SQLDatabaseUpdate"].Message = "Unhandled error during SQL update."
		$OperationStatus["SQLDatabaseUpdate"].Details = $_.Exception.Message
		
		[System.Windows.Forms.MessageBox]::Show(
			"Unhandled error:`r`n$($_.Exception.Message)",
			"Error",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Error
		)
	}
}

# ===================================================================================================
#                                       FUNCTION: Remove_GT_Registry_Values
# ---------------------------------------------------------------------------------------------------
# Description:
#   Removes all registry values starting with 'GT' from specified registry paths.
# ===================================================================================================

function Remove_GT_Registry_Values
{
	# Define registry paths for 32-bit and 64-bit
	$regPath32 = "HKLM:\SOFTWARE\Store Management\Counters"
	$regPath64 = "HKLM:\SOFTWARE\Wow6432Node\Store Management\Counters"
	
	$is64bit = [System.Environment]::Is64BitOperatingSystem
	
	$totalDeletedCount = 0
	$success = $true
	$status = "Successful"
	$message = ""
	
	# Build list of candidate paths (prefer Wow6432Node on 64-bit, but include both if present)
	$pathsToCheck = @()
	
	if ($is64bit)
	{
		if (Test-Path $regPath64) { $pathsToCheck += $regPath64 }
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	else
	{
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	
	if (-not $pathsToCheck -or $pathsToCheck.Count -eq 0)
	{
		return @{
			Success	     = $false
			Status	     = "Failed"
			DeletedCount = 0
			Message	     = "No valid registry paths found for the current environment."
		}
	}
	
	foreach ($path in $pathsToCheck)
	{
		try
		{
			# Get value names from the key by inspecting item properties
			# (Get-ItemProperty returns a PSCustomObject with properties = value names + PS metadata)
			$itemProps = Get-ItemProperty -Path $path -ErrorAction Stop
			
			# Collect only real registry value names starting with GT (exclude PS metadata props)
			$gtNames = @()
			foreach ($p in $itemProps.PSObject.Properties)
			{
				if ($p.Name -eq 'PSPath') { continue }
				if ($p.Name -eq 'PSParentPath') { continue }
				if ($p.Name -eq 'PSChildName') { continue }
				if ($p.Name -eq 'PSDrive') { continue }
				if ($p.Name -eq 'PSProvider') { continue }
				
				if ($p.Name -like 'GT*')
				{
					$gtNames += $p.Name
				}
			}
			
			# Delete each GT* value
			foreach ($name in $gtNames)
			{
				try
				{
					Remove-ItemProperty -Path $path -Name $name -ErrorAction Stop
					$totalDeletedCount++
				}
				catch
				{
					# Keep going but mark overall as partial failure
					$success = $false
					$status = "Failed"
					if ($message) { $message += " | " }
					$message += "Failed to delete '$name' in '$path': $_"
				}
			}
		}
		catch
		{
			$success = $false
			$status = "Failed"
			if ($message) { $message += " | " }
			$message += "Error accessing registry path '$path': $_"
		}
	}
	
	return @{
		Success	     = $success
		Status	     = $status
		DeletedCount = $totalDeletedCount
		Message	     = $message
	}
}

# ===================================================================================================
#                              FUNCTION: Update_SmsHttpsINI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates the [PROCESSORS] section inside SmsHttps.INI to reflect the current machine/store values,
#   and clears the LicenseGUID value in [GENERAL] so a new GUID can be generated after reboot.
#
#   Updates:
#     - Processor key: <STORE><TERMINAL> (3+3 or 4+3 based on StoreNumber length)
#     - REDIRMAIL / REDIRMSG: <STORE>901
#     - [GENERAL] LicenseGUID: clears value -> "LicenseGUID="
#
#   Preserves:
#     - Other parameters on the processor line (TARGETSEND/TARGETRECV/DEADLOCKPRIORITY/etc.)
#     - File encoding + newline style
#
#   Matching:
#     - Finds [PROCESSORS] section
#     - Prefers exact old key if OldStoreNumber + OldMachineName provided
#     - Otherwise matches by terminal suffix
#     - Can create entry if missing (-CreateIfMissing)
#
#   Returns:
#     Hashtable: Success/Message/Path/OldKey/NewKey/Changed/Redir901/LicenseGuidCleared
# ===================================================================================================

function Update_SmsHttpsINI
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$IniPath,
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[switch]$CreateIfMissing
	)
	
	# -----------------------------
	# Resolve INI path
	# -----------------------------
	if ([string]::IsNullOrWhiteSpace($IniPath))
	{
		$candidates = @(
			"\\localhost\storeman\SmsHttps64\SmsHttps.INI",
			"C:\storeman\SmsHttps64\SmsHttps.INI",
			"D:\storeman\SmsHttps64\SmsHttps.INI"
		)
		
		foreach ($c in $candidates)
		{
			if (Test-Path $c)
			{
				$IniPath = $c
				break
			}
		}
	}
	
	if ([string]::IsNullOrWhiteSpace($IniPath) -or -not (Test-Path $IniPath))
	{
		return @{
			Success		       = $false
			Message		       = "SmsHttps.INI not found. Provide -IniPath or ensure it exists in the default locations."
			Path			   = $IniPath
			OldKey			   = $null
			NewKey			   = $null
			Changed		       = $false
			Redir901		   = $null
			LicenseGuidCleared = $false
		}
	}
	
	# -----------------------------
	# Read file with encoding detect
	# -----------------------------
	$bytes = [System.IO.File]::ReadAllBytes($IniPath)
	
	$enc = $null
	if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
	{
		$enc = [System.Text.Encoding]::UTF8
	}
	elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
	{
		$enc = [System.Text.Encoding]::Unicode # UTF-16LE
	}
	elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
	{
		$enc = [System.Text.Encoding]::BigEndianUnicode # UTF-16BE
	}
	else
	{
		$enc = [System.Text.Encoding]::Default # common for INIs
	}
	
	$text = $enc.GetString($bytes)
	
	# Detect newline style
	$nl = "`r`n"
	if ($text -notmatch "`r`n" -and $text -match "`n") { $nl = "`n" }
	
	$lines = $text -split "\r?\n", -1
	
	# -----------------------------
	# Validate StoreNumber (3 or 4 digits, not all zeros)
	# -----------------------------
	$sn = $StoreNumber
	if ($null -eq $sn) { $sn = "" }
	$sn = $sn.Trim()
	
	if ($sn -notmatch '^(?!0{3,4})\d{3,4}$')
	{
		return @{
			Success = $false
			Message = "Invalid StoreNumber '$StoreNumber' (must be 3 or 4 digits, not all zeros)."
			Path    = $IniPath
			OldKey  = $null
			NewKey  = $null
			Changed = $false
			Redir901 = $null
			LicenseGuidCleared = $false
		}
	}
	
	# Store width = length of the CURRENT store number you pass (3 or 4)
	$storeWidth = $sn.Length
	$storeNorm = ([int]$sn).ToString(("D{0}" -f $storeWidth))
	
	# -----------------------------
	# Normalize MachineName -> host and extract terminal (1-3 trailing digits), pad to 3
	# -----------------------------
	$mn = $MachineName
	if ($null -eq $mn) { $mn = "" }
	$mn = $mn.Trim()
	$mn = $mn -replace '^[\\\/]+', ''
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] }
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] }
	$mn = $mn.Trim().ToUpper()
	
	if ($mn -notmatch '(\d{1,3})$')
	{
		return @{
			Success = $false
			Message = "MachineName '$MachineName' must end with 1-3 digits to extract terminal number."
			Path    = $IniPath
			OldKey  = $null
			NewKey  = $null
			Changed = $false
			Redir901 = $null
			LicenseGuidCleared = $false
		}
	}
	
	$newTerminal = ([int]$Matches[1]).ToString("D3")
	if ($newTerminal -eq "000")
	{
		return @{
			Success		       = $false
			Message		       = "Extracted terminal number is 000 (not allowed)."
			Path			   = $IniPath
			OldKey			   = $null
			NewKey			   = $null
			Changed		       = $false
			Redir901		   = $null
			LicenseGuidCleared = $false
		}
	}
	
	# Optional old terminal extraction
	$oldTerminal = $null
	if (-not [string]::IsNullOrWhiteSpace($OldMachineName))
	{
		$omn = $OldMachineName.Trim()
		$omn = $omn -replace '^[\\\/]+', ''
		if ($omn -match '[\\\/]') { $omn = ($omn -split '[\\\/]')[0] }
		if ($omn -match '\.') { $omn = ($omn -split '\.')[0] }
		$omn = $omn.Trim().ToUpper()
		
		if ($omn -match '(\d{1,3})$')
		{
			$oldTerminal = ([int]$Matches[1]).ToString("D3")
			if ($oldTerminal -eq "000") { $oldTerminal = $null }
		}
	}
	
	$newKey = $storeNorm + $newTerminal
	$redirKey901 = $storeNorm + "901"
	
	# -----------------------------
	# Locate [PROCESSORS] section
	# -----------------------------
	$secStart = -1
	for ($i = 0; $i -lt $lines.Length; $i++)
	{
		if ($lines[$i] -match '^\s*\[\s*PROCESSORS\s*\]\s*$')
		{
			$secStart = $i
			break
		}
	}
	
	if ($secStart -lt 0)
	{
		# Still attempt to clear LicenseGUID even if PROCESSORS missing
		# (but report processors failure)
		$processorsError = $true
	}
	else
	{
		$processorsError = $false
	}
	
	$secEnd = $lines.Length
	if ($secStart -ge 0)
	{
		for ($i = $secStart + 1; $i -lt $lines.Length; $i++)
		{
			if ($lines[$i] -match '^\s*\[.*\]\s*$')
			{
				$secEnd = $i
				break
			}
		}
	}
	
	# -----------------------------
	# Find target PROCESSORS line
	# -----------------------------
	$targetIndex = -1
	$targetWs = ""
	$targetRhs = ""
	$foundKey = $null
	
	$oldKeyExact = $null
	if (-not [string]::IsNullOrWhiteSpace($OldStoreNumber) -and $oldTerminal)
	{
		$osn = $OldStoreNumber.Trim()
		if ($osn -match '^(?!0{3,4})\d{3,4}$')
		{
			$osWidth = $osn.Length
			$oldKeyExact = ([int]$osn).ToString(("D{0}" -f $osWidth)) + $oldTerminal
		}
	}
	
	if ($secStart -ge 0)
	{
		# 1) Exact old key match
		if ($oldKeyExact)
		{
			for ($i = $secStart + 1; $i -lt $secEnd; $i++)
			{
				if ($lines[$i] -match '^(?<ws>\s*)(?<key>\d{6,7})\s*=\s*(?<rhs>.*)$')
				{
					if ($matches['key'] -eq $oldKeyExact)
					{
						$targetIndex = $i
						$targetWs = $matches['ws']
						$foundKey = $matches['key']
						$targetRhs = $matches['rhs']
						break
					}
				}
			}
		}
		
		# 2) Match by terminal suffix (old first then new)
		if ($targetIndex -lt 0)
		{
			$termToTry = @()
			if ($oldTerminal) { $termToTry += $oldTerminal }
			$termToTry += $newTerminal
			
			foreach ($t in $termToTry)
			{
				for ($i = $secStart + 1; $i -lt $secEnd; $i++)
				{
					if ($lines[$i] -match '^(?<ws>\s*)(?<key>\d{6,7})\s*=\s*(?<rhs>.*)$')
					{
						$k = $matches['key']
						if ($k.Length -ge 6 -and $k.Substring($k.Length - 3) -eq $t)
						{
							$targetIndex = $i
							$targetWs = $matches['ws']
							$foundKey = $k
							$targetRhs = $matches['rhs']
							break
						}
					}
				}
				if ($targetIndex -ge 0) { break }
			}
		}
	}
	
	# -----------------------------
	# Update PROCESSORS entry
	# -----------------------------
	$processorsChanged = $false
	
	if ($secStart -ge 0)
	{
		if ($targetIndex -ge 0)
		{
			$rhsUpdated = $targetRhs
			
			# REDIRMAIL
			if ($rhsUpdated -match '(?i)\bREDIRMAIL\s*=')
			{
				$rhsUpdated = [System.Text.RegularExpressions.Regex]::Replace(
					$rhsUpdated,
					'(?i)\bREDIRMAIL\s*=\s*\d{6,7}',
					("REDIRMAIL=" + $redirKey901)
				)
			}
			else
			{
				if ([string]::IsNullOrWhiteSpace($rhsUpdated)) { $rhsUpdated = "" }
				if ($rhsUpdated.Length -gt 0 -and $rhsUpdated[0] -ne ',') { $rhsUpdated = "," + $rhsUpdated }
				$rhsUpdated = ("REDIRMAIL=" + $redirKey901) + $rhsUpdated
			}
			
			# REDIRMSG
			if ($rhsUpdated -match '(?i)\bREDIRMSG\s*=')
			{
				$rhsUpdated = [System.Text.RegularExpressions.Regex]::Replace(
					$rhsUpdated,
					'(?i)\bREDIRMSG\s*=\s*\d{6,7}',
					("REDIRMSG=" + $redirKey901)
				)
			}
			else
			{
				if ($rhsUpdated -match '(?i)\bREDIRMAIL\s*=\s*\d{6,7}\s*,?')
				{
					$rhsUpdated = [System.Text.RegularExpressions.Regex]::Replace(
						$rhsUpdated,
						'(?i)\bREDIRMAIL\s*=\s*\d{6,7}\s*,?',
						("$0" + "REDIRMSG=" + $redirKey901 + ",")
					)
					$rhsUpdated = $rhsUpdated -replace ',{2,}', ','
				}
				else
				{
					if ($rhsUpdated.Length -gt 0 -and $rhsUpdated[0] -ne ',') { $rhsUpdated = "," + $rhsUpdated }
					$rhsUpdated = ("REDIRMSG=" + $redirKey901) + $rhsUpdated
				}
			}
			
			$rhsUpdated = $rhsUpdated -replace ',{2,}', ','
			
			$oldLine = $lines[$targetIndex]
			$newLine = $targetWs + $newKey + "=" + $rhsUpdated
			
			if ($oldLine -ne $newLine)
			{
				$lines[$targetIndex] = $newLine
				$processorsChanged = $true
			}
		}
		else
		{
			if ($CreateIfMissing)
			{
				$newEntry = $newKey + "=REDIRMAIL=" + $redirKey901 + ",REDIRMSG=" + $redirKey901 + ",TARGETSEND=,TARGETRECV=,DEADLOCKPRIORITY="
				$insertAt = $secEnd
				
				if ($insertAt -eq $lines.Length -and $lines.Length -gt 0 -and [string]::IsNullOrEmpty($lines[$lines.Length - 1]))
				{
					$insertAt = $lines.Length - 1
				}
				
				$before = @()
				$after = @()
				
				for ($i = 0; $i -lt $insertAt; $i++) { $before += $lines[$i] }
				$before += $newEntry
				for ($i = $insertAt; $i -lt $lines.Length; $i++) { $after += $lines[$i] }
				
				$lines = @($before + $after)
				$processorsChanged = $true
			}
			else
			{
				$processorsError = $true
			}
		}
	}
	
	# -----------------------------
	# Clear LicenseGUID value in [GENERAL]
	# -----------------------------
	$licenseGuidCleared = $false
	
	$genStart = -1
	for ($i = 0; $i -lt $lines.Length; $i++)
	{
		if ($lines[$i] -match '^\s*\[\s*GENERAL\s*\]\s*$')
		{
			$genStart = $i
			break
		}
	}
	
	if ($genStart -ge 0)
	{
		$genEnd = $lines.Length
		for ($i = $genStart + 1; $i -lt $lines.Length; $i++)
		{
			if ($lines[$i] -match '^\s*\[.*\]\s*$')
			{
				$genEnd = $i
				break
			}
		}
		
		for ($i = $genStart + 1; $i -lt $genEnd; $i++)
		{
			# Only change within [GENERAL]
			if ($lines[$i] -match '^(?<ws>\s*)(?i:LicenseGUID)\s*=\s*(?<val>.*)\s*$')
			{
				$ws = $matches['ws']
				$val = $matches['val']
				if (-not [string]::IsNullOrEmpty($val))
				{
					$lines[$i] = $ws + "LicenseGUID="
					$licenseGuidCleared = $true
				}
				break
			}
		}
	}
	
	# -----------------------------
	# Write back only if changes occurred
	# -----------------------------
	$anyChanged = ($processorsChanged -or $licenseGuidCleared)
	
	if ($anyChanged)
	{
		$outText = ($lines -join $nl)
		[System.IO.File]::WriteAllText($IniPath, $outText, $enc)
	}
	
	# -----------------------------
	# Return status
	# -----------------------------
	if ($processorsError -and -not $anyChanged)
	{
		return @{
			Success		       = $false
			Message		       = "Failed to update [PROCESSORS] (not found or no matching entry) and no other changes applied."
			Path			   = $IniPath
			OldKey			   = $foundKey
			NewKey			   = $newKey
			Changed		       = $false
			Redir901		   = $redirKey901
			LicenseGuidCleared = $licenseGuidCleared
		}
	}
	
	return @{
		Success = $true
		Message = ($(if ($anyChanged) { "SmsHttps.INI updated successfully." }
				else { "No changes were necessary." }))
		Path    = $IniPath
		OldKey  = $foundKey
		NewKey  = $newKey
		Changed = $anyChanged
		Redir901 = $redirKey901
		LicenseGuidCleared = $licenseGuidCleared
	}
}

# ===================================================================================================
#                              FUNCTION: Delete_Files/Folders
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes specified files within a directory, supporting wildcards and exclusions.
#   Can be executed synchronously or as a background job to prevent interruption of the main script.
#   Parameters:
#     - Path: The directory path where files will be deleted.
#     - SpecifiedFiles: Specific file names or patterns to delete. Wildcards are supported.
#     - Exclusions: File names or patterns to exclude from deletion. Wildcards are supported.
#     - AsJob: (Optional) Runs the deletion process as a background job.
# ===================================================================================================

function Delete_Files/Folders
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, Position = 0, HelpMessage = "The directory path where files and folders will be deleted.")]
		[ValidateNotNullOrEmpty()]
		[string]$Path,
		[Parameter(Mandatory = $false, HelpMessage = "Specific file or folder patterns to delete within the specified directory. Wildcards supported.")]
		[string[]]$SpecifiedFiles,
		[Parameter(Mandatory = $false, HelpMessage = "File or folder patterns to exclude from deletion. Wildcards supported.")]
		[string[]]$Exclusions,
		[Parameter(Mandatory = $false, HelpMessage = "Run the deletion as a background job.")]
		[switch]$AsJob
	)
	
	if ($AsJob)
	{
		# Define the script block that performs the deletion
		$scriptBlock = {
			param ($Path,
				$SpecifiedFiles,
				$Exclusions)
			
			# Initialize counter for deleted items
			$deletedCount = 0
			
			# Resolve the full path
			$resolvedPath = Resolve-Path -Path $Path -ErrorAction SilentlyContinue
			if (-not $resolvedPath)
			{
				# Write-Log "The specified path '$Path' does not exist." "Red"
				return
			}
			$targetPath = $resolvedPath.ProviderPath
			
			try
			{
				if ($SpecifiedFiles)
				{
					# Delete only specified files and folders
					foreach ($filePattern in $SpecifiedFiles)
					{
						# Retrieve matching items using wildcards
						$matchedItems = Get-ChildItem -Path $targetPath -Filter $filePattern -Recurse -Force -ErrorAction SilentlyContinue
						
						if ($matchedItems)
						{
							foreach ($matchedItem in $matchedItems)
							{
								# Check against exclusions
								$exclude = $false
								if ($Exclusions)
								{
									foreach ($exclusionPattern in $Exclusions)
									{
										if ($matchedItem.Name -like $exclusionPattern)
										{
											$exclude = $true
											# Write-Log "Excluded: $($matchedItem.FullName)" "Yellow"
											break
										}
									}
								}
								
								if (-not $exclude)
								{
									try
									{
										if ($matchedItem.PSIsContainer)
										{
											Remove-Item -Path $matchedItem.FullName -Recurse -Force -ErrorAction Stop
										}
										else
										{
											Remove-Item -Path $matchedItem.FullName -Force -ErrorAction Stop
										}
										$deletedCount++
										# Write-Log "Deleted: $($matchedItem.FullName)" "Green"
									}
									catch
									{
										# Write-Log "Failed to delete $($matchedItem.FullName). Error: $_" "Red"
									}
								}
							}
						}
						else
						{
							# Write-Log "No items matched the pattern: '$filePattern' in '$targetPath'." "Yellow"
						}
					}
				}
				else
				{
					# Delete all files and folders in the path
					$allItems = Get-ChildItem -Path $targetPath -Recurse -Force -ErrorAction SilentlyContinue
					
					foreach ($item in $allItems)
					{
						# Check against exclusions
						$exclude = $false
						if ($Exclusions)
						{
							foreach ($exclusionPattern in $Exclusions)
							{
								if ($item.Name -like $exclusionPattern)
								{
									$exclude = $true
									# Write-Log "Excluded: $($item.FullName)" "Yellow"
									break
								}
							}
						}
						
						if (-not $exclude)
						{
							try
							{
								if ($item.PSIsContainer)
								{
									Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
								}
								else
								{
									Remove-Item -Path $item.FullName -Force -ErrorAction Stop
								}
								$deletedCount++
								# Write-Log "Deleted: $($item.FullName)" "Green"
							}
							catch
							{
								# Write-Log "Failed to delete $($item.FullName). Error: $_" "Red"
							}
						}
					}
				}
				
				# Write-Log "Total items deleted: $deletedCount" "Blue"
				return $deletedCount
			}
			catch
			{
				# Write-Log "An error occurred during the deletion process. Error: $_" "Red"
				return $deletedCount
			}
		}
		
		# Start the background job
		Start-Job -ScriptBlock $scriptBlock -ArgumentList $Path, $SpecifiedFiles, $Exclusions
	}
	else
	{
		# Synchronous execution
		# Initialize counter for deleted items
		$deletedCount = 0
		
		# Resolve the full path
		$resolvedPath = Resolve-Path -Path $Path -ErrorAction SilentlyContinue
		if (-not $resolvedPath)
		{
			#	Write-Log "The specified path '$Path' does not exist." "Red"
			return
		}
		$targetPath = $resolvedPath.ProviderPath
		
		try
		{
			if ($SpecifiedFiles)
			{
				# Delete only specified files and folders
				foreach ($filePattern in $SpecifiedFiles)
				{
					# Retrieve matching items using wildcards
					$matchedItems = Get-ChildItem -Path $targetPath -Filter $filePattern -Recurse -Force -ErrorAction SilentlyContinue
					
					if ($matchedItems)
					{
						foreach ($matchedItem in $matchedItems)
						{
							# Check against exclusions
							$exclude = $false
							if ($Exclusions)
							{
								foreach ($exclusionPattern in $Exclusions)
								{
									if ($matchedItem.Name -like $exclusionPattern)
									{
										$exclude = $true
										#	Write-Log "Excluded: $($matchedItem.FullName)" "Yellow"
										break
									}
								}
							}
							
							if (-not $exclude)
							{
								try
								{
									if ($matchedItem.PSIsContainer)
									{
										Remove-Item -Path $matchedItem.FullName -Recurse -Force -ErrorAction Stop
									}
									else
									{
										Remove-Item -Path $matchedItem.FullName -Force -ErrorAction Stop
									}
									$deletedCount++
									#	Write-Log "Deleted: $($matchedItem.FullName)" "Green"
								}
								catch
								{
									#	Write-Log "Failed to delete $($matchedItem.FullName). Error: $_" "Red"
								}
							}
						}
					}
					else
					{
						#	Write-Log "No items matched the pattern: '$filePattern' in '$targetPath'." "Yellow"
					}
				}
			}
			else
			{
				# Delete all files and folders in the path
				$allItems = Get-ChildItem -Path $targetPath -Recurse -Force -ErrorAction SilentlyContinue
				
				foreach ($item in $allItems)
				{
					# Check against exclusions
					$exclude = $false
					if ($Exclusions)
					{
						foreach ($exclusionPattern in $Exclusions)
						{
							if ($item.Name -like $exclusionPattern)
							{
								$exclude = $true
								#	Write-Log "Excluded: $($item.FullName)" "Yellow"
								break
							}
						}
					}
					
					if (-not $exclude)
					{
						try
						{
							if ($item.PSIsContainer)
							{
								Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
							}
							else
							{
								Remove-Item -Path $item.FullName -Force -ErrorAction Stop
							}
							$deletedCount++
							#	Write-Log "Deleted: $($item.FullName)" "Green"
						}
						catch
						{
							#	Write-Log "Failed to delete $($item.FullName). Error: $_" "Red"
						}
					}
				}
			}
			
			#	Write-Log "Total items deleted: $deletedCount" "Blue"
			return $deletedCount
		}
		catch
		{
			#	Write-Log "An error occurred during the deletion process. Error: $_" "Red"
			return $deletedCount
		}
	}
}

# ===================================================================================================
#                                       SECTION: Main Script Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Orchestrates the execution flow of the script, initializing variables, processing items, and handling user interactions.
# ===================================================================================================

# Get current IP configuration
$currentConfigs = Get_Active_IP_Config

# Get the store number
$Get_Store_Number_From_INI = Get_Store_Number_From_INI -UpdateLabel
$currentStoreNumber = $script:FunctionResults['StoreNumber']

# Get the store name
Get_Store_Name_From_INI
$storeName = $script:FunctionResults['StoreName']

# Get the database connection string
Get_Database_Connection_String
$connectionString = $script:FunctionResults['ConnectionString']

# Set the old machine name variable
$oldMachineName = $currentMachineName

# Clear %Temp% foder on start
# $FilesAndDirsDeleted = Delete_Files/Folders -Path "$TempDir" -Exclusions "MiniGhost.ps1" -AsJob

# Indicate the script has started
Write-Host "Script started" -ForegroundColor Green

# ===================================================================================================
#                                       SECTION: Initialize GUI Components
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and initializes the main graphical user interface (GUI) form and its components.
# ===================================================================================================

# Create the main form (resizable, DPI-aware)
$form = New-Object System.Windows.Forms.Form
$form.Text = "Created by Alex_C.T | Version: $VersionNumber | Revised: $VersionDate"
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.MinimumSize = New-Object System.Drawing.Size(760, 520)
$form.Size = New-Object System.Drawing.Size(880, 580)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.BackColor = [System.Drawing.Color]::White

# Global tooltip provider (per-button tooltips)
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 20000
$toolTip.InitialDelay = 350
$toolTip.ReshowDelay = 150
$toolTip.ShowAlways = $true

# ---------------------------------------------------------------------------------------------------
# Main layout (Banner + Info + Actions)
# ---------------------------------------------------------------------------------------------------
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 3
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(12)
$mainLayout.Margin  = New-Object System.Windows.Forms.Padding(0)
$mainLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize

[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$form.Controls.Add($mainLayout)

# Banner Label
$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Text = "PowerShell Script - Mini Ghost"
$bannerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$bannerLabel.AutoSize = $true
$bannerLabel.TextAlign = 'MiddleCenter'
$bannerLabel.Dock = 'Fill'
$bannerLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)

$mainLayout.Controls.Add($bannerLabel, 0, 0)

# Display current machine name and store number in labels
$script:machineNameLabel = New-Object System.Windows.Forms.Label
$script:machineNameLabel.Text = "Current Machine Name: $currentMachineName"
$script:machineNameLabel.AutoSize = $true
$script:machineNameLabel.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 2)

$script:storeNameLabel = New-Object System.Windows.Forms.Label
$script:storeNameLabel.Text = "Store Name: $storeName"
$script:storeNameLabel.AutoSize = $true
$script:storeNameLabel.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 2)

$script:storeNumberLabel = New-Object System.Windows.Forms.Label
$script:storeNumberLabel.Text = "Store Number: $currentStoreNumber"
$script:storeNumberLabel.AutoSize = $true
$script:storeNumberLabel.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 2)

# Display current IP address in a label
$currentIP = if ($currentConfigs -and $currentConfigs.Count -gt 0) { $currentConfigs[0].IPv4Address.IPAddress }
else { "IP Not Found" }

$script:ipAddressLabel = New-Object System.Windows.Forms.Label
$script:ipAddressLabel.Text = "Current IP Address: $currentIP"
$script:ipAddressLabel.AutoSize = $true
$script:ipAddressLabel.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 6)

# Info group
$infoGroup = New-Object System.Windows.Forms.GroupBox
$infoGroup.Text = "Current Info"
$infoGroup.Dock = 'Top'
$infoGroup.AutoSize = $true
$infoGroup.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$infoGroup.Padding = New-Object System.Windows.Forms.Padding(10)
$infoGroup.Margin  = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)

$infoLayout = New-Object System.Windows.Forms.TableLayoutPanel
$infoLayout.Dock = 'Fill'
$infoLayout.ColumnCount = 1
$infoLayout.RowCount = 4
$infoLayout.AutoSize = $true
$infoLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$infoLayout.Margin = New-Object System.Windows.Forms.Padding(0)

$infoLayout.Controls.Add($machineNameLabel, 0, 0)
$infoLayout.Controls.Add($storeNameLabel, 0, 1)
$infoLayout.Controls.Add($storeNumberLabel, 0, 2)
$infoLayout.Controls.Add($ipAddressLabel, 0, 3)

$infoGroup.Controls.Add($infoLayout)

$mainLayout.Controls.Add($infoGroup, 0, 1)

# ---------------------------------------------------------------------------------------------------
# Actions group (buttons live here; populated after button creation)
# ---------------------------------------------------------------------------------------------------
$actionsGroup = New-Object System.Windows.Forms.GroupBox
$actionsGroup.Text = "Actions"
$actionsGroup.Dock = 'Fill'
$actionsGroup.Padding = New-Object System.Windows.Forms.Padding(10)
$actionsGroup.Margin  = New-Object System.Windows.Forms.Padding(0)

$actionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionsLayout.Dock = 'Fill'
$actionsLayout.ColumnCount = 3
$actionsLayout.RowCount = 3
$actionsLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize

[void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.34)))

[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.34)))

$actionsGroup.Controls.Add($actionsLayout)

$mainLayout.Controls.Add($actionsGroup, 0, 2)

# ===================================================================================================
#									SECTION: GUI Buttons
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and configures buttons on the GUI form for various operations.
# ===================================================================================================

############################################################################
# 1) Change Machine Name Button
############################################################################
$changeMachineNameButton = New-Object System.Windows.Forms.Button
$changeMachineNameButton.Text = "Change Machine Name"
$changeMachineNameButton.Location = New-Object System.Drawing.Point(10, 120)
$changeMachineNameButton.Size = New-Object System.Drawing.Size(150, 35)
$changeMachineNameButton.Add_Click({
		
		# Ensure the script is running with administrative privileges
		if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
		{
			[System.Windows.Forms.MessageBox]::Show(
				"This script must be run as an administrator.",
				"Insufficient Privileges",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			return
		}
		
		# Capture old machine name for INI matching (before rename takes effect)
		$oldMachineName = $env:COMPUTERNAME
		
		# Get the new machine name from the user
		$newMachineNameInput = Get_NEW_Machine_Name
		
		if ($newMachineNameInput -eq $null)
		{
			# Handle cancellation
			if ($operationStatus -and $operationStatus.ContainsKey("MachineNameChange"))
			{
				$operationStatus["MachineNameChange"].Status = "Cancelled"
				$operationStatus["MachineNameChange"].Message = "Machine name change was cancelled by the user."
				$operationStatus["MachineNameChange"].Details = ""
			}
			[System.Windows.Forms.MessageBox]::Show(
				"Machine name change was cancelled.",
				"Cancelled",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			)
			return
		}
		
		# Confirm the change
		$result = [System.Windows.Forms.MessageBox]::Show(
			"Are you sure you want to change the machine name to '$newMachineNameInput'?",
			"Confirm Machine Name Change",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Question
		)
		
		if ($result -ne [System.Windows.Forms.DialogResult]::Yes)
		{
			return
		}
		
		# Determine the store number once (from INI)
		$currentStoreNumber = Get_Store_Number_From_INI
		if (-not $currentStoreNumber)
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Store number not found in startup.ini.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			if ($operationStatus -and $operationStatus.ContainsKey("MachineNameChange"))
			{
				$operationStatus["MachineNameChange"].Status = "Failed"
				$operationStatus["MachineNameChange"].Message = "Store number not found in startup.ini."
				$operationStatus["MachineNameChange"].Details = ""
			}
			return
		}
		
		# Robust machine-number extraction (normalize host; allow 1-3 trailing digits; pad to 3)
		$normalizedHost = $newMachineNameInput
		if ($null -eq $normalizedHost) { $normalizedHost = "" }
		$normalizedHost = $normalizedHost.Trim()
		$normalizedHost = $normalizedHost -replace '^[\\\/]+', ''
		if ($normalizedHost -match '[\\\/]') { $normalizedHost = ($normalizedHost -split '[\\\/]')[0] }
		if ($normalizedHost -match '\.') { $normalizedHost = ($normalizedHost -split '\.')[0] }
		$normalizedHost = $normalizedHost.Trim().ToUpper()
		
		if ($normalizedHost -notmatch '(\d{1,3})$')
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Invalid terminal number in machine name '$newMachineNameInput'. Must end with 1-3 digits (ex: POS6, POS006, ...).",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			if ($operationStatus -and $operationStatus.ContainsKey("MachineNameChange"))
			{
				$operationStatus["MachineNameChange"].Status = "Failed"
				$operationStatus["MachineNameChange"].Message = "Invalid machine name."
				$operationStatus["MachineNameChange"].Details = "Cannot extract machine number."
			}
			return
		}
		
		$machineNumber = ([int]$Matches[1]).ToString("D3")
		
		if ($machineNumber -eq "000")
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Invalid terminal number extracted from '$newMachineNameInput' (000 is not allowed).",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			if ($operationStatus -and $operationStatus.ContainsKey("MachineNameChange"))
			{
				$operationStatus["MachineNameChange"].Status = "Failed"
				$operationStatus["MachineNameChange"].Message = "Invalid machine name."
				$operationStatus["MachineNameChange"].Details = "Extracted terminal number 000."
			}
			return
		}
		
		# Proceed to change machine name
		try
		{
			Rename-Computer -NewName $normalizedHost -Force -ErrorAction Stop
			
			# Assign to script-level variable
			$script:newMachineName = $normalizedHost
			
			# Update machine name label
			if ($machineNameLabel)
			{
				$machineNameLabel.Text = "The machine name will change from: $env:COMPUTERNAME to $script:newMachineName"
				$machineNameLabel.Refresh()
				if ($machineNameLabel.Parent) { $machineNameLabel.Parent.PerformLayout(); $machineNameLabel.Parent.Refresh() }
				[System.Windows.Forms.Application]::DoEvents()
			}
			
			# Update operation status
			if ($operationStatus -and $operationStatus.ContainsKey("MachineNameChange"))
			{
				$operationStatus["MachineNameChange"].Status = "Successful"
				$operationStatus["MachineNameChange"].Message = "Machine name changed successfully."
				$operationStatus["MachineNameChange"].Details = "Machine name changed to '$script:newMachineName'."
			}
			
			# Remove old XF/XW folders (call operator protects weird chars in function names)
			& 'Remove_Old_XF/XW_Folders' -MachineName $script:newMachineName -StoreNumber $currentStoreNumber
			
			# Update startup.ini after changing machine name
			$startupIniPath = "\\localhost\storeman\startup.ini"
			
			# Build TER and DBSERVER values correctly (preserve instance name if present)
			$terValue = "TER=$machineNumber"
			$dbServerValue = $null
			
			if (Test-Path $startupIniPath)
			{
				$content = Get-Content $startupIniPath
				
				# Pull existing DBSERVER= line to preserve instance name if it exists
				$existingDbLine = $null
				foreach ($line in $content)
				{
					if ($line -match '^\s*(?i:DBSERVER)\s*=')
					{
						$existingDbLine = $line
						break
					}
				}
				
				$existingDbServer = $null
				if ($existingDbLine -match '^\s*(?i:DBSERVER)\s*=\s*(.+?)\s*$')
				{
					$existingDbServer = $matches[1].Trim()
				}
				
				# Extract instance name (if any)
				$instance = $null
				if ($existingDbServer -and ($existingDbServer -match '\\'))
				{
					$parts = $existingDbServer.Split('\')
					if ($parts.Count -ge 2 -and $parts[1]) { $instance = $parts[1] }
				}
				
				if ($instance)
				{
					$dbServerValue = "DBSERVER=$($script:newMachineName)\$instance"
				}
				else
				{
					$dbServerValue = "DBSERVER=$($script:newMachineName)"
				}
				
				# Replace TER and DBSERVER (tolerate whitespace/casing)
				$updatedContent = @()
				foreach ($line in $content)
				{
					if ($line -match '^\s*(?i:TER)\s*=')
					{
						$updatedContent += $terValue
					}
					elseif ($line -match '^\s*(?i:DBSERVER)\s*=')
					{
						$updatedContent += $dbServerValue
					}
					else
					{
						$updatedContent += $line
					}
				}
				
				Set-Content -Path $startupIniPath -Value $updatedContent -ErrorAction Stop
				
				if ($operationStatus -and $operationStatus.ContainsKey("StartupIniUpdate"))
				{
					$operationStatus["StartupIniUpdate"].Status = "Successful"
					$operationStatus["StartupIniUpdate"].Message = "startup.ini updated successfully."
					$operationStatus["StartupIniUpdate"].Details = "Updated TER to '$terValue' and DBSERVER to '$dbServerValue'."
				}
			}
			else
			{
				if ($operationStatus -and $operationStatus.ContainsKey("StartupIniUpdate"))
				{
					$operationStatus["StartupIniUpdate"].Status = "Failed"
					$operationStatus["StartupIniUpdate"].Message = "startup.ini file not found."
					$operationStatus["StartupIniUpdate"].Details = "File not found at $startupIniPath."
				}
			}
			
			# -----------------------------------------------------------------------------------------
			# SmsHttps.INI update (ONLY if the file exists; otherwise skip without error)
			# -----------------------------------------------------------------------------------------
			
			# Capture store before any possible "sync store number" logic happened (best-effort)
			$oldStoreNumberForSms = $currentStoreNumber
			
			# Re-read store in case Get_NEW_Machine_Name updated startup.ini store number
			$newStoreNumberForSms = Get_Store_Number_From_INI
			if (-not $newStoreNumberForSms) { $newStoreNumberForSms = $currentStoreNumber }
			
			$smsIniPath = $null
			$smsCandidates = @(
				"\\localhost\storeman\SmsHttps64\SmsHttps.INI",
				"C:\storeman\SmsHttps64\SmsHttps.INI",
				"D:\storeman\SmsHttps64\SmsHttps.INI"
			)
			
			foreach ($p in $smsCandidates)
			{
				if (Test-Path $p)
				{
					$smsIniPath = $p
					break
				}
			}
			
			if (-not $smsIniPath)
			{
				if ($operationStatus)
				{
					$operationStatus["SmsHttpsIniUpdate"] = [pscustomobject]@{
						Status  = "Skipped"
						Message = "SmsHttps.INI not found on this machine. No update performed."
						Details = "Looked for: $($smsCandidates -join ', ')"
					}
				}
			}
			else
			{
				try
				{
					$smsResult = Update_SmsHttpsINI `
													-IniPath $smsIniPath `
													-StoreNumber $newStoreNumberForSms `
													-MachineName $script:newMachineName `
													-OldStoreNumber $oldStoreNumberForSms `
													-OldMachineName $oldMachineName `
													-CreateIfMissing
					
					if ($operationStatus)
					{
						if ($smsResult -and $smsResult.Success)
						{
							$operationStatus["SmsHttpsIniUpdate"] = [pscustomobject]@{
								Status  = "Successful"
								Message = "SmsHttps.INI updated."
								Details = "Path: $smsIniPath | OldKey: $($smsResult.OldKey) | NewKey: $($smsResult.NewKey) | Redir901: $($smsResult.Redir901) | Changed: $($smsResult.Changed) | LicenseGuidCleared: $($smsResult.LicenseGuidCleared)"
							}
						}
						else
						{
							$operationStatus["SmsHttpsIniUpdate"] = [pscustomobject]@{
								Status  = "Failed"
								Message = "SmsHttps.INI update returned failure."
								Details = "Path: $smsIniPath | " + ($(if ($smsResult) { $smsResult.Message }
										else { "No result returned." }))
							}
						}
					}
				}
				catch
				{
					if ($operationStatus)
					{
						$operationStatus["SmsHttpsIniUpdate"] = [pscustomobject]@{
							Status  = "Failed"
							Message = "Unhandled error updating SmsHttps.INI."
							Details = "Path: $smsIniPath | Error: $($_.Exception.Message)"
						}
					}
				}
			}
			
			# SQL update
			$sqlUpdateResult = Update_SQL_Tables_For_Machine_Name_Change -storeNumber $currentStoreNumber -machineName $script:newMachineName -machineNumber $machineNumber
			
			if ($sqlUpdateResult -and $sqlUpdateResult.Success)
			{
				if ($operationStatus -and $operationStatus.ContainsKey("SQLDatabaseUpdate"))
				{
					$operationStatus["SQLDatabaseUpdate"].Status = "Successful"
					$operationStatus["SQLDatabaseUpdate"].Message = "SQL tables updated successfully after machine name change."
					$operationStatus["SQLDatabaseUpdate"].Details = "STO_TAB, TER_TAB, LNK_TAB, and RUN_TAB updated."
				}
			}
			else
			{
				if ($operationStatus -and $operationStatus.ContainsKey("SQLDatabaseUpdate"))
				{
					$operationStatus["SQLDatabaseUpdate"].Status = "Failed"
					$operationStatus["SQLDatabaseUpdate"].Message = "Failed to update SQL tables after machine name change."
					$operationStatus["SQLDatabaseUpdate"].Details = "SQL update result indicated failure."
				}
			}
			
			# Prompt reboot
			$rebootResult = [System.Windows.Forms.MessageBox]::Show(
				"Machine name changed successfully to '$script:newMachineName'. The system will need to reboot for changes to take effect. Do you want to reboot now?",
				"Success",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Information
			)
			
			if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Restart-Computer -Force
			}
		}
		catch
		{
			$errorMessage = $_.Exception.Message
			[System.Windows.Forms.MessageBox]::Show(
				"Error changing machine name: $errorMessage",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			
			if ($operationStatus -and $operationStatus.ContainsKey("MachineNameChange"))
			{
				$operationStatus["MachineNameChange"].Status = "Failed"
				$operationStatus["MachineNameChange"].Message = "Error changing machine name."
				$operationStatus["MachineNameChange"].Details = "Error: $errorMessage"
			}
		}
	})

############################################################################
# 2) Configure Network Button
############################################################################
$configureNetworkButton = New-Object System.Windows.Forms.Button
$configureNetworkButton.Text = "Configure Network"
$configureNetworkButton.Location = New-Object System.Drawing.Point(170, 120)
$configureNetworkButton.Size = New-Object System.Drawing.Size(150, 35)
$configureNetworkButton.Add_Click({
		# Implement the ConfigureNetwork function with GUI elements
		
		if ($currentConfigs -eq $null -or $currentConfigs.Count -eq 0)
		{
			[System.Windows.Forms.MessageBox]::Show("No valid active IP configuration found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
		
		# If multiple adapters are found, ask the user to select one
		if ($currentConfigs.Count -gt 1)
		{
			# Create a form to select the network adapter
			$adapterForm = New-Object System.Windows.Forms.Form
			$adapterForm.Text = "Select Network Adapter"
			$adapterForm.Size = New-Object System.Drawing.Size(400, 200)
			$adapterForm.StartPosition = "CenterParent"
			
			$label = New-Object System.Windows.Forms.Label
			$label.Text = "Select the network adapter:"
			$label.Location = New-Object System.Drawing.Point(10, 10)
			$label.Size = New-Object System.Drawing.Size(380, 20)
			$label.AutoSize = $true
			
			$listBox = New-Object System.Windows.Forms.ListBox
			$listBox.Location = New-Object System.Drawing.Point(10, 40)
			$listBox.Size = New-Object System.Drawing.Size(360, 80)
			
			for ($i = 0; $i -lt $currentConfigs.Count; $i++)
			{
				$adapterName = $currentConfigs[$i].InterfaceAlias
				$ipAddress = $currentConfigs[$i].IPv4Address.IPAddress
				$listBox.Items.Add("$adapterName - IP: $ipAddress")
			}
			
			$okButton = New-Object System.Windows.Forms.Button
			$okButton.Text = "OK"
			$okButton.Location = New-Object System.Drawing.Point(150, 130)
			$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
			
			$adapterForm.AcceptButton = $okButton
			$adapterForm.Controls.AddRange(@($label, $listBox, $okButton))
			$adapterForm.ShowDialog() | Out-Null
			
			$selectedIndex = $listBox.SelectedIndex
			if ($selectedIndex -ge 0 -and $selectedIndex -lt $currentConfigs.Count)
			{
				$currentConfig = $currentConfigs[$selectedIndex]
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("No network adapter selected.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return
			}
		}
		else
		{
			# Only one adapter found
			$currentConfig = $currentConfigs[0]
		}
		
		# Proceed with configuring network
		# Display current IP Address
		$currentIP = $currentConfig.IPv4Address.IPAddress
		$currentGateway = $currentConfig.IPv4DefaultGateway.NextHop
		$adapterName = $currentConfig.InterfaceAlias
		
		# Ask user for network type
		$networkTypeResult = [System.Windows.Forms.MessageBox]::Show("Will the adapter use DHCP?", "Network Configuration", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
		
		if ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Configure DHCP
			try
			{
				netsh interface ip set address name="$adapterName" source=dhcp
				netsh interface ip set dns name="$adapterName" source=dhcp
				[System.Windows.Forms.MessageBox]::Show("DHCP configuration applied.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
				$operationStatus["IPConfiguration"].Status = "Successful"
				$operationStatus["IPConfiguration"].Message = "DHCP configuration applied."
				$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName"
			}
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to configure DHCP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				$operationStatus["IPConfiguration"].Status = "Failed"
				$operationStatus["IPConfiguration"].Message = "Failed to configure DHCP."
				$operationStatus["IPConfiguration"].Details = "Error: $_"
			}
			
		}
		elseif ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::No)
		{
			# Configure Static IP
			# Loop until valid input or user cancels
			$validInput = $false
			while (-not $validInput)
			{
				# Ask for last octet
				$gatewayParts = $currentGateway.Split('.')
				if ($gatewayParts.Length -lt 4)
				{
					[System.Windows.Forms.MessageBox]::Show("Invalid gateway format: $currentGateway", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					$operationStatus["IPConfiguration"].Status = "Failed"
					$operationStatus["IPConfiguration"].Message = "Invalid gateway format."
					$operationStatus["IPConfiguration"].Details = "Gateway: $currentGateway"
					return
				}
				$gatewayPrefix = $gatewayParts[0 .. 2] -join '.'
				$ipForm = New-Object System.Windows.Forms.Form
				$ipForm.Text = "Enter Last Octet for Static IP"
				$ipForm.Size = New-Object System.Drawing.Size(350, 170)
				$ipForm.StartPosition = "CenterParent"
				
				$label = New-Object System.Windows.Forms.Label
				$label.Text = "Enter the last octet (1-254) of the static IP address $gatewayPrefix."
				$label.Location = New-Object System.Drawing.Point(10, 20)
				$label.Size = New-Object System.Drawing.Size(320, 40)
				$label.AutoSize = $true
				
				$textBox = New-Object System.Windows.Forms.TextBox
				$textBox.Location = New-Object System.Drawing.Point(10, 70)
				$textBox.Size = New-Object System.Drawing.Size(320, 20)
				
				$okButton = New-Object System.Windows.Forms.Button
				$okButton.Text = "OK"
				$okButton.Location = New-Object System.Drawing.Point(130, 100)
				$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
				
				$ipForm.AcceptButton = $okButton
				
				$ipForm.Controls.AddRange(@($label, $textBox, $okButton))
				$dialogResult = $ipForm.ShowDialog()
				
				if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
				{
					$lastOctet = $textBox.Text.Trim()
					# Validate last octet
					if ($lastOctet -match '^\d{1,3}$' -and [int]$lastOctet -ge 1 -and [int]$lastOctet -le 254)
					{
						$ipAddress = "$gatewayPrefix.$lastOctet"
						
						# Check if IP is in use
						$pingResult = Test-Connection -ComputerName $ipAddress -Count 1 -Quiet
						if ($pingResult)
						{
							[System.Windows.Forms.MessageBox]::Show("The IP address '$ipAddress' is already in use. Please choose a different one.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
						}
						else
						{
							try
							{
								netsh interface ip set address name="$adapterName" static $ipAddress 255.255.255.0 $currentGateway
								
								# Update the label
								$ipAddressLabel.Text = "IP Address changed from: $currentIP to $ipAddress"
								$operationStatus["IPConfiguration"].Status = "Successful"
								$operationStatus["IPConfiguration"].Message = "Static IP configuration applied."
								$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName, IP: $ipAddress"
								
								# Set DNS
								Set-DnsClientServerAddress -InterfaceAlias $adapterName -ServerAddresses ("8.8.8.8", "8.8.4.4")
								[System.Windows.Forms.MessageBox]::Show("Static IP configuration applied.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
								$operationStatus["IPConfiguration"].Status = "Successful"
								$operationStatus["IPConfiguration"].Message = "Static IP configuration applied."
								$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName, IP: $ipAddress"
								$validInput = $true
							}
							catch
							{
								[System.Windows.Forms.MessageBox]::Show("Failed to configure static IP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								$operationStatus["IPConfiguration"].Status = "Failed"
								$operationStatus["IPConfiguration"].Message = "Failed to configure static IP."
								$operationStatus["IPConfiguration"].Details = "Error: $_"
								$validInput = $true # Exit loop since an error occurred
							}
						}
					}
					else
					{
						[System.Windows.Forms.MessageBox]::Show("Invalid input. Please enter a valid number between 1 and 254.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					}
				}
				else
				{
					# User canceled the dialog
					$operationStatus["IPConfiguration"].Status = "Skipped"
					$operationStatus["IPConfiguration"].Message = "User canceled static IP configuration."
					$operationStatus["IPConfiguration"].Details = "" # Empty details
					$validInput = $true # Exit loop
				}
			}
			
		}
		else
		{
			$operationStatus["IPConfiguration"].Status = "Skipped"
			$operationStatus["IPConfiguration"].Message = "User chose not to configure network settings."
		}
	})

############################################################################
# 3) Update Store Number Button
############################################################################
$updateStoreNumberButton = New-Object System.Windows.Forms.Button
$updateStoreNumberButton.Text = "Update Store Number"
$updateStoreNumberButton.Location = New-Object System.Drawing.Point(330, 120)
$updateStoreNumberButton.Size = New-Object System.Drawing.Size(150, 35)
$updateStoreNumberButton.Add_Click({
		# Get the old store number from startup.ini
		$oldStoreNumber = Get_Store_Number_From_INI
		
		if ($oldStoreNumber -ne $null)
		{
			# Prompt for new store number
			$newStoreNumberInput = Get_NEW_Store_Number
			if ($newStoreNumberInput -ne $null)
			{
				# Show warning before updating
				$warningResult = [System.Windows.Forms.MessageBox]::Show("You are about to change the store number from '$oldStoreNumber' to '$newStoreNumberInput'. Do you want to proceed?", "Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
				
				if ($warningResult -eq [System.Windows.Forms.DialogResult]::Yes)
				{
					# Update startup.ini
					if (Test-Path $startupIniPath)
					{
						$updateSuccess = Update_Store_Number_In_INI -newStoreNumber $newStoreNumberInput
						if ($updateSuccess)
						{
							# Assign to script-level variable
							$script:newStoreNumber = $newStoreNumberInput
							
							# Update the label
							$storeNumberLabel.Text = "Store Number changed from: $currentStoreNumber to $script:newStoreNumber"
							$operationStatus["StoreNumberChange"].Status = "Successful"
							$operationStatus["StoreNumberChange"].Message = "Store number updated in startup.ini."
							$operationStatus["StoreNumberChange"].Details = "Store number changed to '$script:newStoreNumber'."
							
							# Inform the user about the new store number
							[System.Windows.Forms.MessageBox]::Show("Store number successfully changed to '$script:newStoreNumber'.", "Store Number Updated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
							
							# Call the SQL update function
							$sqlUpdateResult = Update_SQL_Tables_For_Store_Number_Change -storeNumber $script:newStoreNumber
							
							if ($sqlUpdateResult.Success)
							{
								$operationStatus["SQLDatabaseUpdate"].Status = "Successful"
								$operationStatus["SQLDatabaseUpdate"].Message = "STD_TAB updated successfully after store number change."
								$operationStatus["SQLDatabaseUpdate"].Details = "STD_TAB updated with new store number."
							}
							else
							{
								$operationStatus["SQLDatabaseUpdate"].Status = "Failed"
								$operationStatus["SQLDatabaseUpdate"].Message = "Failed to update STD_TAB after store number change."
								$operationStatus["SQLDatabaseUpdate"].Details = "Failed commands: $($sqlUpdateResult.FailedCommands -join ', ')"
							}
							
						}
						else
						{
							[System.Windows.Forms.MessageBox]::Show("Failed to update startup.ini.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
							$operationStatus["StoreNumberChange"].Status = "Failed"
							$operationStatus["StoreNumberChange"].Message = "Failed to update store number."
							$operationStatus["StoreNumberChange"].Details = "Error updating startup.ini."
						}
					}
					else
					{
						[System.Windows.Forms.MessageBox]::Show("startup.ini not found at $startupIniPath.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
						$operationStatus["StoreNumberChange"].Status = "Failed"
						$operationStatus["StoreNumberChange"].Message = "Failed to update store number."
						$operationStatus["StoreNumberChange"].Details = "startup.ini not found."
					}
				}
				else
				{
					$operationStatus["StoreNumberChange"].Status = "Cancelled"
					$operationStatus["StoreNumberChange"].Message = "Store number change was cancelled by the user."
					$operationStatus["StoreNumberChange"].Details = "Old store number remains '$oldStoreNumber'."
				}
			}
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show("Store number not found in startup.ini.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			$operationStatus["StoreNumberChange"].Status = "Failed"
			$operationStatus["StoreNumberChange"].Message = "Store number not found."
			$operationStatus["StoreNumberChange"].Details = "startup.ini not found or store number not defined."
		}
	})

############################################################################
# 4) Truncate Tables Button
############################################################################
$truncateTablesButton = New-Object System.Windows.Forms.Button
$truncateTablesButton.Text = "Truncate Tables"
$truncateTablesButton.Location = New-Object System.Drawing.Point(10, 160)
$truncateTablesButton.Size = New-Object System.Drawing.Size(150, 35)
$truncateTablesButton.Add_Click({
		# Define the tables to truncate
		$tablesToTruncate = @(
			"TRS_LOG", "TRS_FIN", "TRS_DPT", "TRS_SUB", "TRS_CLK",
			"TRS_CLT", "TRS_VND", "TRS_ITM",
			"RPT_CLT_D", "RPT_CLT_W", "RPT_CLT_M", "RPT_CLT_N", "RPT_CLT_Y", "RPT_CLT_P", "RPT_CLT_F", "RPT_CLT_ITM_D", "RPT_CLT_ITM_N",
			"RPT_FIN", "RPT_DPT", "RPT_SUB", "RPT_HOU", "RPT_VND",
			"COST_REV", "POS_REV", "OBJ_REV", "PRICE_REV", "REV_HDR", "SAL_REG_SAV", "SAL_HDR_SAV", "SAL_TTL_SAV", "SAL_DET_SAV",
			"RPT_CLK_D", "RPT_CLK_W", "RPT_CLK_M", "RPT_CLK_Y", "RPT_CLK_P", "RPT_CLK_F", "RPT_CLK_N",
			"RPT_ITM_D", "RPT_ITM_W", "RPT_ITM_M", "RPT_ITM_Y", "RPT_ITM_P", "RPT_ITM_F", "RPT_ITM_N",
			"DATA_REG",
			"RPT_CHK", "RENT_TAB", "TIM_TAB", "GAS_INVENT", "GAS_COUNT", "GAS_TRANS",
			"SAL_BAT", "SAL_HDR", "SAL_REG", "SAL_DET", "SAL_TTL",
			"REC_BAT", "REC_HDR", "REC_REG", "REC_TTL", "INV_HDR", "INV_REG", "INV_TTL",
			"RPT_PAY_N", "RPT_PAY_M", "TRS_PAY", "OFR_TAB"
		)
		
		$result = [System.Windows.Forms.MessageBox]::Show("Do you want to truncate the specified tables?", "Truncate Tables", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Truncate tables
			$failedTruncateTables = RPT_Blank -tables $tablesToTruncate
			
			if ($failedTruncateTables.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("Tables truncated successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
				$operationStatus["TableTruncation"].Status = "Successful"
				$operationStatus["TableTruncation"].Message = "Tables truncated successfully."
				$operationStatus["TableTruncation"].Details = "All tables were truncated successfully."
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to truncate some tables: $($failedTruncateTables -join ', ')", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				$operationStatus["TableTruncation"].Status = "Failed"
				$operationStatus["TableTruncation"].Message = "Failed to truncate some tables."
				$operationStatus["TableTruncation"].Details = "Failed tables: $($failedTruncateTables -join ', ')"
			}
		}
		else
		{
			$operationStatus["TableTruncation"].Status = "Skipped"
			$operationStatus["TableTruncation"].Message = "User chose not to truncate tables."
		}
	})

############################################################################
# 6) Registry Cleanup Button
############################################################################
$registryCleanupButton = New-Object System.Windows.Forms.Button
$registryCleanupButton.Text = "Registry Cleanup"
$registryCleanupButton.Location = New-Object System.Drawing.Point(330, 160)
$registryCleanupButton.Size = New-Object System.Drawing.Size(150, 35)
$registryCleanupButton.Add_Click({
		$result = [System.Windows.Forms.MessageBox]::Show("Do you want to delete all registry values starting with 'GT'?", "Registry Cleanup", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			$gtRegistryCleanupResult = Remove_GT_Registry_Values
			
			$registryStatus = $gtRegistryCleanupResult.Status
			
			if ($registryStatus -eq 'Successful')
			{
				$operationStatus["RegistryCleanup"].Status = "Successful"
				$operationStatus["RegistryCleanup"].Message = "GT registry values removed successfully."
				$operationStatus["RegistryCleanup"].Details = "$($gtRegistryCleanupResult.DeletedCount) 'GT' registry keys were deleted."
				[System.Windows.Forms.MessageBox]::Show("GT registry values removed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
			elseif ($registryStatus -eq 'Skipped')
			{
				$operationStatus["RegistryCleanup"].Status = "Skipped"
				$operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
				$operationStatus["RegistryCleanup"].Details = ""
			}
			else
			{
				$operationStatus["RegistryCleanup"].Status = "Failed"
				$operationStatus["RegistryCleanup"].Message = "Failed to remove GT registry values."
				$operationStatus["RegistryCleanup"].Details = $gtRegistryCleanupResult.Message
				[System.Windows.Forms.MessageBox]::Show("Failed to remove GT registry values.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		}
		else
		{
			$operationStatus["RegistryCleanup"].Status = "Skipped"
			$operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
		}
	})

############################################################################
# 9) Update Database Button (simple click call)
############################################################################
$updateSQLDatabaseButton = New-Object System.Windows.Forms.Button
$updateSQLDatabaseButton.Text = "Update SQL Database"
$updateSQLDatabaseButton.Location = New-Object System.Drawing.Point(10, 240)
$updateSQLDatabaseButton.Size = New-Object System.Drawing.Size(150, 35)

$updateSQLDatabaseButton.Add_Click({
		Update_SQL_Database `
								 -CurrentMachineName $currentMachineName `
								 -NewMachineName $script:newMachineName `
								 -OperationStatus $operationStatus
	})

############################################################################
# 10) Summary Button
############################################################################
$summaryButton = New-Object System.Windows.Forms.Button
$summaryButton.Text = "Show Summary"
$summaryButton.Location = New-Object System.Drawing.Point(170, 240)
$summaryButton.Size = New-Object System.Drawing.Size(150, 35)
$summaryButton.Add_Click({
		# Display the summary in a new form
		$summaryForm = New-Object System.Windows.Forms.Form
		$summaryForm.Text = "Operation Summary"
		$summaryForm.Size = New-Object System.Drawing.Size(600, 400)
		$summaryForm.StartPosition = "CenterParent"
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Multiline = $true
		$textBox.ReadOnly = $true
		$textBox.ScrollBars = "Vertical"
		$textBox.Dock = "Fill"
		
		# Build the summary text
		$summaryText = ""
		foreach ($operationKey in $operationStatus.Keys)
		{
			$operation = $operationStatus[$operationKey]
			$status = $operation.Status
			$message = $operation.Message
			$details = $operation.Details
			
			$summaryText += "${operationKey}: $status`r`n"
			if ($message -ne "")
			{
				$summaryText += "  Message: $message`r`n"
			}
			if ($details -ne "")
			{
				$summaryText += "  Details: $details`r`n"
			}
			$summaryText += "`r`n"
		}
		
		$textBox.Text = $summaryText
		$summaryForm.Controls.Add($textBox)
		$summaryForm.ShowDialog()
	})

############################################################################
# 11) Summary Button
############################################################################
$rebootButton = New-Object System.Windows.Forms.Button
$rebootButton.Text = "Reboot System"
$rebootButton.Location = New-Object System.Drawing.Point(330, 240)
$rebootButton.Size = New-Object System.Drawing.Size(150, 35)
$rebootButton.Add_Click({
		$rebootResult = [System.Windows.Forms.MessageBox]::Show("Do you want to reboot now?", "Reboot", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			Restart-Computer -Force
			# Clean Temp Folder
			Delete_Files/Folders -Path "$TempDir" -SpecifiedFiles "MiniGhost.ps1"
		}
	})

############################################################################
# Handle form closing event (X button)
############################################################################
$form.add_FormClosing({
		# Confirmation message box to confirm exit
		$confirmResult = [System.Windows.Forms.MessageBox]::Show(
			"Are you sure you want to exit?",
			"Confirm Exit",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Question
		)
		
		# If the user clicks No, cancel the form close action
		if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes)
		{
			$_.Cancel = $true
		}
		else
		{
			# Proceed with form closing and perform actions
			# Write-Log "Form is closing. Performing cleanup." "green"
			
			# Clean Temp Folder
			Delete_Files/Folders -Path "$TempDir" -SpecifiedFiles "MiniGhost.ps1"
		}
	})

# Apply consistent styling + layout behavior for all main action buttons
$mainButtons = @(
	$updateStoreNumberButton, $configureNetworkButton, $changeMachineNameButton,
	$truncateTablesButton, $registryCleanupButton, $updateSQLDatabaseButton,
	$summaryButton, $rebootButton
)

foreach ($btn in $mainButtons)
{
	$btn.Dock = 'Fill'
	$btn.Margin = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
	$btn.MinimumSize = New-Object System.Drawing.Size(0, 42)
	$btn.UseVisualStyleBackColor = $true
}

# Tooltips (one per button)
$toolTip.SetToolTip($updateStoreNumberButton, "Update the store number (startup.ini) and refresh the on-screen store info. If supported, also syncs related SQL/INI values.")
$toolTip.SetToolTip($configureNetworkButton, "Configure NIC settings (IP/Subnet/Gateway/DNS) using store/lane conventions and validate the current network configuration.")
$toolTip.SetToolTip($changeMachineNameButton, "Rename this computer (supports formats like LANE003 / SCO012 / POS901 and prefixed formats like 0231LANE006).")
$toolTip.SetToolTip($truncateTablesButton, "Run SQL cleanup by truncating selected tables (maintenance routine).")
$toolTip.SetToolTip($registryCleanupButton, "Remove 'GT' registry values (after confirmation) to clean up old configuration leftovers.")
$toolTip.SetToolTip($updateSQLDatabaseButton, "Update SQL records to reflect the new machine name/store number for this node.")
$toolTip.SetToolTip($summaryButton, "View a session summary of which operations succeeded, failed, or were skipped.")
$toolTip.SetToolTip($rebootButton, "Restart the computer (optional cleanup before reboot).")

# Add buttons to the Actions grid (3 columns x 3 rows)
# Row 1
$actionsLayout.Controls.Add($changeMachineNameButton, 0, 0)
$actionsLayout.Controls.Add($configureNetworkButton,   1, 0)
$actionsLayout.Controls.Add($updateStoreNumberButton, 2, 0)

# Row 2 (middle cell intentionally left empty to preserve the classic 3-column layout)
$actionsLayout.Controls.Add($truncateTablesButton,     0, 1)
$actionsLayout.Controls.Add($registryCleanupButton,    2, 1)

# Row 3
$actionsLayout.Controls.Add($updateSQLDatabaseButton,  0, 2)
$actionsLayout.Controls.Add($summaryButton,            1, 2)
$actionsLayout.Controls.Add($rebootButton,             2, 2)

# ===================================================================================================
#                                       SECTION: Show the Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays the main GUI form to the user.
# ===================================================================================================

[void]$form.ShowDialog()

# Indicate the script is closing
Write-Host "Script closing..." -ForegroundColor Yellow

# Close the console to aviod duplicate logging to the richbox
exit
