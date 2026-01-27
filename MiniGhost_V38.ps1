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
$VersionNumber = "1.3.1"
$VersionDate = "2026-01-27"

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
# FUNCTION: Remove_Old_XZ_Folders
# ---------------------------------------------------------------------------------------------------
# Description:
# Removes old X* and Z* folders under \storeman\office that do NOT match the NEW store and allowed
# terminal numbers.
#
# Strict folder name pattern acted upon:
#   ^(X|Z)[A-Za-z](store 3-4 digits)(terminal 3 digits)$
# Examples: XF0231006, XW0242901, ZF1234006, ZW1234901
#
# Keeps ONLY folders that match:
#   - Store = new store (3-digit OR 4-digit representation)
#   - Terminal IN: current lane terminal (from MachineName), 900, 901
#
# Safeguard: refuses to run if extracted terminal is 901 unless -AllowBackoffice is passed.
# Automatic: no prompts.
# ===================================================================================================

function Remove_Old_XZ_Folders
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus,
		[Parameter(Mandatory = $false)]
		[switch]$AllowBackoffice
	)
	
	# -----------------------------
	# Init / status bucket
	# -----------------------------
	$deletedFolders = @()
	$failedToDeleteFolders = @()
	$anyFoldersSeen = $false
	$anyCandidatesMatched = $false
	
	if ($OperationStatus)
	{
		if (-not $OperationStatus.ContainsKey("OldXFoldersDeletion"))
		{
			$OperationStatus["OldXFoldersDeletion"] = [pscustomobject]@{
				Status  = ""
				Message = ""
				Details = ""
			}
		}
	}
	
	# -----------------------------
	# Validate store number
	# -----------------------------
	$storeNumberTrim = $StoreNumber
	if ($null -eq $storeNumberTrim) { $storeNumberTrim = "" }
	$storeNumberTrim = $storeNumberTrim.Trim()
	
	if ($storeNumberTrim -notmatch '^(?!0+$)\d{3,4}$')
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
	
	$storeInt = [int]$storeNumberTrim
	$store3 = $storeInt.ToString("D3")
	$store4 = $storeInt.ToString("D4")
	
	# -----------------------------
	# Pick base path
	# -----------------------------
	$possibleBasePaths = @("\\localhost\storeman\office", "C:\storeman\office", "D:\storeman\office")
	$basePath = $null
	foreach ($p in $possibleBasePaths)
	{
		if (Test-Path $p) { $basePath = $p; break }
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
	
	# -----------------------------
	# Extract machine terminal (last 1-3 digits, padded to 3)
	# -----------------------------
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
	
	# Safeguard for backoffice unless explicitly allowed
	if ($machineNumber -eq "901" -and -not $AllowBackoffice)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Refusing to run because extracted terminal number is 901 (backoffice). Use -AllowBackoffice to override."
			$OperationStatus["OldXFoldersDeletion"].Details = "This could delete many lane folders. Override only if intentional."
		}
		Write-Host "Failed: Refusing to run on terminal 901 (backoffice safeguard). Use -AllowBackoffice to override." -ForegroundColor Red
		return
	}
	
	# Keep set: current lane + 900 + 901
	$keepStores = @($store3, $store4)
	$keepTerminals = @($machineNumber, "900", "901")
	
	# -----------------------------
	# Enumerate X* and Z* folders
	# -----------------------------
	$folders = @()
	
	$xFolders = Get-ChildItem -Path $basePath -Directory -Filter "X*" -ErrorAction SilentlyContinue
	if ($xFolders) { $folders += $xFolders }
	
	$zFolders = Get-ChildItem -Path $basePath -Directory -Filter "Z*" -ErrorAction SilentlyContinue
	if ($zFolders) { $folders += $zFolders }
	
	if (-not $folders -or $folders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Folders Found"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X* or Z* folders were found under '$basePath'."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: No X* or Z* folders found under '$basePath'." -ForegroundColor Cyan
		return
	}
	
	$anyFoldersSeen = $true
	
	foreach ($folder in $folders)
	{
		$folderName = $folder.Name
		
		# STRICT pattern only: (X|Z)(letter)(store 3-4 digits)(terminal 3 digits)
		if ($folderName -match '^(?<prefix>[XZ][A-Za-z])(?<store>\d{3,4})(?<terminal>\d{3})$')
		{
			$folderStore = $matches['store']
			$folderTerminal = $matches['terminal']
			
			$storeOk = ($keepStores -contains $folderStore)
			$terminalOk = ($keepTerminals -contains $folderTerminal)
			
			# Delete if it does NOT match new store AND allowed terminal(s)
			if (-not ($storeOk -and $terminalOk))
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
							Write-Host "Failed to delete folder: $folderName. Error: $($_.Exception.Message)" -ForegroundColor Red
						}
					}
				}
			}
		}
	}
	
	# -----------------------------
	# Results
	# -----------------------------
	$resultMessage = ""
	if ($deletedFolders.Count -gt 0)
	{
		$resultMessage += "Deleted folders:`n$($deletedFolders -join "`n")`n"
	}
	if ($failedToDeleteFolders.Count -gt 0)
	{
		$resultMessage += "Failed to delete folders:`n$($failedToDeleteFolders -join "`n")`n"
	}
	
	if (-not $anyCandidatesMatched)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Matching Folders"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X*/Z* folders needed deletion. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: Nothing to delete. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')." -ForegroundColor Cyan
		return
	}
	
	if ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Successful"
			$OperationStatus["OldXFoldersDeletion"].Message = "Old X*/Z* folders deleted successfully."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Success: Old X*/Z* folders deleted successfully." -ForegroundColor Green
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Green }
	}
	elseif ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -gt 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Partial Failure"
			$OperationStatus["OldXFoldersDeletion"].Message = "Some old X*/Z* folders could not be deleted."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Warning: Some old X*/Z* folders could not be deleted." -ForegroundColor Yellow
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Yellow }
	}
	else
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Failed to delete any old X*/Z* folders."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Error: Failed to delete any old X*/Z* folders." -ForegroundColor Red
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
# 									FUNCTION: Update_INIs
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates ALL required INI files with the new store number.
# Supports both 3-digit and 4-digit store numbers.
# NOTE: No backups are created.
#
# Always-best-effort behavior:
# - Does NOT prompt
# - If MachineName missing -> uses $env:COMPUTERNAME
# - If terminal digits cannot be extracted -> still updates other INIs; skips SmsHttps Processor update
# - If old store cannot be detected -> still forces STORE= and REDIR lines; skips token-replace mapping
# - SmsHttps.INI: enforces exactly ONE processor entry (purges all others), and forces REDIR to <STORE>901
#
# FIX INCLUDED:
# - Startup.ini / Server.ini / SMSStart.ini / INFO_*901_WIN.ini now preserve ORIGINAL encoding + ORIGINAL newline style
#   (same approach as SmsHttps.INI: read bytes -> detect encoding -> detect CRLF/LF -> write back with same)
#
# NEW INCLUDED:
# - Updates Server.ini when present in the SAME folder as Startup.ini (and/or known storeman roots)
#   Updates relevant fields:
#     STORE=, REDIRMAIL=, REDIRMSG=, TER= (if terminal extracted), SERVERNAME= (token replace mapping as applicable)
# ===================================================================================================

function Update_INIs
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$newStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath,
		[Parameter(Mandatory = $false)]
		[string]$GlobalSmsStartIniPath,
		[Parameter(Mandatory = $false)]
		[string]$WinIniPath,
		[Parameter(Mandatory = $false)]
		[string]$SmsHttpsIniPath,
		[Parameter(Mandatory = $false)]
		[switch]$CreateSmsHttpsIfMissing
	)
	
	$success = $true
	
	# If caller didn't supply MachineName, use current machine name (no prompt, no rename)
	if ([string]::IsNullOrWhiteSpace($MachineName))
	{
		$MachineName = $env:COMPUTERNAME
	}
	
	# ------------------------------------------------------------------------------------------------
	# Normalize NEW store (preserve user width for REDIR/STORE lines)
	# ------------------------------------------------------------------------------------------------
	$newStoreTrim = $newStoreNumber
	if ($null -eq $newStoreTrim) { $newStoreTrim = "" }
	$newStoreTrim = $newStoreTrim.Trim()
	
	$newStoreInt = [int]$newStoreTrim
	$newStore3 = $newStoreInt.ToString("D3")
	$newStore4 = $newStoreInt.ToString("D4")
	
	# ------------------------------------------------------------------------------------------------
	# Normalize machine -> extract terminal (last 1-3 digits) padded to 3 (used for SmsHttps key)
	# If not possible, we still update other INIs; just skip SmsHttps processor update.
	# ------------------------------------------------------------------------------------------------
	$mn = $MachineName
	if ($null -eq $mn) { $mn = "" }
	$mn = $mn.Trim()
	$mn = $mn -replace '^[\\\/]+', ''
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] }
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] }
	$mn = $mn.Trim().ToUpper()
	
	$newTerminal = $null
	if ($mn -match '(\d{1,3})$')
	{
		$newTerminal = ([int]$Matches[1]).ToString("D3")
		if ($newTerminal -eq "000") { $newTerminal = $null }
	}
	
	# ===============================================================================================
	# 1) Resolve + Update Startup.ini (required)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($StartupIniPath))
	{
		$startupCandidates = @(
			"\\localhost\storeman\Startup.ini",
			"C:\storeman\Startup.ini",
			"D:\storeman\Startup.ini"
		)
		foreach ($c in $startupCandidates)
		{
			if (Test-Path $c) { $StartupIniPath = $c; break }
		}
	}
	
	if ([string]::IsNullOrWhiteSpace($StartupIniPath) -or -not (Test-Path $StartupIniPath))
	{
		Write-Host "startup.ini not found. Provide -StartupIniPath or ensure it exists in default locations." -ForegroundColor Red
		return $false
	}
	
	$oldStoreDetected = $null
	$startupDir = $null
	try { $startupDir = Split-Path -Path $StartupIniPath -Parent }
	catch { $startupDir = $null }
	
	try
	{
		# --- Read preserving encoding + newline (SmsHttps-style) ---
		$bytes = [System.IO.File]::ReadAllBytes($StartupIniPath)
		
		$encStartup = $null
		if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
		{
			$encStartup = New-Object System.Text.UTF8Encoding($true)
		}
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
		{
			$encStartup = [System.Text.Encoding]::Unicode
		}
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
		{
			$encStartup = [System.Text.Encoding]::BigEndianUnicode
		}
		else
		{
			$encStartup = [System.Text.Encoding]::Default
		}
		
		$textStartup = $encStartup.GetString($bytes)
		
		$nlStartup = "`r`n"
		if ($textStartup -notmatch "`r`n" -and $textStartup -match "`n") { $nlStartup = "`n" }
		
		$startupLines = $textStartup -split "\r?\n", -1
		
		# Detect OLD store number (prefer STORE=)
		foreach ($l in $startupLines)
		{
			if ($l -match '^\s*STORE\s*=\s*(\d{3,4})\s*$')
			{
				$oldStoreDetected = $matches[1]
				break
			}
		}
		
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '^\s*SERVERNAME\s*=\s*(\d{3,4})[A-Za-z]')
				{
					$oldStoreDetected = $matches[1]
					break
				}
			}
		}
		
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '\\\\(\d{3,4})[A-Za-z]')
				{
					$oldStoreDetected = $matches[1]
					break
				}
			}
		}
		
		# If user provided OldStoreNumber, it overrides detection (as long as valid)
		if (-not [string]::IsNullOrWhiteSpace($OldStoreNumber))
		{
			$os = $OldStoreNumber.Trim()
			if ($os -match '^(?!0+$)\d{3,4}$')
			{
				$oldStoreDetected = $os
			}
		}
		
		# If still not detected, do NOT abort - we can still force STORE=/REDIR= lines.
		$doTokenReplace = $true
		if (-not $oldStoreDetected)
		{
			$doTokenReplace = $false
		}
		
		$oldTokens = @()
		$newTokens = @()
		
		if ($doTokenReplace)
		{
			$oldStoreInt = [int]$oldStoreDetected
			$oldStore3 = $oldStoreInt.ToString("D3")
			$oldStore4 = $oldStoreInt.ToString("D4")
			
			# Replace order: 4-digit first, then 3-digit
			$oldTokens = @($oldStore4, $oldStore3)
			$newTokens = @($newStore4, $newStore3)
		}
		
		for ($i = 0; $i -lt $startupLines.Count; $i++)
		{
			$line = $startupLines[$i]
			
			# Force STORE=
			$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
			
			# REDIRMAIL/REDIRMSG=<store>901 (keep 901 suffix)
			$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
			
			# Replace store token wherever it appears (only if we detected old store)
			if ($doTokenReplace)
			{
				for ($t = 0; $t -lt $oldTokens.Count; $t++)
				{
					$oldEsc = [regex]::Escape($oldTokens[$t])
					$newTok = $newTokens[$t]
					
					# before letters: 0242SERVER001
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
					# before 3 digits: 0242901, XF0242901
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
					# standalone token
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?!\d)"), $newTok
				}
			}
			
			$startupLines[$i] = $line
		}
		
		# --- Write preserving encoding + newline ---
		$outTextStartup = ($startupLines -join $nlStartup)
		[System.IO.File]::WriteAllText($StartupIniPath, $outTextStartup, $encStartup)
		
		Write-Host "Updated startup.ini" -ForegroundColor Green
		
		# If old store wasn't detected, keep OldStoreNumber fallback for downstream (optional use)
		if (-not $oldStoreDetected)
		{
			$oldStoreDetected = $newStoreTrim
		}
	}
	catch
	{
		$success = $false
		Write-Host "Failed updating startup.ini: $($_.Exception.Message)" -ForegroundColor Red
		# still continue to other INIs
		if (-not $oldStoreDetected) { $oldStoreDetected = $newStoreTrim }
	}
	
	# ===============================================================================================
	# 1B) Update Server.ini (optional; when present near Startup.ini / storeman root)
	#     - Same folder as Startup.ini is primary
	#     - Also tries common roots if not found there
	#     - Updates: STORE=, REDIRMAIL/REDIRMSG=<store>901, TER= (if terminal extracted)
	#     - Performs token replace mapping like Startup.ini when old store detected
	# ===============================================================================================
	$serverIniCandidates = @()
	
	if (-not [string]::IsNullOrWhiteSpace($startupDir))
	{
		$serverIniCandidates += (Join-Path $startupDir "Server.ini")
	}
	# common roots fallback
	$serverIniCandidates += @(
		"\\localhost\storeman\Server.ini",
		"C:\storeman\Server.ini",
		"D:\storeman\Server.ini"
	)
	
	$serverIniPath = $null
	foreach ($p in $serverIniCandidates)
	{
		if (-not [string]::IsNullOrWhiteSpace($p) -and (Test-Path $p))
		{
			$serverIniPath = $p
			break
		}
	}
	
	if (-not [string]::IsNullOrWhiteSpace($serverIniPath) -and (Test-Path $serverIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($serverIniPath)
			
			$encServer = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{
				$encServer = New-Object System.Text.UTF8Encoding($true)
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{
				$encServer = [System.Text.Encoding]::Unicode
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{
				$encServer = [System.Text.Encoding]::BigEndianUnicode
			}
			else
			{
				$encServer = [System.Text.Encoding]::Default
			}
			
			$textServer = $encServer.GetString($bytes)
			
			$nlServer = "`r`n"
			if ($textServer -notmatch "`r`n" -and $textServer -match "`n") { $nlServer = "`n" }
			
			$serverLines = $textServer -split "\r?\n", -1
			
			$doTokenReplaceS = $false
			$oldTokensS = @()
			$newTokensS = @()
			
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplaceS = $true
				$oldStoreIntS = [int]$oldStoreDetected
				$oldTokensS = @($oldStoreIntS.ToString("D4"), $oldStoreIntS.ToString("D3"))
				$newTokensS = @($newStore4, $newStore3)
			}
			
			for ($i = 0; $i -lt $serverLines.Count; $i++)
			{
				$line = $serverLines[$i]
				
				# STORE=
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				
				# TER= (only if we extracted terminal)
				if ($newTerminal)
				{
					$line = $line -replace '^\s*TER\s*=\s*\d{1,4}\s*$', ("TER=" + $newTerminal)
				}
				
				# REDIRMAIL / REDIRMSG keep 901 suffix
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				# SERVERNAME= line token replace (and any other occurrences) when old store detected
				if ($doTokenReplaceS)
				{
					for ($t = 0; $t -lt $oldTokensS.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokensS[$t])
						$newTok = $newTokensS[$t]
						
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?!\d)"), $newTok
					}
				}
				
				$serverLines[$i] = $line
			}
			
			$outTextServer = ($serverLines -join $nlServer)
			[System.IO.File]::WriteAllText($serverIniPath, $outTextServer, $encServer)
			
			Write-Host "Updated Server.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating Server.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 2) Update Global SMSStart.ini (optional; skip if missing)
	#    FIX: also update/ensure TER=<terminal> inside [SMSSTART]
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath))
	{
		$smsStartCandidates = @(
			"\\localhost\storeman\SMSStart.ini",
			"C:\storeman\SMSStart.ini",
			"D:\storeman\SMSStart.ini"
		)
		foreach ($c in $smsStartCandidates)
		{
			if (Test-Path $c) { $GlobalSmsStartIniPath = $c; break }
		}
	}
	
	if (-not [string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath) -and (Test-Path $GlobalSmsStartIniPath))
	{
		try
		{
			# --- Read preserving encoding + newline (SmsHttps-style) ---
			$bytes = [System.IO.File]::ReadAllBytes($GlobalSmsStartIniPath)
			
			$encSmsStart = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{
				$encSmsStart = New-Object System.Text.UTF8Encoding($true)
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{
				$encSmsStart = [System.Text.Encoding]::Unicode
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{
				$encSmsStart = [System.Text.Encoding]::BigEndianUnicode
			}
			else
			{
				$encSmsStart = [System.Text.Encoding]::Default
			}
			
			$textSmsStart = $encSmsStart.GetString($bytes)
			
			$nlSmsStart = "`r`n"
			if ($textSmsStart -notmatch "`r`n" -and $textSmsStart -match "`n") { $nlSmsStart = "`n" }
			
			$globalLines = $textSmsStart -split "\r?\n", -1
			
			$doTokenReplace = $false
			$oldTokens = @()
			$newTokens = @()
			
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplace = $true
				$oldStoreInt = [int]$oldStoreDetected
				$oldTokens = @($oldStoreInt.ToString("D4"), $oldStoreInt.ToString("D3"))
				$newTokens = @($newStore4, $newStore3)
			}
			
			$out = @()
			$inSmsStartSection = $false
			$terFound = $false
			$terLine = $null
			if ($newTerminal) { $terLine = "TER=$newTerminal" }
			
			foreach ($raw in $globalLines)
			{
				$line = $raw
				
				# Section header?
				if ($line -match '^\s*\[(.+?)\]\s*$')
				{
					# Leaving [SMSSTART] -> ensure TER exists (if we can build it)
					if ($inSmsStartSection -and -not $terFound -and $terLine)
					{
						$out += $terLine
						$terFound = $true
					}
					
					$sectionName = $matches[1].Trim()
					$inSmsStartSection = ($sectionName -ieq 'SMSSTART')
					if ($inSmsStartSection) { $terFound = $false }
					
					$out += $line
					continue
				}
				
				if ($inSmsStartSection)
				{
					# Force STORE= inside [SMSSTART]
					$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
					
					# Keep 901 suffix
					$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$',
					('$1=' + $newStoreTrim + '$2')
					
					# FIX: Update TER= inside [SMSSTART] (only if terminal was extracted)
					if ($terLine -and ($line -match '^\s*(?i:TER)\s*='))
					{
						$line = $terLine
						$terFound = $true
					}
					
					# Replace store token contexts (if present in this section)
					if ($doTokenReplace)
					{
						for ($t = 0; $t -lt $oldTokens.Count; $t++)
						{
							$oldEsc = [regex]::Escape($oldTokens[$t])
							$newTok = $newTokens[$t]
							
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?!\d)"), $newTok
						}
					}
				}
				
				$out += $line
			}
			
			# EOF while still in [SMSSTART] -> ensure TER exists
			if ($inSmsStartSection -and -not $terFound -and $terLine)
			{
				$out += $terLine
			}
			
			# --- Write preserving encoding + newline ---
			$outTextSmsStart = ($out -join $nlSmsStart)
			[System.IO.File]::WriteAllText($GlobalSmsStartIniPath, $outTextSmsStart, $encSmsStart)
			
			Write-Host "Updated SMSStart.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SMSStart.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 3) Update INFO_*901_WIN.ini (optional; skip if missing)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($WinIniPath))
	{
		$officeBases = @("\\localhost\storeman\office", "C:\storeman\office", "D:\storeman\office")
		foreach ($b in $officeBases)
		{
			if (Test-Path $b)
			{
				$f = Get-ChildItem -Path $b -File -Filter "INFO_*901_WIN.ini" -ErrorAction SilentlyContinue | Select-Object -First 1
				if ($f) { $WinIniPath = $f.FullName; break }
			}
		}
	}
	
	if (-not [string]::IsNullOrWhiteSpace($WinIniPath) -and (Test-Path $WinIniPath))
	{
		try
		{
			# --- Read preserving encoding + newline (SmsHttps-style) ---
			$bytes = [System.IO.File]::ReadAllBytes($WinIniPath)
			
			$encWin = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{
				$encWin = New-Object System.Text.UTF8Encoding($true)
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{
				$encWin = [System.Text.Encoding]::Unicode
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{
				$encWin = [System.Text.Encoding]::BigEndianUnicode
			}
			else
			{
				$encWin = [System.Text.Encoding]::Default
			}
			
			$textWin = $encWin.GetString($bytes)
			
			$nlWin = "`r`n"
			if ($textWin -notmatch "`r`n" -and $textWin -match "`n") { $nlWin = "`n" }
			
			$winLines = $textWin -split "\r?\n", -1
			
			$doTokenReplace = $false
			$oldTokens = @()
			$newTokens = @()
			
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplace = $true
				$oldStoreInt = [int]$oldStoreDetected
				$oldTokens = @($oldStoreInt.ToString("D4"), $oldStoreInt.ToString("D3"))
				$newTokens = @($newStore4, $newStore3)
			}
			
			for ($j = 0; $j -lt $winLines.Count; $j++)
			{
				$line = $winLines[$j]
				
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				if ($doTokenReplace)
				{
					for ($t = 0; $t -lt $oldTokens.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokens[$t])
						$newTok = $newTokens[$t]
						
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?!\d)"), $newTok
					}
				}
				
				$winLines[$j] = $line
			}
			
			# --- Write preserving encoding + newline ---
			$outTextWin = ($winLines -join $nlWin)
			[System.IO.File]::WriteAllText($WinIniPath, $outTextWin, $encWin)
			
			Write-Host "Updated INFO_*901_WIN.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating INFO_*901_WIN.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 4) Update SmsHttps.INI (optional; skip if missing)
	#     - Update [PROCESSORS] / [PROCESSOR] to EXACTLY ONE entry matching <STORE><LANE>
	#     - Purge any other processor entries
	#     - Force REDIRMAIL/REDIRMSG=<STORE>901
	#     - Clear LicenseGUID in [GENERAL]
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($SmsHttpsIniPath))
	{
		$smsHttpsCandidates = @(
			"\\localhost\storeman\SmsHttps64\SmsHttps.INI",
			"C:\storeman\SmsHttps64\SmsHttps.INI",
			"D:\storeman\SmsHttps64\SmsHttps.INI"
		)
		foreach ($c in $smsHttpsCandidates)
		{
			if (Test-Path $c) { $SmsHttpsIniPath = $c; break }
		}
	}
	
	if (-not [string]::IsNullOrWhiteSpace($SmsHttpsIniPath) -and (Test-Path $SmsHttpsIniPath))
	{
		try
		{
			# If we can't extract a lane number, skip processor update (but still clear LicenseGUID)
			$canUpdateProcessor = $true
			if ([string]::IsNullOrWhiteSpace($newTerminal))
			{
				$canUpdateProcessor = $false
			}
			
			$bytes = [System.IO.File]::ReadAllBytes($SmsHttpsIniPath)
			
			$enc = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{
				$enc = New-Object System.Text.UTF8Encoding($true)
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{
				$enc = [System.Text.Encoding]::Unicode
			}
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{
				$enc = [System.Text.Encoding]::BigEndianUnicode
			}
			else
			{
				$enc = [System.Text.Encoding]::Default
			}
			
			$text = $enc.GetString($bytes)
			
			$nl = "`r`n"
			if ($text -notmatch "`r`n" -and $text -match "`n") { $nl = "`n" }
			
			$lines = $text -split "\r?\n", -1
			
			$storeWidth = $newStoreTrim.Length
			$storeNorm = ([int]$newStoreTrim).ToString(("D{0}" -f $storeWidth))
			
			$newKey = $null
			if ($canUpdateProcessor)
			{
				$newKey = $storeNorm + $newTerminal
			}
			$redirKey901 = $storeNorm + "901"
			
			# oldTerminal (optional)
			$oldTerminal = $null
			if (-not [string]::IsNullOrWhiteSpace($OldMachineName))
			{
				$omn = $OldMachineName
				if ($null -eq $omn) { $omn = "" }
				$omn = $omn.Trim()
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
			
			$oldKeyExact = $null
			if (-not [string]::IsNullOrWhiteSpace($OldStoreNumber) -and $oldTerminal)
			{
				$osn = $OldStoreNumber.Trim()
				if ($osn -match '^(?!0+$)\d{3,4}$')
				{
					$osWidth = $osn.Length
					$oldKeyExact = ([int]$osn).ToString(("D{0}" -f $osWidth)) + $oldTerminal
				}
			}
			
			# Find section [PROCESSORS] or [PROCESSOR]
			$secStart = -1
			for ($i = 0; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[\s*PROCESSORS?\s*\]\s*$')
				{
					$secStart = $i
					break
				}
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
			
			$processorsChanged = $false
			$licenseGuidCleared = $false
			
			# ---- Processor enforcement (ONLY if we can build the newKey)
			if ($canUpdateProcessor)
			{
				if ($secStart -ge 0)
				{
					$entries = @()
					for ($i = $secStart + 1; $i -lt $secEnd; $i++)
					{
						if ($lines[$i] -match '^(?<ws>\s*)(?<key>\d{6,7})\s*=\s*(?<rhs>.*)$')
						{
							$entries += [pscustomobject]@{
								Index = $i
								Ws    = $matches['ws']
								Key   = $matches['key']
								Rhs   = $matches['rhs']
							}
						}
					}
					
					# pick a source rhs
					$source = $null
					if ($oldKeyExact)
					{
						foreach ($e in $entries) { if ($e.Key -eq $oldKeyExact) { $source = $e; break } }
					}
					if (-not $source -and $oldTerminal)
					{
						foreach ($e in $entries)
						{
							if ($e.Key.Length -ge 3 -and $e.Key.Substring($e.Key.Length - 3) -eq $oldTerminal) { $source = $e; break }
						}
					}
					if (-not $source)
					{
						foreach ($e in $entries)
						{
							if ($e.Key.Length -ge 3 -and $e.Key.Substring($e.Key.Length - 3) -eq $newTerminal) { $source = $e; break }
						}
					}
					if (-not $source -and $entries.Count -gt 0) { $source = $entries[0] }
					
					$rhsUpdated = ""
					$wsToUse = ""
					if ($source)
					{
						$rhsUpdated = $source.Rhs
						$wsToUse = $source.Ws
					}
					
					# Replace any old processor keys inside RHS -> newKey (fixes 2222004... lingering)
					foreach ($e in $entries)
					{
						$rhsUpdated = $rhsUpdated -replace ("(?<!\d)" + [regex]::Escape($e.Key) + "(?!\d)"), $newKey
					}
					if ($oldKeyExact)
					{
						$rhsUpdated = $rhsUpdated -replace ("(?<!\d)" + [regex]::Escape($oldKeyExact) + "(?!\d)"), $newKey
					}
					
					# Strip existing REDIRMAIL/REDIRMSG then enforce correct
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMAIL\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMSG\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					
					$rhsUpdated = $rhsUpdated -replace ',{2,}', ','
					$rhsUpdated = $rhsUpdated.Trim().Trim(',').Trim()
					
					if ([string]::IsNullOrWhiteSpace($rhsUpdated))
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901,TARGETSEND=,TARGETRECV=,DEADLOCKPRIORITY="
					}
					else
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901," + $rhsUpdated
					}
					
					$newProcessorLine = $wsToUse + $newKey + "=" + $rhsUpdated
					
					# Purge all entries; insert ONE
					$insertAt = $secStart + 1
					if ($entries.Count -gt 0)
					{
						$min = $entries[0].Index
						foreach ($e in $entries) { if ($e.Index -lt $min) { $min = $e.Index } }
						$insertAt = $min
					}
					
					$out = @()
					$inserted = $false
					
					for ($i = 0; $i -lt $lines.Length; $i++)
					{
						if ($i -ge ($secStart + 1) -and $i -lt $secEnd)
						{
							if ($i -eq $insertAt -and -not $inserted)
							{
								$out += $newProcessorLine
								$inserted = $true
							}
							
							if ($lines[$i] -match '^\s*\d{6,7}\s*=')
							{
								continue
							}
							
							$out += $lines[$i]
							continue
						}
						
						$out += $lines[$i]
					}
					
					if (-not $inserted)
					{
						$out2 = @()
						for ($i = 0; $i -lt $out.Length; $i++)
						{
							$out2 += $out[$i]
							if ($i -eq $secStart)
							{
								$out2 += $newProcessorLine
								$inserted = $true
							}
						}
						$out = $out2
					}
					
					if (($out -join $nl) -ne ($lines -join $nl))
					{
						$lines = $out
						$processorsChanged = $true
					}
				}
				else
				{
					# Section missing: create only if requested
					if ($CreateSmsHttpsIfMissing)
					{
						$add = @()
						$add += ""
						$add += "[PROCESSORS]"
						$add += ($newKey + "=REDIRMAIL=" + $redirKey901 + ",REDIRMSG=" + $redirKey901 + ",TARGETSEND=,TARGETRECV=,DEADLOCKPRIORITY=")
						$lines = @($lines + $add)
						$processorsChanged = $true
					}
				}
			}
			
			# ---- Clear LicenseGUID in [GENERAL] (always attempt)
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
					if ($lines[$i] -match '^(?<ws>\s*)(?i:LicenseGUID)\s*=\s*(?<val>.*)\s*$')
					{
						$ws = $matches['ws']
						$val = $matches['val']
						if ($val -ne "")
						{
							$lines[$i] = $ws + "LicenseGUID="
							$licenseGuidCleared = $true
						}
						break
					}
				}
			}
			
			if ($processorsChanged -or $licenseGuidCleared)
			{
				$outText = ($lines -join $nl)
				[System.IO.File]::WriteAllText($SmsHttpsIniPath, $outText, $enc)
				Write-Host "Updated SmsHttps.INI" -ForegroundColor Green
			}
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SmsHttps.INI: $($_.Exception.Message)" -ForegroundColor Red
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
		
		$userInput = $userInputRaw
		if ($null -eq $userInput) { $userInput = "" }
		$userInput = $userInput.Trim().ToUpper()
		
		if ([string]::IsNullOrEmpty($userInput))
		{
			[void][System.Windows.Forms.MessageBox]::Show("Machine name cannot be empty.", "Error")
			continue
		}
		
		# Regex: optional 1-4 digits + 2-8 letters + exactly 3 digits
		if ($userInput -match '^(\d{1,4})?([A-Z]{2,8})(\d{3})$')
		{
			$storePrefixRaw = $matches[1] # may be empty
			$namePrefix = $matches[2]
			$terminal = $matches[3]
			
			if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
			
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
			
			# Store width:
			# - If current store is 3/4 digits -> use that length
			# - Else if user typed 3/4 digits -> use that length
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
			
			$storePrefixNorm = $null
			if ($storePrefixRaw)
			{
				$val = [int]$storePrefixRaw
				$storePrefixNorm = $val.ToString(("D{0}" -f $storeWidth))
			}
			
			$currentStoreNorm = $null
			if ($currentStore -match '^\d{3,4}$')
			{
				$val2 = [int]$currentStore
				$currentStoreNorm = $val2.ToString(("D{0}" -f $storeWidth))
			}
			
			$normalizedName = $null
			if ($storePrefixNorm)
			{
				$normalizedName = "$storePrefixNorm$namePrefix$terminal"
			}
			else
			{
				$normalizedName = "$namePrefix$terminal"
			}
			
			# If user included store prefix, optionally sync INIs when mismatch/unknown
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
						$iniParams = @{
							newStoreNumber = $storePrefixNorm
							MachineName    = $normalizedName # <-- IMPORTANT: new lane digits (002)
							OldStoreNumber = $currentStoreNorm # best-effort (may be $null)
							OldMachineName = $env:COMPUTERNAME # old lane digits (004)
						}
						if ($currentStoreNorm) { $iniParams["OldStoreNumber"] = $currentStoreNorm }
						
						$success = Update_INIs @iniParams
						
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
							[void][System.Windows.Forms.MessageBox]::Show("Failed to update store number in INI.", "Error")
							continue
						}
					}
				}
			}
			
			if ($normalizedName -eq $env:COMPUTERNAME.ToUpper())
			{
				[void][System.Windows.Forms.MessageBox]::Show(
					"New name is the same as current computer name ($env:COMPUTERNAME).`nChoose a different name.",
					"Invalid Name"
				)
				continue
			}
			
			return $normalizedName
		}
		
		[void][System.Windows.Forms.MessageBox]::Show(
			"Invalid format.`n`nValid examples:`n  LANE003`n  SCO012`n  0231LANE006`n  1234POS999`n`nRule: [optional 1-4 digits] + [2-8 letters] + [3 digits]",
			"Invalid Machine Name",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Error
		)
	}
}

# ===================================================================================================
#                       FUNCTION: Update_SQL_Tables_For_Machine_Name_Change
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
	$stdTableName = "STD_TAB"
	
	# Prepare SQL commands
	
	# ---------------------------
	# TER_TAB commands
	# ---------------------------
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
	
	# ---------------------------
	# RUN_TAB commands
	# ---------------------------
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
	
	# ---------------------------
	# STO_TAB commands
	# ---------------------------
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
	
	# ---------------------------
	# LNK_TAB commands
	# ---------------------------
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
	
	# ---------------------------
	# STD_TAB commands
	# ---------------------------
	$createViewCommandStd = @"
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
"@
	
	$updateStdTabCommand = @"
UPDATE $stdTableName
SET F1056 = '$storeNumber'
WHERE F1056 NOT IN ('999','901');
"@
	
	$dropViewCommandStd = "DROP VIEW Std_Load;"
	
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
		
		# STD_TAB
		$createViewCommandStd,
		$updateStdTabCommand,
		$dropViewCommandStd
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
		
		# ---------------------------
		# RUN_TAB commands
		# ---------------------------
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
		
		# ---------------------------
		# STO_TAB commands
		# ---------------------------
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
		
		# ---------------------------
		# LNK_TAB commands
		# ---------------------------
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
		
		# ---------------------------
		# STD_TAB commands
		# ---------------------------
		$createViewCommandStd = @"
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
"@
		
		$updateStdTabCommand = @"
UPDATE $stdTableName
SET F1056 = '$storeNumber'
WHERE F1056 NOT IN ('999','901');
"@
		
		$dropViewCommandStd = "DROP VIEW Std_Load;"
		
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
# 1) Change Machine Name Button (UPDATED: always runs Update_INIs in finally)
############################################################################
$changeMachineNameButton = New-Object System.Windows.Forms.Button
$changeMachineNameButton.Text = "Change Machine Name"
$changeMachineNameButton.Location = New-Object System.Drawing.Point(10, 120)
$changeMachineNameButton.Size = New-Object System.Drawing.Size(150, 35)
$changeMachineNameButton.Add_Click({
		
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
		
		$oldMachineName = $env:COMPUTERNAME
		$newMachineNameInput = Get_NEW_Machine_Name
		if ($newMachineNameInput -eq $null)
		{
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
		
		$result = [System.Windows.Forms.MessageBox]::Show(
			"Are you sure you want to change the machine name to '$newMachineNameInput'?",
			"Confirm Machine Name Change",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Question
		)
		if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
		
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
		
		$normalizedHost = [string]$newMachineNameInput
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
		
		$startupIniPath = "\\localhost\storeman\startup.ini"
		
		try
		{
			Rename-Computer -NewName $normalizedHost -Force -ErrorAction Stop
			$script:newMachineName = $normalizedHost
			
			if ($machineNameLabel)
			{
				$machineNameLabel.Text = "The machine name will change from: $env:COMPUTERNAME to $script:newMachineName"
				$machineNameLabel.Refresh()
				if ($machineNameLabel.Parent) { $machineNameLabel.Parent.PerformLayout(); $machineNameLabel.Parent.Refresh() }
				[System.Windows.Forms.Application]::DoEvents()
			}
			
			if ($operationStatus -and $operationStatus.ContainsKey("MachineNameChange"))
			{
				$operationStatus["MachineNameChange"].Status = "Successful"
				$operationStatus["MachineNameChange"].Message = "Machine name changed successfully."
				$operationStatus["MachineNameChange"].Details = "Machine name changed to '$script:newMachineName'."
			}
			
			& 'Remove_Old_XZ_Folders' -MachineName $script:newMachineName -StoreNumber $currentStoreNumber
			
			# Update TER + DBSERVER in startup.ini
			$terValue = "TER=$machineNumber"
			$dbServerValue = $null
			
			if (Test-Path $startupIniPath)
			{
				$content = Get-Content $startupIniPath
				
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
		finally
		{
			# Always run Update_INIs at the end so INIs reflect the latest intended state
			$effectiveMachineName = $env:COMPUTERNAME
			if ($script:newMachineName -and -not [string]::IsNullOrWhiteSpace([string]$script:newMachineName))
			{
				$effectiveMachineName = [string]$script:newMachineName
			}
			
			try
			{
				$null = Update_INIs -newStoreNumber $currentStoreNumber -MachineName $effectiveMachineName -OldStoreNumber $currentStoreNumber -OldMachineName $oldMachineName -CreateSmsHttpsIfMissing -StartupIniPath $startupIniPath
			}
			catch { }
		}
		
		# Reboot prompt only if rename succeeded
		if ($script:newMachineName -and -not [string]::IsNullOrWhiteSpace([string]$script:newMachineName))
		{
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
# 3) Update Store Number Button (UPDATED: always runs Update_INIs using latest intended MachineName)
############################################################################
$updateStoreNumberButton = New-Object System.Windows.Forms.Button
$updateStoreNumberButton.Text = "Update Store Number"
$updateStoreNumberButton.Location = New-Object System.Drawing.Point(330, 120)
$updateStoreNumberButton.Size = New-Object System.Drawing.Size(150, 35)
$updateStoreNumberButton.Add_Click({
		
		# Get old store number
		$oldStoreNumber = Get_Store_Number_From_INI
		
		if ($oldStoreNumber -eq $null)
		{
			[System.Windows.Forms.MessageBox]::Show("Store number not found in startup.ini.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			$operationStatus["StoreNumberChange"].Status = "Failed"
			$operationStatus["StoreNumberChange"].Message = "Store number not found."
			$operationStatus["StoreNumberChange"].Details = "startup.ini not found or store number not defined."
			return
		}
		
		$newStoreNumberInput = Get_NEW_Store_Number
		if ($newStoreNumberInput -eq $null) { return }
		
		$warningResult = [System.Windows.Forms.MessageBox]::Show(
			"You are about to change the store number from '$oldStoreNumber' to '$newStoreNumberInput'. Do you want to proceed?",
			"Warning",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Warning
		)
		
		if ($warningResult -ne [System.Windows.Forms.DialogResult]::Yes)
		{
			$operationStatus["StoreNumberChange"].Status = "Cancelled"
			$operationStatus["StoreNumberChange"].Message = "Store number change was cancelled by the user."
			$operationStatus["StoreNumberChange"].Details = "Old store number remains '$oldStoreNumber'."
			return
		}
		
		# Prefer the "intended" machine name if you set it elsewhere (rename pending), else current computername
		$effectiveMachineName = $env:COMPUTERNAME
		if ($script:newMachineName -and -not [string]::IsNullOrWhiteSpace([string]$script:newMachineName))
		{
			$effectiveMachineName = [string]$script:newMachineName
		}
		
		# Always run Update_INIs (best-effort) so INIs are consistent even if other steps later fail
		$iniOk = $false
		try
		{
			$iniOk = Update_INIs -newStoreNumber $newStoreNumberInput -MachineName $effectiveMachineName -OldStoreNumber $oldStoreNumber -OldMachineName $env:COMPUTERNAME -CreateSmsHttpsIfMissing
		}
		catch
		{
			$iniOk = $false
		}
		
		if ($iniOk)
		{
			$script:newStoreNumber = $newStoreNumberInput
			$storeNumberLabel.Text = "Store Number changed from: $oldStoreNumber to $script:newStoreNumber"
			
			$operationStatus["StoreNumberChange"].Status = "Successful"
			$operationStatus["StoreNumberChange"].Message = "Store number updated in INIs."
			$operationStatus["StoreNumberChange"].Details = "Store number changed to '$script:newStoreNumber'."
			
			[System.Windows.Forms.MessageBox]::Show("Store number successfully changed to '$script:newStoreNumber'.", "Store Number Updated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		}
		else
		{
			$operationStatus["StoreNumberChange"].Status = "Failed"
			$operationStatus["StoreNumberChange"].Message = "Update_INIs failed."
			$operationStatus["StoreNumberChange"].Details = "Best-effort update encountered an error."
			[System.Windows.Forms.MessageBox]::Show("Update_INIs failed (best-effort). Check console/log output.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
		
		# SQL update (keep your existing logic)
		$sqlUpdateResult = Update_SQL_Tables_For_Store_Number_Change -storeNumber $newStoreNumberInput
		
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
		
		<# Run Update_INIs again at the end so INIs always reflect the latest state
		try
		{
			$null = Update_INIs -newStoreNumber $newStoreNumberInput -MachineName $effectiveMachineName -OldStoreNumber $oldStoreNumber -OldMachineName $env:COMPUTERNAME -CreateSmsHttpsIfMissing
		}
		catch { }#>
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
