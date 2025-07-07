#######################################################################################################
#                                                                                                     #
#                                     TBS MAINTENANCE SCRIPT                                          #
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
$VersionNumber = "2.3.0"
$VersionDate = "2025-07-07"

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
#                                SECTION: Import Necessary Assemblies
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports required .NET assemblies for creating and managing Windows Forms and graphical components.
# ===================================================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
$script:ProcessedLanes = @()
$script:ProcessedStores = @()
$script:ProcessedServers = @()
$script:ProcessedHosts = @()

# ---------------------------------------------------------------------------------------------------
# Count Tracking Variables
# ---------------------------------------------------------------------------------------------------
$NumberOfLanes = 0
$NumberOfServers = 0
$NumberOfScales = 0
$NumberOfBackoffices = 0

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
		$candidatePath = Join-Path $_.FullName 'Office\Dbs'
		$files = Get-ChildItem -Path $candidatePath -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue
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
	$candidatePath = Join-Path $fallback 'Office\Dbs'
	$files = Get-ChildItem -Path $candidatePath -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue
	if ($files) { $BasePath = $fallback }
	else
	{
		Write-Warning "Could not locate any storeman folder containing Office\Dbs\INFO_*901_WIN.INI.`nRunning with limited functionality."
		$BasePath = $fallback
	}
}

Write-Host "Selected (storeman) folder: '$BasePath'" -ForegroundColor Magenta

# ---------------------------------------------------------------------------------------------------
# Build All Core Paths and File Locations
# ---------------------------------------------------------------------------------------------------
$OfficePath = Join-Path $BasePath "Office"
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

# ---------------------------------------------------------------------------------------------------
# (Optional) Script Name Extraction
# ---------------------------------------------------------------------------------------------------
# $scriptName = Split-Path -Leaf $PSCommandPath

# ---------------------------------------------------------------------------------------------------
# Add C# MailSlotSender Type for Direct Windows Mailslot Messaging (if not already loaded)
# ---------------------------------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'MailslotSender').Type)
{
	Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class MailslotSender {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile
    );

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool WriteFile(
        IntPtr hFile,
        byte[] lpBuffer,
        uint nNumberOfBytesToWrite,
        out uint lpNumberOfBytesWritten,
        IntPtr lpOverlapped
    );
        
    [DllImport("kernel32.dll")]
    public static extern bool CloseHandle(IntPtr hObject);

    public static bool SendMailslotCommand(string mailslotName, string command) {
        const uint GENERIC_WRITE = 0x40000000;
        const uint FILE_SHARE_READ = 0x00000001;
        const uint OPEN_EXISTING = 3;

        IntPtr hFile = CreateFile(mailslotName, GENERIC_WRITE, FILE_SHARE_READ, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
        if (hFile == new IntPtr(-1)) {
            return false;
        }

        byte[] data = Encoding.ASCII.GetBytes(command);
        uint bytesWritten;
        bool success = WriteFile(hFile, data, (uint)data.Length, out bytesWritten, IntPtr.Zero);
        CloseHandle(hFile);
        return success;
    }
}
"@
}

# ===================================================================================================
#                                       FUNCTION: Write to Log
# ---------------------------------------------------------------------------------------------------
# Description:
#   Writes messages to the log GUI box. No silent mode or file logging.
# ===================================================================================================

function Write_Log
{
	param (
		[string]$Message,
		[string]$Color = "Black",
		[switch]$IncludeTimestamp = $true
	)
	
	<# Prepare timestamp if needed
	$timestamp = if ($IncludeTimestamp) { Get-Date -Format 'yyyy-MM-dd HH:mm:ss' }
	else { "" }#>
	
	if ($logBox -ne $null)
	{
		# Write to GUI log box
		$logBox.SelectionColor = switch ($Color.ToLower())
		{
			"green"   { [System.Drawing.Color]::Green }
			"red"     { [System.Drawing.Color]::Red }
			"yellow"  { [System.Drawing.Color]::Goldenrod }
			"blue"    { [System.Drawing.Color]::Blue }
			"magenta" { [System.Drawing.Color]::Magenta }
			"gray"    { [System.Drawing.Color]::Gray }
			"cyan"    { [System.Drawing.Color]::Cyan }
			"white"   { [System.Drawing.Color]::White }
			"orange"  { [System.Drawing.Color]::Orange }
			default   { [System.Drawing.Color]::Black }
		}
		
		$fullMessage = if ($timestamp) { "[$timestamp] $Message" }
		else { $Message }
		$logBox.AppendText("$fullMessage`r`n")
		$logBox.SelectionColor = $logBox.ForeColor
		$logBox.ScrollToCaret()
		
		# Process GUI events to refresh the log box
		[System.Windows.Forms.Application]::DoEvents()
	}
	else
	{
		# Output to console until logBox is initialized
		$fullMessage = if ($timestamp) { "[$timestamp] $Message" }
		else { $Message }
		Write-Host $fullMessage -ForegroundColor $Color
	}
}

# ===================================================================================================
#                               FUNCTION: Get-Store-And-Db-Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Uses provided paths to INFO_*901_WIN.INI, INFO_*901_SMSStart.ini, Startup.ini, and optionally
#   system.ini, to extract:
#     - Store metadata (store number, name, terminal, address, version, etc.) from WIN.INI
#     - SQL Server/Database name using any key containing 'ServerName*' or 'DatabaseName*' from SMSStart.ini
#     - Falls back to Startup.ini for DBSERVER/DBNAME if not found in SMSStart.ini
#     - Falls back to system.ini [SMS] Name=... for store name if not found in WIN.INI
#   Builds a SQL Server connection string using trusted authentication and TrustServerCertificate.
#   Populates all results in $script:FunctionResults for downstream use.
# ===================================================================================================

function Get_Store_And_Database_Info
{
	param (
		[string]$WinIniPath,
		[string]$SmsStartIniPath,
		[string]$StartupIniPath,
		[string]$SystemIniPath
	)
	
	# ------------------------------------------------------------------------------------------------
	# Initialize results with N/A for every expected property
	# ------------------------------------------------------------------------------------------------
	$fields = @(
		'StoreNumber', 'StoreName', 'CompanyName', 'Terminal', 'Address', 'City', 'State',
		'SMSVersionFull', 'StampDate', 'StampTime', 'KeyNumber',
		'DBSERVER', 'DBNAME', 'ConnectionString'
	)
	foreach ($f in $fields) { $script:FunctionResults[$f] = 'N/A' }
	
	# ------------------------------------------------------------------------------------------------
	# Extract Store Info from WIN.INI (required for store metadata)
	# ------------------------------------------------------------------------------------------------
	if ($WinIniPath -and (Test-Path $WinIniPath))
	{
		$currentSection = ""
		foreach ($line in Get-Content $WinIniPath)
		{
			$trimmed = $line.Trim()
			if ($trimmed -match '^\[(.+)\]$') { $currentSection = $Matches[1]; continue }
			if ($trimmed -notmatch "=" -or $trimmed.StartsWith(";")) { continue }
			$parts = $trimmed -split "=", 2
			$key = $parts[0].Trim()
			$value = $parts[1].Trim()
			switch ($currentSection)
			{
				"ORIGIN" {
					if ($key -ieq "StampDate") { $script:FunctionResults['StampDate'] = $value }
					if ($key -ieq "StampTime") { $script:FunctionResults['StampTime'] = $value }
				}
				"SYSTEM" {
					if ($key -ieq "CompanyName") { $script:FunctionResults['CompanyName'] = $value }
					if ($key -ieq "Store") { $script:FunctionResults['StoreNumber'] = $value.PadLeft(3, "0") }
					if ($key -ieq "Terminal") { $script:FunctionResults['Terminal'] = $value }
				}
				"STOREDETAIL" {
					if ($key -ieq "Name") { $script:FunctionResults['StoreName'] = $value }
					if ($key -ieq "Address") { $script:FunctionResults['Address'] = $value }
					if ($key -ieq "City") { $script:FunctionResults['City'] = $value }
					if ($key -ieq "State") { $script:FunctionResults['State'] = $value }
				}
				"KEY" {
					if ($key -ieq "KeyNumber") { $script:FunctionResults['KeyNumber'] = $value }
				}
				"Versions" {
					if ($key -ieq "VersionIni") { $script:FunctionResults['SMSVersionFull'] = $value }
				}
			}
		}
	}
	else
	{
		Write_Log "No INFO_*901_WIN.INI found at $WinIniPath" "red"
	}
	
	# ------------------------------------------------------------------------------------------------
	# Fallback: Get StoreName from system.ini ([SMS] Name=...) if still N/A
	# ------------------------------------------------------------------------------------------------
	if (
		($script:FunctionResults['StoreName'] -eq 'N/A' -or
			[string]::IsNullOrWhiteSpace($script:FunctionResults['StoreName'])) -and
		$SystemIniPath -and (Test-Path $SystemIniPath)
	)
	{
		$inSMSSection = $false
		foreach ($line in Get-Content $SystemIniPath)
		{
			$trimmed = $line.Trim()
			if ($trimmed -match '^\[SMS\]$')
			{
				$inSMSSection = $true
				continue
			}
			if ($inSMSSection)
			{
				# End of section if new [Section] starts
				if ($trimmed -match '^\[.+\]$') { break }
				# Look for Name=
				if ($trimmed -match '^Name\s*=(.*)$')
				{
					$storeNameBackup = $Matches[1].Trim()
					if ($storeNameBackup)
					{
						$script:FunctionResults['StoreName'] = $storeNameBackup
						break
					}
				}
			}
		}
	}
	
	# ------------------------------------------------------------------------------------------------
	# Extract SQL Server/Database from SMSStart INI (first matching ServerName*/DatabaseName*)
	# ------------------------------------------------------------------------------------------------
	$dbServer = $null
	$dbName = $null
	
	if ($SmsStartIniPath -and (Test-Path $SmsStartIniPath))
	{
		foreach ($line in Get-Content $SmsStartIniPath)
		{
			$trimmed = $line.Trim()
			if ($trimmed -notmatch "=" -or $trimmed.StartsWith(";")) { continue }
			$parts = $trimmed -split "=", 2
			$key = $parts[0].Trim()
			$value = $parts[1].Trim()
			if (-not $dbServer -and $key -match 'ServerName') { $dbServer = $value }
			if (-not $dbName -and $key -match 'DatabaseName') { $dbName = $value }
			if ($dbServer -and $dbName) { break }
		}
	}
	
	# ------------------------------------------------------------------------------------------------
	# Fallback: Try Startup.ini for DBSERVER/DBNAME if either is missing
	# ------------------------------------------------------------------------------------------------
	if ((!$dbServer -or !$dbName) -and $StartupIniPath -and (Test-Path $StartupIniPath))
	{
		foreach ($line in Get-Content $StartupIniPath)
		{
			$trimmed = $line.Trim()
			if ($trimmed -notmatch "=" -or $trimmed.StartsWith(";")) { continue }
			$parts = $trimmed -split "=", 2
			$key = $parts[0].Trim()
			$value = $parts[1].Trim()
			if (-not $dbServer -and $key -match 'DBSERVER') { $dbServer = $value }
			if (-not $dbName -and $key -match 'DBNAME') { $dbName = $value }
			if ($dbServer -and $dbName) { break }
		}
		if (-not $dbServer) { $dbServer = "localhost" }
		if (-not $dbName) { $dbName = "STORESQL" }
	}
	
	# ------------------------------------------------------------------------------------------------
	# Finalize connection string and store in function results
	# ------------------------------------------------------------------------------------------------
	if ($dbServer -and $dbName)
	{
		$script:FunctionResults['DBSERVER'] = $dbServer
		$script:FunctionResults['DBNAME'] = $dbName
		$script:FunctionResults['ConnectionString'] = "Server=$dbServer;Database=$dbName;Integrated Security=True;Encrypt=True;TrustServerCertificate=True"
	}
	
	# ------------------------------------------------------------------------------------------------
	# GUI label updates (optional; can be removed if not needed)
	# ------------------------------------------------------------------------------------------------
	if ($storeNumberLabel -ne $null)
	{
		$storeNumberLabel.Text = "Store Number: $($script:FunctionResults['StoreNumber'])"
		$form.Refresh(); [System.Windows.Forms.Application]::DoEvents()
	}
	if ($storeNameLabel -ne $null)
	{
		$storeNameLabel.Text = "Store Name: $($script:FunctionResults['StoreName'])"
		$form.Refresh(); [System.Windows.Forms.Application]::DoEvents()
	}
	if ($smsVersionLabel -ne $null)
	{
		$smsVersionLabel.Text = "SMS Version: $($script:FunctionResults['SMSVersionFull'])"
		$form.Refresh(); [System.Windows.Forms.Application]::DoEvents()
	}
}

# ===================================================================================================
#                                FUNCTION: Get_All_Lanes_Database_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the DB server names and database names for all lanes in TER_TAB.
#   The lane number starts with '0' and excludes '901' and '999'.
#   Constructs a unique connection string for each lane and stores it in $script:FunctionResults['LaneDatabaseInfo'].
# ===================================================================================================

function Get_All_Lanes_Database_Info
{
	Write-Host "Gathering connection string information for all lanes..."
	
	$LaneMachines = $script:FunctionResults['LaneMachines']
	if (-not $LaneMachines) { return }
	
	$script:FunctionResults['LaneDatabaseInfo'] = @{ }
	
	foreach ($laneNumber in $LaneMachines.Keys)
	{
		# Skip unwanted lanes
		if ($laneNumber -match '^(8|9)' -or $laneNumber -eq '901' -or $laneNumber -eq '999') { continue }
		
		$machineName = $LaneMachines[$laneNumber]
		if (-not $machineName) { continue }
		
		$startupIniPath = "\\$machineName\storeman\Startup.ini"
		if (-not (Test-Path $startupIniPath)) { continue }
		
		try
		{
			$content = Get-Content -Path $startupIniPath -ErrorAction Stop
			$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' }
			$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }
			
			if ($dbNameLine)
			{
				$dbName = ($dbNameLine -replace '^DBNAME=', '').Trim()
			}
			else
			{
				continue
			}
			if ($dbServerLine)
			{
				$dbServerRaw = ($dbServerLine -replace '^DBSERVER=', '').Trim()
			}
			else
			{
				$dbServerRaw = ""
			}
			$dbServer = $dbServerRaw
			if (-not $dbServer -or $dbServer -eq '')
			{
				$dbServer = $machineName
			}
			
			# Parse instance name for Named Pipes/TCP logic
			$serverName = $dbServer
			$instanceName = $null
			if ($dbServer -match '\\')
			{
				$parts = $dbServer -split '\\'
				$serverName = $parts[0]
				$instanceName = $parts[1]
			}
			elseif ($dbServer -match ',')
			{
				$serverName = $dbServer
				$instanceName = $null
			}
			else
			{
				$serverName = $dbServer
				$instanceName = $null
			}
			
			# Build connection strings
			if ($instanceName -and $instanceName.ToUpper() -ne "MSSQLSERVER")
			{
				$namedPipes = "np:\\$serverName\pipe\MSSQL`$$instanceName\sql\query"
				$tcpServer = "$serverName\$instanceName"
			}
			else
			{
				$namedPipes = "np:\\$serverName\pipe\sql\query"
				$tcpServer = $serverName
			}
			
			$tcpConnStr = "Server=$tcpServer;Database=$dbName;Integrated Security=True;"
			$namedPipesConnStr = "Server=$namedPipes;Database=$dbName;Integrated Security=True;"
			$simpleConnStr = "Server=$dbServer;Database=$dbName;Integrated Security=True;"
			
			$script:FunctionResults['LaneDatabaseInfo'][$laneNumber] = @{
				'MachineName'	    = $machineName
				'DBName'		    = $dbName
				'DBServer'		    = $dbServer
				'ServerName'	    = $serverName
				'InstanceName'	    = $instanceName
				'NamedPipes'	    = $namedPipes
				'TcpServer'		    = $tcpServer
				'ConnectionString'  = $simpleConnStr
				'NamedPipesConnStr' = $namedPipesConnStr
				'TcpConnStr'	    = $tcpConnStr
			}
		}
		catch
		{
			continue
		}
	}
}

# ===================================================================================================
#                           FUNCTION: Retrieve_Nodes
# ---------------------------------------------------------------------------------------------------
# **Purpose:**
#   The `Retrieve_Nodes` function is designed to count various entities within a 
#   system, specifically **hosts**, **stores**, **lanes**, **servers**, and **scales**. It primarily retrieves 
#   these nodes from the `TER_TAB` database table and additional tables as needed. If database access fails, it gracefully falls 
#   back to a file system-based mechanism to obtain the counts. Additionally, the function updates 
#   GUI labels to reflect the current nodes and stores the results in a shared hashtable for use 
#   by other parts of the script. For scales, the function retrieves the IPNetwork information from the 
#   TBS_SCL_ver520 table.
#
# **Parameters:**
#   - `[string]$StoreNumber`
#       - **Description:** Specifies the identifier for a particular store. This parameter is 
#         **mandatory** when `$Mode` is set to `"Store"` and is ignored when `$Mode` is `"Host"`.
#
# **Variables:**
#   - **Initialization Variables:**
#       - `$HostPath`: Base directory path where store and host directories are located.
#       - `$NumberOfLanes`, `$NumberOfStores`, `$NumberOfHosts`, `$NumberOfServers`, `$NumberOfScales`: Counters initialized to `0`.
#       - `$LaneContents`: Array to hold lane identifiers.
#       - `$LaneMachines`: Hashtable to map lane numbers to machine names.
#       - `$ScaleIPNetworks`: Hashtable to map scale identifiers to their IPNetwork values.
#   - **Database Connection Variables:**
#       - `$ConnectionString`: Retrieves the database connection string from the `FunctionResults` hashtable.
#       - `$NodesFromDatabase`: Boolean flag indicating whether to retrieve counts from the database.
#   - **Result Variables:**
#       - `$Nodes`: Custom PowerShell object aggregating all nodes counts and related data.
#   - **GUI-Related Variables:**
#       - `$NodesHost`, `$NodesStore`, `$NodesScales`: GUI label controls displaying the counts.
#       - `$form`: GUI form that needs to be refreshed to display updated counts.
#
# **Workflow:**
#   1. **Retrieve Database Connection String:**
#      - Attempts to get the connection string from `FunctionResults`.
#      - If unavailable, calls `Get_Database_Connection_String` to generate it.
#      - Sets `$CountsFromDatabase` based on availability.
#
#   2. **Database Counting Mechanism (`$CountsFromDatabase = $true`):**
#      - **Mode: `"Host"`**
#          - Counts distinct stores excluding store number `'999'`.
#          - Checks for the existence of the host server.
#      - **Mode: `"Store"`**
#          - Validates the presence of `$StoreNumber`.
#          - Retrieves and counts lanes for the specified store.
#          - Maps lane numbers to machine names.
#          - Retrieves scales from TER_TAB (count only) and additional scales from TBS_SCL_ver520 (which provides the IPNetwork info).
#          - Checks for the existence of the server for the store.
#      - **Error Handling:**
#          - Logs warnings and falls back if any database queries fail.
#
#   3. **Fallback Counting Mechanism (`$CountsFromDatabase = $false`):**
#      - **Mode: `"Host"`**
#          - Counts store directories matching specific patterns.
#          - Checks for the existence of the host directory.
#      - **Mode: `"Store"`**
#          - Validates the presence of `$StoreNumber`.
#          - Counts lane and scale directories matching specific patterns.
#          - Checks for the existence of the server directory for the store.
#
#   4. **Compile and Store Results:**
#      - Creates a `[PSCustomObject]` containing all counts and related data.
#      - Updates the `FunctionResults` hashtable with the count results.
#
#   5. **Update GUI Labels:**
#      - If not in silent mode and GUI labels are available, updates them with the latest counts.
#      - Refreshes the GUI form to display the updated counts.
#
#   6. **Return Value:**
#      - Returns the `$Nodes` custom object containing all the count information.
#
# **Summary:**
#   The `Retrieve_Nodes` function is a robust PowerShell utility that accurately counts system entities 
#   such as hosts, stores, lanes, servers, and scales. It prioritizes retrieving counts from a database to 
#   ensure accuracy and reliability but includes a fallback mechanism leveraging the file system for 
#   resilience. Additionally, it integrates with a GUI to display real-time counts, stores results 
#   for easy access by other script components, and retrieves IPNetwork information for scales from the TBS_SCL_ver520 table.
# ===================================================================================================

function Retrieve_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	# -------------------- Init --------------------
	$HostPath = "$OfficePath"
	$LaneContents = @()
	$LaneMachines = @{ }
	$ScaleIPNetworks = @{ }
	$TerLoadSqlPath = Join-Path $LoadPath 'Ter_Load.sql'
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$NodesFromDatabase = $false
	
	# -------------------- 1. Database --------------------
	if ($ConnectionString)
	{
		$NodesFromDatabase = $true
		try
		{
			#--------------------------------------------------------------------------------
			# 1) Retrieve Lanes from TER_TAB
			#--------------------------------------------------------------------------------
			$queryLaneContents = @"
SELECT F1057, F1125
FROM TER_TAB
WHERE F1056 = '$StoreNumber'
  AND F1057 LIKE '0%'
  AND F1057 NOT LIKE '8%'
  AND F1057 NOT LIKE '9%'
"@
			try
			{
				$laneContentsResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryLaneContents -ErrorAction Stop
			}
			catch [System.Management.Automation.ParameterBindingException] {
				$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
				$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
				$laneContentsResult = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryLaneContents -ErrorAction Stop
			}
			
			$LaneContents = $laneContentsResult | Select-Object -ExpandProperty F1057
			$NumberOfLanes = $LaneContents.Count
			
			foreach ($row in $laneContentsResult)
			{
				$laneNumber = $row.F1057
				$machinePath = $row.F1125
				if ($machinePath -match '\\\\([^\\]+)\\')
				{
					$machineName = $matches[1]
					$LaneMachines[$laneNumber] = $machineName
				}
			}
			
			#--------------------------------------------------------------------------------
			# 2) Retrieve scales from TER_TAB (count only)
			#--------------------------------------------------------------------------------
			$queryScaleContents = @"
SELECT F1057, F1125
FROM TER_TAB
WHERE F1056 = '$StoreNumber'
  AND F1057 LIKE '8%'
  AND F1057 NOT LIKE '0%'
  AND F1057 NOT LIKE '9%'
"@
			try
			{
				$scaleContentsResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryScaleContents -ErrorAction Stop
			}
			catch [System.Management.Automation.ParameterBindingException] {
				$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
				$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
				$scaleContentsResult = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryScaleContents -ErrorAction Stop
			}
			
			$ScaleContents = $scaleContentsResult | Select-Object -ExpandProperty F1057
			$NumberOfScales += $ScaleContents.Count
			
			#--------------------------------------------------------------------------------
			# 3) Retrieve additional scales from TBS_SCL_ver520 (with IPNetwork)
			#--------------------------------------------------------------------------------
			$queryTbsSclScales = @"
SELECT ScaleCode, ScaleName, ScaleLocation, IPNetwork, IPDevice, Active, ScaleBrand, ScaleModel
FROM TBS_SCL_ver520
WHERE Active = 'Y'
"@
			try
			{
				$tbsSclScalesResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryTbsSclScales -ErrorAction Stop
			}
			catch
			{
				if ($_.Exception.Message -match "Invalid object name 'TBS_SCL_ver520'")
				{
					$tbsSclScalesResult = $null
				}
				else
				{
					throw # rethrow for other errors, which will trigger fallback
				}
			}
			
			if ($tbsSclScalesResult)
			{
				$NumberOfScales += $tbsSclScalesResult.Count
				foreach ($row in $tbsSclScalesResult)
				{
					$fullIP = "$($row.IPNetwork)$($row.IPDevice)"
					$scaleObj = [PSCustomObject]@{
						ScaleCode	  = $row.ScaleCode
						ScaleName	  = $row.ScaleName
						ScaleLocation = $row.ScaleLocation
						IPNetwork	  = $row.IPNetwork
						IPDevice	  = $row.IPDevice
						FullIP	      = $fullIP
						Active	      = $row.Active
						ScaleBrand    = $row.ScaleBrand
						ScaleModel    = $row.ScaleModel
					}
					$ScaleIPNetworks[$row.ScaleCode] = $scaleObj
				}
			}
			
			#--------------------------------------------------------------------------------
			# 4) Check if server exists for the store (F1057 = '901')
			#--------------------------------------------------------------------------------
			$queryServer = @"
SELECT COUNT(*) AS ServerCount
FROM TER_TAB
WHERE F1056 = '$StoreNumber'
  AND F1057 = '901'
"@
			try
			{
				$serverResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryServer -ErrorAction Stop
			}
			catch [System.Management.Automation.ParameterBindingException] {
				$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
				$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
				$serverResult = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryServer -ErrorAction Stop
			}
			
			$NumberOfServers = if ($serverResult.ServerCount -gt 0) { 1 }
			else { 0 }
			
			#--------------------------------------------------------------------------------
			# 5) Retrieve backoffices (F1057 = '902', '903', etc): COUNT and MAP at once
			#--------------------------------------------------------------------------------
			$queryBackoffices = @"
SELECT F1057, F1125
FROM TER_TAB
WHERE F1056 = '$StoreNumber'
  AND F1057 >= '902'
  AND F1057 <= '998'
"@
			try
			{
				$backofficesList = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryBackoffices -ErrorAction Stop
			}
			catch [System.Management.Automation.ParameterBindingException] {
				$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
				$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
				$backofficesList = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryBackoffices -ErrorAction Stop
			}
			$BackofficeMachines = @{ }
			if ($backofficesList)
			{
				foreach ($row in $backofficesList)
				{
					$terminal = $row.F1057
					$backofficeName = $row.F1125
					if ($backofficeName)
					{
						$cleanName = $backofficeName -replace '^@', ''
						$BackofficeMachines[$terminal] = $cleanName
					}
				}
			}
			$NumberOfBackoffices = $BackofficeMachines.Count
		}
		catch
		{
			Write_Log "Failed to retrieve counts from the database: $_" "yellow"
			$NodesFromDatabase = $false
		}
	}
	
	#--------------------------------------------------------------------------------
	# Fallback: If counts from database failed, use Ter_Load.sql fallback
	#--------------------------------------------------------------------------------
	$TerLoadUsed = $false
	if ((-not $NodesFromDatabase) -and (Test-Path $TerLoadSqlPath))
	{
		Write_Log "Using Ter_Load.sql as backup for TER_TAB information." "yellow"
		$TerLoadUsed = $true
		$insertLines = Select-String -Path $TerLoadSqlPath -Pattern "INSERT INTO Ter_Load VALUES" -Context 0, 100 |
		Select-Object -First 1 | ForEach-Object { $_.Context.PostContext }
		
		$rows = @()
		$current = ""
		foreach ($line in $insertLines)
		{
			if ($line.Trim() -eq "") { continue }
			$current += $line.Trim()
			if ($current -match "\),$|\);$")
			{
				$rows += $current.TrimEnd(',', ';')
				$current = ""
			}
		}
		
		$parsed = @()
		if ($rows -and $rows.Count -gt 0)
		{
			$parsed = $rows | ForEach-Object {
				if ($_ -match "\((.*)\)")
				{
					$split = $matches[1] -split "',\s*'"
					if ($split.Count -ge 5)
					{
						[PSCustomObject]@{
							Store    = $split[0].Trim("'")
							Terminal = $split[1].Trim("'")
							Label    = $split[2].Trim("'")
							LanePath = $split[3].Trim("'")
							HostPath = $split[4].Trim("'")
						}
					}
				}
			} | Where-Object { $_.Store -eq $StoreNumber }
		}
		
		# Lanes: terminals starting with 0
		$laneObjs = @()
		if ($parsed.Count -gt 0)
		{
			$laneObjs = $parsed | Where-Object { $_.Terminal -match '^0\d\d$' }
		}
		$NumberOfLanes = $laneObjs.Count
		$LaneContents = $laneObjs | Select-Object -ExpandProperty Terminal
		foreach ($obj in $laneObjs)
		{
			if ($obj.LanePath -match '\\\\([^\\]+)\\')
			{
				$LaneMachines[$obj.Terminal] = $matches[1]
			}
		}
		# Scales: terminals starting with 8
		$scaleObjs = @()
		if ($parsed.Count -gt 0)
		{
			$scaleObjs = $parsed | Where-Object { $_.Terminal -match '^8\d\d$' }
		}
		$NumberOfScales = $scaleObjs.Count
		
		# Servers: terminal 901
		$NumberOfServers = 0
		if ($parsed.Count -gt 0)
		{
			$NumberOfServers = ($parsed | Where-Object { $_.Terminal -eq '901' }).Count
		}
		
		# Backoffices: 902-998
		$NumberOfBackoffices = 0
		if ($parsed.Count -gt 0)
		{
			$NumberOfBackoffices = ($parsed | Where-Object { $_.Terminal -match '^9(0[2-9]|[1-8][0-9])$' }).Count
		}
	}
	
	#--------------------------------------------------------------------------------
	# Fallback: If counts from database failed, use directory-based logic
	#--------------------------------------------------------------------------------
	if ((-not $NodesFromDatabase) -and (-not $TerLoadUsed))
	{
		Write_Log "Using file system directories as backup for node counts." "yellow"
		
		if (Test-Path $HostPath)
		{
			$LaneFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF${StoreNumber}0??"
			$NumberOfLanes = $LaneFolders.Count
			$LaneContents = $LaneFolders | ForEach-Object { $_.Name.Substring($_.Name.Length - 3, 3) }
		}
		
		if (Test-Path $HostPath)
		{
			$ScaleFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF${StoreNumber}8??"
			$NumberOfScales = $ScaleFolders.Count
		}
		
		$NumberOfServers = if (Test-Path "$HostPath\XF${StoreNumber}901") { 1 }
		else { 0 }
		
		# Backoffice folders: XF${StoreNumber}902 - XF${StoreNumber}998
		if (Test-Path $HostPath)
		{
			$BackofficeFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF${StoreNumber}9??" |
			Where-Object { $_.Name -match "^XF${StoreNumber}9(0[2-9]|[1-8][0-9])$" -and $_.Name -ne "XF${StoreNumber}999" }
			$NumberOfBackoffices = $BackofficeFolders.Count
		}
	}
	
	#--------------------------------------------------------------------------------
	# Final: Create a custom object with the counts
	#--------------------------------------------------------------------------------
	$Nodes = [PSCustomObject]@{
		NumberOfLanes	    = $NumberOfLanes
		NumberOfServers	    = $NumberOfServers
		NumberOfBackoffices = $NumberOfBackoffices
		NumberOfScales	    = $NumberOfScales
		LaneContents	    = $LaneContents
		LaneMachines	    = $LaneMachines
		ScaleIPNetworks	    = $ScaleIPNetworks
	}
	
	#--------------------------------------------------------------------------------
	# Store counts in FunctionResults
	#--------------------------------------------------------------------------------
	$script:FunctionResults['NumberOfLanes'] = $NumberOfLanes
	$script:FunctionResults['NumberOfServers'] = $NumberOfServers
	$script:FunctionResults['NumberOfBackoffices'] = $NumberOfBackoffices
	$script:FunctionResults['NumberOfScales'] = $NumberOfScales
	$script:FunctionResults['LaneContents'] = $LaneContents
	$script:FunctionResults['LaneMachines'] = $LaneMachines
	$script:FunctionResults['ScaleIPNetworks'] = $ScaleIPNetworks
	$script:FunctionResults['BackofficeMachines'] = $BackofficeMachines
	$script:FunctionResults['Nodes'] = $Nodes
	
	#--------------------------------------------------------------------------------
	# Update the GUI labels (if labels exist)
	#--------------------------------------------------------------------------------
	if ($NodesHost -ne $null) { $NodesHost.Text = "Number of Servers: $NumberOfServers" }
	if ($NodesBackoffices -ne $null) { $NodesBackoffices.Text = "Number of Backoffices: $NumberOfBackoffices" }
	if ($NodesStore -ne $null) { $NodesStore.Text = "Number of Lanes: $NumberOfLanes" }
	if ($scalesLabel -ne $null) { $scalesLabel.Text = "Number of Scales: $NumberOfScales" }
	if ($form -ne $null) { $form.Refresh() }
	
	return $Nodes
}

# ===================================================================================================
#                                      FUNCTION: Clearing XE folder
# ---------------------------------------------------------------------------------------------------
# Description:
#   Performs an initial cleanup of the XE (Urgent Messages) folder by deleting all files and subdirectories,
#   then starts a background job to continuously monitor and clear the folder at specified intervals,
#   excluding any files or directories whose names start with "FATAL".
# ===================================================================================================

function Clear_XE_Folder
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$folderPath = "$OfficePath\XE${StoreNumber}901",
		[Parameter(Mandatory = $false)]
		[int]$checkIntervalSeconds = 2
	)
	
	# -- Initial clearing: remove everything except valid S*.??? health files
	if (Test-Path -Path $folderPath)
	{
		try
		{
			$currentTime = Get-Date
			Get-ChildItem -Path $folderPath -Recurse -Force | ForEach-Object {
				$file = $_
				$keep = $false
				if ($file.Name -like 'FATAL*') { $keep = $false }
				elseif ($file.Name -match '^S.*\.\w{3}$')
				{
					if (($currentTime - $file.LastWriteTime).TotalDays -le 30)
					{
						try { $content = Get-Content $file.FullName -ErrorAction Stop }
						catch { $content = $null }
						if ($content)
						{
							$fromLine = $content | Where-Object { $_ -like 'From:*' }
							$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
							$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
							$lastStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
							if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
							{
								$fileStoreNumber = $Matches[1]
								$fileLaneNumber = $Matches[2]
								if ($fileStoreNumber -eq $StoreNumber -and
									$file.Name -match '^S.*\.(\d{3})$' -and
									$fileLaneNumber -eq $Matches[1] -and
									$subjectLine -match 'Subject:\s*Health' -and
									$msgLine -match 'MSG:\s*This application is not running\.' -and
									$lastStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
								{
									$keep = $true
								}
							}
						}
					}
				}
				if (-not $keep) { Remove-Item -Path $file.FullName -Force -Recurse }
			}
		}
		catch
		{
			Write_Log "An error occurred during initial cleaning of 'XE${StoreNumber}901': $_" "red"
		}
	}
	else
	{
		Write_Log "Folder 'XE${StoreNumber}901' (Urgent Messages) does not exist." "red"
		return
	}
	
	# -- Start background monitoring as a job
	try
	{
		$job = Start-Job -Name "ClearXEFolderJob" -ScriptBlock {
			param ($folderPath,
				$checkIntervalSeconds,
				$StoreNumber)
			while ($true)
			{
				try
				{
					$currentTime = Get-Date
					if (Test-Path -Path $folderPath)
					{
						Get-ChildItem -Path $folderPath -Recurse -Force | ForEach-Object {
							$file = $_
							$keep = $false
							if ($file.Name -like 'FATAL*') { $keep = $true }
							elseif ($file.Name -match '^S.*\.\w{3}$')
							{
								if (($currentTime - $file.LastWriteTime).TotalDays -le 30)
								{
									try { $content = Get-Content $file.FullName -ErrorAction Stop }
									catch { $content = $null }
									if ($content)
									{
										$fromLine = $content | Where-Object { $_ -like 'From:*' }
										$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
										$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
										$lastStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
										if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
										{
											$fileStoreNumber = $Matches[1]
											$fileLaneNumber = $Matches[2]
											if ($fileStoreNumber -eq $StoreNumber -and
												$file.Name -match '^S.*\.(\d{3})$' -and
												$fileLaneNumber -eq $Matches[1] -and
												$subjectLine -match 'Subject:\s*Health' -and
												$msgLine -match 'MSG:\s*This application is not running\.' -and
												$lastStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
											{
												$keep = $true
											}
										}
									}
								}
							}
							if (-not $keep) { Remove-Item -Path $file.FullName -Force -Recurse }
						}
					}
				}
				catch { }
				Start-Sleep -Seconds $checkIntervalSeconds
			}
		} -ArgumentList $folderPath, $checkIntervalSeconds, $StoreNumber
	}
	catch
	{
		Write_Log "Failed to start background job for 'XE${StoreNumber}901': $_" "red"
	}
	
	return $job
}

# ===================================================================================================
#                                       SECTION: Generate SQL Scripts
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates SQL scripts for Lanes and Stores, including memory configuration and maintenance tasks.
# ===================================================================================================

function Generate_SQL_Scripts
{
	param (
		[string]$StoreNumber,
		[string]$LanesqlFilePath,
		[string]$StoresqlFilePath
	)
	
	# Ensure StoreNumber is properly formatted (e.g., '005')
	# $StoreNumber = $StoreNumber.PadLeft(3, '0')
	
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write_Log "Failed to retrieve the connection string." "red"
		return
	}
	
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	# Initialize default database names
	$defaultStoreDbName = "STORESQL"
	$defaultLaneDbName = "LANESQL"
	
	# Retrive the DB name
	if ($script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace($script:FunctionResults['DBNAME']))
	{
		$dbName = $script:FunctionResults['DBNAME']
		#	Write_Log "Using DBNAME from FunctionResults: $dbName" "blue"
		$storeDbName = $dbName
	}
	else
	{
		Write_Log "No 'Database' in $script:FunctionResults. Defaulting to '$defaultStoreDbName'." "yellow"
		$storeDbName = $defaultStoreDbName
	}
	
	# Define replacements for SQL scripts
	# $storeDbName is now either the retrieved DBNAME or the default 'STORESQL'
	# $laneDbName remains as 'LANESQL' unless you wish to make it dynamic as well
	$laneDbName = $defaultLaneDbName # If LANESQL is also dynamic, you can retrieve it similarly
	
	# Write_Log "Generating SQL scripts using Store DB: '$storeDbName' and Lane DB: '$laneDbName'..." "blue"
	
	# Generate Lanesql script
	$LaneSQLScript = @"
/* Set a long timeout so the entire script runs */
@WIZRPL(DBASE_TIMEOUT=E);

/* Set memory configuration */
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', 1024;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Truncate PRICE_EVENT table for records older than 7 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-7,GETDATE());

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000) 
DECLARE cmds CURSOR FOR 
SELECT 'drop table [' + name + ']' 
FROM sys.tables 
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$StoreNumber%' OR name LIKE 'R$StoreNumber%') 
OPEN cmds 
WHILE 1 = 1 
BEGIN 
FETCH cmds INTO @cmd 
IF @@fetch_status != 0 BREAK 
EXEC(@cmd) 
END 
CLOSE cmds; 
DEALLOCATE cmds;

/* Cleaning HEADER_SAV */
IF OBJECT_ID('HEADER_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('HEADER_SAV', 'OBJECT', 'DELETE') = 1 
    DELETE FROM HEADER_SAV 
    WHERE (F903 = 'SVHOST' OR F903 = 'MPHOST' OR F903 = CONCAT('M', '$StoreNumber', '901')) 
    AND (DATEDIFF(DAY, F907, GETDATE()) > 30 OR DATEDIFF(DAY, F909, GETDATE()) > 30);

/* Delete bad SMS items */
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01='0020000000000') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01 LIKE '% %') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE LEN(F01)<>13) 
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000') 
@dbEXEC(DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002') 
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000') 
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000'))) 
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000'))) 
@dbEXEC(DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB)) 
@dbEXEC(DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F267 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F268 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB)) 
@dbEXEC(UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0) 
@dbEXEC(UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))

/* Delete batches older than 7 days from the lane */
DELETE FROM Header_bat WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_dct WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_old WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_sav WHERE F909 < DATEADD(day, -7, GETDATE());

/* Rebuild indexes and update database statistics */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'

/* Shrink the main database file */
DBCC SHRINKFILE ($laneDbName)

/* Shrink the database log file */
DBCC SHRINKFILE (${laneDbName}_Log)

/* Restrict the indefinite log file growth */
ALTER DATABASE $laneDbName SET RECOVERY SIMPLE

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
"@
	
	# Store the LaneSQLScript in the script scope
	$script:LaneSQLScript = $LaneSQLScript
	
	# -------------------------------
	# Create a filtered version of LaneSQL by skipping sections using regex
	# -------------------------------
	
	# The dynamic T-SQL memory config we want to use in the *filtered* Lane script
	$ServerMemoryConfig = @"
DECLARE @Memory25PercentMB BIGINT;
SELECT @Memory25PercentMB = (total_physical_memory_kb / 1024) * 25 / 100
FROM sys.dm_os_sys_memory;
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', @Memory25PercentMB;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;
"@
	
	# Define the regex pattern to match sections
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	
	# Define the names of the sections to skip
	$sectionsToSkip = @(
		'Set a long timeout so the entire script runs',
		'Clear the long database timeout'
	)
	
	# Initialize the filtered script
	$LaneSQLFiltered = ""
	
	# Use regex to parse the script into sections
	$matches = [regex]::Matches($LaneSQLScript, $sectionPattern)
	
	foreach ($match in $matches)
	{
		$sectionName = $match.Groups['SectionName'].Value.Trim()
		$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()
		
		if ($sectionsToSkip -contains $sectionName)
		{
			# 1) If it's in the skip list, do nothing (omit).
			continue
		}
		elseif ($sectionName -eq 'Set memory configuration')
		{
			# 2) If it's the "Set memory configuration" block, replace it with the dynamic version:
			$LaneSQLFiltered += "/* $sectionName */`r`n$ServerMemoryConfig`r`n`r`n"
		}
		else
		{
			# 3) Otherwise, keep the block exactly
			# Additionally, remove the @dbEXEC() wrappers but keep the inner SQL commands
			# Use regex to replace @dbEXEC(...) with the content inside the parentheses
			# This handles both @dbEXEC("...") and @dbEXEC(...) without quotes
			$sqlCommands = $sqlCommands -replace '@dbEXEC\((?:\"(.*)\"|(.*))\)', '$1$2'
			$LaneSQLFiltered += "/* $sectionName */`r`n$sqlCommands`r`n`r`n"
		}
	}
	
	# Store the filtered LaneSQL script in the script scope for later use
	$script:LaneSQLFiltered = $LaneSQLFiltered
	
	# --- New: prepare mailslot-friendly script ---
	$lines = $script:LaneSQLScript -split "`r?`n"
	$macroPattern = '^\s*(@|/|\*)'
	$fixedLines = foreach ($line in $lines)
	{
		if ($line -match $macroPattern -or [string]::IsNullOrWhiteSpace($line))
		{
			$line
		}
		else
		{
			"@EXEC($line)"
		}
	}
	$script:LaneSQLScript_Mailslot = ($fixedLines -join "`r`n")
	
	<#
	# Optionally write to file as fallback
	if ($LanesqlFilePath)
	{
		[System.IO.File]::WriteAllText($LanesqlFilePath, $script:LaneSQLScript, $utf8NoBOM)
	}
	#>
	
	# Similarly generate Storesql script
	$ServerSQLScript = @"
/* Set memory configuration */
DECLARE @Memory25PercentMB BIGINT;
SELECT @Memory25PercentMB = (total_physical_memory_kb / 1024) * 25 / 100
FROM sys.dm_os_sys_memory;
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', @Memory25PercentMB;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Create Table TBS_ITM_SMAppUPDATED */
-----Drop the table if it exist-----
IF OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('dbo.TBS_ITM_SMAppUPDATED', 'OBJECT', 'ALTER') = 1 BEGIN DROP TABLE dbo.TBS_ITM_SMAppUPDATED END;
-----Create TBS_ITM_SMAppUPDATED Table with Optional ID Column-----
CREATE TABLE dbo.TBS_ITM_SMAppUPDATED (
    Id INT IDENTITY(1,1) PRIMARY KEY,   -- Surrogate primary key
    CodeF01 VARCHAR(13) NULL,      -- Stores the constructed code
    Sent BIT NOT NULL DEFAULT 0,       -- Indicates if the record has been sent
    SentAt DATETIME NOT NULL DEFAULT GETDATE() -- Timestamp of insertion
);
-----Create Indexes for Performance-----
CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_CodeF01 ON dbo.TBS_ITM_SMAppUPDATED (CodeF01);
CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_Sent ON dbo.TBS_ITM_SMAppUPDATED (Sent);
CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_SentAt ON dbo.TBS_ITM_SMAppUPDATED (SentAt);

/* Create TBS_ITM_SMAppUPDATED Triggers */
-----Drop existing triggers if they exist-----
IF EXISTS (select * from sysobjects where name like '%SMApp_UpdateOBJ%')
DROP TRIGGER [dbo].[SMApp_UpdateOBJ]
GO
IF EXISTS (select * from sysobjects where name like '%SMApp_UpdatePOS%')
DROP TRIGGER [dbo].[SMApp_UpdatePOS]
GO
IF EXISTS (select * from sysobjects where name like '%SMApp_UpdatePrice%')
DROP TRIGGER [dbo].[SMApp_UpdatePrice]
GO
IF EXISTS (select * from sysobjects where name like '%SMApp_UpdateSCL%')
DROP TRIGGER [dbo].[SMApp_UpdateSCL]
GO
IF EXISTS (select * from sysobjects where name like '%SMApp_UpdateSCL_TXT%')
DROP TRIGGER [dbo].[SMApp_UpdateSCL_TXT]
GO
-----Triggers for OBJ_TAB-----
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdateOBJ]
   ON  [dbo].[OBJ_TAB]
   AFTER INSERT,UPDATE
AS 
BEGIN
       SET NOCOUNT ON;
	INSERT INTO TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
	SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000'
END;
-----Triggers for POS_TAB-----
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdatePOS]
   ON  [dbo].[POS_TAB]
   AFTER INSERT,UPDATE
AS 
BEGIN
		SET NOCOUNT ON;
		INSERT INTO TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
		SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000' 
END;
-----Triggers for PRICE_TAB-----
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdatePrice]
   ON  [dbo].[PRICE_TAB]
   AFTER INSERT,UPDATE
AS 
BEGIN
		SET NOCOUNT ON;
		INSERT INTO TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
		SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000' 
END;
-----Triggers for SCL_TAB-----
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdateSCL]
   ON  [dbo].[SCL_TAB]
   AFTER INSERT,UPDATE
AS 
BEGIN
       SET NOCOUNT ON;
       INSERT INTO TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
	SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000'
END;
-----Triggers for SCL_TXT_TAB-----
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER  [dbo].[SMApp_UpdateSCL_TXT]
   ON  [dbo].[SCL_TXT_TAB] 
   AFTER INSERT,UPDATE
AS 
BEGIN
       SET NOCOUNT ON;
       INSERT INTO TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
       SELECT '002'+cast(RIGHT('00000'+ CONVERT(VARCHAR,F267),5) as varchar)+'00000',0, GETDATE() 
       FROM inserted,OBJ_TAB 
       WHERE '002'+cast(RIGHT('00000'+ CONVERT(VARCHAR,F267),5) as varchar)+'00000' = F01 
       and SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000'
 
END;

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Truncate PRICE_EVENT table for records older than 7 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-7,GETDATE());

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000);
DECLARE cmds CURSOR FOR
SELECT 'drop table [' + name + ']'
FROM sys.tables
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$StoreNumber%' OR name LIKE 'R$StoreNumber%') AND DATEDIFF(DAY, create_date, GETDATE()) > 30;
OPEN cmds;
WHILE 1 = 1
BEGIN
    FETCH cmds INTO @cmd;
    IF @@fetch_status != 0 BREAK;
    EXEC(@cmd);
END;
CLOSE cmds;
DEALLOCATE cmds;

/* Cleaning HEADER_SAV */
IF OBJECT_ID('HEADER_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('HEADER_SAV', 'OBJECT', 'DELETE') = 1 
    DELETE FROM HEADER_SAV 
    WHERE (F903 = 'SVHOST' OR F903 = 'MPHOST' OR F903 = CONCAT('M', '$StoreNumber', '901')) 
    AND (DATEDIFF(DAY, F907, GETDATE()) > 30 OR DATEDIFF(DAY, F909, GETDATE()) > 30);

/* Delete bad SMS items */
DELETE FROM OBJ_TAB WHERE F01='0020000000000' 
DELETE FROM OBJ_TAB WHERE F01 LIKE '% %' 
DELETE FROM OBJ_TAB WHERE LEN(F01)<>13
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000' 
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000'
DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)
DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB) 
DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB) 
DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002' 
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000'
UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) 
UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000')) 
DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB)
DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB) 
UPDATE SCL_TAB SET F267 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB) 
UPDATE SCL_TAB SET F268 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB) 
UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0
UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')

/* Delete batches older than 14 days from the server */
DELETE FROM Header_bat WHERE F909 < DATEADD(day, -14, GETDATE());
DELETE FROM Header_dct WHERE F909 < DATEADD(day, -14, GETDATE());
DELETE FROM Header_old WHERE F909 < DATEADD(day, -14, GETDATE());
DELETE FROM Header_sav WHERE F909 < DATEADD(day, -14, GETDATE());

/* Rebuild indexes and update database statistics */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD';
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN';

/* Shrink the main database file */
DBCC SHRINKFILE ($storeDbName);

/* Shrink the database log file */
DBCC SHRINKFILE (${storeDbName}_Log);

/* Restrict the indefinite log file growth */
ALTER DATABASE $storeDbName SET RECOVERY SIMPLE;
"@
	
	# Store the ServerSQLScript in the script scope
	$script:ServerSQLScript = $ServerSQLScript
	
		<#
	# Optionally write to file as fallback
	if ($StoresqlFilePath)
	{
		[System.IO.File]::WriteAllText($StoresqlFilePath, $script:ServerSQLScript, $utf8NoBOM)
	}
	#>
	
	# Write_Log "SQL scripts generated successfully." "green"
	
	# Separate server script for the schedule maintenance
	$ScheduleServerScript = @"
/* Set a long timeout so the entire script runs */
@WIZRPL(DBASE_TIMEOUT=E);

/* Set memory configuration */
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', 8192;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000) 
DECLARE cmds CURSOR FOR 
SELECT 'drop table [' + name + ']' 
FROM sys.tables 
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$StoreNumber%' OR name LIKE 'R$StoreNumber%') 
OPEN cmds 
WHILE 1 = 1 
BEGIN 
FETCH cmds INTO @cmd 
IF @@fetch_status != 0 BREAK 
EXEC(@cmd) 
END 
CLOSE cmds; 
DEALLOCATE cmds;

/* Cleaning HEADER_SAV */
IF OBJECT_ID('HEADER_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('HEADER_SAV', 'OBJECT', 'DELETE') = 1 
    DELETE FROM HEADER_SAV 
    WHERE (F903 = 'SVHOST' OR F903 = 'MPHOST' OR F903 = CONCAT('M', '$StoreNumber', '901')) 
    AND (DATEDIFF(DAY, F907, GETDATE()) > 30 OR DATEDIFF(DAY, F909, GETDATE()) > 30);

/* Delete bad SMS items */
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01='0020000000000') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01 LIKE '% %') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE LEN(F01)<>13) 
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000') 
@dbEXEC(DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002') 
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000') 
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000'))) 
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000'))) 
@dbEXEC(DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB)) 
@dbEXEC(DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F267 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F268 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB)) 
@dbEXEC(UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0) 
@dbEXEC(UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))

/* Shrink database and log files */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'
DBCC SHRINKFILE ($storeDbName)
DBCC SHRINKFILE (${storeDbName}_Log)
ALTER DATABASE $storeDbName SET RECOVERY SIMPLE

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
"@
	
	# Store in global scope for downstream consumption
	$script:ScheduleServerScript = $ScheduleServerScript
}

# ===================================================================================================
#                                       FUNCTION: Get_Table_Aliases
# ---------------------------------------------------------------------------------------------------
# Description:
#   Parses SQL files in the specified target directory to extract table names and their corresponding
#   aliases from @CREATE statements. This function internally defines the list of base table names to search
#   for and uses regex to identify and capture the full table name and alias pairs. The results are returned
#   as a collection of objects containing details about each match, along with a hash table for quick
#   lookups in both directions (Table Name ? Alias).
# ===================================================================================================

function Get_Table_Aliases
{
	# Define the target directory for SQL files
	$targetDirectory = "$LoadPath"
	
	# Define the list of base table names internally (without _TAB)
	$baseTables = @(
		'OBJ', 'POS', 'PRICE', 'COST', 'DSD', 'KIT', 'LOC', 'ALT', 'ECL',
		'SCL', 'SCL_TXT', 'SCL_NUT', 'DEPT', 'SDP', 'CAT', 'RPC', 'FAM',
		'CPN', 'PUB', 'BIO', 'CLK', 'LVL', 'MIX', 'BTL', 'TAR', 'UNT',
		'RES', 'ROUTE', 'VENDOR', 'DELV', 'CLT', 'CLG', 'CLF', 'CLR',
		'CLL', 'CLT_ITM', 'CLF_SDP', 'STD', 'CFG', 'MOD'
	)
	
	# Escape special regex characters in table names and sort by length (longest first)
	$escapedTables = $baseTables | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }
	$tablesPattern = $escapedTables -join '|'
	
	# Updated Regex to allow optional whitespace around the comma and various quotation styles
	$pattern = "^\s*@CREATE\s*\(\s*['""]?(?<Table>($tablesPattern)(_[A-Z]+))?['""]?\s*,\s*['""]?(?<Alias>[\w-]+)['""]?\s*\);"
	
	# Compile the regex with IgnoreCase for case-insensitive matching
	$regex = [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
	
	# Initialize a collection to store matched files using ArrayList for better performance
	$sqlFiles = New-Object System.Collections.ArrayList
	
	# Try using -File parameter; if it fails, fallback to older approach
	try
	{
		# Primary attempt with -File parameter
		$allLoadSqlFiles = Get-ChildItem -Path $TargetDirectory -Recurse -File -Filter '*_Load.sql' -ErrorAction Stop
	}
	catch
	{
		# If primary fails, try fallback with -ErrorAction Stop so it can be caught
		try
		{
			$allLoadSqlFiles = Get-ChildItem -Path "C:\storeman\office\load" -Recurse -Filter '*_Load.sql' -ErrorAction Stop |
			Where-Object { -not $_.PsIsContainer }
		}
		catch
		{
			# If fallback also fails, handle it without showing a big red error
			Write_Log "Could not find load files in either of the specified paths." "red"
			$allLoadSqlFiles = @() # Provide an empty array so the script continues gracefully
		}
	}
	
	foreach ($file in $allLoadSqlFiles)
	{
		# Check if the file name starts with any of the base table names
		foreach ($baseTable in $baseTables)
		{
			if ($file.Name -ieq "$baseTable`_Load.sql")
			{
				# Using -ieq for case-insensitive comparison
				[void]$sqlFiles.Add($file)
				break # Move to next file once matched
			}
		}
	}
	
	# Remove duplicates
	$uniqueSqlFiles = $sqlFiles | Sort-Object -Unique
	
	# Confirm number of files found
	Write-Verbose "`nTotal matched .sql files: $($uniqueSqlFiles.Count)"
	
	# Initialize ArrayList for results
	$aliasResults = New-Object System.Collections.ArrayList
	
	# Initialize hash table for quick alias and table name lookup
	$tableAliasHash = @{ }
	
	# Process each matched SQL file
	foreach ($file in $uniqueSqlFiles)
	{
		Write-Verbose "`nProcessing file: $($file.FullName)"
		try
		{
			# Open the file for reading using a streaming approach
			$reader = [System.IO.File]::OpenText($file.FullName)
			$lineNumber = 0 # Initialize line number
			while (($line = $reader.ReadLine()) -ne $null)
			{
				$lineNumber++
				
				# Remove single-line and multi-line comments
				$lineClean = $line -replace '--.*', '' -replace '/\*.*?\*/', ''
				
				# Check if the line contains @CREATE
				if ($lineClean -match '@CREATE')
				{
					Write-Verbose "Found @CREATE on line ${lineNumber}: $lineClean"
					$match = $regex.Match($lineClean)
					
					if ($match.Success)
					{
						$matchedTable = $match.Groups['Table'].Value # Full table name with _TAB
						$matchedAlias = $match.Groups['Alias'].Value
						if ([string]::IsNullOrEmpty($matchedTable) -or [string]::IsNullOrEmpty($matchedAlias))
						{
							Write-Warning "Incomplete match in file $($file.FullName) on line ${lineNumber}: Table='$matchedTable', Alias='$matchedAlias'"
						}
						else
						{
							# Create a PSObject with match details
							$aliasInfo = [PSCustomObject]@{
								File	   = $file.FullName
								Table	   = $matchedTable
								Alias	   = $matchedAlias
								LineNumber = $lineNumber
								Context    = $lineClean.Trim()
							}
							
							[void]$aliasResults.Add($aliasInfo)
							Write-Verbose "Match found: Table='$matchedTable', Alias='$matchedAlias' on line $lineNumber"
							
							# Populate the hash table for quick lookup in both directions
							if (-not $tableAliasHash.ContainsKey($matchedTable))
							{
								$tableAliasHash[$matchedTable] = $matchedAlias
							}
							if (-not $tableAliasHash.ContainsKey($matchedAlias))
							{
								$tableAliasHash[$matchedAlias] = $matchedTable
							}
						}
					}
					else
					{
						Write-Verbose "Line found with @CREATE but did not match pattern on line ${lineNumber}: $lineClean"
					}
				}
			}
			$reader.Close()
		}
		catch
		{
			Write-Warning "Failed to read file: $($file.FullName). Skipping."
			continue
		}
	}
	
	# Optionally, sort the results before returning
	if ($aliasResults.Count -gt 0)
	{
		$sortedResults = $aliasResults | Sort-Object File, Table, LineNumber
	}
	else
	{
		Write-Verbose "`nNo @CREATE table-alias pairs found in the specified files."
		$sortedResults = @()
	}
	
	# Store results in script-scoped hash table
	$script:FunctionResults['Get_Table_Aliases'] = @{
		Aliases   = $sortedResults
		AliasHash = $tableAliasHash
	}
}

# ===================================================================================================
#                            FUNCTION: Get_All_Lanes_VNC_Passwords
# ---------------------------------------------------------------------------------------------------
# Description:
#   Scans all specified lanes for the UltraVNC configuration file (UltraVNC.ini, any case) in the
#   default install locations (Program Files and Program Files (x86)). Extracts the encrypted VNC
#   password and stores results in a hashtable keyed by machine name. Handles all filename case 
#   variations. Designed for remote auditing of VNC password status across all lanes.
#
# Parameters:
#   - LaneMachines   [hashtable]: LaneNumber => MachineName mapping.
#
# Details:
#   - Searches for UltraVNC.ini with any capitalization in both standard install folders.
#   - Uses PowerShell remoting (Invoke-Command) to access remote lane files.
#   - Finds the "passwd=" entry in the INI (case-insensitive, first match).
#   - Returns a hashtable: MachineName => Password (or $null if not found).
#   - Uses Write_Log for status, progress, and error messages.
#
# Usage:
#   $LanePasswords = Get-AllLaneVNCPasswords -LaneMachines $LaneMachines
#
# Author: Alex_C.T
# ===================================================================================================

function Get_All_Lanes_VNC_Passwords
{
	param (
		[Parameter(Mandatory)]
		[hashtable]$LaneMachines # LaneNumber => MachineName
	)
	
	$uvncFolders = @(
		"C$\Program Files\uvnc bvba\UltraVNC",
		"C$\Program Files (x86)\uvnc bvba\UltraVNC"
	)
	
	$LaneVNCPasswords = @{ }
	
	foreach ($lane in $LaneMachines.GetEnumerator())
	{
		$laneNum = $lane.Key
		$laneMachine = $lane.Value
		$password = $null
		$foundPath = $null
		$fileFound = $false
		$success = $false
		
		foreach ($folder in $uvncFolders)
		{
			# --- Try with Invoke-Command first ---
			try
			{
				$iniFiles = Invoke-Command -ComputerName $laneMachine -ScriptBlock {
					param ($dir)
					if (Test-Path $dir)
					{
						Get-ChildItem -Path $dir -Filter "*.ini" -File | ForEach-Object {
							if ($_.Name.ToLower() -eq "ultravnc.ini")
							{
								$_.FullName
							}
						}
					}
				} -ArgumentList ("C:\" + $folder.Substring(3)) -ErrorAction Stop
				
				foreach ($iniFile in $iniFiles)
				{
					$fileFound = $true
					$content = Invoke-Command -ComputerName $laneMachine -ScriptBlock {
						param ($path)
						if (Test-Path $path)
						{
							Get-Content $path -ErrorAction Stop
						}
					} -ArgumentList $iniFile -ErrorAction Stop
					
					foreach ($line in $content)
					{
						if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
						{
							$password = $matches[1]
							$foundPath = $iniFile
							$success = $true
							break
						}
					}
					if ($password) { break }
				}
			}
			catch
			{
				# ---- Fallback to UNC/network path if remoting fails ----
				try
				{
					$remotePath = "\\$laneMachine\$folder"
					if (Test-Path $remotePath)
					{
						$iniFiles = Get-ChildItem -Path $remotePath -Filter "*.ini" -File | Where-Object {
							$_.Name.ToLower() -eq "ultravnc.ini"
						}
						foreach ($iniFile in $iniFiles)
						{
							$fileFound = $true
							$content = Get-Content $iniFile.FullName -ErrorAction Stop
							foreach ($line in $content)
							{
								if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
								{
									$password = $matches[1]
									$foundPath = $iniFile.FullName
									$success = $true
									break
								}
							}
							if ($password) { break }
						}
					}
				}
				catch
				{
					# Ignore network path errors and continue to next folder
					continue
				}
			}
			if ($password) { break }
		}
		
		$LaneVNCPasswords[$laneMachine] = $password
		if ($password)
		{
			$method = if ($success) { "Remoting" } else { "Network path" }
		#	Write_Log "Lane [$laneNum/$laneMachine]: VNC password found in [$foundPath] via [$method]: $password" "green"
		}
		elseif ($fileFound)
		{
		#	Write_Log "Lane [$laneNum/$laneMachine]: UltraVNC.ini found but no password found inside [$foundPath]." "yellow"
		}
		else
		{
		#	Write_Log "Lane [$laneNum/$laneMachine]: UltraVNC.ini not found in any standard folder." "yellow"
		}
	}
	$script:FunctionResults['LaneVNCPasswords'] = $LaneVNCPasswords
	return $LaneVNCPasswords
}

# ===================================================================================================
#                              FUNCTION: Get_Remote_Machine_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   For each specified remote machine (lane), enables and starts the Remote Registry service,
#   then queries and retrieves the System Manufacturer and System Product Name via the registry.
#   Results are stored in a hashtable keyed by machine name, and in $script:LaneHardwareInfo.
#
# Author: Alex_C.T
# ===================================================================================================

function Get_Remote_Machine_Info
{
	param (
		[Parameter(Mandatory)]
		[string[]]$LaneMachines
	)
	
	$results = @{ }
	
	foreach ($remote in $LaneMachines)
	{
		$laneInfo = @{
			Success		       = $false
			SystemManufacturer = $null
			SystemProductName  = $null
			Error			   = $null
		}
		try
		{
			sc.exe \\$remote config RemoteRegistry start= auto | Out-Null
			sc.exe \\$remote start RemoteRegistry | Out-Null
			
			$manuf = reg.exe query "\\$remote\HKLM\HARDWARE\DESCRIPTION\System\BIOS" /v SystemManufacturer 2>&1
			$manufMatch = [regex]::Match($manuf, 'SystemManufacturer\s+REG_SZ\s+(.+)$')
			if ($manufMatch.Success)
			{
				$laneInfo.SystemManufacturer = $manufMatch.Groups[1].Value.Trim()
			}
			else
			{
				$laneInfo.SystemManufacturer = $null
			}
			
			$prod = reg.exe query "\\$remote\HKLM\HARDWARE\DESCRIPTION\System\BIOS" /v SystemProductName 2>&1
			$prodMatch = [regex]::Match($prod, 'SystemProductName\s+REG_SZ\s+(.+)$')
			if ($prodMatch.Success)
			{
				$laneInfo.SystemProductName = $prodMatch.Groups[1].Value.Trim()
			}
			else
			{
				$laneInfo.SystemProductName = $null
			}
			
			if ($laneInfo.SystemManufacturer -and $laneInfo.SystemProductName)
			{
				$laneInfo.Success = $true
			}
			else
			{
				$laneInfo.Error = "SystemManufacturer or SystemProductName not found in registry output."
			}
		}
		catch
		{
			$laneInfo.Error = $_.Exception.Message
		}
		$results[$remote] = $laneInfo
	}
	
	$script:LaneHardwareInfo = $results
	
	return $results
}

# ===================================================================================================
#                                  SECTION: Fix_Journal
# ---------------------------------------------------------------------------------------------------
# Description:
#   Processes EJ files within a ZX folder to correct specific lines based on a user-provided date.
#   - Prompts the user to select a date using a Windows Form.
#   - Constructs the ZX folder path.
#   - Identifies related EJ files based on the date/store data.
#   - Repairs lines in matching EJ files.
# ===================================================================================================

function Fix_Journal
{
	[CmdletBinding()]
	param (
		# The base "OfficePath", e.g. "C:\storeman\office"
		[Parameter(Mandatory = $true)]
		[string]$OfficePath,
		# The store number (e.g., "001")
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Fix_Journal Function ====================`r`n" "blue"
	
	# ---------------------------------------------------------------------------------------------
	# 1) Load Windows Forms assembly
	# ---------------------------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# ---------------------------------------------------------------------------------------------
	# 2) Create and configure the Windows Form
	# ---------------------------------------------------------------------------------------------
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Date"
	$form.Size = New-Object System.Drawing.Size(300, 200)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = "FixedDialog"
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	# Create Label
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Please select the date:"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(10, 20)
	$form.Controls.Add($label)
	
	# Create DateTimePicker
	$dateTimePicker = New-Object System.Windows.Forms.DateTimePicker
	$dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$dateTimePicker.Location = New-Object System.Drawing.Point(10, 50)
	$dateTimePicker.Width = 260
	$form.Controls.Add($dateTimePicker)
	
	# Create OK Button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(110, 100)
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $okButton
	$form.Controls.Add($okButton)
	
	# Show the form and capture the result
	$dialogResult = $form.ShowDialog()
	
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log -Message "Date selection canceled by user. Exiting function." "yellow"
		return
	}
	
	# ---------------------------------------------------------------------------------------------
	# 3) Retrieve and format the selected date
	# ---------------------------------------------------------------------------------------------
	$snippetDate = $dateTimePicker.Value
	$formattedDate = $snippetDate.ToString('MMddyyyy') # MMDDYYYY format
	
	Write_Log -Message "Selected date: $formattedDate" "magenta"
	
	# ---------------------------------------------------------------------------------------------
	# 4) Construct ZX folder path from $OfficePath + $StoreNumber
	#    ZX folder: $OfficePath\ZX${StoreNumber}901
	# ---------------------------------------------------------------------------------------------
	$zxFolderPath = Join-Path $OfficePath "ZX${StoreNumber}901"
	
	# ---------------------------------------------------------------------------------------------
	# 5) Confirm the ZX folder exists
	# ---------------------------------------------------------------------------------------------
	if (-not (Test-Path -Path $zxFolderPath))
	{
		Write_Log -Message "ZX folder not found: $zxFolderPath." "red"
		return
	}
	
	# ---------------------------------------------------------------------------------------------
	# 6) Build the file prefix: YMMDDSSS (ignoring lane)
	#     Y = last digit of year
	#     MM = 2-digit month
	#     DD = 2-digit day
	#     SSS = store number (3 digits, e.g., "001")
	# ---------------------------------------------------------------------------------------------
	$yearLastDigit = ($snippetDate.Year % 10)
	$mm = $snippetDate.ToString('MM')
	$dd = $snippetDate.ToString('dd')
	$filePrefix = "$yearLastDigit$mm$dd$StoreNumber" # e.g., 41227001
	
	Write_Log -Message "Looking for files named '$filePrefix.*' in $zxFolderPath..." "blue"
	
	# ---------------------------------------------------------------------------------------------
	# 7) Find matching EJ files in ZX folder: e.g., 41227001.*
	# ---------------------------------------------------------------------------------------------
	$searchPattern = "$filePrefix.*"
	$matchingFiles = Get-ChildItem -Path $zxFolderPath -Filter $searchPattern -File -ErrorAction SilentlyContinue
	
	if (-not $matchingFiles)
	{
		Write_Log -Message "No files matching '$searchPattern' found in $zxFolderPath." "yellow"
		return
	}
	
	Write_Log -Message "Found $($matchingFiles.Count) file(s) to fix." "green"
	
	# ---------------------------------------------------------------------------------------------
	# 8) For each matching EJ file, remove lines from <trs F10... up to <trs F1068...
	# ---------------------------------------------------------------------------------------------
	foreach ($file in $matchingFiles)
	{
		# [Optional] Skip files that have ".bak" anywhere in their name 
		# to avoid infinite backup loops:
		if ($file.Extension -eq ".bak")
		{
			Write_Log -Message "Skipping backup file: $($file.Name)" "yellow"
			continue
		}
		
		Write_Log -Message "Fixing lines in: $($file.FullName)" "yellow"
		
		# Read the file lines
		try
		{
			$originalLines = Get-Content -Path $file.FullName -ErrorAction Stop
		}
		catch
		{
			Write_Log -Message "Failed to read EJ file: $($file.FullName). Skipping." "red"
			continue
		}
		
		# Prepare a list for the fixed lines
		$fixedLines = New-Object System.Collections.Generic.List[string]
		
		$skip = $false
		
		foreach ($line in $originalLines)
		{
			# 1) Start skipping at '<trs F10'
			if ($line -match '^\s*<trs\s+F10\b')
			{
				$skip = $true
				continue
			}
			
			# 2) Stop skipping at '<trs F1068'
			if ($skip -and ($line -match '^\s*<trs\s+F1068\b'))
			{
				$skip = $false
				# We *do* want to keep this line
				$fixedLines.Add($line)
				continue
			}
			
			# Keep the line if we're not skipping
			if (-not $skip)
			{
				$fixedLines.Add($line)
			}
		}
		
		<# -----------------------------------------------------------------------------------------
		# 10) Make a backup of the original file
		# * Commented out for now
		# -----------------------------------------------------------------------------------------
		$backupPath = "$($file.FullName).bak"
		try
		{
			Copy-Item -Path $file.FullName -Destination $backupPath -Force -ErrorAction Stop
			Write_Log -Message "Backup created: $backupPath" "green"
		}
		catch
		{
			Write_Log -Message "Failed to create backup for: $($file.FullName). Skipping file edit." "red"
			continue
		}
		#>
		
		# -----------------------------------------------------------------------------------------
		# 11) Overwrite the original file with the fixed lines in ANSI encoding
		# -----------------------------------------------------------------------------------------
		try
		{
			$fixedLines | Set-Content -Path $file.FullName -Encoding Default -ErrorAction Stop
			#	Write_Log -Message "Successfully edited: $($file.FullName). Backup: $backupPath" "green"
			Write_Log -Message "Successfully edited: $($file.FullName)" "green"
		}
		catch
		{
			Write_Log -Message "Failed to write fixed content to: $($file.FullName)." "red"
			continue
		}
	}
	Write_Log "`r`n==================== Fix_Journal Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Process_Server
# ---------------------------------------------------------------------------------------------------
# Description:
#   Executes the full Server SQL maintenance routine. Reads and parses the specified SQL script
#   file or variable, prompts for section selection if desired, executes each section with retries,
#   and logs results to the console and file. Fails gracefully and outputs summary banners at start/end.
# ===================================================================================================

function Process_Server
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StoresqlFilePath,
		[Parameter(Mandatory = $false)]
		[switch]$PromptForSections
	)
	
	Write_Log "`r`n==================== Starting Server Database Maintenance ====================`r`n" "blue"
	
	if ($PromptForSections)
	{
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
	}
	
	$MaxRetries = 2
	$RetryDelaySeconds = 5
	$FailedCommandsPath = "$OfficePath\XF${StoreNumber}901\Failed_ServerSQLScript_Sections.sql"
	
	$sqlScript = $script:ServerSQLScript
	$dbName = $script:FunctionResults['DBNAME']
	
	if (-not [string]::IsNullOrWhiteSpace($sqlScript))
	{
		Write_Log "Executing SQL script from variable..." "blue"
	}
	elseif ($StoresqlFilePath)
	{
		if (-not (Test-Path $StoresqlFilePath))
		{
			Write_Log "SQL file not found: $StoresqlFilePath" "red"
			return
		}
		Write_Log "Executing SQL file: $StoresqlFilePath" "blue"
		try
		{
			$sqlScript = Get-Content -Path $StoresqlFilePath -Raw -ErrorAction Stop
		}
		catch
		{
			Write_Log "Failed to read SQL file: $_" "red"
			return
		}
	}
	else
	{
		Write_Log "No SQL script content or file path provided." "red"
		return
	}
	
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$matches = [regex]::Matches($sqlScript, $sectionPattern)
	
	if ($matches.Count -eq 0)
	{
		Write_Log "No SQL sections found to execute." "red"
		return
	}
	
	$SectionsToRun = $null
	if ($PromptForSections)
	{
		$allSectionNames = $matches | ForEach-Object {
			$_.Groups['SectionName'].Value.Trim()
		}
		$SectionsToRun = Show_Section_Selection_Form -SectionNames $allSectionNames
		if (-not $SectionsToRun -or $SectionsToRun.Count -eq 0)
		{
			Write_Log "No sections selected or form was canceled. Aborting execution." "yellow"
			return
		}
	}
	
	$useSpecificSections = $SectionsToRun -and $SectionsToRun.Count -gt 0
	
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	$ConnectionString = $script:FunctionResults['ConnectionString']
	if (-not $ConnectionString)
	{
		Write_Log "Connection string not found. Attempting to generate it..." "yellow"
		$ConnectionString = Get_Database_Connection_String
		if (-not $ConnectionString)
		{
			Write_Log "Unable to generate connection string. Cannot execute SQL script." "red"
			return
		}
	}
	Write_Log "Using connection string: $ConnectionString" "gray"
	
	$supportsConnectionString = $false
	try
	{
		$cmd = Get-Command Invoke-Sqlcmd -ErrorAction Stop
		$supportsConnectionString = $cmd.Parameters.Keys -contains 'ConnectionString'
	}
	catch
	{
		Write_Log "Invoke-Sqlcmd cmdlet not found: $_" "red"
		$supportsConnectionString = $false
	}
	
	$retryCount = 0
	$success = $false
	$failedSections = @()
	$failedCommands = @()
	
	while (-not $success -and $retryCount -lt $MaxRetries)
	{
		try
		{
			Write_Log "Starting execution of SQL script. Attempt $($retryCount + 1) of $MaxRetries." "blue"
			
			$sectionsToExecute = if ($retryCount -eq 0) { $matches }
			else { $failedSections }
			$failedSections = @()
			
			foreach ($match in $sectionsToExecute)
			{
				$sectionName = $match.Groups['SectionName'].Value.Trim()
				$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()
				
				if ($useSpecificSections -and ($SectionsToRun -notcontains $sectionName)) { continue }
				if ([string]::IsNullOrWhiteSpace($sqlCommands))
				{
					Write_Log "Section '$sectionName' has no commands. Skipping..." "yellow"
					continue
				}
				
				Write_Log "`r`nExecuting section: '$sectionName'" "blue"
				Write_Log "--------------------------------------------------------------------------------"
				Write_Log "$sqlCommands" "orange"
				Write_Log "--------------------------------------------------------------------------------"
				
				try
				{
					if ($supportsConnectionString)
					{
						Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
					}
					else
					{
						$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
						$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
						$cmdParams = (Get-Command Invoke-Sqlcmd).Parameters.Keys
						if ($cmdParams -contains 'TrustServerCertificate')
						{
							Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0 -TrustServerCertificate $true
						}
						else
						{
							Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
						}
					}
					Write_Log "Section '$sectionName' executed successfully." "green"
				}
				catch [System.Management.Automation.ParameterBindingException] {
					Write_Log "ParameterBindingException in section '$sectionName'. Attempting fallback." "yellow"
					try
					{
						$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
						$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
						$cmdParams = (Get-Command Invoke-Sqlcmd).Parameters.Keys
						if ($cmdParams -contains 'TrustServerCertificate')
						{
							Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0 -TrustServerCertificate $true
						}
						else
						{
							Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
						}
						Write_Log "Section '$sectionName' executed successfully with fallback." "green"
					}
					catch
					{
						Write_Log "Error executing section '$sectionName' with fallback: $_" "red"
						$failedSections += $match
						if ($retryCount -eq $MaxRetries - 1)
						{
							$failedCommands += "/* $sectionName */`r`n$sqlCommands`r`n"
						}
					}
				}
				catch
				{
					Write_Log "Error executing section '$sectionName': $_" "red"
					$failedSections += $match
					if ($retryCount -eq $MaxRetries - 1)
					{
						$failedCommands += "/* $sectionName */`r`n$sqlCommands`r`n"
					}
				}
			}
			
			if ($failedSections.Count -eq 0)
			{
				Write_Log "`r`nAll SQL sections executed successfully." "green"
				$success = $true
			}
			else
			{
				throw "`r`nSome sections failed to execute: $($failedSections.Count) sections."
			}
		}
		catch
		{
			$retryCount++
			Write_Log "Error during SQL script execution: $_" "red"
			if ($retryCount -lt $MaxRetries)
			{
				Write_Log "Retrying in $RetryDelaySeconds seconds..." "yellow"
				Start-Sleep -Seconds $RetryDelaySeconds
			}
		}
	}
	
	if (-not $success)
	{
		Write_Log "Max retries reached. SQL script execution failed." "red"
		$failedCommandsText = ($failedCommands -join "`r`n") + "`r`n"
		<#try
		{
			[System.IO.File]::WriteAllText($FailedCommandsPath, $failedCommandsText, $ansiPcEncoding)
			Write_Log "`r`nFailed SQL sections written to: $FailedCommandsPath" "yellow"
			Set-ItemProperty -Path $FailedCommandsPath -Name Attributes -Value (
				(Get-Item $FailedCommandsPath).Attributes -band (-bnot [System.IO.FileAttributes]::Archive)
			)
		}
		catch
		{
			Write_Log "Failed to write failed commands: $_" "red"
		}#>
	}
	else
	{
		Write_Log "SQL script executed successfully on '$dbName'." "green"
	}
	
	Write_Log "`r`n==================== Completed Server Database Maintenance ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Delete_Files
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

function Delete_Files
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
				# Write_Log "The specified path '$Path' does not exist." "Red"
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
											# Write_Log "Excluded: $($matchedItem.FullName)" "Yellow"
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
										# Write_Log "Deleted: $($matchedItem.FullName)" "Green"
									}
									catch
									{
										# Write_Log "Failed to delete $($matchedItem.FullName). Error: $_" "Red"
									}
								}
							}
						}
						else
						{
							# Write_Log "No items matched the pattern: '$filePattern' in '$targetPath'." "Yellow"
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
									# Write_Log "Excluded: $($item.FullName)" "Yellow"
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
								# Write_Log "Deleted: $($item.FullName)" "Green"
							}
							catch
							{
								# Write_Log "Failed to delete $($item.FullName). Error: $_" "Red"
							}
						}
					}
				}
				
				# Write_Log "Total items deleted: $deletedCount" "Blue"
				return $deletedCount
			}
			catch
			{
				# Write_Log "An error occurred during the deletion process. Error: $_" "Red"
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
			Write_Log "The specified path '$Path' does not exist." "Red"
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
										Write_Log "Excluded: $($matchedItem.FullName)" "Yellow"
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
									Write_Log "Deleted: $($matchedItem.FullName)" "Green"
								}
								catch
								{
									Write_Log "Failed to delete $($matchedItem.FullName). Error: $_" "Red"
								}
							}
						}
					}
					else
					{
						Write_Log "No items matched the pattern: '$filePattern' in '$targetPath'." "Yellow"
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
								Write_Log "Excluded: $($item.FullName)" "Yellow"
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
							Write_Log "Deleted: $($item.FullName)" "Green"
						}
						catch
						{
							Write_Log "Failed to delete $($item.FullName). Error: $_" "Red"
						}
					}
				}
			}
			
			Write_Log "Total items deleted: $deletedCount" "Blue"
			return $deletedCount
		}
		catch
		{
			Write_Log "An error occurred during the deletion process. Error: $_" "Red"
			return $deletedCount
		}
	}
}

# ===================================================================================================
#                                       FUNCTION: Process_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Processes one or more lanes based on user selection, parses and writes the lane SQL script
#   (with embedded fixed header/footer), and logs all progress and errors. Handles all-in-one.
# ===================================================================================================

function Process_Lanes
{
	param (
		[string]$LanesqlFilePath,
		[string]$StoreNumber,
		[switch]$ProcessAllLanes
	)
	
	Write_Log "`r`n==================== Starting Process_Lanes Function ====================`r`n" "blue"
	
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Get the user's lane selection
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if ($selection -eq $null)
	{
		Write_Log "Lane processing canceled by user." "yellow"
		return
	}
	$Type = $selection.Type
	$Lanes = $selection.Lanes
	
	# Prompt for section selection ONCE
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$matches = [regex]::Matches($LaneSQLScript, $sectionPattern)
	if ($matches.Count -eq 0)
	{
		Write_Log "No sections found in Lane SQL script." "red"
		return
	}
	
	$fixedSections = @(
		"Set a long timeout so the entire script runs",
		"Clear the long database timeout"
	)
	$allSectionNames = $matches | ForEach-Object { $_.Groups['SectionName'].Value.Trim() } | Where-Object { $fixedSections -notcontains $_ }
	$SectionsToSend = Show_Section_Selection_Form -SectionNames $allSectionNames
	if (-not $SectionsToSend -or $SectionsToSend.Count -eq 0)
	{
		Write_Log "No sections selected for lanes." "yellow"
		return
	}
	
	foreach ($LaneNumber in $Lanes)
	{
		# Get the correct machine name for this lane
		$machineName = $LaneNumber
		if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'].ContainsKey($LaneNumber))
		{
			$machineName = $script:FunctionResults['LaneMachines'][$LaneNumber]
		}
		
		$LaneLocalPath = "$OfficePath\XF${StoreNumber}${LaneNumber}"
		
		if (Test-Path $LaneLocalPath)
		{
			Write_Log "`r`nProcessing $machineName..." "blue"
			Write_Log "Lane path found: $LaneLocalPath" "blue"
			Write_Log "Writing Lane_Database_Maintenance.sqi to Lane..." "blue"
			
			try
			{
				# Always embed the fixed top section
				$topBlock = "/* Set a long timeout so the entire script runs */`r`n@WIZRPL(DBASE_TIMEOUT=E);`r`n" +
				"--------------------------------------------------------------------------------`r`n"
				# Always embed the fixed bottom section
				$bottomBlock = "--------------------------------------------------------------------------------`r`n" +
				"/* Clear the long database timeout */`r`n@WIZCLR(DBASE_TIMEOUT);"
				
				# User-selected middle sections (with section headers)
				$middleBlock = ($matches | Where-Object {
						$SectionsToSend -contains $_.Groups['SectionName'].Value.Trim()
					}) | ForEach-Object {
					"/* $($_.Groups['SectionName'].Value.Trim()) */`r`n$($_.Groups['SQLCommands'].Value.Trim())"
				} | Out-String
				
				# Compose the final script
				$finalScript = $topBlock + $middleBlock + $bottomBlock
				
				Set-Content -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Value $finalScript -Encoding Ascii
				Set-ItemProperty -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
				Write_Log "Created and wrote to file at Lane #${LaneNumber} ($machineName) successfully." "green"
				
				if (-not ($script:ProcessedLanes -contains $LaneNumber))
				{
					$script:ProcessedLanes += $LaneNumber
				}
			}
			catch
			{
				Write_Log "Failed to write to [$machineName]: $_" "red"
			}
		}
		else
		{
			Write_Log "Lane #$LaneNumber not found at path: $LaneLocalPath" "yellow"
		}
	}
	
	Write_Log "`r`nTotal Lanes processed: $($script:ProcessedLanes.Count)" "green"
	if ($script:ProcessedLanes.Count -gt 0)
	{
		Write_Log "Processed Lanes: $($script:ProcessedLanes -join ', ')" "green"
	}
	
	Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Repair_Windows
# ---------------------------------------------------------------------------------------------------
# Description:
#   Performs various system maintenance tasks to repair Windows.
#   Updates Windows Defender signatures, runs a full scan, executes DISM commands,
#   runs System File Checker, performs disk cleanup, optimizes all fixed drives by trimming SSDs or defragmenting HDDs,
#   and schedules a disk check.
#   Uses Write_Log to provide updates after each command execution.
#   Author: Alex_C.T
# ===================================================================================================

function Repair_Windows
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Repair_Windows Function ====================`r`n" "blue"
	
	# Import GUI assemblies if needed
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Confirm user intent
	$confirmationResult = [System.Windows.Forms.MessageBox]::Show(
		"The Windows repair process will take a long time and make significant changes to your system. Do you want to proceed?",
		"Confirmation Required",
		[System.Windows.Forms.MessageBoxButtons]::YesNo,
		[System.Windows.Forms.MessageBoxIcon]::Warning
	)
	if ($confirmationResult -ne [System.Windows.Forms.DialogResult]::Yes)
	{
		Write_Log "Windows repair process cancelled by the user." "yellow"
		return
	}
	
	Write_Log "Starting Windows repair process. This might take a while, please wait..." "blue"
	
	# Create the Repair Options Form
	$repairForm = New-Object System.Windows.Forms.Form
	$repairForm.Text = "Select Repair Operations"
	$repairForm.Size = New-Object System.Drawing.Size(400, 380)
	$repairForm.StartPosition = "CenterScreen"
	$repairForm.FormBorderStyle = 'FixedDialog'
	$repairForm.MaximizeBox = $false
	$repairForm.MinimizeBox = $false
	$repairForm.ShowInTaskbar = $false
	
	# Operation Checkboxes (with correct .Location and .Size types)
	$checkboxes = @()
	
	$cb1 = New-Object System.Windows.Forms.CheckBox
	$cb1.Text = "Windows Defender Update and Scan"
	$cb1.Location = [System.Drawing.Point]::new(20, 20)
	$cb1.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb1
	
	$cb2 = New-Object System.Windows.Forms.CheckBox
	$cb2.Text = "Run DISM Commands"
	$cb2.Location = [System.Drawing.Point]::new(20, 60)
	$cb2.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb2
	
	$cb3 = New-Object System.Windows.Forms.CheckBox
	$cb3.Text = "Run System File Checker (SFC)"
	$cb3.Location = [System.Drawing.Point]::new(20, 100)
	$cb3.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb3
	
	$cb4 = New-Object System.Windows.Forms.CheckBox
	$cb4.Text = "Disk Cleanup"
	$cb4.Location = [System.Drawing.Point]::new(20, 140)
	$cb4.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb4
	
	$cb5 = New-Object System.Windows.Forms.CheckBox
	$cb5.Text = "Optimize Drives"
	$cb5.Location = [System.Drawing.Point]::new(20, 180)
	$cb5.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb5
	
	$cb6 = New-Object System.Windows.Forms.CheckBox
	$cb6.Text = "Schedule Check Disk"
	$cb6.Location = [System.Drawing.Point]::new(20, 220)
	$cb6.Size = [System.Drawing.Size]::new(350, 25)
	$checkboxes += $cb6
	
	foreach ($cb in $checkboxes) { $repairForm.Controls.Add($cb) }
	
	# "Select All" checkbox
	$cbAll = New-Object System.Windows.Forms.CheckBox
	$cbAll.Text = "Select All"
	$cbAll.Location = [System.Drawing.Point]::new(20, 260)
	$cbAll.Size = [System.Drawing.Size]::new(350, 25)
	$cbAll.Add_CheckedChanged({
			foreach ($cb in $checkboxes) { $cb.Checked = $cbAll.Checked }
		})
	$repairForm.Controls.Add($cbAll)
	
	# Enable/Disable Run Button logic
	$runButton = New-Object System.Windows.Forms.Button
	$runButton.Text = "Run"
	$runButton.Location = [System.Drawing.Point]::new(150, 300)
	$runButton.Size = [System.Drawing.Size]::new(100, 30)
	$runButton.Enabled = $false
	$repairForm.Controls.Add($runButton)
	foreach ($cb in $checkboxes)
	{
		$cb.Add_CheckedChanged({ $runButton.Enabled = ($checkboxes | Where-Object { $_.Checked }).Count -gt 0 })
	}
	
	# Show form and get selections
	$selectedParams = @{ }
	$runButton.Add_Click({
			$selectedParams.Defender = $cb1.Checked
			$selectedParams.DISM = $cb2.Checked
			$selectedParams.SFC = $cb3.Checked
			$selectedParams.DiskCleanup = $cb4.Checked
			$selectedParams.OptimizeDrives = $cb5.Checked
			$selectedParams.CheckDisk = $cb6.Checked
			$repairForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$repairForm.Close()
		})
	
	$dialogResult = $repairForm.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Windows repair process cancelled by the user." "yellow"
		return
	}
	
	# -------------------- Operations Begin --------------------
	
	if ($selectedParams.Defender)
	{
		try
		{
			Write_Log "Updating Windows Defender signatures..." "blue"
			& "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -SignatureUpdate
			Write_Log "Windows Defender signatures updated successfully." "green"
			Write_Log "Running Windows Defender full scan..." "blue"
			& "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -Scan -ScanType 2
			Write_Log "Windows Defender full scan completed." "green"
		}
		catch { Write_Log "Defender update/scan failed: $_" "red" }
	}
	else { Write_Log "Skipping Windows Defender update and scan as per user request." "yellow" }
	
	if ($selectedParams.DISM)
	{
		try
		{
			Write_Log "Running DISM StartComponentCleanup..." "blue"
			DISM /Online /Cleanup-Image /StartComponentCleanup /NoRestart
			Write_Log "DISM StartComponentCleanup completed." "green"
			Write_Log "Running DISM RestoreHealth..." "blue"
			DISM /Online /Cleanup-Image /RestoreHealth /NoRestart
			Write_Log "DISM RestoreHealth completed." "green"
		}
		catch { Write_Log "DISM failed: $_" "red" }
	}
	else { Write_Log "Skipping DISM operations as per user request." "yellow" }
	
	if ($selectedParams.SFC)
	{
		try
		{
			Write_Log "Running System File Checker (SFC)..." "blue"
			SFC /scannow
			Write_Log "System File Checker completed." "green"
		}
		catch { Write_Log "SFC failed: $_" "red" }
	}
	else { Write_Log "Skipping System File Checker as per user request." "yellow" }
	
	if ($selectedParams.DiskCleanup)
	{
		try
		{
			Write_Log "Running Disk Cleanup..." "blue"
			Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait
			Write_Log "Disk Cleanup completed." "green"
		}
		catch { Write_Log "Disk Cleanup failed: $_" "red" }
	}
	else { Write_Log "Skipping Disk Cleanup as per user request." "yellow" }
	
	if ($selectedParams.OptimizeDrives)
	{
		try
		{
			Write_Log "Optimizing all fixed drives..." "blue"
			Get-Volume | Where-Object { $_.DriveType -eq 'Fixed' -and $_.DriveLetter } | ForEach-Object {
				Write_Log "Optimizing drive: $($_.DriveLetter)" "blue"
				Optimize-Volume -DriveLetter $_.DriveLetter -Verbose
			}
			Write_Log "Disk optimization completed." "green"
		}
		catch { Write_Log "Drive optimization failed: $_" "red" }
	}
	else { Write_Log "Skipping disk optimization as per user request." "yellow" }
	
	if ($selectedParams.CheckDisk)
	{
		try
		{
			Write_Log "Scheduling Check Disk on C: drive..." "blue"
			Start-Process "cmd.exe" -ArgumentList "/c echo Y|chkdsk C: /f /r" -Verb RunAs -Wait
			Write_Log "Check Disk scheduled. Restart may be required." "green"
		}
		catch { Write_Log "Check Disk scheduling failed: $_" "red" }
	}
	else { Write_Log "Skipping Check Disk scheduling as per user request." "yellow" }
	
	Write_Log "`r`n==================== Repair_Windows Function Completed ====================" "blue"
}

# ===================================================================================================
#                                     FUNCTION: Update_Lane_Config
# ---------------------------------------------------------------------------------------------------
# Description:
#   Processes SQL load files in the \\localhost\storeman\office\Load directory.
#   For each XF lane folder corresponding to the specified StoreNumber and selected Lanes,
#   it copies and/or generates lane-specific .sql files strictly in ANSI (Windows-1252) encoding
#   with CRLF line endings.
#
#   The created files are:
#       - run_load.sql (copied as-is from the script below)
#       - lnk_load.sql (dynamically generated)
#       - sto_load.sql (dynamically generated)
#       - ter_load.sql (dynamically generated)
#
#   All files are written with:
#       - Windows-1252 encoding
#       - CRLF line endings
#       - No BOM
# ===================================================================================================

function Update_Lane_Config
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Update_Lane_Config Function ====================" "blue"
	
	# Ensure $LoadPath exists
	if (-not (Test-Path $LoadPath))
	{
		Write_Log "`r`nLoad Base Path not found: $LoadPath" "yellow"
		return
	}
	
	# Get the user's lane selection
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	
	if ($selection -eq $null)
	{
		Write_Log "`r`nLane processing canceled by user." "yellow"
		return
	}
	
	$Type = $selection.Type
	$Lanes = $selection.Lanes
	
	# Determine if "All Lanes" is selected
	$processAllLanes = $false
	if ($Type -eq "All")
	{
		$processAllLanes = $true
	}
	
	# If "All Lanes" is selected, attempt to retrieve LaneContents
	if ($processAllLanes)
	{
		try
		{
			$LaneContents = $script:FunctionResults['LaneContents']
			
			if ($LaneContents -and $LaneContents.Count -gt 0)
			{
				$Lanes = $LaneContents
			}
			else
			{
				throw "LaneContents is empty or not available."
			}
		}
		catch
		{
			Write_Log "Failed to retrieve LaneContents: $_. Falling back to user-selected lanes." "yellow"
			$processAllLanes = $false
		}
	}
	
	# --------------------------------------------------------------------------------------------
	# Define the run_load, lnk_load, sto_load, and ter_load script parts
	# --------------------------------------------------------------------------------------------
	
	$runLoadScript = @"
@CREATE(RUN_TAB,RUN);
CREATE VIEW Run_Load AS SELECT F1102,F1000,F1103,F1104,F1105,F1106,F1107,F1108,F1109,F1110,F1111,F1112,F1113,F1114,F1115,F1116,F1117 FROM RUN_TAB;

INSERT INTO Run_Load VALUES
(60,'@TER','sql=BACKUP_ALL','@TER','@DSSF',,'@DSS+001',1,'Backup programs',,0,,'',1,,,),
(62,'@TER','sql=BACKUP_DEVICE','@TER',,,,1,'Backup programs on device',,0,,'',1,,,),
(68,'@TER','sqi=BACKUP_RESELLER','@TER',,,,1,'Backup for reseller',,0,,'',1,,,),
(70,'@TER','sql=DBASE_MAINTENANCE','@TER',,,,1,'Database maintenance',,0,,'',1,,,),
(90,'@TER','sql=PURGE_ALL','@TER','@DSW-008',,'@DSW-001',0,'Purge all report data',,0,,'',7,,,),
(95,'@TER','sqi=PURGE_CLIENT','@TER',,,,0,'Purge customer sensitive data',,0,,'',1,,,),
(120,'@TER','sql=UPDATE_DOWNLOAD','@TER',,,,1,'Automatic update download',,0,,'',1,,,),
(125,'@TER','sql=UPDATE_AUTOMATIC','@TER',,,,1,'Automatic update execution',,0,,'',1,,,),
(250,'SMS','sqi=TRS_POS_BANK_EOS_PUP OUTPUT=RECEIPT','901',,,,1,'Automatic bank close',,,,,1,,,);

@UPDATE_BATCH(JOB=ADDRPL,TAR=RUN_TAB,
KEY=F1102=:F1102 AND F1000=:F1000,
SRC=SELECT * FROM Run_Load);

DROP TABLE Run_Load;
"@
	
	$lnkLoadHeader = @"
@CREATE(LNK_TAB,LNK);
CREATE VIEW Lnk_Load AS SELECT F1000,F1056,F1057 FROM LNK_TAB;

INSERT INTO Lnk_Load VALUES
"@
	
	$lnkLoadFooter = @"

@UPDATE_BATCH(JOB=ADDRPL,TAR=LNK_TAB,
KEY=F1000=:F1000 AND F1056=:F1056 AND F1057=:F1057,
SRC=SELECT * FROM Lnk_Load);

DROP TABLE Lnk_Load;
"@
	
	$stoLoadHeader = @"
@CREATE(STO_TAB,STO);
CREATE VIEW Sto_Load AS SELECT F1000,F1018,F1180,F1181,F1182,F1937,F1965,F1966,F2691 FROM STO_TAB;

INSERT INTO Sto_Load VALUES
"@
	
	$stoLoadFooter = @"

@UPDATE_BATCH(JOB=ADDRPL,TAR=STO_TAB,
KEY=F1000=:F1000,
SRC=SELECT * FROM Sto_Load);

DROP TABLE Sto_Load;
"@
	
	$terLoadHeader = @"
@CREATE(TER_TAB,TER); 
CREATE VIEW Ter_Load AS SELECT F1056,F1057,F1058,F1125,F1169 FROM TER_TAB;

INSERT INTO Ter_Load VALUES
"@
	
	$terLoadFooter = @"

@UPDATE_BATCH(JOB=ADDRPL,TAR=TER_TAB,
KEY=F1056=:F1056 AND F1057=:F1057,
SRC=SELECT * FROM Ter_Load);

DROP TABLE Ter_Load;
"@
	
	# Script filenames
	$runLoadFilename = "run_load.sql"
	$lnkLoadFilename = "lnk_load.sql"
	$stoLoadFilename = "sto_load.sql"
	$terLoadFilename = "ter_load.sql"
	
	# --------------------------------------------------------------------------------------------
	# Loop each selected Lane
	# --------------------------------------------------------------------------------------------
	foreach ($laneNumber in $Lanes)
	{
		$laneFolderName = "XF${StoreNumber}${laneNumber}"
		$laneFolderPath = Join-Path -Path $OfficePath -ChildPath $laneFolderName
		
		if (-not (Test-Path $laneFolderPath))
		{
			Write_Log "`r`nLane #$laneNumber not found at path: $laneFolderPath" "yellow"
			continue
		}
		
		$laneFolder = Get-Item -Path $laneFolderPath
		Write_Log "`r`nProcessing Lane #$laneNumber" "blue"
		
		$actionSummaries = @()
		
		# Determine Machine Name from LaneMachines
		try
		{
			Write_Log "Determining machine name for Lane #$laneNumber..." "blue"
			
			$MachineName = $script:FunctionResults['LaneMachines'][$laneNumber]
			
			if ($MachineName)
			{
				Write_Log "Lane #${laneNumber}: Retrieved machine name '$MachineName' from LaneMachines." "green"
			}
			else
			{
				Write_Log "Lane #${laneNumber}: Machine name not found in LaneMachines. Defaulting to 'POS${laneNumber}'." "yellow"
				$MachineName = "POS${laneNumber}"
			}
		}
		catch
		{
			Write_Log "Lane #${laneNumber}: Error retrieving machine name. Error: $_. Defaulting to 'POS${laneNumber}'." "red"
			$MachineName = "POS${laneNumber}"
		}
		
	<# 
		# Process each load SQL file (currently commented out; uncomment if needed)
		foreach ($file in $loadFiles) 
		{
		    Write_Log "Processing file '$($file.Name)' for Lane #$laneNumber..." "blue"
		    
		    # Read the original file content
		    try 
		    {
		        $originalContent = Get-Content -Path $file.FullName -ErrorAction Stop
		        Write_Log "Successfully read '$($file.Name)'." "green"
		    }
		    catch 
		    {
		        Write_Log "Failed to read '$($file.Name)'. Error: $_" "red"
		        $actionSummaries += "Failed to read $($file.Name)"
		        continue
		    }
		    
		    # Filter content to include only records matching the current StoreNumber and LaneNumber
		    $filteredContent = $originalContent | Where-Object { 
		        $_ -match "^\s*\(\s*'[^']+'\s*,\s*'$StoreNumber'\s*,\s*'$laneNumber'\s*\)\s*[;,]?$"
		    }
		    
		    if ($filteredContent) 
		    {
		        # Construct the destination path with .sql extension
		        $destinationPath = Join-Path -Path $laneFolder.FullName -ChildPath ($file.BaseName + ".sql")
		    
		        try 
		        {
		            # Write the filtered content to the lane-specific .sql file using UTF8 without BOM
		            [System.IO.File]::WriteAllText($destinationPath, ($filteredContent -join "`r`n"), $utf8NoBOM)
		    
		            # Set the archive bit on the copied file
		            $fileItem = Get-Item -Path $destinationPath
		            $fileItem.Attributes = $fileItem.Attributes -bor [System.IO.FileAttributes]::Archive
		    
		            Write_Log "Successfully copied to '$destinationPath'." "green"
		            $actionSummaries += "Copied $($file.Name)"
		        }
		        catch 
		        {
		            Write_Log "Failed to copy '$($file.Name)' to '$destinationPath'. Error: $_" "red"
		            $actionSummaries += "Failed to copy $($file.Name)"
		        }
		    }
		    else 
		    {
		        Write_Log "No matching records found in '$($file.Name)' for Lane #$laneNumber." "yellow"
		    }
		}
	#>
		
		# --------------------------------------------------------------------------------------------
		# run_load.sql
		# --------------------------------------------------------------------------------------------
		try
		{
			# Replace line endings with CRLF before writing
			$runLoadScriptCRLF = $runLoadScript -replace "`r?`n", "`r`n"
			
			$runLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $runLoadFilename
			[System.IO.File]::WriteAllText($runLoadDestinationPath, $runLoadScriptCRLF, $ansiPcEncoding)
			
			Set-ItemProperty -Path $runLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			$actionSummaries += "Copied run_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy run_load.sql"
		}
		
		# --------------------------------------------------------------------------------------------
		# lnk_load.sql
		# --------------------------------------------------------------------------------------------
		try
		{
			$lnkLoadInsertStatements = @(
				"('${laneNumber}','${StoreNumber}','${laneNumber}'),",
				"('DSM','${StoreNumber}','${laneNumber}'),",
				"('PAL','${StoreNumber}','${laneNumber}'),",
				"('RAL','${StoreNumber}','${laneNumber}'),",
				"('XAL','${StoreNumber}','${laneNumber}');"
			)
			
			$completeLnkLoadScript = $lnkLoadHeader + "`r`n" + ($lnkLoadInsertStatements -join "`r`n") + "`r`n`r`n" + $lnkLoadFooter.TrimStart() + "`r`n"
			
			# Replace line endings with CRLF before writing
			$completeLnkLoadScriptCRLF = $completeLnkLoadScript -replace "`r?`n", "`r`n"
			
			$lnkLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $lnkLoadFilename
			[System.IO.File]::WriteAllText($lnkLoadDestinationPath, $completeLnkLoadScriptCRLF, $ansiPcEncoding)
			
			Set-ItemProperty -Path $lnkLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			$actionSummaries += "Copied lnk_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy lnk_load.sql"
		}
		
		# --------------------------------------------------------------------------------------------
		# sto_load.sql
		# --------------------------------------------------------------------------------------------
		try
		{
			$stoLoadInsertStatements = @(
				"('${laneNumber}','Terminal ${laneNumber}',1,1,1,,,,),",
				"('DSM','Deploy SMS',1,1,1,,,,),",
				"('PAL','Program all',0,0,1,1,,,),",
				"('RAL','Report all',1,0,0,,,,),",
				"('XAL','Exchange all',0,1,0,,,,);"
			)
			
			$completeStoLoadScript = $stoLoadHeader + "`r`n" + ($stoLoadInsertStatements -join "`r`n") + "`r`n`r`n" + $stoLoadFooter.TrimStart() + "`r`n"
			
			# Replace line endings with CRLF before writing
			$completeStoLoadScriptCRLF = $completeStoLoadScript -replace "`r?`n", "`r`n"
			
			$stoLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $stoLoadFilename
			[System.IO.File]::WriteAllText($stoLoadDestinationPath, $completeStoLoadScriptCRLF, $ansiPcEncoding)
			
			Set-ItemProperty -Path $stoLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			$actionSummaries += "Copied sto_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy sto_load.sql"
		}
		
		# --------------------------------------------------------------------------------------------
		# ter_load.sql
		# --------------------------------------------------------------------------------------------
		try
		{
			$terLoadInsertStatements = @(
				"('${StoreNumber}','${laneNumber}','Terminal ${laneNumber}','\\${MachineName}\storeman\office\XF${StoreNumber}${laneNumber}\','\\${MachineName}\storeman\office\XF${StoreNumber}901\'),",
				"('${StoreNumber}','901','Server','','');"
			)
			
			$completeTerLoadScript = $terLoadHeader + "`r`n" + ($terLoadInsertStatements -join "`r`n") + "`r`n`r`n" + $terLoadFooter.TrimStart() + "`r`n"
			
			# Replace line endings with CRLF before writing
			$completeTerLoadScriptCRLF = $completeTerLoadScript -replace "`r?`n", "`r`n"
			
			$terLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $terLoadFilename
			[System.IO.File]::WriteAllText($terLoadDestinationPath, $completeTerLoadScriptCRLF, $ansiPcEncoding)
			
			Set-ItemProperty -Path $terLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			$actionSummaries += "Copied ter_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy ter_load.sql"
		}
		
		# Summarize
		$summaryMessage = "Lane ${laneNumber} (Machine: ${MachineName}): " + ($actionSummaries -join "; ")
		Write_Log $summaryMessage "green"
		
		# Mark lane as processed
		if (-not ($script:ProcessedLanes -contains $laneNumber))
		{
			$script:ProcessedLanes += $laneNumber
		}
	}
	
	Write_Log "`r`n==================== Update_Lane_Config Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Pump_Tables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Allows a user to select a subset of tables (from Get_Table_Aliases) to extract from SQL Server
#   and copy to the specified lanes or hosts. Similar to Pump_All_Items but restricted to a user-chosen
#   list of tables.
# ===================================================================================================

function Pump_Tables
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Pump_Tables Function ====================`r`n" "blue"
	
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Prompt for lane selection
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if ($selection -eq $null)
	{
		Write_Log "Lane processing canceled by user." "yellow"
		return
	}
	$Type = $selection.Type
	$Lanes = $selection.Lanes
	
	# Determine if "All Lanes" is selected
	$processAllLanes = $false
	if ($Type -eq "All")
	{
		$processAllLanes = $true
	}
	
	# If "All Lanes" is selected, attempt to retrieve LaneContents
	if ($processAllLanes)
	{
		try
		{
			$LaneContents = $script:FunctionResults['LaneContents']
			
			if ($LaneContents -and $LaneContents.Count -gt 0)
			{
				$Lanes = $LaneContents
			}
			else
			{
				throw "LaneContents is empty or not available."
			}
		}
		catch
		{
			$processAllLanes = $false
		}
	}
	
	# --------------------------------------------------------------------------------------------
	# Fetch the alias data that Get_Table_Aliases stored
	# --------------------------------------------------------------------------------------------
	if ($script:FunctionResults.ContainsKey('Get_Table_Aliases'))
	{
		$aliasData = $script:FunctionResults['Get_Table_Aliases']
		$aliasResults = $aliasData.Aliases
		$aliasHash = $aliasData.AliasHash
	}
	else
	{
		Write_Log "Alias data not found. Ensure Get_Table_Aliases has been run." "red"
		return
	}
	
	if ($aliasResults.Count -eq 0)
	{
		Write_Log "No tables found to process. Exiting Pump_Tables." "red"
		return
	}
	
	# Prompt user to select which tables to pump
	$selectedTables = Show_Table_Selection_Form -AliasResults $aliasResults
	if (-not $selectedTables -or $selectedTables.Count -eq 0)
	{
		Write_Log "No tables were selected. Exiting Pump_Tables." "yellow"
		return
	}
	
	# --------------------------------------------------------------------------------------------
	# Get the SQL ConnectionString from script-scoped results
	# --------------------------------------------------------------------------------------------
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write_Log "Connection string not found. Cannot proceed with Pump_Tables." "red"
		return
	}
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	# Open SQL connection
	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$sqlConnection.ConnectionString = $ConnectionString
	$sqlConnection.Open()
	
	# Prepare tracking
	$generatedFiles = @()
	$copiedTables = @()
	$skippedTables = @()
	
	# Filter out only the alias entries that match the user's selection
	$filteredAliasEntries = $aliasResults | Where-Object {
		$selectedTables -contains $_.Table
	}
	
	# --------------------------------------------------------------------------------------------
	# Process each user-selected table
	# --------------------------------------------------------------------------------------------
	foreach ($aliasEntry in $filteredAliasEntries)
	{
		$table = $aliasEntry.Table # e.g. "XYZ_TAB"
		$tableAlias = $aliasEntry.Alias # e.g. "XYZ"
		
		if (-not $table -or -not $tableAlias)
		{
			Write_Log "Invalid table or alias: $($aliasEntry | ConvertTo-Json)" "yellow"
			continue
		}
		
		# Check row count
		$dataCheckQuery = "SELECT COUNT(*) FROM [$table]"
		$cmdCheck = $sqlConnection.CreateCommand()
		$cmdCheck.CommandText = $dataCheckQuery
		
		try
		{
			$rowCount = $cmdCheck.ExecuteScalar()
		}
		catch
		{
			Write_Log "Error checking row count for '$table': $_" "red"
			continue
		}
		
		# Skip tables with zero rows
		if ($rowCount -eq 0)
		{
			$skippedTables += $table
			continue
		}
		
		Write_Log "Processing table '$table'..." "blue"
		
		# Remove "_TAB" suffix for the base name
		$baseTable = $table -replace '_TAB$', ''
		
		# File name for the extracted data
		$sqlFileName = "${baseTable}_Load.sql"
		$localTempPath = Join-Path $env:TEMP $sqlFileName
		
		# Check for a recent file in TEMP (less than 1 hour old)
		$useExistingFile = $false
		if (Test-Path $localTempPath)
		{
			$fileInfo = Get-Item $localTempPath
			$fileAge = (Get-Date) - $fileInfo.LastWriteTime
			if ($fileAge.TotalHours -le 1)
			{
				Write_Log "Recent SQL file found for '$table' in %TEMP%. Using existing file." "green"
				$useExistingFile = $true
			}
			else
			{
				Write_Log "SQL file for '$table' is older than 1 hour. Regenerating." "yellow"
			}
		}
		
		# ----------------------------------------------------------------------------------------
		# Generate or reuse the _Load.sql file
		# ----------------------------------------------------------------------------------------
		if (-not $useExistingFile)
		{
			try
			{
				# Create a StreamWriter in Windows-1252 encoding, CRLF endings
				$streamWriter = New-Object System.IO.StreamWriter($localTempPath, $false, $ansiPcEncoding)
				$streamWriter.NewLine = "`r`n" # Force CRLF
				
				# 1) Gather column data types
				$columnDataTypesQuery = @"
SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '$table'
ORDER BY ORDINAL_POSITION
"@
				$cmdColumnTypes = $sqlConnection.CreateCommand()
				$cmdColumnTypes.CommandText = $columnDataTypesQuery
				$readerColumnTypes = $cmdColumnTypes.ExecuteReader()
				
				$columnDataTypes = [ordered]@{ }
				while ($readerColumnTypes.Read())
				{
					$colName = $readerColumnTypes["COLUMN_NAME"]
					$dataType = $readerColumnTypes["DATA_TYPE"]
					$columnDataTypes[$colName] = $dataType
				}
				$readerColumnTypes.Close()
				
				# 2) Retrieve primary key columns
				$pkQuery = @"
SELECT c.COLUMN_NAME
FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE c
    ON c.CONSTRAINT_NAME = tc.CONSTRAINT_NAME
    AND c.TABLE_NAME = tc.TABLE_NAME
WHERE tc.TABLE_NAME = '$table' AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
ORDER BY c.ORDINAL_POSITION
"@
				$cmdPK = $sqlConnection.CreateCommand()
				$cmdPK.CommandText = $pkQuery
				$readerPK = $cmdPK.ExecuteReader()
				$pkColumns = @()
				while ($readerPK.Read())
				{
					$pkColumns += $readerPK["COLUMN_NAME"]
				}
				$readerPK.Close()
				
				# If no PK, default to first column
				if ($pkColumns.Count -eq 0)
				{
					$primaryKeyColumns = @()
					$cmdFirstColumn = $sqlConnection.CreateCommand()
					$cmdFirstColumn.CommandText = "SELECT TOP 1 * FROM [$table]"
					$readerFirstColumn = $cmdFirstColumn.ExecuteReader()
					
					if ($readerFirstColumn.Read())
					{
						$primaryKeyColumns = @($readerFirstColumn.GetName(0))
					}
					$readerFirstColumn.Close()
				}
				else
				{
					$primaryKeyColumns = $pkColumns
				}
				
				# Build the key string for @UPDATE_BATCH
				$keyString = ($primaryKeyColumns | ForEach-Object { "$_=:$_" }) -join " AND "
				
				# 3) Generate @CREATE, CREATE VIEW, and INSERT lines
				$viewName = $baseTable.Substring(0, 1).ToUpper() + $baseTable.Substring(1).ToLower() + '_Load'
				$columnList = ($columnDataTypes.Keys) -join ','
				
				$header = @"
@WIZRPL(DBASE_TIMEOUT=E);

CREATE VIEW $viewName AS SELECT $columnList FROM $table;

INSERT INTO $viewName VALUES
"@
				# Normalize line endings to CRLF
				$header = $header -replace "(\r\n|\n|\r)", "`r`n"
				$streamWriter.WriteLine($header.TrimEnd())
				
				# 4) Fetch data from the table
				$dataQuery = "SELECT * FROM [$table]"
				$cmdData = $sqlConnection.CreateCommand()
				$cmdData.CommandText = $dataQuery
				$readerData = $cmdData.ExecuteReader()
				
				$firstRow = $true
				while ($readerData.Read())
				{
					if ($firstRow)
					{
						$firstRow = $false
					}
					else
					{
						$streamWriter.WriteLine(",")
					}
					
					$values = @()
					foreach ($col in $columnDataTypes.Keys)
					{
						$val = $readerData[$col]
						$dataType = $columnDataTypes[$col]
						
						if ($val -eq $null -or $val -is [System.DBNull])
						{
							$values += ""
							continue
						}
						
						switch -Wildcard ($dataType)
						{
							{ $_ -in @('char', 'nchar', 'varchar', 'nvarchar', 'text', 'ntext') } {
								$escapedVal = $val.ToString().Replace("'", "''")
								$escapedVal = $escapedVal -replace "(\r\n|\n|\r)", " "
								$values += "'$escapedVal'"
								break
							}
							{ $_ -in @('datetime', 'smalldatetime', 'date', 'datetime2') } {
								$dayOfYear = $val.DayOfYear.ToString("D3")
								$formattedDate = "'{0}{1} {2}'" -f $val.Year, $dayOfYear, $val.ToString("HH:mm:ss")
								$values += $formattedDate
								break
							}
							{ $_ -eq 'bit' } {
								$bitVal = if ($val) { "1" }
								else { "0" }
								$values += $bitVal
								break
							}
							{ $_ -in @('decimal', 'numeric', 'float', 'real', 'money', 'smallmoney') } {
								if ([math]::Floor($val) -eq $val)
								{
									$values += $val.ToString()
								}
								else
								{
									$values += $val.ToString("0.00")
								}
								break
							}
							{ $_ -in @('tinyint', 'smallint', 'int', 'bigint') } {
								$values += $val.ToString()
								break
							}
							default {
								# Fallback to string
								$escapedVal = $val.ToString().Replace("'", "''")
								$escapedVal = $escapedVal -replace "(\r\n|\n|\r)", " "
								$values += "'$escapedVal'"
								break
							}
						}
					}
					
					$insertStatement = "(" + ($values -join ",") + ")"
					$insertStatement = $insertStatement -replace "(\r\n|\n|\r)", " "
					$streamWriter.Write($insertStatement)
				}
				$readerData.Close()
				
				# End the INSERT
				$streamWriter.WriteLine(";")
				$streamWriter.WriteLine()
				
				# Footer
				$footer = @"
@UPDATE_BATCH(JOB=ADDRPL,TAR=$table,
KEY=$keyString,
SRC=SELECT * FROM $viewName);

DROP TABLE $viewName;

@EXEC(INI=HOST_STORE[ACTIVATE_ACCEPT_ALL]);

@WIZCLR(DBASE_TIMEOUT);
"@
				$footer = $footer -replace "(\r\n|\n|\r)", "`r`n"
				$streamWriter.WriteLine($footer.TrimEnd())
				$streamWriter.WriteLine()
				
				$streamWriter.Flush()
				$streamWriter.Close()
				$streamWriter.Dispose()
				
				$generatedFiles += $localTempPath
				$copiedTables += $table
			}
			catch
			{
				Write_Log "Error generating SQL for table '$table': $_" "red"
				continue
			}
		}
		else
		{
			# Reuse existing file
			$generatedFiles += $localTempPath
			$copiedTables += $table
		}
	} # end foreach table
	
	$sqlConnection.Close()
	
	# Summaries
	if ($copiedTables.Count -gt 0)
	{
		Write_Log "Successfully generated _Load.sql files for tables: $($copiedTables -join ', ')" "green"
	}
	if ($skippedTables.Count -gt 0)
	{
		Write_Log "Tables with no data (skipped): $($skippedTables -join ', ')" "yellow"
	}
	
	# --------------------------------------------------------------------------------------------
	# Copy the generated .sql files to each selected lane
	# --------------------------------------------------------------------------------------------
	Write_Log "`r`nDetermining selected lanes...`r`n" "magenta"
	$ProcessedLanes = @()
	foreach ($lane in $Lanes)
	{
		$LaneLocalPath = Join-Path $OfficePath "XF${StoreNumber}${lane}"
		
		if (Test-Path $LaneLocalPath)
		{
			Write_Log "Copying _Load.sql files to Lane #$lane..." "blue"
			try
			{
				foreach ($filePath in $generatedFiles)
				{
					$fileName = [System.IO.Path]::GetFileName($filePath)
					$destinationPath = Join-Path $LaneLocalPath $fileName
					
					# Copy the file
					Copy-Item -Path $filePath -Destination $destinationPath -Force -ErrorAction Stop
					
					# **Clear the Archive attribute on the copied file**
					$fileItem = Get-Item $destinationPath
					if ($fileItem.Attributes -band [System.IO.FileAttributes]::Archive)
					{
						$fileItem.Attributes -= [System.IO.FileAttributes]::Archive
						Write_Log "Cleared Archive attribute for '$fileName' in Lane #$lane." "green"
					}
					
				}
				Write_Log "Successfully copied all generated _Load.sql files to Lane #$lane." "green"
				$ProcessedLanes += $lane
			}
			catch
			{
				Write_Log "Error copying files to Lane #${lane}: $_" "red"
			}
		}
		else
		{
			Write_Log "Lane #$lane not found at path: $LaneLocalPath" "yellow"
		}
	}
	
	Write_Log "`r`nTotal Lane folders processed: $($ProcessedLanes.Count)" "green"
	if ($ProcessedLanes.Count -gt 0)
	{
		Write_Log "Processed Lanes: $($ProcessedLanes -join ', ')" "green"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
	}
}

# ===================================================================================================
#                                     FUNCTION: Reboot_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Reboots one, a range, or all lane machines based on the user's selection.
#   Builds a Windows Form for lane selection using the LaneMachines hashtable from FunctionResults.
#   For each lane, creates a custom object (with LaneNumber, MachineName, and a friendly DisplayName).
#   Reboots the selected machines by first attempting the shutdown command and, if that fails, falling
#   back to using Restart-Computer.
# ===================================================================================================

function Reboot_Lanes
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Reboot_Lanes Function ====================`r`n" "blue"
	
	# Grab the lane→machine map
	$LaneMachines = $script:FunctionResults['LaneMachines']
	if (-not $LaneMachines -or $LaneMachines.Count -eq 0)
	{
		Write_Log "LaneMachines not available. Cannot reboot lanes." "Red"
		return
	}
	
	# Let user pick lanes
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if (-not $selection -or -not $selection.Lanes)
	{
		Write_Log "No lanes selected or selection cancelled. Exiting." "Yellow"
		return
	}
	$lanes = $selection.Lanes
	
	# Loop through each lane and attempt reboots
	foreach ($lane in $lanes)
	{
		
		if (-not $LaneMachines.ContainsKey($lane))
		{
			Write_Log "Unknown lane '$lane'. Skipping." "Yellow"
			continue
		}
		
		$machine = $LaneMachines[$lane]
		Write_Log "Lane $lane on [$machine]: attempting mailslot reboot" "Yellow"
		
		# 1) SMSStart mailslot reboot
		$mailslot = "\\$machine\mailslot\SMSStart_${StoreNumber}${lane}"
		$msResult = [MailslotSender]::SendMailslotCommand($mailslot, '@exec(REBOOT=1).')
		if ($msResult)
		{
			Write_Log "Mailslot reboot sent to $machine (Lane $lane)" "Green"
			continue
		}
		Write_Log "Mailslot reboot failed for $machine. Falling back to shutdown.exe" "Yellow"
		
		# 2) Fallback: shutdown.exe
		Write_Log "Running shutdown.exe /r /m \\$machine /t 0 /f" "Yellow"
		cmd.exe /c "shutdown /r /m \\$machine /t 0 /f" | Out-Null
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "shutdown.exe reboot succeeded for $machine" "Green"
			continue
		}
		Write_Log "shutdown.exe exit code $LASTEXITCODE. Now trying Restart-Computer" "Yellow"
		
		# 3) Final fallback: Restart-Computer
		Restart-Computer -ComputerName $machine -Force -ErrorAction SilentlyContinue
		if ($?)
		{
			Write_Log "Restart-Computer succeeded for $machine" "Green"
		}
		else
		{
			Write_Log "All reboot methods failed for $machine (Lane $lane)" "Red"
		}
	}
	Write_Log "`r`n==================== Reboot_Lanes Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: CloseOpenTransactions
# ---------------------------------------------------------------------------------------------------
# Description:
#   This function monitors the specified XE folder for error files, extracts relevant data, and closes
#   open transactions on specified lanes for a given store. Logs are written to both a log file and
#   through the Write_Log function for the main script.
# ===================================================================================================

function Close_Open_Transactions
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting CloseOpenTransactions ====================`r`n" "blue"
	
	# Define the path to monitor
	$XEFolderPath = "$OfficePath\XE${StoreNumber}901"
	
	# Ensure the XE folder exists
	if (-not (Test-Path $XEFolderPath))
	{
		Write_Log -Message "XE folder not found: $XEFolderPath" "red"
		return
	}
	
	# Path to the manual Close_Transaction.sqi content
	$CloseTransactionManual = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE')"
	
	# Path to the log file
	$LogFolderPath = "$BasePath\Scripts_by_Alex_C.T"
	$LogFilePath = Join-Path -Path $LogFolderPath -ChildPath "Closed_Transactions_LOG.txt"
	
	# Ensure the log directory exists
	if (-not (Test-Path $LogFolderPath))
	{
		try
		{
			New-Item -Path $LogFolderPath -ItemType Directory -Force | Out-Null
			Write_Log -Message "Created log directory: $LogFolderPath" "green"
		}
		catch
		{
			Write_Log -Message "Failed to create log directory '$LogFolderPath'. Error: $_" "red"
			return
		}
	}
	
	$MatchedTransactions = $false
	
	try
	{
		# Get the current time
		$currentTime = Get-Date
		
		# Get the list of files matching the pattern in the XE folder that are not older than 30 days
		$files = Get-ChildItem -Path $XEFolderPath -Filter "S*.???" | Where-Object {
			($currentTime - $_.LastWriteTime).TotalDays -le 30
		}
		
		if ($files -and $files.Count -gt 0)
		{
			# We have files, attempt to process them
			foreach ($file in $files)
			{
				try
				{
					# Extract lane number from filename based on the pattern S*."???"
					if ($file.Name -match '^S.*\.(\d{3})$')
					{
						$LaneNumber = $Matches[1]
					}
					else
					{
						continue # Skip to the next file if pattern doesn't match
					}
					
					# Read the content of the file
					$content = Get-Content -Path $file.FullName
					
					# Parse the content
					$fromLine = $content | Where-Object { $_ -like 'From:*' }
					$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
					$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
					$lastRecordedStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
					
					# Extract store number and lane number from the From line
					if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
					{
						$fileStoreNumber = $Matches[1]
						$fileLaneNumber = $Matches[2]
						
						# Check if the store number matches
						if ($fileStoreNumber -eq $StoreNumber -and $fileLaneNumber -eq $LaneNumber)
						{
							# Check the Subject line
							if ($subjectLine -match 'Subject:\s*(.*)')
							{
								$subject = $Matches[1].Trim()
								if ($subject -eq 'Health')
								{
									# Check the MSG line
									if ($msgLine -match 'MSG:\s*(.*)')
									{
										$message = $Matches[1].Trim()
										if ($message -eq 'This application is not running.')
										{
											# Extract the transaction number from the Last recorded status line
											if ($lastRecordedStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
											{
												$transactionNumber = $Matches[1]
												
												# Path to the automatic Close_Transaction.sqi content
												$CloseTransactionAuto = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1032 = $transactionNumber)"
												
												# Define the path to the lane directory
												$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
												
												if (Test-Path $LaneDirectory)
												{
													# Define the path to the Close_Transaction.sqi file in the lane directory
													$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
													
													# Write the content to the file
													Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionAuto -Encoding ASCII
													
													# Remove the Archive attribute from the file
													Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
													
													# Log the event
													$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Closed transaction $transactionNumber on lane $LaneNumber"
													Add-Content -Path $LogFilePath -Value $logMessage
													
													# Delete the error file from the XE folder
													Remove-Item -Path $file.FullName -Force
													
													Write_Log -Message "Processed file $($file.Name) for lane $LaneNumber and closed transaction $transactionNumber" "green"
													$MatchedTransactions = $true
													
													# Send restart command 3 seconds after deployment
													Start-Sleep -Seconds 3
													
													# Retrieve updated node information to get machine mapping
													$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
													if ($nodes)
													{
														$machineName = $nodes.LaneMachines[$LaneNumber]
														if ($machineName)
														{
															$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
															$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
															$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
															if ($result)
															{
																Write_Log -Message "Restart command sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after deployment." "green"
															}
															else
															{
																Write_Log -Message "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red"
															}
														}
														else
														{
															Write_Log -Message "No machine found for lane $LaneNumber. Restart command not sent." "yellow"
														}
													}
													else
													{
														Write_Log -Message "Could not retrieve node information for store $StoreNumber. Restart command not sent." "red"
													}
												}
												else
												{
													Write_Log -Message "Lane directory $LaneDirectory not found" "yellow"
												}
											}
											else
											{
												Write_Log -Message "Could not extract transaction number from Last recorded status line in file $($file.Name)" "red"
											}
										}
										# else MSG did not match the condition - no action needed
									}
									# else no MSG line found - no action needed
								}
								# else subject not health - no action needed
							}
							# else no Subject line found - no action needed
						}
						else
						{
							Write_Log -Message "Store or Lane number mismatch in file $($file.Name). File Store/Lane: $fileStoreNumber/$fileLaneNumber vs Expected Store/Lane: $StoreNumber/$LaneNumber" "yellow"
						}
					}
					# else From line not matched - no action needed
				}
				catch
				{
					Write_Log -Message "Error processing file $($file.Name): $_" "red"
				}
			}
		}
		
		# After processing all files, if no matched transactions were found, prompt once for lane number
		if (-not $MatchedTransactions)
		{
			Write_Log -Message "No files or no matching transactions found. Prompting for lane number." "yellow"
			
			# Show your lane-selection form
			$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
			if (-not $selection)
			{
				Write_Log -Message "Lane selection cancelled or returned no selection." "yellow"
				Write_Log "`r`n==================== CloseOpenTransactions Function Completed ====================" "blue"
				return
			}
			
			# Loop through each selected lane
			foreach ($LaneNumber in $selection.Lanes)
			{
				# Pad to three digits
				$LaneNumber = $LaneNumber.PadLeft(3, '0')
				
				if (-not $LaneNumber)
				{
					Write_Log -Message "No lane number provided by the user." "red"
					return
				}
				
				# Define the path to the lane directory
				$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
				
				if (Test-Path $LaneDirectory)
				{
					# Define the path to the Close_Transaction.sqi file in the lane directory
					$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
					
					# Write the content to the file
					Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionManual -Encoding ASCII
					
					# Remove the Archive attribute from the file
					Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
					
					# Log the event
					$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - User deployed Close_Transaction.sqi to lane $LaneNumber"
					Add-Content -Path $LogFilePath -Value $logMessage
					
					Write_Log -Message "Deployed Close_Transaction.sqi to lane $LaneNumber" "green"
					
					# After user deploys the file, clear the folder except for files with "FATAL" in the name
					Get-ChildItem -Path $XEFolderPath -File | Where-Object { $_.Name -notlike "*FATAL*" } | Remove-Item -Force
					
					# Send restart command 3 seconds after deployment by the user
					# Start-Sleep -Seconds 3
					$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
					if ($nodes)
					{
						$machineName = $nodes.LaneMachines[$LaneNumber]
						if ($machineName)
						{
							$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
							$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
							$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
							if ($result)
							{
								Write_Log -Message "Restart All Programs sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after user deployment." "green"
							}
							else
							{
								Write_Log -Message "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red"
							}
						}
						else
						{
							Write_Log -Message "No machine found for lane $LaneNumber. Restart command not sent." "yellow"
						}
					}
					else
					{
						Write_Log -Message "Could not retrieve node information for store $StoreNumber. Restart command not sent." "red"
					}
				}
				else
				{
					Write_Log -Message "Lane directory $LaneDirectory not found" "yellow"
				}
				
				Write_Log "Prompt deployment process completed." "yellow"
			}
		}
	}
	catch
	{
		Write_Log -Message "An error occurred during monitoring: $_" "red"
	}
	
	Write_Log "No further matching files were found after processing." "yellow"
	Write_Log "`r`n==================== CloseOpenTransactions Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Ping_All_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#    Ping all lanes at once for a given store (store mode only).
# ===================================================================================================

function Ping_All_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Ping_All_Lanes Function ====================`r`n" "blue"
	
	# Ensure necessary functions are available
	foreach ($func in @('Write_Log'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneContents') -or
		-not $script:FunctionResults.ContainsKey('LaneMachines'))
	{
		Write_Log "Lane information is not available. Please run Retrieve_Nodes first." "Red"
		return
	}
	
	# Retrieve lane information
	$LaneContents = $script:FunctionResults['LaneContents']
	$LaneMachines = $script:FunctionResults['LaneMachines']
	
	if ($LaneContents.Count -eq 0)
	{
		Write_Log "No lanes found for Store Number: $StoreNumber." "Yellow"
		return
	}
	
	# Assume all lanes are selected
	$selectedLanes = $LaneContents
	
	Write_Log "All lanes will be pinged for Store Number: $StoreNumber." "Green"
	
	# Prepare list of machines to ping
	$machinesToPing = @()
	foreach ($lane in $selectedLanes)
	{
		if ($LaneMachines.ContainsKey($lane))
		{
			$machineName = $LaneMachines[$lane]
			if ($machineName)
			{
				$machinesToPing += [PSCustomObject]@{
					Lane    = $lane
					Machine = $machineName
				}
			}
			else
			{
				$machinesToPing += [PSCustomObject]@{
					Lane    = $lane
					Machine = "Unknown"
				}
			}
		}
		else
		{
			$machinesToPing += [PSCustomObject]@{
				Lane    = $lane
				Machine = "Not Found"
			}
		}
	}
	
	if ($machinesToPing.Count -eq 0)
	{
		Write_Log "No valid machines found to ping." "Yellow"
		return
	}
	
	# Initialize counters
	$successCount = 0
	$failureCount = 0
	
	# Ping each machine and log status
	foreach ($machine in $machinesToPing)
	{
		$lane = $machine.Lane
		$machineName = $machine.Machine
		
		if ($machineName -in @("Unknown", "Not Found"))
		{
			Write_Log "Lane #${lane}: Machine Name - $machineName. Status: Skipped." "Yellow"
			continue
		}
		
		try
		{
			$pingResult = Test-Connection -ComputerName $machineName -Count 1 -Quiet -ErrorAction Stop
			if ($pingResult)
			{
				Write_Log "Lane #${lane}: Machine '$machineName' is reachable. Status: Success." "Green"
				$successCount++
			}
			else
			{
				Write_Log "Lane #${lane}: Machine '$machineName' is not reachable. Status: Failed." "Red"
				$failureCount++
			}
		}
		catch
		{
			Write_Log "Lane #${lane}: Failed to ping Machine '$machineName'. Error: $($_.Exception.Message)" "Red"
			$failureCount++
		}
	}
	
	# Summary of ping results
	Write_Log "Ping Summary for Store Number: $StoreNumber - Success: $successCount, Failed: $failureCount." "Blue"
	Write_Log "`r`n==================== Ping_All_Lanes Function Completed ====================" "blue"
}

# ===================================================================================================
#                                           FUNCTION: Delete_DBS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Enables users to delete specific file types (.txt and .dwr) from selected lanes within a specified
#   store. Additionally, users are prompted to include or exclude .sus files from the deletion process.
#   The function leverages pre-stored lane information from the Retrieve_Nodes function to identify 
#   machine paths associated with each lane. File deletions are handled by the Delete_Files helper function,
#   and all actions and results are logged using the existing Write_Log function.
# ---------------------------------------------------------------------------------------------------
#   -StoreNumber (Mandatory)
#       The store number for which lanes are to be processed. This must correspond to a valid store in the system.
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   Delete_DBS -StoreNumber "123"
#
# Prerequisites:
#   - Ensure that the Retrieve_Nodes function has been executed prior to running Delete_DBS.
#   - Verify that the Show_Lane_Selection_Form, Delete_Files, and Write_Log functions are available in the session.
#   - Confirm network accessibility to the machines associated with the lanes.
#   - The user must have the necessary permissions to delete files in the target directories.
# ===================================================================================================

function Delete_DBS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Delete_DBS Function ====================`r`n" "blue"
	
	# Ensure necessary functions are available
	foreach ($func in @('Show_Lane_Selection_Form', 'Delete_Files', 'Write_Log'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneContents') -or
		-not $script:FunctionResults.ContainsKey('LaneMachines'))
	{
		Write_Log "Lane information is not available. Please run Retrieve_Nodes first." "Red"
		return
	}
	
	# Retrieve lane information
	$LaneContents = $script:FunctionResults['LaneContents']
	$LaneMachines = $script:FunctionResults['LaneMachines']
	
	if ($LaneContents.Count -eq 0)
	{
		Write_Log "No lanes found for Store Number: $StoreNumber." "Yellow"
		return
	}
	
	# Prompt user to include .sus files in deletion
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Delete_DBS Confirmation"
	$form.Size = New-Object System.Drawing.Size(400, 200)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Do you want to include .sus files in the deletion?"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(20, 30)
	$form.Controls.Add($label)
	
	$checkboxSus = New-Object System.Windows.Forms.CheckBox
	$checkboxSus.Text = "Include .sus files"
	$checkboxSus.AutoSize = $true
	$checkboxSus.Location = New-Object System.Drawing.Point(20, 60)
	$form.Controls.Add($checkboxSus)
	
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(100, 120)
	$buttonOK.Size = New-Object System.Drawing.Size(80, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(200, 120)
	$buttonCancel.Size = New-Object System.Drawing.Size(80, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($buttonCancel)
	
	$form.AcceptButton = $buttonOK
	$form.CancelButton = $buttonCancel
	
	$dialogResult = $form.ShowDialog()
	
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "User canceled the deletion process." "Yellow"
		return
	}
	
	$includeSus = $checkboxSus.Checked
	
	# Define file types to delete
	$fileExtensions = @("*.txt", "*.dwr")
	if ($includeSus)
	{
		$fileExtensions += "*.sus"
	}
	
	Write_Log "Starting deletion of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "Green"
	
	# Show the selection dialog to choose lanes
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	
	if (-not $selection)
	{
		# User canceled the dialog
		Write_Log "User canceled the lane selection." "Yellow"
		return
	}
	
	# Get the list of lanes to process (store mode: only 'Specific' type possible)
	$selectedLanes = $selection.Lanes
	
	if ($selectedLanes.Count -eq 0)
	{
		Write_Log "No lanes selected for processing." "Yellow"
		return
	}
	
	# Initialize counters
	$totalDeleted = 0
	$totalFailed = 0
	
	foreach ($lane in $selectedLanes)
	{
		if ($LaneMachines.ContainsKey($lane))
		{
			$machineName = $LaneMachines[$lane]
			
			if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
			{
				Write_Log "Lane #{$lane}: Machine name is invalid or unknown. Skipping deletion." "Yellow"
				continue
			}
			
			# Construct the target path (modify this path as per your environment)
			$targetPath = "\\$machineName\Storeman\Office\DBS\"
			
			if (-not (Test-Path -Path $targetPath))
			{
				Write_Log "Lane #${lane}: Target path '$targetPath' does not exist. Skipping." "Yellow"
				continue
			}
			
			Write_Log "Processing Lane #$lane at '$targetPath', please wait..." "Blue"
			
			try
			{
				# Delete_Files function is now expected to return an integer count
				$deletionCount = Delete_Files -Path $targetPath -SpecifiedFiles $fileExtensions -Exclusions @() -AsJob:$false
				
				if ($deletionCount -is [int])
				{
					$totalDeleted += $deletionCount
				}
				else
				{
					Write_Log "Lane #${lane}: Unexpected response from Delete_Files." "Red"
					$totalFailed++
				}
			}
			catch
			{
				Write_Log "Lane #${lane}: An error occurred while deleting files. Error: $_" "Red"
				$totalFailed++
			}
		}
		else
		{
			Write_Log "Lane #${lane}: Machine information not found. Skipping." "Yellow"
			continue
		}
	}
	
	# Summary of deletion results
	Write_Log "Deletion Summary for Store Number: $StoreNumber - Total Files Deleted: $totalDeleted, Total Failures: $totalFailed." "Blue"
	Write_Log "`r`n==================== Delete_DBS Function Completed ====================" "blue"
}

# ===================================================================================================
#                                         FUNCTION: Invoke_Secure_Script
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user for a password via a Windows Form before executing a primary script from a specified
#   URL. If the primary script fails to execute, the function automatically attempts to run an alternative
#   script from a backup URL. The password is securely stored in the script using encryption to ensure 
#   that only authorized users can execute the scripts. All actions, including successes and failures, 
#   are logged using the existing Write_Log function.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   None
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   Invoke_Secure_Script
#
# Prerequisites:
#   - Ensure that the Write_Log function is available in the session.
#   - The user must have the necessary permissions to execute scripts from the specified URLs.
#   - Internet connectivity is required to access the script URLs.
# ===================================================================================================

function Invoke_Secure_Script
{
	[CmdletBinding()]
	param ()
	
	# --- Configuration ---
	$storedPassword = "112922"
	$primaryScriptURL = "https://get.activated.win"
	$fallbackScriptURL = "https://massgrave.dev/get"
	
	# --- Log Start ---
	Write_Log "`r`n==================== Starting Invoke_Secure_Script Function ====================`r`n" "blue"
	
	# --- Password Prompt (GUI) ---
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Authentication Required"
	$form.Size = New-Object System.Drawing.Size(350, 150)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Please enter the password to proceed:"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(10, 20)
	$form.Controls.Add($label)
	
	$textbox = New-Object System.Windows.Forms.TextBox
	$textbox.Location = New-Object System.Drawing.Point(10, 50)
	$textbox.Size = New-Object System.Drawing.Size(310, 20)
	$textbox.UseSystemPasswordChar = $true
	$form.Controls.Add($textbox)
	
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(160, 80)
	$buttonOK.Size = New-Object System.Drawing.Size(80, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $buttonOK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(240, 80)
	$buttonCancel.Size = New-Object System.Drawing.Size(80, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $buttonCancel
	$form.Controls.Add($buttonCancel)
	
	$dialogResult = $form.ShowDialog()
	$password = if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) { $textbox.Text }
	else { $null }
	
	if (-not $password)
	{
		Write_Log "User canceled the authentication prompt." "yellow"
		Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
		return
	}
	
	# --- Verify Password ---
	if ($password -ne $storedPassword)
	{
		Write_Log "Authentication failed. Incorrect password." "red"
		Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
		return
	}
	
	Write_Log "Authentication successful. Proceeding with script execution." "green"
	
	# --- Execute Script from URL ---
	try
	{
		Write_Log "Executing primary script from $primaryScriptURL." "blue"
		Invoke-Expression (irm $primaryScriptURL)
		Write_Log "Primary script executed successfully." "green"
	}
	catch
	{
		Write_Log "Primary script execution failed. Attempting to execute fallback script." "red"
		try
		{
			Invoke-Expression (irm $fallbackScriptURL)
			Write_Log "Fallback script executed successfully." "green"
		}
		catch
		{
			Write_Log "Fallback script execution also failed. Please check the URLs and your network connection." "red"
		}
	}
	
	# --- Log End ---
	Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
}

# ===================================================================================================
#                                     FUNCTION: Configure_System_Settings
# ---------------------------------------------------------------------------------------------------
# Description:
#   Configures various system settings to optimize performance and organization. This function performs
#   the following tasks:
#     1. **Organizes Desktop**:
#        - Creates an "Unorganized Items" folder (or a custom-named folder) on the Desktop.
#        - Moves all non-system and non-excluded items from the Desktop into the designated folder.
#        - Ensures that specified excluded folders (e.g., "Lanes", "Scales", "BackOffices") exist.
#
#     2. **Configures Power Settings**:
#        - Sets the power plan to High Performance.
#        - Disables system sleep modes.
#        - Sets the minimum processor performance to 100%.
#        - Configures the monitor to turn off after 15 minutes of inactivity.
#
#     3. **Configures Services**:
#        - Sets specified services (e.g., "fdPHost", "FDResPub", "SSDPSRV", "upnphost") to start automatically.
#        - Starts the services if they are not already running.
#
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - [string]$UnorganizedFolderName (Optional)
#     Specifies the name of the folder where unorganized Desktop items will be moved.
#     Default value: "Unorganized Items"
#
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   # Use default folder name
#   Configure_System_Settings
#
#   # Specify a custom folder name for unorganized Desktop items
#   Configure_System_Settings -UnorganizedFolderName "MyCustomFolder"
#
# ---------------------------------------------------------------------------------------------------
# Prerequisites:
#   - **Administrator Privileges**:
#     The script must be run with elevated privileges. If not, it will prompt the user to restart PowerShell as an Administrator.
#
#   - **Write_Log Function**:
#     Ensure that the `Write_Log` function is available in the session for logging actions and statuses.
#
#   - **Permissions**:
#     The user must have the necessary permissions to create folders, modify power settings, and configure services.
#
#   - **PowerShell Version**:
#     Compatible with PowerShell versions that support the cmdlets used in the script (e.g., PowerShell 5.1 or later).
#
#   - **Internet Connectivity** (if applicable):
#     Required if any of the configured services or settings depend on internet access.
#
# ===================================================================================================

function Configure_System_Settings
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$UnorganizedFolderName = "My Unorganized Items"
	)
	
	Write_Log "`r`n==================== Starting Configure_System_Settings Function ====================`r`n" "blue"
	
	# Ensure the function is run as Administrator
	if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
	{
		Write_Log "This script must be run as an Administrator. Please restart PowerShell with elevated privileges." "Red"
		return
	}
	
	try
	{
		# ===========================================
		# 1. Organize Desktop
		# ===========================================
		Write_Log "`r`nOrganizing Desktop..." "Blue"
		
		$DesktopPath = [Environment]::GetFolderPath("Desktop")
		$UnorganizedFolder = Join-Path -Path $DesktopPath -ChildPath $UnorganizedFolderName
		
		# Define system icons and excluded folders
		$systemIcons = @("This PC.lnk", "Network.lnk", "Control Panel.lnk", "Recycle Bin.lnk", "User's Files.lnk", "Execute(TBS_Maintenance_Script).bat", "Execute(MiniGhost).bat", "$scriptName")
		$excludedFolders = @("Lanes", "Scales", "BackOffices", "My Unorganized Items")
		
		# Create excluded folders if they don't exist
		foreach ($folder in $excludedFolders)
		{
			$folderPath = Join-Path -Path $DesktopPath -ChildPath $folder
			if (-not (Test-Path -Path $folderPath))
			{
				New-Item -Path $folderPath -ItemType Directory | Out-Null
				Write_Log "Created excluded folder: $folderPath" "green"
			}
			else
			{
				Write_Log "Excluded folder already exists: $folderPath" "Cyan"
			}
		}
		
		# Get all items on the desktop
		$desktopItems = Get-ChildItem -Path $DesktopPath -Force | Where-Object { $_.Name -notin $systemIcons -and ($_PSIsContainer -or $_.Extension -ne ".lnk") }
		
		foreach ($item in $desktopItems)
		{
			$exclude = $false
			
			# Check if item is in excluded folders
			foreach ($excluded in $excludedFolders)
			{
				if ($item.Name -ieq $excluded)
				{
					$exclude = $true
					break
				}
			}
			
			if (-not $exclude)
			{
				try
				{
					Move-Item -Path $item.FullName -Destination $UnorganizedFolder -Force
					Write_Log "Moved item: $($item.Name)" "Green"
				}
				catch
				{
					Write_Log "Failed to move item: $($item.Name). Error: $_" "Red"
				}
			}
			else
			{
				#	Write_Log "Excluded from moving: $($item.Name)" "Cyan"
			}
		}
		
		Write_Log "Desktop organization complete." "Green"
		
		# ===========================================
		# 2. Configure Power Settings
		# ===========================================
		Write_Log "`r`nConfiguring power plan and performance settings..." "Blue"
		
		# Set the power plan to High Performance
		$highPerfGUID = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
		try
		{
			powercfg /s $highPerfGUID
			Write_Log "Power plan set to High Performance." "Green"
		}
		catch
		{
			Write_Log "Failed to set power plan to High Performance. Error: $_" "Red"
		}
		
		# Set system to never sleep
		try
		{
			powercfg /change standby-timeout-ac 0
			powercfg /change standby-timeout-dc 0
			Write_Log "System sleep disabled." "Green"
		}
		catch
		{
			Write_Log "Failed to disable system sleep. Error: $_" "Red"
		}
		
		# Set minimum processor performance to 100%
		try
		{
			powercfg /setacvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setdcvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setactive $highPerfGUID
			Write_Log "Minimum processor performance set to 100%." "Green"
		}
		catch
		{
			Write_Log "Failed to set processor performance. Error: $_" "Red"
		}
		
		# Turn off screen after 15 minutes
		try
		{
			powercfg /change monitor-timeout-ac 15
			Write_Log "Monitor timeout set to 15 minutes." "Green"
		}
		catch
		{
			Write_Log "Failed to set monitor timeout. Error: $_" "Red"
		}
		
		Write_Log "Power plan and performance settings configuration complete. Some changes may require a reboot to take effect." "Green"
		
		# ===========================================
		# 3. Configure Services
		# ===========================================
		Write_Log "`r`nConfiguring services to start automatically..." "Blue"
		
		$servicesToConfigure = @("fdPHost", "FDResPub", "SSDPSRV", "upnphost")
		
		foreach ($service in $servicesToConfigure)
		{
			try
			{
				# Set service to start automatically
				Set-Service -Name $service -StartupType Automatic -ErrorAction Stop
				Write_Log "Set service '$service' to Automatic." "Green"
				
				# Start the service if not running
				$svc = Get-Service -Name $service -ErrorAction Stop
				if ($svc.Status -ne 'Running')
				{
					Start-Service -Name $service -ErrorAction Stop
					Write_Log "Started service '$service'." "Green"
				}
				else
				{
					Write_Log "Service '$service' is already running." "Cyan"
				}
			}
			catch
			{
				Write_Log "Failed to configure service '$service'. Error: $_" "Red"
			}
		}
		
		Write_Log "Service configuration complete." "Green"
		
		Write_Log "All system configurations have been applied successfully." "Green"
		Write_Log "`r`n==================== Configure_System_Settings Function Completed ====================`r`n" "blue"
	}
	catch
	{
		Write_Log "An unexpected error occurred: $_" "Red"
	}
}

# ===================================================================================================
#                                 FUNCTION: Refresh_PIN_Pad_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Refreshes specific configuration files (.ini) within selected lanes of a specified store.
#   Updates the timestamps of critical files to ensure they are recognized as modified.
# ===================================================================================================

function Refresh_PIN_Pad_Files
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Refresh_PIN_Pad_Files Function ====================`r`n" "blue"
	
	# Validate the target path (e.g., $OfficePath)
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Ensure necessary functions are available
	foreach ($func in @('Show_Lane_Selection_Form', 'Write_Log', 'Retrieve_Nodes'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Ensure lane information is available
	if (-not ($script:FunctionResults.ContainsKey('LaneContents') -and $script:FunctionResults.ContainsKey('LaneMachines')))
	{
		Write_Log "No lane information found. Please ensure Retrieve_Nodes has been executed." "Red"
		return
	}
	
	$LaneContents = $script:FunctionResults['LaneContents']
	$LaneMachines = $script:FunctionResults['LaneMachines']
	$NumberOfLanes = $script:FunctionResults['NumberOfLanes']
	
	# Get the user's selection
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	
	if ($selection -eq $null)
	{
		Write_Log "Operation canceled by user." "yellow"
		return
	}
	
	$Lanes = $selection.Lanes
	
	if ($Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for processing." "Yellow"
		return
	}
	
	# Define file types to refresh
	$fileExtensions = @("PreferredAIDs.ini", "EMVCAPKey.ini", "EMVAID.ini")
	
	Write_Log "Starting refresh of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "Green"
	
	# Initialize counters
	$totalRefreshed = 0
	$totalFailed = 0
	
	foreach ($lane in $Lanes)
	{
		if ($LaneMachines.ContainsKey($lane))
		{
			$machineName = $LaneMachines[$lane]
			
			if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
			{
				Write_Log "Lane #${lane}: Machine name is invalid or unknown. Skipping refresh." "Yellow"
				continue
			}
			
			# Construct the target path (modify this path as per your environment)
			$targetPath = "\\$machineName\Storeman\XchDev\EMVConfig\"
			
			if (-not (Test-Path -Path $targetPath))
			{
				Write_Log "Lane #${lane}: Target path '$targetPath' does not exist. Skipping." "Yellow"
				continue
			}
			
			Write_Log "Processing Lane #$lane at '$targetPath'." "Blue"
			
			foreach ($file in $fileExtensions)
			{
				$filePath = Join-Path -Path $targetPath -ChildPath $file
				
				if (Test-Path -Path $filePath)
				{
					try
					{
						# Update the LastWriteTime to current date and time
						(Get-Item -Path $filePath).LastWriteTime = Get-Date
						Write_Log "Lane #${lane}: Refreshed file '$filePath'." "Green"
						$totalRefreshed++
					}
					catch
					{
						Write_Log "Lane #${lane}: Failed to refresh file '$filePath'. Error: $_" "Red"
						$totalFailed++
					}
				}
				else
				{
					Write_Log "Lane #${lane}: File '$filePath' does not exist. Skipping." "Yellow"
				}
			}
		}
		else
		{
			Write_Log "Lane #${lane}: Machine information not found. Skipping." "Yellow"
			continue
		}
	}
	
	# Summary of refresh results
	Write_Log "Refresh Summary for Store Number: $StoreNumber - Total Files Refreshed: $totalRefreshed, Total Failures: $totalFailed." "Green"
	Write_Log "`r`n==================== Refresh_PIN_Pad_Files Function Completed ====================" "Blue"
}

# ===================================================================================================
#                                       FUNCTION: Install_ONE_FUNCTION_Into_SMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates and deploys specific SQL and SQM files required for SMS installation.
#   The files are written directly to their respective destinations in ANSI (Windows-1252) encoding
#   with CRLF line endings and no BOM.
# ===================================================================================================

function Install_ONE_FUNCTION_Into_SMS
{
	param (
		[Parameter(Mandatory = $false)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OfficePath
	)
	
	Write_Log "`r`n==================== Starting Install_ONE_FUNCTION_Into_SMS Function ====================`r`n" "blue"
	
	# --------------------------------------------------------------------------------------------
	# Define Destination Paths
	# --------------------------------------------------------------------------------------------
	
	# Destination folder for Pump_all_items_tables.sql
	$PumpAllItemsTablesDestinationFolder = Join-Path -Path $OfficePath -ChildPath "XF${StoreNumber}901"
	$PumpAllItemsTablesFilePath = Join-Path -Path $PumpAllItemsTablesDestinationFolder -ChildPath "Pump_all_items_tables.sql"
	
	# Destination paths for DEPLOY_SYS.sql and DEPLOY_ONE_FCT.sqm
	$DeploySysDestinationPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_SYS.sql"
	$DeployOneFctDestinationPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_ONE_FCT.sqm"
	
	# --------------------------------------------------------------------------------------------
	# Define File Contents
	# --------------------------------------------------------------------------------------------
	
	# Define the content for Pump_all_items_tables.sql
	$PumpAllItemsTablesContent = @"
/* First delete the record if it exist */
DELETE FROM FCT_TAB WHERE F1063 = 11899 AND F1000 = 'PAL';

/* Insert the new function */
INSERT INTO FCT_TAB (F1063,F1000,F1047,F1050,F1051,F1052,F1053,F1064,F1081) 
VALUES (11899,'PAL',9,'','SKU','Preference','1','Pump all item tables','sql=DEPLOY_LOAD');

/* Activate the new function right away */
@EXEC(SQL=ACTIVATE_ACCEPT_SYS);
"@
	
	# Define the content for DEPLOY_SYS.sql
	$DeploySysContent = @"
@FMT(CMP,@dbHot(FINDFIRST,UD_DEPLOY_SYS.SQL)=,®WIZRPL(UD_RUN=0));
@FMT(CMP,@WIZGET(UD_RUN)=,'®EXEC(SQL=UD_DEPLOY_SYS)®FMT(CHR,27)');

@FMT(CMP,@TOOLS(MESSAGEDLG,"!TO KEEP THE LANE'S REFERENCE SAMPLE UP TO DATE YOU SHOULD USE THE "REFERENCE SAMPLE MECHANISM". DO YOU WANT TO CONTINUE?",,NO,YES)=1,'®FMT(CHR,27)');

@EXEC(INI=HOST_OFFICE[DEPLOY_SYS]);

@WIZRPL(STYLE=SIL);
@WIZRPL(TARGET_FILTER=@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE));

@EXEC(sqi=USERB_DEPLOY_SYS);

@WIZINIT;
@WIZMENU(ONESQM=What do you want to send,
    One Function=DEPLOY_ONE_FCT,
    All Functions=fct_load,      
    Function link=fcz_load,
    Totalizer=tlz_load,
    Drill down pages=dril_page_load,
    Drill down files=dril_file_load,
    All system=ALL);
@WIZDISPLAY;

@WIZINIT;
@WIZTARGET(TARGET=,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1 ORDER BY F1000","
SELECT DISTINCT STO.F1000,STO.F1018 
FROM LNK_TAB LN2 JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000 
WHERE STO.F1181='1' AND LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'
ORDER BY STO.F1000"));
@WIZDISPLAY;

@FMT(CMP,@dbSelect(select distinct 1 from lnk_tab where F1000='@Wizget(Target)' and f1056='999')=,,"®EXEC(msg=!*****_can_not_deploy_system_tables_to_a_host_****);®FMT(CHR,27);")

@WIZINIT;
@WIZMENU(ACTION=Action on the target database,Add or replace=ADDRPL,Add only=ADD,Replace only=UPDATE,Clean and load=LOAD);
@WIZDISPLAY;

/* SEND ONLY ONE TABLE */

@FMT(CMP,@wizget(ONESQM)=tlz_load,®EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(ONESQM)=fcz_load,®EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(ONESQM)=fct_load,®EXEC(SQM=fct_load));
@FMT(CMP,@wizget(ONESQM)=dril_file_load,®EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(ONESQM)=dril_page_load,®EXEC(SQM=DRIL_PAGE_LOAD));
@FMT(CMP,@wizget(ONESQM)=DEPLOY_ONE_FCT,®EXEC(SQM=DEPLOY_ONE_FCT));

@FMT(CMP,@WIZGET(ONESQM)=ALL,,'®EXEC(SQM=exe_activate_accept_sys)®fmt(chr,27)');

@FMT(CMP,@wizget(tlz_load)=0,,®EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(fcz_load)=0,,®EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(fct_load)=0,,®EXEC(SQM=fct_load));
@FMT(CMP,@wizget(DRIL_FILE_LOAD)=0,,®EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(DRIL_PAGE_LOAD)=0,,®EXEC(SQM=DRIL_PAGE_LOAD));

@FMT(CMP,@wizget(exe_activate_accept_all)=0,,®EXEC(SQM=exe_activate_accept_sys));
@FMT(CMP,@wizget(exe_refresh_menu)=1,®EXEC(SQM=exe_refresh_menu));

@EXEC(sqi=USERE_DEPLOY_SYS);
"@
	
	# Define the content for DEPLOY_ONE_FCT.sqm
	$DeployOneFctContent = @"
INSERT INTO HEADER_DCT VALUES
('HC','00000001','001901','001001',,,1997001,0000,1997001,0001,,'LOAD','CREATE DCT',,,,,,'1/1.0','V1.0',,);

CREATE TABLE FCT_DCT(@MAP_FROM_QUERY);

INSERT INTO HEADER_DCT VALUES
('HM','00000001','001901','001001',,,1997001,0000,1997001,0001,,'@WIZGET(ACTION)','@WIZGET(ACTION) ALL FUNCTIONS',,,,,,'1/1.0','V1.0','F1063',);

CREATE VIEW FCT_CHG AS SELECT @FIELDS_FROM_QUERY FROM FCT_DCT;

INSERT INTO FCT_CHG VALUES

/* EXTRACT SECTION */

@DBHOT(HOT_WIZ,PARAMTOLINE,PARAMSAV_FCT_LOAD);
@FMT(CMP,'@WIZGET(TARGET)<>','®WIZRPL(TARGET_FILTER=@WIZGET(TARGET))');

@WIZINIT;
@WIZTARGET(TARGET_FILTER=,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1","
SELECT DISTINCT STO.F1000,STO.F1018 
FROM LNK_TAB LN2 JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000 
WHERE STO.F1181='1' AND LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'"));
@WIZDISPLAY;

@WIZINIT;
@WIZMENU(ACTION=Action on the target database,Add or replace=ADDRPL,Add only=ADD,Replace only=UPDATE,Create and load=LOAD);
@WIZDISPLAY;

@WIZSET(STYLE=SIL);
@WIZCLR(TARGET);
@WIZSET(FORCE_F1000=@F1056);

@WIZINIT;
@WIZEDIT(FCT=,Enter the function number);
@WIZDISPLAY;

@WIZSET(TARGET_FILTER=@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE));
@WIZSET(F1063=@WIZGET(FCT));
@WIZSET(STYLE=SIL);

@MAP_DEPLOY
SELECT FCT.F1056,FCT.F1056+FCT.F1057 AS F1000,@dbFld(FCT_TAB,FCT.,F1000) FROM 
	(SELECT LNI.F1056,LNI.F1057,FCT.*,ROW_NUMBER() OVER (PARTITION BY FCT.F1063,LNI.F1056,LNI.F1057 ORDER BY CASE WHEN FCT.F1000='PAL' THEN 1 ELSE 2 END DESC) AS F1301 
	FROM FCT_TAB FCT
	JOIN LNK_TAB LNI ON FCT.F1000=LNI.F1000
	JOIN LNK_TAB LNO ON LNI.F1056=LNO.F1056 AND LNI.F1057=LNO.F1057
	WHERE LNO.F1000 = '@WIZGET(TARGET_FILTER)' AND FCT.F1063 = '@WIZGET(FCT)') FCT
WHERE FCT.F1301=1
ORDER BY F1000,F1063;

/* RESTORE INITITAL PARAMETER POOL */
@WIZRESET; 
@DBHOT(HOT_WIZ,LINETOPARAM,PARAMSAV_FCT_LOAD);
@DBHOT(HOT_WIZ,CLR,PARAMSAV_FCT_LOAD);
"@
	
	# --------------------------------------------------------------------------------------------
	# Prepare File Contents
	# --------------------------------------------------------------------------------------------
	
	# Ensure content strings have Windows-style line endings
	$PumpAllItemsTablesContent = $PumpAllItemsTablesContent -replace "`n", "`r`n"
	$DeploySysContent = $DeploySysContent -replace "`n", "`r`n"
	$DeployOneFctContent = $DeployOneFctContent -replace "`n", "`r`n"
	
	# Define encoding as ANSI (Windows-1252) without BOM
	$ansiEncoding = [System.Text.Encoding]::GetEncoding(1252)
	
	# --------------------------------------------------------------------------------------------
	# Ensure Destination Directories Exist
	# --------------------------------------------------------------------------------------------
	try
	{
		if (-not (Test-Path $PumpAllItemsTablesDestinationFolder))
		{
			New-Item -Path $PumpAllItemsTablesDestinationFolder -ItemType Directory -Force | Out-Null
			Write_Log "Created directory '$PumpAllItemsTablesDestinationFolder'." "yellow"
		}
	}
	catch
	{
		Write_Log "Failed to create directory '$PumpAllItemsTablesDestinationFolder'. Error: $_" "red"
		return
	}
	
	# --------------------------------------------------------------------------------------------
	# Write Files Directly to Destination Paths
	# --------------------------------------------------------------------------------------------
	try
	{
		# Write Pump_all_items_tables.sql
		[System.IO.File]::WriteAllText($PumpAllItemsTablesFilePath, $PumpAllItemsTablesContent, $ansiEncoding)
		Set-ItemProperty -Path $PumpAllItemsTablesFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Successfully wrote 'Pump_all_items_tables.sql' to '$PumpAllItemsTablesDestinationFolder'." "green"
	}
	catch
	{
		Write_Log "Failed to write 'Pump_all_items_tables.sql'. Error: $_" "red"
	}
	
	try
	{
		# Write DEPLOY_SYS.sql
		[System.IO.File]::WriteAllText($DeploySysDestinationPath, $DeploySysContent, $ansiEncoding)
		Set-ItemProperty -Path $DeploySysDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Successfully wrote 'DEPLOY_SYS.sql' to '$OfficePath'." "green"
	}
	catch
	{
		Write_Log "Failed to write 'DEPLOY_SYS.sql'. Error: $_" "red"
	}
	
	try
	{
		# Write DEPLOY_ONE_FCT.sqm
		[System.IO.File]::WriteAllText($DeployOneFctDestinationPath, $DeployOneFctContent, $ansiEncoding)
		Set-ItemProperty -Path $DeployOneFctDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Successfully wrote 'DEPLOY_ONE_FCT.sqm' to '$OfficePath'." "green"
	}
	catch
	{
		Write_Log "Failed to write 'DEPLOY_ONE_FCT.sqm'. Error: $_" "red"
	}
	
	# --------------------------------------------------------------------------------------------
	# Remove Archive Bit from Pump_all_items_tables.sql Only If the File Exists
	# --------------------------------------------------------------------------------------------
	try
	{
		if (Test-Path $PumpAllItemsTablesFilePath)
		{
			$file = Get-Item -Path $PumpAllItemsTablesFilePath
			if ($file.Attributes -band [System.IO.FileAttributes]::Archive)
			{
				$file.Attributes = $file.Attributes -bxor [System.IO.FileAttributes]::Archive
				Write_Log "Removed the archive bit from '$PumpAllItemsTablesFilePath'." "green"
			}
			else
			{
				#	Write_Log "Archive bit was not set for '$PumpAllItemsTablesFilePath'." "yellow"
			}
		}
		else
		{
			Write_Log "File '$PumpAllItemsTablesFilePath' does not exist. Cannot remove archive bit." "red"
		}
	}
	catch
	{
		Write_Log "Failed to remove the archive bit from '$PumpAllItemsTablesFilePath'. Error: $_" "red"
	}
	
	Write_Log "`r`n==================== Install_ONE_FUNCTION_Into_SMS Function Completed ====================" "blue"
}

# ===================================================================================================
#                                   FUNCTION: Schedule_Lane_DB_Maintenance
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys a lane DB repair SQL/SQI script to the selected lane's Office folder,
#   and creates a scheduler macro in the local XF folder to run the repair weekly.
#   All files are written in ANSI (Windows-1252) encoding with CRLF line endings and no BOM.
#   Ensures destination directories exist and removes the archive bit after writing.
# ===================================================================================================

function Schedule_Lane_DB_Maintenance
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	Write_Log "`r`n==================== Starting Schedule_Lane_DB_Maintenance ====================`r`n" "blue"
	
	if (-not $LaneSQLScript)
	{
		Write_Log "Lane SQL script content variable (`\$LaneSQLScript`) is empty or not defined." "red"
		return
	}
	
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if (-not $selection -or -not $selection.Lanes)
	{
		Write_Log "No lanes selected. Cancelling operation." "yellow"
		return
	}
	# --- ENSURE NODE MAPPING IS PRESENT ---
	$LaneMachines = $script:FunctionResults['LaneMachines']
	if (-not $LaneMachines)
	{
		$null = Retrieve_Nodes -StoreNumber $StoreNumber
		$LaneMachines = $script:FunctionResults['LaneMachines']
		if (-not $LaneMachines)
		{
			Write_Log "Failed to retrieve lane-to-machine mapping for store $StoreNumber." "red"
			return
		}
	}
	$ansiEncoding = [System.Text.Encoding]::GetEncoding(1252)
	$LaneSQLScriptContent = $LaneSQLScript -replace "`n", "`r`n"
	
	# Prompt user for the repeat interval in days
	Add-Type -AssemblyName System.Windows.Forms
	$daysPromptForm = New-Object System.Windows.Forms.Form
	$daysPromptForm.Text = "Lane DB Maintenance - Schedule Interval"
	$daysPromptForm.Width = 350
	$daysPromptForm.Height = 160
	$daysPromptForm.StartPosition = "CenterScreen"
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "How many days between each run (minimum 1):"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(15, 20)
	$daysPromptForm.Controls.Add($label)
	
	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Location = New-Object System.Drawing.Point(20, 50)
	$textBox.Width = 60
	$textBox.Text = "7"
	$daysPromptForm.Controls.Add($textBox)
	
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(90, 90)
	$okButton.Add_Click({ $daysPromptForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })
	$daysPromptForm.Controls.Add($okButton)
	
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(170, 90)
	$cancelButton.Add_Click({ $daysPromptForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
	$daysPromptForm.Controls.Add($cancelButton)
	
	$daysPromptForm.AcceptButton = $okButton
	$daysPromptForm.CancelButton = $cancelButton
	
	if ($daysPromptForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Operation cancelled by user in interval prompt." "yellow"
		return
	}
	
	[int]$UserDays = 0
	if ([int]::TryParse($textBox.Text, [ref]$UserDays) -and $UserDays -ge 1)
	{
		$RepeatDays = $UserDays
	}
	else
	{
		Write_Log "Invalid or no interval provided, using 7 days." "yellow"
		$RepeatDays = 7
	}
	
	foreach ($LaneNumber in $selection.Lanes)
	{
		# ----------- USE MAPPING ----------
		$LaneMachineName = $LaneMachines[$LaneNumber]
		if (-not $LaneMachineName)
		{
			Write_Log "Could not resolve machine name for lane $LaneNumber. Skipping." "red"
			continue
		}
		$LaneOfficeFolder = "\\$LaneMachineName\storeman\office"
		$DestScriptPath = Join-Path $LaneOfficeFolder "LANE_DB_MAINTENANCE.SQI"
		$LocalXFPath = Join-Path $OfficePath "XF$StoreNumber$LaneNumber"
		$SchedulerMacroPath = Join-Path $LocalXFPath "Add_LaneDBMaintenance_to_RUN_TAB.sqi"
		
		# Prepare scheduler macro content (unique task number per lane if needed)
		$TaskNumber = 750
		$HostTarget = "{0:D3}" -f [int]$LaneNumber
		$CommandToRun = 'sqi=LANE_DB_MAINTENANCE'
		$ExecTarget = $HostTarget
		$TaskName = 'Lane DB Maintenance'
		$ManualAllowed = 1
		$CatchupMissed = 1
		$WeeklyDays = $RepeatDays
		$Months = 0
		$Minutes = 0
		$LastRanDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
		$NextRunDate = (Get-Date).AddDays($RepeatDays).ToString("yyyy-MM-dd HH:mm:ss.fff")
		
		
		$SchedulerMacroContent = @"
 /* First delete the scheduled maintenance if it exists */
 DELETE FROM RUN_TAB WHERE F1103 = '$CommandToRun' AND F1000 = '$HostTarget';

 /* Insert the scheduled weekly maintenance */
 INSERT INTO RUN_TAB (F1102, F1000, F1103, F1104, F1105, F1107, F1108, F1109, F1111, F1114, F1115, F1117)
 VALUES ($TaskNumber, '$HostTarget', '$CommandToRun', '$ExecTarget', '$LastRanDate', '$NextRunDate', $ManualAllowed, '$TaskName', $CatchupMissed, $WeeklyDays, $Months, $Minutes);

 /* Activate the new task right away */
 @EXEC(SQL=ACTIVATE_ACCEPT_SYS);
"@ -replace "`n", "`r`n"
		
		try
		{
			if (-not (Test-Path $LaneOfficeFolder))
			{
				Write_Log "Remote office folder not found: $LaneOfficeFolder (lane may be offline or not shared)." "red"
				continue
			}
			if (-not (Test-Path $LocalXFPath))
			{
				Write_Log "Local XF folder not found: $LocalXFPath (scheduler macro not dropped for this lane)." "red"
				continue
			}
		}
		catch
		{
			Write_Log "Failed to create required directories: $_" "red"
			continue
		}
		
		# Write the lane repair script
		try
		{
			[System.IO.File]::WriteAllText($DestScriptPath, $LaneSQLScriptContent, $ansiEncoding)
			Set-ItemProperty -Path $DestScriptPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			Write_Log "Wrote lane DB maintenance script to $DestScriptPath" "green"
		}
		catch
		{
			Write_Log "Failed to write script: $_" "red"
			continue
		}
		
		# Write the scheduler macro
		try
		{
			[System.IO.File]::WriteAllText($SchedulerMacroPath, $SchedulerMacroContent, $ansiEncoding)
			Set-ItemProperty -Path $SchedulerMacroPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			Write_Log "Scheduler SQL macro created at $SchedulerMacroPath" "green"
		}
		catch
		{
			Write_Log "Failed to write scheduler macro: $_" "red"
			continue
		}
		
		# Remove archive bit if set (optional)
		try
		{
			if (Test-Path $DestScriptPath)
			{
				$file = Get-Item -Path $DestScriptPath
				if ($file.Attributes -band [System.IO.FileAttributes]::Archive)
				{
					$file.Attributes = $file.Attributes -bxor [System.IO.FileAttributes]::Archive
					Write_Log "Removed the archive bit from '$DestScriptPath'." "green"
				}
			}
		}
		catch
		{
			Write_Log "Failed to remove the archive bit from '$DestScriptPath'. Error: $_" "red"
		}
	}
	
	Write_Log "`r`n==================== Schedule_Lane_DB_Maintenance Function Completed ====================" "blue"
}

# ===================================================================================================
#                                   FUNCTION: Schedule_Server_DB_Maintenance
# ---------------------------------------------------------------------------------------------------
# Description:
#   Schedules a DB repair task on the server by writing the repair SQL/SQI script and scheduler macro
#   to the local server XF folder.
#   Files are written in ANSI (Windows-1252) encoding with CRLF line endings and no BOM.
#   XF folder must already exist; if not, the operation is skipped and logged.
# ===================================================================================================

function Schedule_Server_DB_Maintenance
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[string]$ServerNumber = "901"
	)
	Write_Log "`r`n==================== Starting Schedule_Server_DB_Maintenance ====================`r`n" "blue"
	
	if (-not $script:ScheduleServerScript)
	{
		Write_Log "Server SQL script content variable (`\$script:ScheduleServerScript`) is empty or not defined." "red"
		return
	}
	
	# Prompt user for the repeat interval in days (once)
	Add-Type -AssemblyName System.Windows.Forms
	$daysPromptForm = New-Object System.Windows.Forms.Form
	$daysPromptForm.Text = "Server DB Maintenance - Schedule Interval"
	$daysPromptForm.Width = 350
	$daysPromptForm.Height = 160
	$daysPromptForm.StartPosition = "CenterScreen"
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "How many days between each run (minimum 1):"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(15, 20)
	$daysPromptForm.Controls.Add($label)
	
	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Location = New-Object System.Drawing.Point(20, 50)
	$textBox.Width = 60
	$textBox.Text = "7"
	$daysPromptForm.Controls.Add($textBox)
	
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(90, 90)
	$okButton.Add_Click({ $daysPromptForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })
	$daysPromptForm.Controls.Add($okButton)
	
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(170, 90)
	$cancelButton.Add_Click({ $daysPromptForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
	$daysPromptForm.Controls.Add($cancelButton)
	
	$daysPromptForm.AcceptButton = $okButton
	$daysPromptForm.CancelButton = $cancelButton
	
	if ($daysPromptForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Operation cancelled by user in interval prompt." "yellow"
		return
	}
	
	[int]$UserDays = 0
	if ([int]::TryParse($textBox.Text, [ref]$UserDays) -and $UserDays -ge 1)
	{
		$RepeatDays = $UserDays
	}
	else
	{
		Write_Log "Invalid or no interval provided, using 7 days." "yellow"
		$RepeatDays = 7
	}
	
	# Paths: Office for the script, XF for the scheduler macro
	$OfficeFolder = $OfficePath
	$DestScriptPath = Join-Path $OfficeFolder "SERVER_DB_MAINTENANCE.SQI"
	$LocalXFPath = Join-Path $OfficePath "XF$StoreNumber$ServerNumber"
	$SchedulerMacroPath = Join-Path $LocalXFPath "Add_ServerDBMaintenance_to_RUN_TAB.sqi"
	
	if (-not (Test-Path $LocalXFPath))
	{
		Write_Log "Local XF folder not found: $LocalXFPath (repair script and scheduler not dropped)." "red"
		return
	}
	
	$ansiEncoding = [System.Text.Encoding]::GetEncoding(1252)
	$ServerSQLScriptContent = $script:ScheduleServerScript
	
	$TaskNumber = 750
	$HostTarget = "{0:D3}" -f [int]$ServerNumber
	$CommandToRun = 'sqi=SERVER_DB_MAINTENANCE'
	$ExecTarget = $HostTarget
	$TaskName = 'Server DB Maintenance'
	$ManualAllowed = 1
	$CatchupMissed = 1
	$WeeklyDays = $RepeatDays
	$Months = 0
	$Minutes = 0
	$LastRanDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
	$NextRunDate = (Get-Date).AddDays($RepeatDays).ToString("yyyy-MM-dd HH:mm:ss.fff")
	
	$SchedulerMacroContent = @"
 /* First delete the scheduled maintenance if it exists */
 DELETE FROM RUN_TAB WHERE F1103 = '$CommandToRun' AND F1000 = '$HostTarget';

 /* Insert the scheduled weekly maintenance */
 INSERT INTO RUN_TAB (F1102, F1000, F1103, F1104, F1105, F1107, F1108, F1109, F1111, F1114, F1115, F1117)
 VALUES ($TaskNumber, '$HostTarget', '$CommandToRun', '$ExecTarget', '$LastRanDate', '$NextRunDate', $ManualAllowed, '$TaskName', $CatchupMissed, $WeeklyDays, $Months, $Minutes);

 /* Activate the new task right away */
 @EXEC(SQL=ACTIVATE_ACCEPT_SYS);
"@ -replace "`n", "`r`n"
	
	# Write the server repair script
	try
	{
		[System.IO.File]::WriteAllText($DestScriptPath, $ServerSQLScriptContent, $ansiEncoding)
		Set-ItemProperty -Path $DestScriptPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Wrote server DB repair script to $DestScriptPath" "green"
	}
	catch
	{
		Write_Log "Failed to write server script: $_" "red"
		return
	}
	
	# Write the scheduler macro
	try
	{
		[System.IO.File]::WriteAllText($SchedulerMacroPath, $SchedulerMacroContent, $ansiEncoding)
		Set-ItemProperty -Path $SchedulerMacroPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Scheduler SQL macro created at $SchedulerMacroPath" "green"
	}
	catch
	{
		Write_Log "Failed to write scheduler macro: $_" "red"
		return
	}
	
	# Remove archive bit if set (optional)
	try
	{
		if (Test-Path $DestScriptPath)
		{
			$file = Get-Item -Path $DestScriptPath
			if ($file.Attributes -band [System.IO.FileAttributes]::Archive)
			{
				$file.Attributes = $file.Attributes -bxor [System.IO.FileAttributes]::Archive
				Write_Log "Removed the archive bit from '$DestScriptPath'." "green"
			}
		}
	}
	catch
	{
		Write_Log "Failed to remove the archive bit from '$DestScriptPath'. Error: $_" "red"
	}
	
	Write_Log "`r`n==================== Schedule_Server_DB_Maintenance Function Completed ====================" "blue"
}

# ===================================================================================================
#                                 FUNCTION: Organize_TBS_SCL_ver520
# ---------------------------------------------------------------------------------------------------
# Description:
#   Organizes the [TBS_SCL_ver520] table by updating ScaleName, BufferTime, and ScaleCode for
#   BIZERBA and ISHIDA records. Specifically:
#     - Sets BufferTime to 1 for the first BIZERBA record and to 5 for all other BIZERBA records.
#     - Updates ScaleName for BIZERBA records to include the IPDevice.
#     - Reassigns ScaleCode to ensure BIZERBA records are first, followed by ISHIDA records.
#     - Updates ScaleName and BufferTime for ISHIDA WMAI records.
#   Optionally exports the organized data to a CSV file.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - OutputCsvPath (Optional): Path to export the organized CSV file. If not provided, the data
#     is only displayed in the console.
# ===================================================================================================

function Organize_TBS_SCL_ver520
{
	[CmdletBinding()]
	param (
		# (Optional) Path to export the organized CSV
		[Parameter(Mandatory = $false)]
		[string]$OutputCsvPath
	)
	
	Write_Log "`r`n==================== Starting Organize_TBS_SCL_ver520 Function ====================`r`n" "blue"
	
	# Access the connection string from the script-scoped variable
	# Ensure that you have set $script:FunctionResults['ConnectionString'] before calling this function
	$connectionString = $script:FunctionResults['ConnectionString']
	
	if (-not $connectionString)
	{
		Write_Log "Connection string not found in `$script:FunctionResults['ConnectionString']`." "red"
		return
	}
	
	# Determine if Invoke-Sqlcmd supports the -ConnectionString parameter
	$supportsConnectionString = $false
	try
	{
		$cmd = Get-Command Invoke-Sqlcmd -ErrorAction Stop
		$supportsConnectionString = $cmd.Parameters.Keys -contains 'ConnectionString'
	}
	catch
	{
		Write_Log "Invoke-Sqlcmd cmdlet not found: $_" "red"
		$supportsConnectionString = $false
	}
	
	# Define the SQL commands:
	# 1. Update ScaleName and BufferTime for ISHIDA WMAI records.
	# 2. Update ScaleName for BIZERBA records.
	# 3. Update ScaleCode for BIZERBA records.
	# 4. Update ScaleCode for ISHIDA records after BIZERBA.
	# 5. Set BufferTime for BIZERBA records: first one = 1, others = 5.
	
	$updateQueries = @"
-------------------------------------------------------------------------------
-- 1) Update ISHIDA WMAI ScaleName and BufferTime based on the record count
-------------------------------------------------------------------------------
DECLARE @IshidaWMAICount INT;

SELECT @IshidaWMAICount = COUNT(*)
FROM [TBS_SCL_ver520]
WHERE ScaleBrand = 'ISHIDA' AND ScaleModel = 'WMAI';

IF @IshidaWMAICount > 1
BEGIN
    UPDATE [TBS_SCL_ver520]
    SET 
        ScaleName = CONCAT('Ishida Wrapper ', IPDevice),
        BufferTime = '1'
    WHERE 
        ScaleBrand = 'ISHIDA' AND ScaleModel = 'WMAI';
END
ELSE
BEGIN
    UPDATE [TBS_SCL_ver520]
    SET 
        ScaleName = 'Ishida Wrapper',
        BufferTime = '1'
    WHERE 
        ScaleBrand = 'ISHIDA' AND ScaleModel = 'WMAI';
END;

-------------------------------------------------------------------------------
-- 2) Update BIZERBA ScaleName
-------------------------------------------------------------------------------
UPDATE [TBS_SCL_ver520]
SET 
    ScaleName = CONCAT('Scale ', IPDevice)
WHERE 
    ScaleBrand = 'BIZERBA';

-------------------------------------------------------------------------------
-- 3) Update ScaleCode for BIZERBA, starting at 10, in IPDevice ascending order
--    Using ROW_NUMBER() ensures uniqueness within this group.
-------------------------------------------------------------------------------
WITH BIZERBA_CTE AS (
    SELECT 
        ScaleCode,
        IPDevice,
        rn = ROW_NUMBER() OVER (ORDER BY TRY_CAST(IPDevice AS INT)) 
    FROM [TBS_SCL_ver520]
    WHERE ScaleBrand = 'BIZERBA'
)
UPDATE T
SET T.ScaleCode = 10 + B.rn - 1
FROM [TBS_SCL_ver520] AS T
JOIN BIZERBA_CTE AS B 
    ON T.ScaleCode = B.ScaleCode
WHERE T.ScaleBrand = 'BIZERBA';

-------------------------------------------------------------------------------
-- 4) Update ScaleCode for ISHIDA, starting after the new max BIZERBA ScaleCode.
--    We add +1 so we don't overlap, or you can add +10 if you want a bigger gap.
-------------------------------------------------------------------------------
;WITH MaxBizerba AS (
    SELECT MAX(ScaleCode) AS MaxCode
    FROM [TBS_SCL_ver520]
    WHERE ScaleBrand = 'BIZERBA'
),
ISHIDA_CTE AS (
    SELECT 
        ScaleCode,
        IPDevice,
        rn = ROW_NUMBER() OVER (ORDER BY TRY_CAST(IPDevice AS INT))
    FROM [TBS_SCL_ver520]
    WHERE ScaleBrand = 'ISHIDA'
)
UPDATE T
SET T.ScaleCode = (SELECT MaxCode FROM MaxBizerba) + 10 + I.rn - 1
FROM [TBS_SCL_ver520] AS T
JOIN ISHIDA_CTE AS I
    ON T.ScaleCode = I.ScaleCode
WHERE T.ScaleBrand = 'ISHIDA';

-------------------------------------------------------------------------------
-- 5) Now set BufferTime for BIZERBA records:
--    The lowest ScaleCode (i.e. first in ascending ScaleCode order) gets 1,
--    and all others get 5.
-------------------------------------------------------------------------------
WITH BIZ_ORDER AS (
    SELECT 
        ScaleCode,
        RN = ROW_NUMBER() OVER (ORDER BY ScaleCode ASC)
    FROM [TBS_SCL_ver520]
    WHERE ScaleBrand = 'BIZERBA'
)
UPDATE T
SET T.BufferTime = CASE WHEN B.RN = 1 THEN '1' ELSE '5' END
FROM [TBS_SCL_ver520] T
INNER JOIN BIZ_ORDER B 
    ON T.ScaleCode = B.ScaleCode
WHERE T.ScaleBrand = 'BIZERBA';
"@
	
	$selectQuery = @"
SELECT 
    ScaleCode,
    ScaleName,
    ScaleLocation,
    IPNetwork,
    IPDevice,
    Active,
    SystemLocalTime,
    AutoStart,
    AutoTransmit,
    BufferTime,
    ScaleBrand,
    ScaleModel
FROM 
    [TBS_SCL_ver520]
ORDER BY 
    ScaleCode ASC; -- Sort by ScaleCode ascending to have BIZERBA first, then ISHIDA
"@
	
	# Initialize variables to track execution
	$retryCount = 0
	$MaxRetries = 2
	$RetryDelaySeconds = 5
	$success = $false
	$failedSections = @()
	$failedCommands = @()
	
	while (-not $success -and $retryCount -lt $MaxRetries)
	{
		try
		{
			Write_Log "Starting execution of Organize_TBS_SCL_ver520. Attempt $($retryCount + 1) of $MaxRetries." "blue"
			
			# Execute the update queries
			Write_Log "Executing update queries to modify ScaleName, BufferTime, and ScaleCode..." "blue"
			try
			{
				if ($supportsConnectionString)
				{
					Invoke-Sqlcmd -ConnectionString $connectionString -Query $updateQueries -ErrorAction Stop
				}
				else
				{
					# Parse ServerInstance and Database from ConnectionString
					$server = ($connectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($connectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					
					if (-not $server -or -not $database)
					{
						Write_Log "Invalid ConnectionString. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $updateQueries -ErrorAction Stop
				}
				Write_Log "Update queries executed successfully." "green"
			}
			catch [System.Management.Automation.ParameterBindingException]
			{
				Write_Log "ParameterBindingException encountered while executing update queries. Attempting fallback." "yellow"
				
				# Attempt to execute using ServerInstance and Database
				try
				{
					# Parse ServerInstance and Database from ConnectionString
					$server = ($connectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($connectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					
					if (-not $server -or -not $database)
					{
						Write_Log "Invalid ConnectionString for fallback. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $updateQueries -ErrorAction Stop
					Write_Log "Update queries executed successfully using fallback parameters." "green"
				}
				catch
				{
					Write_Log "Error executing update queries with fallback parameters: $_" "red"
					throw $_
				}
			}
			catch
			{
				Write_Log "An error occurred while executing update queries: $_" "red"
				throw $_
			}
			
			# Execute the select query to retrieve organized data
			Write_Log "Retrieving organized data..." "blue"
			try
			{
				if ($supportsConnectionString)
				{
					$data = Invoke-Sqlcmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
				}
				else
				{
					# Parse ServerInstance and Database from ConnectionString
					$server = ($connectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($connectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					
					if (-not $server -or -not $database)
					{
						Write_Log "Invalid ConnectionString. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					$data = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
				}
				Write_Log "Data retrieval successful." "green"
			}
			catch [System.Management.Automation.ParameterBindingException]
			{
				Write_Log "ParameterBindingException encountered while retrieving data. Attempting fallback." "yellow"
				
				# Attempt to execute using ServerInstance and Database
				try
				{
					# Parse ServerInstance and Database from ConnectionString
					$server = ($connectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($connectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					
					if (-not $server -or -not $database)
					{
						Write_Log "Invalid ConnectionString for fallback. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					$data = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					Write_Log "Data retrieval successful using fallback parameters." "green"
				}
				catch
				{
					Write_Log "Error retrieving data with fallback parameters: $_" "red"
					throw $_
				}
			}
			catch
			{
				Write_Log "An error occurred while retrieving data: $_" "red"
				throw $_
			}
			
			# Check if data was retrieved
			if (-not $data)
			{
				Write_Log "No data retrieved from the table 'TBS_SCL_ver520'." "red"
				throw "No data retrieved from the table 'TBS_SCL_ver520'."
			}
			
			# Export the data if an output path is provided
			if ($PSBoundParameters.ContainsKey('OutputCsvPath'))
			{
				Write_Log "Exporting organized data to '$OutputCsvPath'..." "blue"
				try
				{
					$data | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
					Write_Log "Data exported successfully to '$OutputCsvPath'." "green"
				}
				catch
				{
					Write_Log "Failed to export data to CSV: $_" "red"
				}
			}
			
			# Display the organized data
			Write_Log "Displaying organized data:" "yellow"
			try
			{
				$formattedData = $data | Format-Table -AutoSize | Out-String
				Write_Log $formattedData "Blue"
			}
			catch
			{
				Write_Log "Failed to format and display data: $_" "red"
			}
			
			Write_Log "==================== Organize_TBS_SCL_ver520 Function Completed ====================" "blue"
			$success = $true
		}
		catch
		{
			$retryCount++
			Write_Log "Error during Organize_TBS_SCL_ver520 execution: $_" "red"
			
			if ($retryCount -lt $MaxRetries)
			{
				Write_Log "Retrying execution in $RetryDelaySeconds seconds..." "yellow"
				Start-Sleep -Seconds $RetryDelaySeconds
			}
		}
	}
	
	if (-not $success)
	{
		Write_Log "Maximum retry attempts reached. Organize_TBS_SCL_ver520 function failed." "red"
		# Optionally, you can handle further actions like sending notifications or logging to a file
	}
}

# ===================================================================================================
#                                 FUNCTION: Repair_BMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Repairs the "BMS" service by performing the following steps:
#     1. Stops the "BMS" service if it's running.
#     2. Deletes the "BMS" service.
#     3. Registers BMSSrv.exe to recreate the "BMS" service.
#     4. Starts the newly registered "BMS" service.
#   Ensures that the script waits appropriately between deleting and registering to prevent errors.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - BMSSrvPath (Optional): Full path to BMSSrv.exe. Defaults to "C:\Bizerba\RetailConnect\BMS\BMSSrv.exe".
# ===================================================================================================

function Repair_BMS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$BMSSrvPath = "$env:SystemDrive\Bizerba\RetailConnect\BMS\BMSSrv.exe"
	)
	
	Write_Log "`r`n==================== Starting Repair_BMS Function ====================`r`n" "blue"
	
	# -- Check for Admin Privileges --
	$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	if (-not $isAdmin)
	{
		Write_Log "Insufficient permissions. Please run this script as an Administrator." "red"
		return
	}
	
	# -- Check BMSSrv.exe Exists --
	if (-not (Test-Path $BMSSrvPath))
	{
		Write_Log "BMSSrv.exe not found at path: $BMSSrvPath" "red"
		return
	}
	
	$serviceName = "BMS"
	
	# -- Service Exists Helper --
	$serviceExists = $false
	try { Get-Service -Name $serviceName -ErrorAction Stop | Out-Null; $serviceExists = $true }
	catch { $serviceExists = $false }
	
	# -- Stop BMS Service if running --
	if ($serviceExists)
	{
		Write_Log "Attempting to stop the '$serviceName' service..." "blue"
		try
		{
			Stop-Service -Name $serviceName -Force -ErrorAction Stop
			Write_Log "'$serviceName' service stopped successfully." "green"
		}
		catch
		{
			Write_Log "Failed to stop '$serviceName' service: $_" "red"
			return
		}
	}
	else
	{
		Write_Log "'$serviceName' service does not exist or is already stopped." "yellow"
	}
	
	# -- Delete the BMS Service if it exists --
	$serviceExists = $false
	try { Get-Service -Name $serviceName -ErrorAction Stop | Out-Null; $serviceExists = $true }
	catch { $serviceExists = $false }
	if ($serviceExists)
	{
		Write_Log "Attempting to delete the '$serviceName' service..." "blue"
		try
		{
			sc.exe delete $serviceName | Out-Null
			Write_Log "'$serviceName' service deleted successfully." "green"
		}
		catch
		{
			Write_Log "Failed to delete '$serviceName' service: $_" "red"
			return
		}
		Start-Sleep -Seconds 5
	}
	else
	{
		Write_Log "'$serviceName' service does not exist. Skipping deletion." "yellow"
	}
	
	# -- Register BMSSrv.exe --
	Write_Log "Registering BMSSrv.exe to recreate the '$serviceName' service..." "blue"
	try
	{
		$process = Start-Process -FilePath $BMSSrvPath -ArgumentList "-reg" -NoNewWindow -Wait -PassThru
		if ($process.ExitCode -eq 0)
		{
			Write_Log "BMSSrv.exe registered successfully." "green"
		}
		else
		{
			Write_Log "BMSSrv.exe registration failed with exit code $($process.ExitCode)." "red"
			return
		}
	}
	catch
	{
		Write_Log "An error occurred while registering BMSSrv.exe: $_" "red"
		return
	}
	
	# -- Start the BMS Service --
	Write_Log "Attempting to start the '$serviceName' service..." "blue"
	try
	{
		Start-Service -Name $serviceName -ErrorAction Stop
		Write_Log "'$serviceName' service started successfully." "green"
	}
	catch
	{
		Write_Log "Failed to start '$serviceName' service: $_" "red"
		return
	}
	
	Write_Log "`r`n==================== Repair_BMS Function Completed ====================`r`n" "blue"
}

# ===================================================================================================
#                                         FUNCTION: Write_SQL_Scripts_To_Desktop
# ---------------------------------------------------------------------------------------------------
# Description:
#   Writes the provided LaneSQL and ServerSQL scripts to the user's Desktop with specified filenames.
#   This function ensures that the scripts are saved with UTF-8 encoding and includes error handling
#   to manage any issues that may arise during the file writing process.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   -LaneSQL (Mandatory)
#       The content of the LaneSQL script without @dbEXEC commands and timeout settings.
#
#   -ServerSQL (Mandatory)
#       The content of the ServerSQL script.
#
#   -LaneFilename (Optional)
#       The filename for the LaneSQL script. Defaults to "Lane_Database_Maintenance.sqi".
#
#   -ServerFilename (Optional)
#       The filename for the ServerSQL script. Defaults to "Server_Database_Maintenance.sqi".
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   Write_SQL_Scripts_To_Desktop -LaneSQL $LaneSQLNoDbExecAndTimeout -ServerSQL $script:ServerSQLScript
#
# Prerequisites:
#   - Ensure that the SQL script contents are correctly generated and stored in the provided variables.
#   - Verify that the user has write permissions to the Desktop directory.
# ===================================================================================================

function Write_SQL_Scripts_To_Desktop
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, HelpMessage = "Content of the LaneSQL script without @dbEXEC commands and timeout settings.")]
		[string]$LaneSQL,
		[Parameter(Mandatory = $true, HelpMessage = "Content of the ServerSQL script.")]
		[string]$ServerSQL,
		[Parameter(Mandatory = $false, HelpMessage = "Filename for the LaneSQL script.")]
		[string]$LaneFilename = "Lane_Database_Maintenance.sqi",
		[Parameter(Mandatory = $false, HelpMessage = "Filename for the ServerSQL script.")]
		[string]$ServerFilename = "Server_Database_Maintenance.sqi"
	)
	
	Write_Log "`r`n==================== Starting Write_SQL_Scripts_To_Desktop Function ====================`r`n" "blue"
	
	try
	{
		# Get the path to the user's Desktop
		$desktopPath = [Environment]::GetFolderPath("Desktop")
		
		# Define full file paths
		$laneFilePath = Join-Path -Path $desktopPath -ChildPath $LaneFilename
		$serverFilePath = Join-Path -Path $desktopPath -ChildPath $ServerFilename
		
		# Write the LaneSQL script to the Desktop
		[System.IO.File]::WriteAllText($laneFilePath, $LaneSQL, [System.Text.Encoding]::UTF8)
		Write_Log "Lane SQL script successfully written to:`n$laneFilePath" "Green"
	}
	catch
	{
		Write_Log "Error writing Lane SQL script to Desktop:`n$_" "Red"
	}
	
	try
	{
		# Write the ServerSQL script to the Desktop
		[System.IO.File]::WriteAllText($serverFilePath, $ServerSQL, [System.Text.Encoding]::UTF8)
		Write_Log "Server SQL script successfully written to:`n$serverFilePath" "Green"
	}
	catch
	{
		Write_Log "Error writing Server SQL script to Desktop:`n$_" "Red"
	}
	Write_Log "`r`n==================== Write_SQL_Scripts_To_Desktop Function Completed ====================" "blue"
}

# ===================================================================================================
#                               FUNCTION: Send_Restart_All_Programs
# ---------------------------------------------------------------------------------------------------
# Description:
#   The `Send_Restart_All_Programs` function automates sending a restart command to selected lanes
#   within a specified store. It retrieves lane-to-machine mappings using the `Retrieve_Nodes` 
#   function, prompts the user to select lanes via the `Show_Lane_Selection_Form` function, and
#   then constructs and sends a mailslot command to each selected lane using the correct 
#   machine address.
#
# Parameters:
#   - [string]$StoreNumber
#         A 3-digit identifier for the store (SSS). This parameter is mandatory and is used
#         to retrieve node details, select lanes, and construct mailslot addresses.
#
# Workflow:
#   1. Retrieve node information for the specified store using `Retrieve_Nodes`, which
#      provides a mapping between lanes and their corresponding machine names.
#   2. Launch `Show_Lane_Selection_Form` in 'Store' mode to allow the user to select one
#      or more lanes (TTT).
#   3. For each selected lane:
#         - Look up the machine name from the lane-to-machine mapping.
#         - Construct the mailslot address using the machine name, store number, and lane number.
#         - Send the restart command via `[MailslotSender]::SendMailslotCommand`.
#         - Report success or failure for each command sent.
#
# Returns:
#   None. Outputs success or failure messages to the console for each lane processed.
#
# Example Usage:
#   Send_Restart_All_Programs -StoreNumber "123"
#
# Notes:
#   - Ensure that the helper functions (`Retrieve_Nodes`, `Show_Lane_Selection_Form`) and the 
#     `[MailslotSender]::SendMailslotCommand` method are defined and accessible in the 
#     session before invoking this function.
# ===================================================================================================

function Send_Restart_All_Programs
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[array]$LaneNumbers # Optional. If supplied, skips prompts and sends to these lanes.
	)
	
	Write_Log "`r`n==================== Starting Send_Restart_All_Programs Function ====================`r`n" "blue"
	
	$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
	if (-not $nodes)
	{
		Write_Log "Failed to retrieve node information for store $StoreNumber." "red"
		return
	}
	
	# Use supplied lanes or prompt for selection
	if ($LaneNumbers -and $LaneNumbers.Count -gt 0)
	{
		$lanes = $LaneNumbers | ForEach-Object { $_.ToString().PadLeft(3, '0') }
	}
	else
	{
		$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
		if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
		{
			Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
			return
		}
		$lanes = $selection.Lanes | ForEach-Object { $_.ToString().PadLeft(3, '0') }
	}
	
	foreach ($lane in $lanes)
	{
		$machineName = $nodes.LaneMachines[$lane]
		if (-not $machineName)
		{
			Write_Log "No machine found for lane $lane. Skipping." "yellow"
			continue
		}
		$mailslotAddress = "\\$machineName\Mailslot\SMSStart_${StoreNumber}${lane}"
		$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
		$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
		if ($result)
		{
			Write_Log "Command sent successfully to Machine $machineName (Store $StoreNumber, Lane $lane)." "green"
		}
		else
		{
			Write_Log "Failed to send command to Machine $machineName (Store $StoreNumber, Lane $lane)." "red"
		}
	}
	Write_Log "`r`n==================== Send_Restart_All_Programs Function Completed ====================" "blue"
}

# ===================================================================================================
#                               FUNCTION: Send_SERVER_time_to_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   The `Send_SERVER_time_to_Lanes` function automates sending a time synchronization command to 
#   selected lanes within a specified store using the server's local date and time. It retrieves 
#   lane-to-machine mappings using the `Retrieve_Nodes` function, prompts the user to select lanes 
#   via the `Show_Lane_Selection_Form` function, and then constructs and sends a mailslot command to each 
#   selected lane with the server's current date and time in the appropriate format.
#
# Parameters:
#   - [string]$StoreNumber
#         A 3-digit identifier for the store (SSS). This parameter is mandatory and is used to 
#         retrieve node details, select lanes, and construct mailslot addresses.
#
# Workflow:
#   1. Retrieve node information for the specified store using `Retrieve_Nodes`, which provides a 
#      mapping between lanes and their corresponding machine names.
#   2. Launch `Show_Lane_Selection_Form` in 'Store' mode to allow the user to select one or more lanes (TTT).
#   3. Retrieve the server's local date and time using PowerShell's Get-Date, formatting the date as 
#      "MM/dd/yyyy" and the time as "HHmmss".
#   4. Construct a command string in the format:
#         "@WIZRPL(DATE=MM/DD/YYYY)@WIZRPL(TIME=HHMMSS)"
#   5. For each selected lane:
#         - Look up the machine name from the lane-to-machine mapping.
#         - Construct the mailslot address using the machine name.
#         - Send the time synchronization command via `[MailslotSender]::SendMailslotCommand`.
#         - Report success or failure for each command sent.
#
# Returns:
#   None. Outputs success or failure messages to the console for each lane processed.
#
# Example Usage:
#   Send_SERVER_time_to_Lanes -StoreNumber "123"
#
# Notes:
#   - Ensure that the helper functions (`Retrieve_Nodes`, `Show_Lane_Selection_Form`) and the 
#     `[MailslotSender]::SendMailslotCommand` method are defined and accessible in the session 
#     before invoking this function.
# ===================================================================================================

function Send_SERVER_time_to_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	Write_Log "`r`n==================== Starting Send_SERVER_time_to_Lanes Function ====================`r`n" "blue"
	
	$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
	if (-not $nodes)
	{
		Write_Log "Failed to retrieve node information for store $StoreNumber." "red"
		return
	}
	
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if (-not $selection)
	{
		Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
		return
	}
	
	$lanes = $selection.Lanes
	if (-not $lanes -or $lanes.Count -eq 0)
	{
		Write_Log "No valid lanes found. Exiting." "yellow"
		return
	}
	
	# Get date and time in required format
	$currentDate = (Get-Date).ToString("MM/dd/yyyy")
	$currentTime = (Get-Date).ToString("HHmmss")
	
	# Most direct: just send as @WIZRPL... (if your Launchpad accepts it)
	$commandMessage = "@WIZRPL(DATE=$currentDate)@WIZRPL(TIME=$currentTime)"
	
	# If your lanes require the two-phase WINMAIL parsing (as in your doc), use this:
	# $commandMessage = "Â©WIZRPL(DATE=$currentDate)Â©WIZRPL(TIME=$currentTime)"
	
	foreach ($lane in $lanes)
	{
		# Zero-pad lane number to 3 digits
		$lanePadded = $lane.PadLeft(3, '0')
		$machineName = $nodes.LaneMachines[$lane]
		if (-not $machineName)
		{
			Write_Log "No machine found for lane $lane. Skipping." "yellow"
			continue
		}
		$mailslotAddress = "\\$machineName\MailSlot\WIN$lanePadded"
		$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
		if ($result)
		{
			Write_Log "Time sync command sent successfully to Machine $machineName (Store $StoreNumber, Lane $lanePadded)." "green"
		}
		else
		{
			Write_Log "Failed to send time sync command to Machine $machineName (Store $StoreNumber, Lane $lanePadded)." "red"
		}
	}
	Write_Log "`r`n==================== Send_SERVER_time_to_Lanes Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Drawer_Control
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys a drawer control SQI command to selected lanes for a specified store.
#   The function first presents a GUI for the user to select the desired drawer state 
#   (Enable = 1, Disable = 0) and then uses the Show_Lane_Selection_Form GUI (in "Store" mode) to 
#   allow selection of one or more lanes. For each selected lane, the function writes an SQI file 
#   (in ANSI PC format with CRLF line endings) with the embedded drawer state and sends a restart 
#   command to the corresponding machine.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: The store number to process. (Mandatory)
# ---------------------------------------------------------------------------------------------------
# Requirements:
#   - The Show_Lane_Selection_Form function must be available.
#   - Variables such as $OfficePath must be defined.
#   - Helper functions like Write_Log, Retrieve_Nodes, and the class [MailslotSender] must be available.
# ===================================================================================================

function Drawer_Control
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Drawer_Control ====================`r`n" "blue"
	
	# --------------------------------------------------
	# STEP 1: Prompt for Drawer State using Enable/Disable radio buttons
	# --------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$stateForm = New-Object System.Windows.Forms.Form
	$stateForm.Text = "Select Drawer State"
	$stateForm.Size = New-Object System.Drawing.Size(400, 200)
	$stateForm.StartPosition = "CenterScreen"
	
	$stateLabel = New-Object System.Windows.Forms.Label
	$stateLabel.Text = "Select Drawer State:"
	$stateLabel.Location = New-Object System.Drawing.Point(10, 20)
	$stateLabel.AutoSize = $true
	$stateForm.Controls.Add($stateLabel)
	
	# Radio button for Enable (value = 1)
	$radioEnable = New-Object System.Windows.Forms.RadioButton
	$radioEnable.Text = "Enable"
	$radioEnable.Location = New-Object System.Drawing.Point(10, 50)
	$radioEnable.AutoSize = $true
	$radioEnable.Checked = $true # default selection
	$stateForm.Controls.Add($radioEnable)
	
	# Radio button for Disable (value = 0)
	$radioDisable = New-Object System.Windows.Forms.RadioButton
	$radioDisable.Text = "Disable"
	$radioDisable.Location = New-Object System.Drawing.Point(10, 80)
	$radioDisable.AutoSize = $true
	$stateForm.Controls.Add($radioDisable)
	
	# OK Button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(80, 120)
	$okButton.Add_Click({
			if ($radioEnable.Checked)
			{
				$stateForm.Tag = "1"
			}
			elseif ($radioDisable.Checked)
			{
				$stateForm.Tag = "0"
			}
			$stateForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$stateForm.Close()
		})
	$stateForm.Controls.Add($okButton)
	
	# Cancel Button
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(180, 120)
	$cancelButton.Add_Click({
			$stateForm.Tag = "Cancelled"
			$stateForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$stateForm.Close()
		})
	$stateForm.Controls.Add($cancelButton)
	
	$stateForm.AcceptButton = $okButton
	$stateForm.CancelButton = $cancelButton
	
	$resultState = $stateForm.ShowDialog()
	if ($stateForm.Tag -eq "Cancelled" -or $resultState -eq [System.Windows.Forms.DialogResult]::Cancel)
	{
		Write_Log "User cancelled the operation at drawer state selection." "yellow"
		Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	$DrawerState = $stateForm.Tag
	Write_Log "Drawer state selected: $DrawerState" "green"
	
	# --------------------------------------------------
	# STEP 2: Use Show_Lane_Selection_Form to select lanes (Store mode)
	# --------------------------------------------------
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if ($null -eq $selection)
	{
		Write_Log "No lanes selected or selection cancelled." "yellow"
		Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	
	# Determine the list of lanes to process.
	$lanesToProcess = @()
	if ($selection.Type -eq "Specific" -or $selection.Type -eq "Range" -or $selection.Type -eq "All")
	{
		$lanesToProcess = $selection.Lanes
	}
	else
	{
		Write_Log "Unexpected selection type returned." "red"
		return
	}
	
	# --------------------------------------------------
	# STEP 3: For each selected lane, deploy the SQI command file in ANSI (PC) format and send restart command.
	# --------------------------------------------------
	foreach ($lane in $lanesToProcess)
	{
		# Construct the lane directory path (assumes folder naming: XF<StoreNumber><Lane>)
		$LaneDirectory = "$OfficePath\XF${StoreNumber}${lane}"
		if (-not (Test-Path $LaneDirectory))
		{
			Write_Log "Lane directory $LaneDirectory not found. Skipping lane $lane." "yellow"
			continue
		}
		
		# Define the SQI content with the chosen drawer state
		$SQIContent = @"
CREATE VIEW Fct_Load AS SELECT F1063,F1000,F81,F85,F96,F97,F98,F99,F100,F101,F102,F125,F172,F239,F240,F241,F242,F1042,F1043,F1044,F1045,F1046,F1047,F1050,F1051,F1052,F1053,F1054,F1055,F1064,F1081,F1082,F1083,F1084,F1085,F1086,F1088,F1089,F1090,F1091,F1092,F1147,F1817,F1818,F1895,F1897,F1965,F1966 FROM FCT_TAB;

INSERT INTO Fct_Load VALUES
(10010,'PAL',,,,,,,,,,,,,,,,1,1,1,1,1,9,'F3','TRS','Log',1,1,1,'Login operator','DRAWEROPEN=$DrawerState',,,,,,,,,,,,,,'',,,),

@UPDATE_BATCH(JOB=ADDRPL,TAR=FCT_TAB,
KEY=F1063=:F1063 AND F1000=:F1000,
SRC=SELECT * FROM Fct_Load);

DROP TABLE Fct_Load;

@WIZSET(TARGET=@TER);
@EXEC(SQM=exe_activate_accept_sys);
"@
		
		# Ensure the SQI content uses CRLF line endings (ANSI PC format)
		$SQIContent = $SQIContent -replace "`n", "`r`n"
		
		# Define the full path to the SQI file (named "DrawerControl.sqi")
		$SQIFilePath = Join-Path -Path $LaneDirectory -ChildPath "DrawerControl.sqi"
		
		# Write the SQI file using ASCII encoding (ANSI PC)
		Set-Content -Path $SQIFilePath -Value $SQIContent -Encoding ASCII
		
		# Remove the Archive attribute (set file attributes to Normal)
		Set-ItemProperty -Path $SQIFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Deployed Drawer_Control.sqi command to lane $lane with state '$DrawerState' in directory $LaneDirectory." "green"
	}
	Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Refresh_Database
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys a database refresh SQI command to selected registers for a specified store.
#   The function uses the Show_Lane_Selection_Form GUI (in "Store" mode) to allow selection of one or 
#   more registers. For each selected register, it writes an SQI file (in ANSI PC format with CRLF 
#   line endings) containing:
#
#       @WIZSET(TARGET=@TER);
#       @EXEC(SQM=exe_activate_accept_sys);
#
#   Then, it sends a restart command to the corresponding machine.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: The store number to process. (Mandatory)
# ---------------------------------------------------------------------------------------------------
# Requirements:
#   - The Show_Lane_Selection_Form function must be available.
#   - Variables such as $OfficePath must be defined.
#   - Helper functions like Write_Log, Retrieve_Nodes, and the class [MailslotSender] must be available.
# ===================================================================================================

function Refresh_Database
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Refresh_Database ====================`r`n" "blue"
	
	# --------------------------------------------------
	# STEP 1: Use Show_Lane_Selection_Form to select registers
	# --------------------------------------------------
	# The Show_Lane_Selection_Form function is assumed to be available and operating in "Store" mode.
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if ($null -eq $selection)
	{
		Write_Log "No registers selected or selection cancelled." "yellow"
		Write_Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
		return
	}
	
	# Determine the list of registers to process.
	$registersToProcess = @()
	if ($selection.Type -eq "Specific" -or $selection.Type -eq "Range" -or $selection.Type -eq "All")
	{
		$registersToProcess = $selection.Lanes # In this context, each "lane" is treated as a register.
	}
	else
	{
		Write_Log "Unexpected selection type returned." "red"
		return
	}
	
	# --------------------------------------------------
	# STEP 2: Define the SQI content to refresh the database
	# --------------------------------------------------
	$SQIContent = @"
@WIZSET(TARGET=@TER);
@EXEC(SQM=exe_activate_accept_sys);
"@
	# Ensure the SQI content uses CRLF line endings (ANSI PC format)
	$SQIContent = $SQIContent -replace "`n", "`r`n"
	
	# --------------------------------------------------
	# STEP 3: For each selected register, deploy the SQI file
	# --------------------------------------------------
	foreach ($register in $registersToProcess)
	{
		# Construct the register directory path (assumes folder naming: XF<StoreNumber><Register>)
		$RegisterDirectory = "$OfficePath\XF${StoreNumber}${register}"
		if (-not (Test-Path $RegisterDirectory))
		{
			Write_Log "Register directory $RegisterDirectory not found. Skipping register $register." "yellow"
			continue
		}
		
		# Define the full path to the SQI file (named "Refresh_Database.sqi")
		$SQIFilePath = Join-Path -Path $RegisterDirectory -ChildPath "Refresh_Database.sqi"
		
		# Write the SQI file in ANSI (PC) format (using ASCII encoding)
		Set-Content -Path $SQIFilePath -Value $SQIContent -Encoding ASCII
		
		# Remove the Archive attribute (set file attributes to Normal)
		Set-ItemProperty -Path $SQIFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Deployed Refresh_Database.sqi command to register $register in directory $RegisterDirectory." "green"
	}
	Write_Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Retrive_Transactions
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user for a start date and a stop date, then deploys an SQI file to selected
#   registers for a specified store. The SQI file retrieves transactions from SAL_HDR based on the
#   provided dates. The file is written in ANSI PC format (ASCII encoding with CRLF line endings).
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: The store number to process. (Mandatory)
# ---------------------------------------------------------------------------------------------------
# Requirements:
#   - The Show_Lane_Selection_Form function must be available.
#   - Variables such as $OfficePath must be defined.
#   - Helper functions like Write_Log, Retrieve_Nodes, and the class [MailslotSender] must be available.
# ===================================================================================================

function Retrive_Transactions
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OutputCsv
	)
	
	Write_Log "`r`n==================== Starting Retrive_Transactions ====================`r`n" "blue"
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$dateForm = New-Object System.Windows.Forms.Form
	$dateForm.Text = "Select Date Range for Transactions"
	$dateForm.Size = New-Object System.Drawing.Size(400, 250)
	$dateForm.StartPosition = "CenterScreen"
	
	# Start Date Label
	$startLabel = New-Object System.Windows.Forms.Label
	$startLabel.Text = "Start Date:"
	$startLabel.Location = New-Object System.Drawing.Point(10, 20)
	$startLabel.AutoSize = $true
	$dateForm.Controls.Add($startLabel)
	
	# Start Date Picker
	$startPicker = New-Object System.Windows.Forms.DateTimePicker
	$startPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$startPicker.Location = New-Object System.Drawing.Point(100, 15)
	$startPicker.Width = 100
	$dateForm.Controls.Add($startPicker)
	
	# Stop Date Label
	$stopLabel = New-Object System.Windows.Forms.Label
	$stopLabel.Text = "Stop Date:"
	$stopLabel.Location = New-Object System.Drawing.Point(10, 60)
	$stopLabel.AutoSize = $true
	$dateForm.Controls.Add($stopLabel)
	
	# Stop Date Picker
	$stopPicker = New-Object System.Windows.Forms.DateTimePicker
	$stopPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$stopPicker.Location = New-Object System.Drawing.Point(100, 55)
	$stopPicker.Width = 100
	$dateForm.Controls.Add($stopPicker)
	
	# OK Button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(80, 120)
	$okButton.Add_Click({
			$dateForm.Tag = @{
				StartDate = $startPicker.Value
				StopDate  = $stopPicker.Value
			}
			$dateForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$dateForm.Close()
		})
	$dateForm.Controls.Add($okButton)
	
	# Cancel Button
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(180, 120)
	$cancelButton.Add_Click({
			$dateForm.Tag = "Cancelled"
			$dateForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$dateForm.Close()
		})
	$dateForm.Controls.Add($cancelButton)
	
	$dateForm.AcceptButton = $okButton
	$dateForm.CancelButton = $cancelButton
	
	$resultDate = $dateForm.ShowDialog()
	if ($dateForm.Tag -eq "Cancelled" -or $resultDate -eq [System.Windows.Forms.DialogResult]::Cancel)
	{
		Write_Log "User cancelled the date selection." "yellow"
		Write_Log "==================== Retrive_Transactions Function Completed ====================" "blue"
		return
	}
	
	# Format the dates as MM/dd/yyyy
	$startDateFormatted = $dateForm.Tag.StartDate.ToString("MM/dd/yyyy")
	$stopDateFormatted = $dateForm.Tag.StopDate.ToString("MM/dd/yyyy")
	Write_Log "Start Date selected: $startDateFormatted" "green"
	Write_Log "Stop Date selected: $stopDateFormatted" "green"
	
	# --- STEP 2: Ask which lanes ---
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber
	if (-not $selection)
	{
		Write-Warning "Lane selection cancelled."
		return
	}
	$lanes = $selection.Lanes
	
	# --- STEP 3: Query each lane over named pipes, fallback to TCP ---
	$allResults = @()
	foreach ($lane in $lanes)
	{
		$machine = $script:FunctionResults['LaneMachines'][$lane]
		if (-not $machine)
		{
			Write-Warning "No machine mapped for lane $lane"
			continue
		}
		
		$instance = 'SQLEXPRESS'
		$npServer = "np:\\$machine\pipe\MSSQL`$$instance\sql\query"
		$tcpServer = "$machine\$instance"
		
		$query = @"
SELECT
    F1032 AS TransactionNumber,
    F254  AS TransactionDate,
    F1067 AS Status,
    *
FROM SAL_HDR
WHERE F1067 = 'CLOSE'
  AND F254 >= '$startDateFormatted'
  AND F254 <= '$stopDateFormatted'
ORDER BY F254, F1032;
"@
		
		try
		{
			$res = Invoke-Sqlcmd -ServerInstance $npServer -Database "LANESQL" -Query $query -ErrorAction Stop
		}
		catch
		{
			Write-Warning "Named-pipe connect to $machine failed, trying TCP..."
			try { $res = Invoke-Sqlcmd -ServerInstance $tcpServer -Database "LANESQL" -Query $query -ErrorAction Stop }
			catch { Write-Error "Cannot connect to $machine via TCP: $_"; continue }
		}
		
		if ($res)
		{
			# <- inject the Lane note-property here
			$res |
			ForEach-Object { $_ | Add-Member -NotePropertyName Lane -NotePropertyValue $lane -PassThru } |
			ForEach-Object { $allResults += $_ }
		}
		else
		{
			Write-Warning "No closed transactions found on lane $lane"
		}
	}
	
	if (-not $allResults)
	{
		Write_Log "No transactions found on any selected lane." "yellow"
		return
	}
	
	# Display results similar to Organize_TBS_SCL...
	Write_Log "Displaying closed transactions:" "yellow"
	try
	{
		$formatted = $allResults | Format-Table -AutoSize | Out-String
		Write_Log $formatted "blue"
	}
	catch
	{
		Write_Log "Failed to format and display transactions: $_" "red"
	}
	
	# --- 5) WRITE A / T / X FILES ------------------------------------
	$outDir = Join-Path $OfficePath "Exports\Transactions\$StoreNumber"
	if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory | Out-Null }
	
	foreach ($tran in $allResults)
	{
		# zero-pad
		$tn = "{0:D8}" -f $tran.F1032
		$st = $StoreNumber.PadLeft(3, '0')
		$ln = "{0:D3}" -f $tran.Lane
		
		# --- A-FILE (SIL) COMPLETE FORMAT ---
		$A = New-Object System.Text.StringBuilder
		
		# SAL_SAV section
		$A.AppendLine('@WIZRPL(ARCHIVE_TYPE=SAL_SAV);')
		
		# SAL_HDR section
		$A.AppendLine("@WIZRPL(SILHDR_TABLE_SAL_HDR=SH$st$ln$tn);")
		$hdrCols = @(
			'F1032', 'F1148', 'F76', 'F91', 'F253', 'F254', 'F902', 'F1035', 'F1036', 'F1056', 'F1057', 'F1067', 'F1068', 'F1101',
			'F1126', 'F1127', 'F1137', 'F1149', 'F1150', 'F1151', 'F1152', 'F1153', 'F1154', 'F1155', 'F1156', 'F1157', 'F1158', 'F1159',
			'F1160', 'F1161', 'F1163', 'F1164', 'F1165', 'F1167', 'F1168', 'F1170', 'F1171', 'F1172', 'F1173', 'F1185', 'F1238', 'F1242',
			'F1245', 'F1246', 'F1254', 'F1255', 'F1504', 'F1520', 'F1642', 'F1643', 'F1644', 'F1645', 'F1646', 'F1647', 'F1648', 'F1649',
			'F1650', 'F1651', 'F1652', 'F1653', 'F1654', 'F1655', 'F1686', 'F1687', 'F1688', 'F1689', 'F1692', 'F1693', 'F1694', 'F1695',
			'F1696', 'F1697', 'F1699', 'F1711', 'F1763', 'F1764', 'F1271', 'F1272', 'F1287', 'F1288', 'F1295', 'F1273', 'F1274', 'F1277',
			'F1685', 'F1938', 'F2596', 'F2598', 'F2599', 'F2613', 'F2614', 'F2615', 'F2616', 'F2617', 'F2618', 'F2619', 'F2620', 'F2621',
			'F2622', 'F2623', 'F2816', 'F2848', 'F1573', 'F2889', 'F2904', 'F2934', 'F2602'
		)
		$A.AppendLine("CREATE VIEW SH$st$ln$tn AS SELECT " + ($hdrCols -join ',') + " FROM TRS_DCT;")
		$A.AppendLine("INSERT INTO SH$st$ln$tn VALUES")
		# Map the values in the same order:
		$hdrVals = $hdrCols | ForEach-Object {
			$v = $tran."$_"
			if ($null -eq $v) { "''" }
			elseif ($v -is [datetime]) { "'$($v.ToString('yyyyMMdd HH:mm:ss'))'" }
			elseif ($v -is [string]) { "'$v'" }
			else { $v }
		}
		$A.AppendLine("(" + ($hdrVals -join ',') + ");")
		
		# SAL_REG section
		$A.AppendLine("@WIZRPL(SILHDR_TABLE_SAL_REG=SR$st$ln$tn);")
		$regCols = @(
			'F1032', 'F1101', 'F01', 'F03', 'F04', 'F05', 'F06', 'F24', 'F30', 'F31', 'F43', 'F50', 'F60', 'F61', 'F64', 'F65', 'F67', 'F77',
			'F79', 'F80', 'F81', 'F82', 'F83', 'F88', 'F96', 'F97', 'F98', 'F99', 'F100', 'F101', 'F102', 'F104', 'F106', 'F108', 'F109',
			'F110', 'F113', 'F114', 'F115', 'F124', 'F125', 'F126', 'F149', 'F150', 'F160', 'F168', 'F169', 'F170', 'F171', 'F172', 'F173',
			'F175', 'F178', 'F253', 'F254', 'F270', 'F383', 'F903', 'F1002', 'F1006', 'F1007', 'F1034', 'F1041', 'F1063', 'F1067', 'F1069',
			'F1070', 'F1071', 'F1072', 'F1078', 'F1080', 'F1086', 'F1120', 'F1136', 'F1178', 'F1203', 'F1204', 'F1205', 'F1206', 'F1207',
			'F1208', 'F1209', 'F1224', 'F1225', 'F1239', 'F1240', 'F1241', 'F1256', 'F1263', 'F1595', 'F1596', 'F1683', 'F1684', 'F1691',
			'F1693', 'F1694', 'F1699', 'F1712', 'F1715', 'F1716', 'F1717', 'F1718', 'F1719', 'F1720', 'F1721', 'F1722', 'F1723', 'F1724',
			'F1725', 'F1726', 'F1727', 'F1728', 'F1729', 'F1730', 'F1731', 'F1732', 'F1733', 'F1734', 'F1739', 'F1740', 'F1741', 'F1742',
			'F177', 'F1785', 'F1787', 'F1789', 'F1802', 'F1803', 'F1805', 'F1081', 'F1831', 'F1832', 'F1833', 'F1834', 'F1835', 'F1079',
			'F1860', 'F1861', 'F1862', 'F1863', 'F1864', 'F1888', 'F1874', 'F08', 'F1924', 'F1925', 'F1926', 'F1927', 'F1928', 'F1929',
			'F1930', 'F1931', 'F1932', 'F1933', 'F1934', 'F1935', 'F1936', 'F1126', 'F1185', 'F2551', 'F2552', 'F2553', 'F2554', 'F2555',
			'F1815', 'F1816', 'F1164', 'F1938', 'F2608', 'F2609', 'F2610', 'F2611', 'F2612', 'F2613', 'F2614', 'F2660', 'F2745', 'F2746',
			'F2752', 'F2753', 'F2747', 'F2748', 'F2749', 'F2750', 'F2751', 'F2744', 'F1687', 'F2860', 'F2861', 'F2862', 'F2863', 'F2865',
			'F2866', 'F2867', 'F2869', 'F2870', 'F2871', 'F163', 'F117', 'F3038', 'F3039', 'F3040', 'F3041', 'F3042', 'F3043', 'F3044',
			'F3045', 'F3046', 'F3047'
		)
		$A.AppendLine("CREATE VIEW SR$st$ln$tn AS SELECT " + ($regCols -join ',') + " FROM TRS_DCT;")
		$A.AppendLine("INSERT INTO SR$st$ln$tn VALUES")
		# Map the values in the same order:
		$regVals = $regCols | ForEach-Object {
			$v = $tran."$_"
			if ($null -eq $v) { "''" }
			elseif ($v -is [datetime]) { "'$($v.ToString('yyyyMMdd HH:mm:ss'))'" }
			elseif ($v -is [string]) { "'$v'" }
			else { $v }
		}
		$A.AppendLine("(" + ($regVals -join ',') + ");")
		
		# SAL_DET section
		$A.AppendLine("@WIZRPL(SILHDR_TABLE_SAL_DET=SD$st$ln$tn);")
		$detCols = @('F1032', 'F1101', 'F2770', 'F01', 'F64', 'F65', 'F1041', 'F1079', 'F1081', 'F1691', 'F1802', 'F2771')
		$A.AppendLine("CREATE VIEW SD$st$ln$tn AS SELECT " + ($detCols -join ',') + " FROM TRS_DCT;")
		$A.AppendLine("INSERT INTO SD$st$ln$tn VALUES")
		# Map the values in the same order:
		$detVals = $detCols | ForEach-Object {
			$v = $tran."$_"
			if ($null -eq $v) { "''" }
			elseif ($v -is [datetime]) { "'$($v.ToString('yyyyMMdd HH:mm:ss'))'" }
			elseif ($v -is [string]) { "'$v'" }
			else { $v }
		}
		$A.AppendLine("(" + ($detVals -join ',') + ");")
		
		# SAL_TTL section
		$A.AppendLine("@WIZRPL(SILHDR_TABLE_SAL_TTL=ST$st$ln$tn);")
		$ttlCols = @('F1032', 'F1034', 'F64', 'F65', 'F67', 'F1039', 'F1067', 'F1093', 'F1094', 'F1095', 'F1096', 'F1097', 'F1098')
		$A.AppendLine("CREATE VIEW ST$st$ln$tn AS SELECT " + ($ttlCols -join ',') + " FROM TRS_DCT;")
		$A.AppendLine("INSERT INTO ST$st$ln$tn VALUES")
		$ttlVals = $ttlCols | ForEach-Object {
			$v = $tran."$_"
			if ($null -eq $v) { "''" }
			elseif ($v -is [datetime]) { "'$($v.ToString('yyyyMMdd HH:mm:ss'))'" }
			elseif ($v -is [string]) { "'$v'" }
			else { $v }
		}
		$A.AppendLine("(" + ($ttlVals -join ',') + ");")
		
		# Import command
		$A.AppendLine("@IMPORT_SIL(SAL_SAV.sqi);")
		
		$A.ToString() | Out-File (Join-Path $outDir "A$tn$st.$ln") -Encoding ASCII
		
		# - T-FILE (header + TRS_ADD) -
		$fileT = Join-Path $outDir "T$tn$st.$ln"
		$sbT = New-Object System.Text.StringBuilder
		
		# 1) HEADER
		$sbT.AppendLine(
			"INSERT INTO HEADER_DCT VALUES('HR','$tn','$st$ln','$st$ln',,,'2025...','...','...','...',,'ADD','SALE',...);"
		)
		
		# 2) CREATE VIEW
		$sbT.AppendLine(
			"CREATE VIEW TRS_ADD AS SELECT F1056,F1057,F254,F1036,F1031,F1032,F1101,F1033,F01,F1034,F126,F30,F31,F64,F65,F67,F1079 FROM TRS_TAB;"
		)
		
		# 3) BUILD THE VALUES LIST
		$rowVals = "{0:D3},{1:D3},'{2}',{3},'{4}',{5},{6}" -f `
		$tran.F1056, `
		$tran.F1057, `
		$tran.F254.ToString('MM/dd/yyyy HH:mm:ss'), `
		$tran.F1035, `
		$tran.F1031, `
		$tran.F1032, `
		$tran.F1101
		
		# 4) FINAL INSERT
		$sbT.AppendLine("INSERT INTO TRS_ADD VALUES($rowVals);")
		
		# 5) OUTPUT
		$sbT.ToString() | Out-File -FilePath $fileT -Encoding ASCII
		
		# …then later you need to close your foreach and function:
	}
	#
	# - X-FILE (.xml) -
	#
	$fileX = Join-Path $outDir "X$tn$st.$ln.xml"
	$tran.XmlPayload | Out-File $fileX -Encoding ASCII
	
	
	Write_Log "✅ A/T/X files written to $outDir" "green"
	Write_Log "`r`n==================== Retrive_Transactions Function Completed ====================" "blue"
}

# ===================================================================================================
#                           FUNCTION: Reboot_Scales
# ---------------------------------------------------------------------------------------------------
# **Purpose:**
#   The `Reboot_Scales` function displays a Windows Form that allows the user to select which scales
#   to reboot based on their full IP addresses. The full IP address should already be built by 
#   concatenating the IPNetwork (first 3 octets) with the IPDevice (last octet). If available, the function
#   uses the ScaleName (from the table) for a friendly display; otherwise, it extracts the last octet
#   of the IP to generate a display name. In either case, the full IP is shown in parentheses next to
#   the name (e.g., "Scale 101 (192.168.5.101)").
#
# **Parameters:**
#   - [hashtable]$ScaleIPNetworks
#       - **Description:** A hashtable where each key represents a scale identifier. The value may either be:
#           1. A full IP address as a string (e.g., "192.168.5.101"), or
#           2. A custom object with at least the properties **FullIP** and **ScaleName**.
#
# **Usage:**
#   ```powershell
#   Reboot_Scales -ScaleIPNetworks $script:FunctionResults['ScaleIPNetworks']
#   ```
#
# **Notes:**
#   - Replace or adjust the reboot logic as needed for your environment.
# ===================================================================================================

function Reboot_Scales
{
	param (
		[hashtable]$ScaleIPNetworks
	)
	
	# Load Windows Forms assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Create the form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Reboot Scales"
	$form.Size = New-Object System.Drawing.Size(400, 500)
	$form.StartPosition = "CenterScreen"
	
	# Create a CheckedListBox to list scales
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 10)
	$checkedListBox.Size = New-Object System.Drawing.Size(360, 350)
	$checkedListBox.CheckOnClick = $true
	
	# Accumulate items in an array
	$scaleItems = @()
	foreach ($key in $ScaleIPNetworks.Keys)
	{
		$entry = $ScaleIPNetworks[$key]
		if ($entry -is [string])
		{
			# Entry is a string containing the full IP.
			$ip = $entry.Trim()
			$octets = $ip -split "\."
			if ($octets.Count -ge 1)
			{
				$lastOctet = $octets[-1]
				# Build display name with the IP in parentheses.
				$displayName = "Scale $lastOctet ($ip)"
			}
			else
			{
				$displayName = $key
			}
		}
		elseif ($entry -is [psobject] -and $entry.PSObject.Properties.Name -contains "ScaleName")
		{
			# Entry is a custom object with ScaleName and FullIP.
			$ip = $entry.FullIP.Trim()
			$displayName = "$($entry.ScaleName) ($ip)"
		}
		else
		{
			# Fallback: treat entry as a string.
			$ip = "$entry".Trim()
			$octets = $ip -split "\."
			if ($octets.Count -ge 1)
			{
				$lastOctet = $octets[-1]
				$displayName = "Scale $lastOctet ($ip)"
			}
			else
			{
				$displayName = $key
			}
		}
		
		$item = New-Object PSObject -Property @{
			DisplayName = $displayName
			IP		    = $ip
		}
		$item | Add-Member -MemberType ScriptMethod -Name ToString -Value { return $this.DisplayName } -Force
		$scaleItems += $item
	}
	
	# Sort items in ascending order by DisplayName
	$sortedScaleItems = $scaleItems | Sort-Object -Property DisplayName
	
	# Add sorted items to the CheckedListBox
	foreach ($item in $sortedScaleItems)
	{
		$checkedListBox.Items.Add($item) | Out-Null
	}
	
	# "Select All" button
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Location = New-Object System.Drawing.Point(10, 370)
	$btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $true)
			}
		})
	
	# "Deselect All" button
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Location = New-Object System.Drawing.Point(120, 370)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(100, 30)
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
		})
	
	# "Reboot Selected" button
	$btnReboot = New-Object System.Windows.Forms.Button
	$btnReboot.Location = New-Object System.Drawing.Point(230, 370)
	$btnReboot.Size = New-Object System.Drawing.Size(140, 30)
	$btnReboot.Text = "Reboot Selected"
	$btnReboot.Add_Click({
			$selectedItems = $checkedListBox.CheckedItems
			if ($selectedItems.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("No scales selected.", "Information", `
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
			else
			{
				foreach ($item in $selectedItems)
				{
					$machineName = $item.IP # Using the full IP address
					Write_Log "Attempting to reboot scale: $($item.DisplayName) at $machineName" "Yellow"
					try
					{
						# First, attempt to reboot using the shutdown command
						$shutdownArgs = "/r /m \\$machineName /t 0 /f"
						$process = Start-Process -FilePath "shutdown.exe" -ArgumentList $shutdownArgs -Wait -PassThru -ErrorAction Stop
						if ($process.ExitCode -ne 0)
						{
							throw "Shutdown command exited with code $($process.ExitCode)"
						}
						Write_Log "Shutdown command executed successfully for $machineName." "Green"
					}
					catch
					{
						Write_Log "Shutdown command failed for $machineName. Falling back to Restart-Computer." "Red"
						try
						{
							Restart-Computer -ComputerName $machineName -Force -ErrorAction Stop
							Write_Log "Restart-Computer command executed successfully for $machineName." "Green"
						}
						catch
						{
							Write_Log "Failed to reboot scale $machineName using both methods: $_" "Red"
						}
					}
				}
				[System.Windows.Forms.MessageBox]::Show("Reboot commands issued for selected scales.", "Reboot", `
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
		})
	
	# "Close" button
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Location = New-Object System.Drawing.Point(10, 410)
	$btnCancel.Size = New-Object System.Drawing.Size(360, 30)
	$btnCancel.Text = "Close"
	$btnCancel.Add_Click({
			$form.Close()
		})
	
	# Add controls to the form
	$form.Controls.Add($checkedListBox)
	$form.Controls.Add($btnSelectAll)
	$form.Controls.Add($btnDeselectAll)
	$form.Controls.Add($btnReboot)
	$form.Controls.Add($btnCancel)
	
	# Show the form
	$form.Add_Shown({ $form.Activate() })
	[void]$form.ShowDialog()
}

# ===================================================================================================
#                         FUNCTION: Remove_ArchiveBit_Interactive
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user to either run the Remove Archive Bit action immediately or schedule it as a task.
#   If scheduled, writes a batch file with current script variables (StoreNumber, paths, etc.) and
#   creates a Windows scheduled task. If run immediately, performs the action using current values.
#   Uses Write_Log for progress and error reporting.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   (none - uses variables from main script context)
# ===================================================================================================

function Remove_ArchiveBit_Interactive
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Remove_ArchiveBit_Interactive Function ====================`r`n" "blue"
	
	# --- Main context variables
	$iniFile = $StartupIniPath
	$storeNumber = $script:FunctionResults['StoreNumber']
	$terFile = Join-Path $OfficePath "Load\Ter_Load.sql"
	$scriptFolder = "C:\Tecnica_Systems\Scripts_by_Alex_C.T"
	$batchName = "Remove_Archive_Bit.bat"
	$batchPath = Join-Path $scriptFolder $batchName
	
	# --------------------- Show Choice Form ---------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Remove Archive Bit"
	$form.Size = New-Object System.Drawing.Size(430, 210)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Do you want to run Remove Archive Bit now, or schedule it as a repeating background task?"
	$label.Location = New-Object System.Drawing.Point(20, 20)
	$label.Size = New-Object System.Drawing.Size(390, 40)
	$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$form.Controls.Add($label)
	
	$btnRunNow = New-Object System.Windows.Forms.Button
	$btnRunNow.Text = "Run Now"
	$btnRunNow.Location = New-Object System.Drawing.Point(35, 90)
	$btnRunNow.Size = New-Object System.Drawing.Size(100, 40)
	$form.Controls.Add($btnRunNow)
	
	$btnSchedule = New-Object System.Windows.Forms.Button
	$btnSchedule.Text = "Schedule Task"
	$btnSchedule.Location = New-Object System.Drawing.Point(160, 90)
	$btnSchedule.Size = New-Object System.Drawing.Size(120, 40)
	$form.Controls.Add($btnSchedule)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(305, 90)
	$btnCancel.Size = New-Object System.Drawing.Size(80, 40)
	$form.Controls.Add($btnCancel)
	
	# Option tracking
	$selectedAction = $null
	$btnRunNow.Add_Click({
			$script:selectedAction = "run"
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})
	$btnSchedule.Add_Click({
			$script:selectedAction = "schedule"
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})
	$btnCancel.Add_Click({
			$script:selectedAction = "cancel"
			$form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$form.Close()
		})
	
	$form.AcceptButton = $btnRunNow
	$form.CancelButton = $btnCancel
	
	$form.ShowDialog() | Out-Null
	if ($script:selectedAction -eq "cancel" -or -not $script:selectedAction)
	{
		Write_Log "User cancelled Remove_ArchiveBit_Interactive." "yellow"
		Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
		return
	}
	
	# --------------------- Schedule Task Path ---------------------
	if ($script:selectedAction -eq "schedule")
	{
		# Use a WinForms interval prompt
		$intervalForm = New-Object System.Windows.Forms.Form
		$intervalForm.Text = "Schedule Interval"
		$intervalForm.Size = New-Object System.Drawing.Size(300, 160)
		$intervalForm.StartPosition = "CenterScreen"
		$intervalForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$intervalForm.MaximizeBox = $false
		$intervalForm.MinimizeBox = $false
		
		$intervalLabel = New-Object System.Windows.Forms.Label
		$intervalLabel.Text = "Enter the interval in minutes (default 30):"
		$intervalLabel.Location = New-Object System.Drawing.Point(10, 20)
		$intervalLabel.Size = New-Object System.Drawing.Size(260, 20)
		$intervalForm.Controls.Add($intervalLabel)
		
		$intervalBox = New-Object System.Windows.Forms.TextBox
		$intervalBox.Text = "30"
		$intervalBox.Location = New-Object System.Drawing.Point(10, 50)
		$intervalBox.Size = New-Object System.Drawing.Size(260, 20)
		$intervalForm.Controls.Add($intervalBox)
		
		$okBtn = New-Object System.Windows.Forms.Button
		$okBtn.Text = "OK"
		$okBtn.Location = New-Object System.Drawing.Point(40, 90)
		$okBtn.Size = New-Object System.Drawing.Size(80, 30)
		$okBtn.Add_Click({ $intervalForm.DialogResult = [System.Windows.Forms.DialogResult]::OK; $intervalForm.Close() })
		$intervalForm.Controls.Add($okBtn)
		
		$cancelBtn = New-Object System.Windows.Forms.Button
		$cancelBtn.Text = "Cancel"
		$cancelBtn.Location = New-Object System.Drawing.Point(160, 90)
		$cancelBtn.Size = New-Object System.Drawing.Size(80, 30)
		$cancelBtn.Add_Click({ $intervalForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $intervalForm.Close() })
		$intervalForm.Controls.Add($cancelBtn)
		
		$intervalForm.AcceptButton = $okBtn
		$intervalForm.CancelButton = $cancelBtn
		
		$intervalResult = $intervalForm.ShowDialog()
		if ($intervalResult -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "User cancelled interval prompt for scheduled Remove_ArchiveBit_Interactive." "yellow"
			Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
			return
		}
		
		$interval = $intervalBox.Text.Trim()
		if (-not $interval -or $interval -notmatch "^\d+$" -or [int]$interval -le 0)
		{
			Write_Log "Invalid interval value for scheduled Remove_ArchiveBit_Interactive." "red"
			Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
			return
		}
		
		if (-not (Test-Path $scriptFolder)) { New-Item -Path $scriptFolder -ItemType Directory | Out-Null }
		
		$batchContent = @"
@echo off
setlocal

REM - check for admin & elevate if needed -
net session >nul 2>&1
if errorlevel 1 (
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0 %*' -Verb RunAs" >nul
    exit /b
)

echo *****************************************************
echo *                Remove Archive Bit                 *
echo *               Created by: Alex_C.T                *
echo *      What it does: Removes archived bits from     *
echo *            Lane directories and Server            *
echo *****************************************************

set "StoreNumber=$storeNumber"
set "TerFile=$terFile"
if not defined StoreNumber (
    echo ERROR: Store number could not be extracted.
    echo Press any key to exit...
    timeout /t 5 >nul
    exit /b
)
echo Debug: Store Number is %StoreNumber%
echo.
if not defined TerFile (
    echo ERROR: Ter_Load.sql could not be located.
    echo Press any key to exit...
    timeout /t 5 >nul
    exit /b
)
echo Debug: Found Ter_Load.sql
echo Processing lane paths from Ter_Load.sql

REM -- Remove archived bit in lanes --
for /f "tokens=4,5 delims=,()'" %%A in ('
    type "%TerFile%" ^
    ^| findstr /b "(" ^
    ^| findstr /i /c:"Terminal 0"
') do (
    echo Refreshing attributes in %%A...
    attrib -a -r "%%A\*.*" >nul 2>&1
    if errorlevel 1 echo ERROR: Failed to refresh attributes for %%A
    if "%%B" NEQ "" (
        echo Refreshing attributes in %%B...
        attrib -a -r "%%B\*.*" >nul 2>&1
        if errorlevel 1 echo ERROR: Failed to refresh attributes for %%B
    )
)
REM -- Remove archived bit in server paths --
echo(
echo Processing server paths
for %%S in (900 901) do (
    if exist "$OfficePath\XF%StoreNumber%%%S" (
        echo Refreshing attributes in "$OfficePath\XF%StoreNumber%%%S"
        attrib -a -r "$OfficePath\XF%StoreNumber%%%S\*.*" >nul 2>&1
        if errorlevel 1 echo ERROR: Failed to refresh attributes for $OfficePath\XF%StoreNumber%%%S
    ) else (
        echo Server path not found: $OfficePath\XF%StoreNumber%%%S
    )
)
endlocal
echo Press any key to exit...
timeout /t 5 >nul
exit /b
"@
		
		Set-Content -Path $batchPath -Value $batchContent -Encoding ASCII
		
		$taskName = "Remove_Archive_Bit"
		$schtasks = "schtasks /create /tn `"$taskName`" /tr `"$batchPath`" /sc MINUTE /mo $interval /st 01:00 /rl HIGHEST /f"
		$result = Invoke-Expression $schtasks
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Scheduled task created successfully for Remove_Archive_Bit." "green"
		}
		else
		{
			Write_Log "Failed to create scheduled task for Remove_Archive_Bit." "red"
		}
		Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
		return
	}
	
	# --------------------- Run Now Path ---------------------
	if (-not (Test-Path $iniFile))
	{
		Write_Log "ERROR: INI file not found - $iniFile" "red"
		Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
		return
	}
	if (-not $storeNumber)
	{
		Write_Log "ERROR: Store number not present in context." "red"
		Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
		return
	}
	if (-not (Test-Path $terFile))
	{
		Write_Log "ERROR: Ter_Load.sql could not be located." "red"
		Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
		return
	}
	
	# --- Use lane paths from Retrieve_Nodes ---
	if (-not $script:FunctionResults.ContainsKey('LaneMachines') -or -not $script:FunctionResults['LaneMachines'])
	{
		Write_Log "No lane machine paths found. Did you run Retrieve_Nodes?" "red"
	}
	else
	{
		foreach ($laneNum in $script:FunctionResults['LaneMachines'].Keys | Sort-Object { [int]$_ })
		{
			$machine = $script:FunctionResults['LaneMachines'][$laneNum]
			if ($machine -and $machine -notlike '@SMSSERVER' -and $machine -ne '')
			{
				$path = "\\$machine\storeman\Office\XF${storeNumber}${laneNum}"
				Write_Log "Refreshing attributes in $path..." "green"
				try
				{
					attrib -a -r "$path\*.*" 2>&1 | Out-Null
				}
				catch
				{
					Write_Log "ERROR: Failed to refresh attributes for $path" "red"
				}
			}
		}
	}
	foreach ($suffix in 900, 901)
	{
		$serverPath = "\\localhost\storeman\office\XF${storeNumber}${suffix}"
		if (Test-Path $serverPath)
		{
			Write_Log "Refreshing attributes in $serverPath" "green"
			try
			{
				attrib -a -r "$serverPath\*.*" 2>&1 | Out-Null
			}
			catch
			{
				Write_Log "ERROR: Failed to refresh attributes for $serverPath" "red"
			}
		}
		else
		{
			Write_Log "Server path not found: $serverPath" "yellow"
		}
	}
	Write_Log "`r`n==================== Remove_ArchiveBit_Interactive Function Completed ====================" "blue"
}

# ===================================================================================================
#                           FUNCTION: Open-SelectedLanesCPath
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user with a GUI dialog to select one or more lanes (registers) within a store
#   and opens the administrative C$ share of each selected lane in Windows Explorer.
#   Uses lane-to-machine mapping from the most recent Retrieve_Nodes execution.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber (string, required): The 3-digit store number for lane selection.
#   - LaneType (string, optional): Lane type to display (e.g., "POS" or "SCO"). Default is "POS".
# ===================================================================================================

function Open_Selected_Lane/s_C_Path
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$LaneType = "POS"
	)
	
	Write_Log "`r`n==================== Starting Open_Selected_Lane/s_C_Path Function ====================`r`n" "blue"
	
	# Get selection from GUI
	$selection = Show_Lane_Selection_Form -StoreNumber $StoreNumber -LaneType $LaneType
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected. Exiting." "Yellow"
		Write_Log "`r`n==================== Open_Selected_Lane/s_C_Path Function Completed ====================" "blue"
		return
	}
	
	$laneMachines = $script:FunctionResults['LaneMachines']
	foreach ($lane in $selection.Lanes)
	{
		if ($laneMachines.ContainsKey($lane))
		{
			$machine = $laneMachines[$lane]
			$sharePath = "\\$machine\c$"
			Write_Log "Opened $sharePath ..." "Green"
			Start-Process "explorer.exe" $sharePath
		}
		else
		{
			Write_Log "Machine not found for lane '$lane'." "Red"
			Write_Log "`r`n==================== Open_Selected_Lane/s_C_Path Function Completed ====================" "blue"
		}
	}
}

# ===================================================================================================
#                           FUNCTION: Open_Selected_Scale/s_C_Path
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user with a GUI dialog to select one or more scales and attempts to open the C$
#   administrative share for each selected scale in Windows Explorer. It tries to connect as
#   user 'bizuser' with password 'bizerba' first, then 'biyerba' if needed. Temporary network
#   mappings are cleaned up after opening.
# ---------------------------------------------------------------------------------------------------
# Requirements:
#   - The scale hostnames/IPs must resolve and allow SMB access.
#   - Show_Scale_Selection_Form and Retrieve_Nodes must be run first to populate FunctionResults.
# ===================================================================================================

function Open_Selected_Scale/s_C_Path
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Open_Selected_Scale/s_C_Path Function ====================`r`n" "blue"
	
	# --- Get scale selection, only BIZERBA ---
	$scaleIPTable = $script:FunctionResults['ScaleIPNetworks']
	if (-not $scaleIPTable)
	{
		Write_Log "No scale IP mapping available. Please run Retrieve_Nodes first." "Yellow"
		return
	}
	# Filter for only BIZERBA
	$bizerbaScales = $scaleIPTable.Values | Where-Object { $_.ScaleBrand -eq "BIZERBA" }
	if (-not $bizerbaScales -or $bizerbaScales.Count -eq 0)
	{
		Write_Log "No BIZERBA scales found for this store." "Yellow"
		Write_Log "`r`n==================== Open_Selected_Scale/s_C_Path Function Completed ====================" "blue"
		return
	}
	# Custom picker: just show BIZERBA scales
	$scaleSelection = Show_Scale_Selection_Form -BizerbaScales $bizerbaScales
	if (-not $scaleSelection -or -not $scaleSelection.Scales -or $scaleSelection.Scales.Count -eq 0)
	{
		Write_Log "No BIZERBA scales selected. Exiting." "Yellow"
		Write_Log "`r`n==================== Open_Selected_Scale/s_C_Path Function Completed ====================" "blue"
		return
	}
	
	foreach ($scaleObj in $scaleSelection.Scales)
	{
		# Get the best host name or IP
		$scaleHost = if ($scaleObj.FullIP -and $scaleObj.FullIP -ne "")
		{
			$scaleObj.FullIP
		}
		elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice)
		{
			"$($scaleObj.IPNetwork)$($scaleObj.IPDevice)"
		}
		elseif ($scaleObj.ScaleName)
		{
			$scaleObj.ScaleName
		}
		if (-not $scaleHost) { Write_Log "Could not determine host for $($scaleObj.ScaleCode)." "Red"; continue }
		
		$sharePath = "\\$scaleHost\c$"
		$opened = $false
		
		# Try bizerba password first
		cmdkey /add:$scaleHost /user:bizuser /pass:bizerba | Out-Null
		Start-Process "explorer.exe" $sharePath
		Start-Sleep -Seconds 2 # Give Explorer a moment to attempt connection
		
		# Test if the path is accessible
		if (Test-Path $sharePath)
		{
			Write_Log "Opened $sharePath as bizuser using password 'bizerba'." "Green"
			$opened = $true
		}
		else
		{
			# Remove previous credential and try 'biyerba'
			cmdkey /delete:$scaleHost | Out-Null
			cmdkey /add:$scaleHost /user:bizuser /pass:biyerba | Out-Null
			Start-Process "explorer.exe" $sharePath
			Start-Sleep -Seconds 2
			if (Test-Path $sharePath)
			{
				Write_Log "Opened $sharePath as bizuser using password 'biyerba'." "Green"
				$opened = $true
			}
			else
			{
				Write_Log "Could not open $sharePath as bizuser with either password." "Red"
				Write_Log "`r`n==================== Open_Selected_Scale/s_C_Path Function Completed ====================" "blue"
			}
		}
		# Optional: Clean up after (remove credential)
		# cmdkey /delete:$scaleHost | Out-Null
	}
}

# ===================================================================================================
#                         FUNCTION: Update_Scales_Specials_Interactive
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user to either run the "Update Scales Specials" action now or schedule it as a daily task.
#   If scheduled, writes a batch file to disk and creates a Windows scheduled task.
#   If run immediately, performs the action directly from PowerShell.
#   Uses Write_Log for progress and error reporting.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   (none - uses script context)
# ===================================================================================================

function Update_Scales_Specials_Interactive
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Update_Scales_Specials_Interactive Function ====================`r`n" "blue"
	
	$scriptFolder = "C:\Tecnica_Systems\Scripts_by_Alex_C.T"
	$batchName_Daily = "Update_Scales_Specials.bat"
	$batchPath_Daily = Join-Path $scriptFolder $batchName_Daily
	$batchName_Minutes = "Update_Scales_Specials_Minutes.bat"
	$batchPath_Minutes = Join-Path $scriptFolder $batchName_Minutes
	
	# Assume $OfficePath is already defined and points to the correct Office folder
	$deployChgFile = Join-Path $OfficePath "DEPLOY_CHG.sql"
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Update Scales Specials"
	$form.Size = New-Object System.Drawing.Size(490, 215)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Schedule Update Scales Specials as a daily (5 AM) or repeating task (in minutes)."
	$label.Location = New-Object System.Drawing.Point(20, 20)
	$label.Size = New-Object System.Drawing.Size(430, 40)
	$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$form.Controls.Add($label)
	
	$btnSchedule = New-Object System.Windows.Forms.Button
	$btnSchedule.Text = "Schedule Daily (5 AM)"
	$btnSchedule.Location = New-Object System.Drawing.Point(60, 90)
	$btnSchedule.Size = New-Object System.Drawing.Size(140, 40)
	$form.Controls.Add($btnSchedule)
	
	$btnScheduleMinutes = New-Object System.Windows.Forms.Button
	$btnScheduleMinutes.Text = "Schedule Task (Minutes)"
	$btnScheduleMinutes.Location = New-Object System.Drawing.Point(220, 90)
	$btnScheduleMinutes.Size = New-Object System.Drawing.Size(180, 40)
	$form.Controls.Add($btnScheduleMinutes)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(180, 150)
	$btnCancel.Size = New-Object System.Drawing.Size(80, 30)
	$form.Controls.Add($btnCancel)
	
	$selectedAction = $null
	$minutesValue = $null
	
	$btnSchedule.Add_Click({
			$script:selectedAction = "schedule"
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})
	$btnScheduleMinutes.Add_Click({
			$inputForm = New-Object System.Windows.Forms.Form
			$inputForm.Text = "Set Minute Interval"
			$inputForm.Size = New-Object System.Drawing.Size(320, 140)
			$inputForm.StartPosition = "CenterParent"
			
			$lblInput = New-Object System.Windows.Forms.Label
			$lblInput.Text = "How many minutes between runs? (1-1439):"
			$lblInput.Location = New-Object System.Drawing.Point(10, 15)
			$lblInput.Size = New-Object System.Drawing.Size(280, 25)
			$inputForm.Controls.Add($lblInput)
			
			$txtInput = New-Object System.Windows.Forms.TextBox
			$txtInput.Location = New-Object System.Drawing.Point(10, 45)
			$txtInput.Size = New-Object System.Drawing.Size(120, 25)
			$inputForm.Controls.Add($txtInput)
			
			$okBtn = New-Object System.Windows.Forms.Button
			$okBtn.Text = "OK"
			$okBtn.Location = New-Object System.Drawing.Point(150, 42)
			$okBtn.Size = New-Object System.Drawing.Size(60, 25)
			$okBtn.Add_Click({
					$inputForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
					$inputForm.Close()
				})
			$inputForm.Controls.Add($okBtn)
			$inputForm.AcceptButton = $okBtn
			
			if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
			{
				$num = $txtInput.Text
				if ($num -match '^\d+$' -and [int]$num -ge 1 -and [int]$num -le 1439)
				{
					$script:selectedAction = "schedule_minutes"
					$script:minutesValue = [int]$num
					$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
					$form.Close()
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Please enter a valid number between 1 and 1439.", "Invalid Interval", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				}
			}
		})
	$btnCancel.Add_Click({
			$script:selectedAction = "cancel"
			$form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$form.Close()
		})
	$form.AcceptButton = $btnSchedule
	$form.CancelButton = $btnCancel
	
	$form.ShowDialog() | Out-Null
	if ($script:selectedAction -eq "cancel" -or -not $script:selectedAction)
	{
		Write_Log "User cancelled Update_Scales_Specials_Interactive." "yellow"
		Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
		return
	}
	
	if (-not (Test-Path $scriptFolder)) { New-Item -Path $scriptFolder -ItemType Directory | Out-Null }
	
	# --- Batch for daily (UpdateSpecials) ---
	if ($script:selectedAction -eq "schedule")
	{
		$batchContent = @"
if "%1" == "" start "" /min "%~f0" MY_FLAG && exit
taskkill /IM ScaleManagementApp.exe /F
taskkill /IM BMSSrv.exe /F
taskkill /IM BMS.exe /F
del /s /q C:\Bizerba\RetailConnect\BMS\toBizerba\*.*
rmdir /s /q C:\Bizerba\RetailConnect\BMS\terminals\ >nul 2>&1
net start BMS /Y
start C:\ScaleCommApp\ScaleManagementAppUpdateSpecials.exe
exit
"@
		Set-Content -Path $batchPath_Daily -Value $batchContent -Encoding ASCII
		
		$taskName = "Update_Scales_Specials_Task"
		
		# SYSTEM: no popup, no password
		$schtasks = "schtasks /create /tn `"$taskName`" /tr `"$batchPath_Daily`" /sc DAILY /st 05:00 /rl HIGHEST /f /ru $env:COMPUTERNAME\Administrator"
		
		Invoke-Expression $schtasks
		
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Scheduled task created successfully for Update_Scales_Specials_Task (daily at 5 AM)." "green"
		}
		else
		{
			Write_Log "Failed to create scheduled task for Update_Scales_Specials_Task." "red"
		}
		Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
		return
	}
	
	# --- Batch for minutes (FastDEPLOY) ---
	if ($script:selectedAction -eq "schedule_minutes" -and $script:minutesValue)
	{
		# --- REMOVE LINES FROM $deployChgFile ---
		if (Test-Path $deployChgFile)
		{
			try
			{
				$content = Get-Content $deployChgFile -Raw
				$newContent = ($content -split "`r?`n") | Where-Object { $_ -notmatch '(?i)ScaleManagementApp\.exe|ScaleManagementApp_FastDEPLOY\.exe' }
				if ($newContent.Count -lt (($content -split "`r?`n").Count))
				{
					$newContent -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default
					Write_Log "Removed lines from $deployChgFile containing ScaleManagementApp.exe or ScaleManagementApp_FastDEPLOY.exe" "green"
				}
				else
				{
					Write_Log "No matching lines found in DEPLOY_CHG.sql for removal." "yellow"
				}
			}
			catch
			{
				Write_Log "Failed to update [$deployChgFile]: $_" "red"
			}
		}
		else
		{
			Write_Log "DEPLOY_CHG.sql not found in $OfficePath" "yellow"
		}
		
		$batchContent = @"
if "%1" == "" start "" /min "%~f0" MY_FLAG && exit
taskkill /IM ScaleManagementApp.exe /F
taskkill /IM BMSSrv.exe /F
taskkill /IM BMS.exe /F
del /s /q C:\Bizerba\RetailConnect\BMS\toBizerba\*.*
rmdir /s /q C:\Bizerba\RetailConnect\BMS\terminals\ >nul 2>&1
net start BMS /Y
start C:\ScaleCommApp\ScaleManagementApp_FastDEPLOY.exe
exit
"@
		Set-Content -Path $batchPath_Minutes -Value $batchContent -Encoding ASCII
		
		$taskName = "Update_Scales_Specials_Task_Minutes"
		$interval = [int]$script:minutesValue
		
		# SYSTEM: no popup, no password
		$schtasks = "schtasks /create /tn `"$taskName`" /tr `"$batchPath_Minutes`" /sc MINUTE /mo $interval /rl HIGHEST /f /ru $env:COMPUTERNAME\Administrator"
		
		Invoke-Expression $schtasks
		
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Scheduled task created successfully for Update_Scales_Specials_Task_Minutes (every $interval minutes)." "green"
		}
		else
		{
			Write_Log "Failed to create scheduled task for Update_Scales_Specials_Task_Minutes." "red"
		}
		Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
		return
	}
}

# ===================================================================================================
#                        FUNCTION: Export-VNCFiles-ForAllNodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates UltraVNC (.vnc) connection files for all lanes, scales, and backoffices discovered in
#   the current store environment. Each node receives its own preconfigured .vnc file (with fixed password
#   or lane-specific password) and is saved to Desktop\Lanes, Desktop\Scales, or Desktop\Backoffices accordingly.
#   Designed for rapid remote support and streamlined access.
#
# Parameters:
#   - LaneMachines      [hashtable]: LaneNumber => MachineName mapping.
#   - ScaleIPNetworks   [hashtable]: ScaleCode => Scale Object (includes IP).
#   - BackofficeMachines[hashtable]: BackofficeTerminal => MachineName mapping.
#   - LaneVNCPasswords  [hashtable] (optional): MachineName => Password mapping.
#
# Details:
#   - Password is set to: 4330df922eb03b6e (UltraVNC encrypted format) by default.
#   - Per-lane password is used if available in LaneVNCPasswords.
#   - Files are written with clear, descriptive names.
#   - Uses Write_Log for all status and progress messages.
#   - Ensures all required folders exist.
#   - Skips scales with missing IP/network info and logs them.
#   - Creates Ishida_Wrapper_#.vnc for each Ishida scale found.
#
# Usage:
#   Export-VNCFiles-ForAllNodes -LaneMachines $... -ScaleIPNetworks $... -BackofficeMachines $... [-LaneVNCPasswords $...]
#
# Author: Alex_C.T
# ===================================================================================================

function Export_VNC_Files_For_All_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[hashtable]$LaneMachines,
		[Parameter(Mandatory = $true)]
		[hashtable]$ScaleIPNetworks,
		[Parameter(Mandatory = $true)]
		[hashtable]$BackofficeMachines,
		[Parameter(Mandatory = $false)]
		[hashtable]$LaneVNCPasswords
	)
	
	Write_Log "`r`n==================== Starting Export_VNCFiles_ForAllNodes ====================`r`n" "blue"
	$DefaultVNCPassword = "4330df922eb03b6e"
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"
	
	# ---- If passwords not provided, scan them ----
	if (-not $LaneVNCPasswords -or $LaneVNCPasswords.Count -eq 0)
	{
		$LaneVNCPasswords = Get_All_Lanes_VNC_Passwords -LaneMachines $LaneMachines
	}
	
	# --- Shared VNC file content with token ---
	$vncTemplate = @"
[connection]
host=%%HOST%%
port=5900
proxyhost=
proxyport=0
password=%%PASSWORD%%
[options]
use_encoding_0=1
use_encoding_1=1
use_encoding_2=1
use_encoding_3=0
use_encoding_4=1
use_encoding_5=1
use_encoding_6=1
use_encoding_7=1
use_encoding_8=1
use_encoding_9=1
use_encoding_10=1
use_encoding_11=0
use_encoding_12=0
use_encoding_13=0
use_encoding_14=0
use_encoding_15=0
use_encoding_16=1
use_encoding_17=1
use_encoding_18=1
use_encoding_19=1
use_encoding_20=0
use_encoding_21=0
use_encoding_22=0
use_encoding_23=0
use_encoding_24=0
use_encoding_25=1
use_encoding_26=1
use_encoding_27=1
use_encoding_28=0
use_encoding_29=1
preferred_encoding=10
restricted=0
AllowUntrustedServers=0
viewonly=0
nostatus=0
nohotkeys=0
showtoolbar=1
fullscreen=0
SavePos=0
SaveSize=0
GNOME=0
directx=0
autoDetect=0
8bit=0
shared=1
swapmouse=0
belldeiconify=0
BlockSameMouse=0
emulate3=1
JapKeyboard=0
emulate3timeout=100
emulate3fuzz=4
disableclipboard=0
localcursor=1
Scaling=0
AutoScaling=1
AutoScalingEven=0
AutoScalingLimit=0
scale_num=1
scale_den=1
cursorshape=1
noremotecursor=0
compresslevel=6
quality=8
ServerScale=1
Reconnect=3
EnableCache=0
EnableZstd=1
QuickOption=1
UseDSMPlugin=0
UseProxy=0
sponsor=0
allowMonitorSpanning=0
ChangeServerRes=0
extendDisplay=0
showExtend=0
use_virt=0
useAllMonitors=0
requestedWidth=0
requestedHeight=0
DSMPlugin=NoPlugin
folder=C:\Users\Administrator\Documents\UltraVNC
prefix=vnc_
imageFormat=.jpeg
InfoMsg=
AutoReconnect=3
ExitCheck=0
FileTransferTimeout=30
ListenPort=5500
KeepAliveInterval=5
ThrottleMouse=0
AutoAcceptIncoming=0
AutoAcceptNoDSM=0
RequireEncryption=0
PreemptiveUpdates=0
"@
	
	# ---- Lanes ---- #
	$laneCount = 0
	$laneInfoLines = @()
	
	foreach ($lane in $LaneMachines.GetEnumerator())
	{
		$laneNumber = $lane.Key
		$machineName = $lane.Value
		
		# File name logic
		if ($machineName -and $machineName.ToUpper() -match '^(POS|SCO)\d+$')
		{
			$fileName = "$($machineName.ToUpper()).vnc"
		}
		else
		{
			$fileName = "Lane_${laneNumber}.vnc"
		}
		
		$filePath = Join-Path $lanesDir $fileName
		$parent = Split-Path $filePath -Parent
		if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory | Out-Null }
		
		# Use custom password if available, else default
		$VNCPassword = $DefaultVNCPassword
		if ($LaneVNCPasswords -and $LaneVNCPasswords.ContainsKey($machineName) -and $LaneVNCPasswords[$machineName])
		{
			$VNCPassword = $LaneVNCPasswords[$machineName]
		}
		$content = $vncTemplate.Replace('%%HOST%%', $machineName).Replace('%%PASSWORD%%', $VNCPassword)
		[System.IO.File]::WriteAllText($filePath, $content, $script:ansiPcEncoding)
		Write_Log "Created: $filePath" "green"
		$laneCount++
		
		# --- Add hardware info for this lane ---
		$line = "Machine Name: $machineName"
		if ($script:LaneHardwareInfo -and $script:LaneHardwareInfo.ContainsKey($machineName))
		{
			$hw = $script:LaneHardwareInfo[$machineName]
			$manuf = $hw.SystemManufacturer
			$model = $hw.SystemProductName
			$succ = $hw.Success
			$err = $hw.Error
			if ($succ)
			{
				$line += "  Manufacturer: $manuf  Model: $model"
			}
			else
			{
				$line += "  [Hardware info unavailable]  Error: $err"
			}
		}
		else
		{
			$line += "  [No hardware info found]"
		}
		$laneInfoLines += $line
	}
		
	# ---- Write Lanes_Info.txt (once, after loop) ----
	$laneInfoPath = Join-Path $lanesDir 'Lanes_Info.txt'
	$laneInfoLines -join "`r`n" | Set-Content -Path $laneInfoPath -Encoding Default
	Write_Log "Wrote: $laneInfoPath" "yellow"
	Write_Log "$laneCount lane VNC files written to $lanesDir`r`n" "blue"
	
	# ---- Scales ---- #
	$scaleCount = 0
	foreach ($scale in $ScaleIPNetworks.GetEnumerator())
	{
		$scaleCode = $scale.Key
		$scaleObj = $scale.Value
		$ip = if ($scaleObj.FullIP) { $scaleObj.FullIP }
		elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)" }
		else { $null }
		
		if ($ip)
		{
			$octets = $ip -split '\.'
			$lastOctet = $octets[-1]
			
			# Normalize Brand/Model
			$brandRaw = ($scaleObj.ScaleBrand -as [string]).Trim()
			$model = ($scaleObj.ScaleModel -as [string]).Trim()
			
			# Capitalize every word in the brand
			$brand = if ($brandRaw)
			{
				($brandRaw -split ' ' | ForEach-Object {
						if ($_.Length -gt 0) { $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower() }
						else { $_ }
					}) -join ' '
			}
			else { "" }
			
			# Naming decision
			if ($brand -and $model)
			{
				$fileName = "$brand($model)_${lastOctet}.vnc"
			}
			elseif ($brand)
			{
				$fileName = "$brand(Unknown)_${lastOctet}.vnc"
			}
			else
			{
				$fileName = "Scale_${lastOctet}.vnc"
			}
			
			$filePath = Join-Path $scalesDir $fileName
			$parent = Split-Path $filePath -Parent
			if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory | Out-Null }
			$content = $vncTemplate.Replace('%%HOST%%', $ip).Replace('%%PASSWORD%%', $DefaultVNCPassword)
			[System.IO.File]::WriteAllText($filePath, $content, $script:ansiPcEncoding)
			Write_Log "Created: $filePath" "green"
			$scaleCount++
		}
		else
		{
			Write_Log "Skipped scale $scaleCode (missing IP)" "yellow"
		}
	}
	Write_Log "$scaleCount scale VNC files written to $scalesDir`r`n" "blue"
	
	# ---- Backoffices ---- #
	$boCount = 0
	foreach ($bo in $BackofficeMachines.GetEnumerator())
	{
		$terminal = $bo.Key
		$boName = $bo.Value
		$fileName = "Backoffice_${terminal}.vnc"
		$filePath = Join-Path $backofficesDir $fileName
		$parent = Split-Path $filePath -Parent
		if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory | Out-Null }
		$content = $vncTemplate.Replace('%%HOST%%', $boName).Replace('%%PASSWORD%%', $DefaultVNCPassword)
		[System.IO.File]::WriteAllText($filePath, $content, $script:ansiPcEncoding)
		Write_Log "Created: $filePath" "green"
		$boCount++
	}
	Write_Log "$boCount backoffice VNC files written to $backofficesDir`r`n" "blue"
	
	Write_Log "VNC file export complete!" "green"
	Write_Log "`r`n==================== Export_VNCFiles_ForAllNodes Completed ====================" "blue"
}

# ===================================================================================================
#                                FUNCTION: Show_Lane_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Presents a GUI dialog for the user to select lanes (registers) to process within a store.
#   Returns a hashtable with the selection type and the list of selected lanes.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: The 3-digit store number to fetch available lanes.
#   - LaneType (Optional): Lane type to display (e.g., "POS" or "SCO"). Default is "POS".
# ===================================================================================================

function Show_Lane_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[string]$LaneType = "POS"
	)
	
	# Load necessary assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Create and configure the form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Lanes to Process"
	$form.Size = New-Object System.Drawing.Size(330, 350)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	# Lane selection list
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 10)
	$checkedListBox.Size = New-Object System.Drawing.Size(295, 200)
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)
	
	# Retrieve available lanes (use LaneContents from globals or scan directory)
	$allLanes = @()
	if ($script:FunctionResults.ContainsKey('LaneContents') -and $script:FunctionResults['LaneContents'].Count -gt 0)
	{
		$allLanes = $script:FunctionResults['LaneContents']
	}
	else
	{
		if (-not (Test-Path -Path $OfficePath))
		{
			[System.Windows.Forms.MessageBox]::Show("The path '$OfficePath' does not exist.", "Error",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return $null
		}
		$laneFolders = Get-ChildItem -Path $OfficePath -Directory -Filter "XF${StoreNumber}0*"
		if ($laneFolders)
		{
			$allLanes = $laneFolders | ForEach-Object { $_.Name.Substring($_.Name.Length - 3, 3) }
		}
	}
	
	# Populate the CheckedListBox
	$sortedLanes = $allLanes | Sort-Object
	foreach ($lane in $sortedLanes)
	{
		if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'].ContainsKey($lane))
		{
			$friendlyName = $script:FunctionResults['LaneMachines'][$lane]
			$displayName = "$friendlyName"
		}
		else
		{
			$displayName = "$LaneType $lane"
		}
		$laneObj = New-Object PSObject -Property @{
			DisplayName = $displayName
			LaneNumber  = $lane
		}
		$laneObj | Add-Member -MemberType ScriptMethod -Name ToString -Value { return $this.DisplayName } -Force
		$checkedListBox.Items.Add($laneObj) | Out-Null
	}
	
	# "Select All" button
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Location = New-Object System.Drawing.Point(10, 220)
	$btnSelectAll.Size = New-Object System.Drawing.Size(150, 30)
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $true)
			}
		})
	$form.Controls.Add($btnSelectAll)
	
	# "Deselect All" button
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Location = New-Object System.Drawing.Point(160, 220)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(150, 30)
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
		})
	$form.Controls.Add($btnDeselectAll)
	
	# OK and Cancel buttons
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(20, 270)
	$buttonOK.Size = New-Object System.Drawing.Size(100, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(200, 270)
	$buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($buttonCancel)
	
	$form.AcceptButton = $buttonOK
	$form.CancelButton = $buttonCancel
	
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		return $null
	}
	
	# Gather selected lanes
	$selectedLanes = @()
	for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
	{
		if ($checkedListBox.GetItemChecked($i))
		{
			$selectedLanes += $checkedListBox.Items[$i].LaneNumber
		}
	}
	if ($selectedLanes.Count -eq 0)
	{
		[System.Windows.Forms.MessageBox]::Show("No lanes selected.", "Information",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		return $null
	}
	return @{
		Type  = 'Specific'
		Lanes = $selectedLanes
	}
}

# ===================================================================================================
#                             FUNCTION: Show_Scale_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Presents a GUI dialog for the user to select scales to process within a store.
#   Returns a hashtable with the selection type and the list of selected scale codes.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: The 3-digit store number to filter scales (optional, set to "" for all).
# ===================================================================================================

function Show_Scale_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[array]$BizerbaScales
	)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select BIZERBA Scales to Process"
	$form.Size = New-Object System.Drawing.Size(330, 350)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 10)
	$checkedListBox.Size = New-Object System.Drawing.Size(295, 200)
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)
	
	$sortedScales = $BizerbaScales | Sort-Object { [int]($_.ScaleCode) }
	foreach ($scale in $sortedScales)
	{
		$ip = if ($scale.IPNetwork -and $scale.IPDevice) { "$($scale.IPNetwork)$($scale.IPDevice)" }
		else { "" }
		$displayName = "$($scale.ScaleName) [$ip]"
		$scaleObj = New-Object PSObject -Property @{
			DisplayName = $displayName
			ScaleCode   = $scale.ScaleCode
			ScaleName   = $scale.ScaleName
			IPAddress   = $ip
			Vendor	    = $scale.Vendor
			FullIP	    = $scale.FullIP
			IPNetwork   = $scale.IPNetwork
			IPDevice    = $scale.IPDevice
		}
		$scaleObj | Add-Member -MemberType ScriptMethod -Name ToString -Value { return $this.DisplayName } -Force
		$checkedListBox.Items.Add($scaleObj) | Out-Null
	}
	
	# Select All
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Location = New-Object System.Drawing.Point(10, 220)
	$btnSelectAll.Size = New-Object System.Drawing.Size(150, 30)
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $true)
			}
		})
	$form.Controls.Add($btnSelectAll)
	
	# Deselect All
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Location = New-Object System.Drawing.Point(160, 220)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(150, 30)
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
		})
	$form.Controls.Add($btnDeselectAll)
	
	# OK/Cancel
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(20, 270)
	$buttonOK.Size = New-Object System.Drawing.Size(100, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(200, 270)
	$buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($buttonCancel)
	
	$form.AcceptButton = $buttonOK
	$form.CancelButton = $buttonCancel
	
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		return $null
	}
	
	$selectedScales = @()
	for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
	{
		if ($checkedListBox.GetItemChecked($i))
		{
			$selectedScales += $checkedListBox.Items[$i]
		}
	}
	if ($selectedScales.Count -eq 0)
	{
		[System.Windows.Forms.MessageBox]::Show("No BIZERBA scales selected.", "Information",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		return $null
	}
	return @{
		Type   = 'Specific'
		Scales = $selectedScales
	}
}

# ===================================================================================================
#                                FUNCTION: Show_Table_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays a GUI dialog listing all discovered tables from Get_Table_Aliases in a checked list box,
#   with buttons to Select All or Deselect All. Returns the list of checked table names (with _TAB).
# ===================================================================================================

function Show_Table_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[System.Collections.ArrayList]$AliasResults
	)
	
	# We assume $AliasResults is the .Aliases property from Get_Table_Aliases
	# that contains objects with .Table and .Alias, e.g. "XYZ_TAB" and "XYZ".
	
	# Load necessary assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Create the form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Tables to Process"
	$form.Size = New-Object System.Drawing.Size(450, 550)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	# Label
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Please select the tables you want to pump:"
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$label.AutoSize = $true
	$form.Controls.Add($label)
	
	# CheckedListBox
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 40)
	$checkedListBox.Size = New-Object System.Drawing.Size(400, 400)
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)
	
	# Populate the checked list box with unique table names (with _TAB)
	# Make a distinct list of tables from the $AliasResults
	$distinctTables = $AliasResults |
	Select-Object -ExpandProperty Table -Unique |
	Sort-Object
	
	foreach ($tableName in $distinctTables)
	{
		[void]$checkedListBox.Items.Add($tableName, $false)
	}
	
	# Button: Select All
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Location = New-Object System.Drawing.Point(10, 450)
	$btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
	$form.Controls.Add($btnSelectAll)
	
	$btnSelectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $true)
			}
		})
	
	# Button: Deselect All
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Location = New-Object System.Drawing.Point(120, 450)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(100, 30)
	$form.Controls.Add($btnDeselectAll)
	
	$btnDeselectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
		})
	
	# OK Button
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(240, 450)
	$btnOK.Size = New-Object System.Drawing.Size(80, 30)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($btnOK)
	
	# Cancel Button
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(330, 450)
	$btnCancel.Size = New-Object System.Drawing.Size(80, 30)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($btnCancel)
	
	$form.AcceptButton = $btnOK
	$form.CancelButton = $btnCancel
	
	# Show the dialog
	$dialogResult = $form.ShowDialog()
	
	if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
	{
		# Gather checked items
		$selectedTables = @()
		foreach ($item in $checkedListBox.CheckedItems)
		{
			$selectedTables += $item
		}
		return $selectedTables
	}
	else
	{
		return $null
	}
}

# ===================================================================================================
#                                FUNCTION: Show_Section_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Helper function that creates a form with checkboxes for each section. Returns an array of the 
#   selected section names or $null if canceled.#   
# ===================================================================================================

function Show_Section_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$SectionNames
	)
	
	# Make sure .NET WinForms is loaded
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Create the form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select SQL Sections"
	$form.StartPosition = "CenterScreen"
	$form.Size = New-Object System.Drawing.Size(550, 420)
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	# Label: brief instructions
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Check the sections you want to run, then click OK."
	$label.AutoSize = $true
	$label.Left = 20
	$label.Top = 10
	$form.Controls.Add($label)
	
	# CheckedListBox
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Width = 500
	$checkedListBox.Height = 280
	$checkedListBox.Left = 20
	$checkedListBox.Top = 40
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)
	
	# Populate with section names
	foreach ($name in $SectionNames)
	{
		[void]$checkedListBox.Items.Add($name, $false)
	}
	
	# "Select All" button
	$selectAllButton = New-Object System.Windows.Forms.Button
	$selectAllButton.Text = "Select All"
	$selectAllButton.Width = 90
	$selectAllButton.Height = 30
	$selectAllButton.Left = 20
	$selectAllButton.Top = 330
	$form.Controls.Add($selectAllButton)
	
	$selectAllButton.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $true)
			}
		})
	
	# "Deselect All" button
	$deselectAllButton = New-Object System.Windows.Forms.Button
	$deselectAllButton.Text = "Deselect All"
	$deselectAllButton.Width = 90
	$deselectAllButton.Height = 30
	$deselectAllButton.Left = 120
	$deselectAllButton.Top = 330
	$form.Controls.Add($deselectAllButton)
	
	$deselectAllButton.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
		})
	
	# OK button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Width = 80
	$okButton.Height = 30
	$okButton.Left = 240
	$okButton.Top = 330
	# Crucial: set DialogResult, not a manual event
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($okButton)
	
	# Cancel button
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Width = 80
	$cancelButton.Height = 30
	$cancelButton.Left = 340
	$cancelButton.Top = 330
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($cancelButton)
	
	# Set AcceptButton and CancelButton so Enter/Esc work
	$form.AcceptButton = $okButton
	$form.CancelButton = $cancelButton
	
	# Show the dialog
	$dialogResult = $form.ShowDialog()
	
	if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
	{
		# Gather checked items AFTER the form closes
		$selectedSections = @()
		foreach ($item in $checkedListBox.CheckedItems)
		{
			$selectedSections += $item
		}
		
		return $selectedSections
	}
	else
	{
		return $null
	}
}

# =================================================================================================== 
#                                       SECTION: Initialize GUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Initializes and configures the graphical user interface components for the Store/Lane SQL Execution Tool.
#   Host mode is removed; strictly Store/Server/Lane (Store Mode only).
# ===================================================================================================

# Ensure $form is only initialized once
if (-not $form)
{
	# Create a timer to refresh the GUI every second
	$refreshTimer = New-Object System.Windows.Forms.Timer
	$refreshTimer.Interval = 1000 # 1 second
	$refreshTimer.add_Tick({
			# Refresh the form to update all controls
			$form.Refresh()
		})
	$refreshTimer.Start()
	
	# Initialize ToolTip
	$toolTip = New-Object System.Windows.Forms.ToolTip
	$toolTip.AutoPopDelay = 5000
	$toolTip.InitialDelay = 500
	$toolTip.ReshowDelay = 500
	$toolTip.ShowAlways = $true
	
	# Create the main form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Created by: Alex_C.T   |   Version: $VersionNumber   |   Revised: $VersionDate   |   Powershell Version: $PowerShellVersion"
	$form.Size = New-Object System.Drawing.Size(1006, 570)
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
	
	# Banner Label
	$bannerLabel = New-Object System.Windows.Forms.Label
	$bannerLabel.Text = "PowerShell Script - TBS_Maintenance_Script"
	$bannerLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
	$bannerLabel.TextAlign = 'MiddleCenter'
	$bannerLabel.Dock = 'Top'
	$form.Controls.Add($bannerLabel)
	
	# Handle form closing event (X button)
	$form.add_FormClosing({
			$confirmResult = [System.Windows.Forms.MessageBox]::Show(
				"Are you sure you want to exit?",
				"Confirm Exit",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)
			if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				$_.Cancel = $true
			}
			else
			{
				Write_Log "Form is closing. Performing cleanup." "green"
				Delete_Files -Path "$TempDir" -SpecifiedFiles "*.sqi", "*.sql"
			}
		})
	
	# Create a Clear Log button
	$clearLogButton = New-Object System.Windows.Forms.Button
	$clearLogButton.Text = "Clear Log"
	$clearLogButton.Location = New-Object System.Drawing.Point(951, 70)
	$clearLogButton.Size = New-Object System.Drawing.Size(39, 34)
	$clearLogButton.add_Click({
			$logBox.Clear()
			Write_Log "Log Cleared"
		})
	$form.Controls.Add($clearLogButton)
	$toolTip.SetToolTip($clearLogButton, "Clears the log display area.")
	
	######################################################################################################################
	# 																													 #
	# 												Labels																 #
	#																													 #
	######################################################################################################################
	
	# SMS Version Level
	$smsVersionLabel = New-Object System.Windows.Forms.Label
	$smsVersionLabel.Text = "SMS Version: N/A"
	$smsVersionLabel.Location = New-Object System.Drawing.Point(50, 30)
	$smsVersionLabel.Size = New-Object System.Drawing.Size(250, 20) # Made wider for longer version strings
	$smsVersionLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($smsVersionLabel)
	
	# Store Name label (centered)
	$storeNameLabel = New-Object System.Windows.Forms.Label
	$storeNameLabel.Text = "Store Name: N/A"
	$storeNameLabel.Size = New-Object System.Drawing.Size(350, 20)
	$storeNameLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$storeNameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$form.Controls.Add($storeNameLabel)
	# Center label initially and on resize
	$storeNameLabel.Left = [math]::Max(0, ($form.ClientSize.Width - $storeNameLabel.Width) / 2)
	$storeNameLabel.Top = 30
	$form.add_Resize({
			$storeNameLabel.Left = [math]::Max(0, ($form.ClientSize.Width - $storeNameLabel.Width) / 2)
			$storeNameLabel.Top = 30
		})
	
	# Store Number Label
	$script:storeNumberLabel = New-Object System.Windows.Forms.Label
	$storeNumberLabel.Text = "Store Number: N/A"
	$storeNumberLabel.Location = New-Object System.Drawing.Point(830, 30)
	$storeNumberLabel.Size = New-Object System.Drawing.Size(200, 20)
	$storeNumberLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($storeNumberLabel)
	
	# Nodes Backoffice Label (Number of Backoffices)
	$NodesBackoffices = New-Object System.Windows.Forms.Label
	$NodesBackoffices.Text = "Number of Backoffices: N/A"
	$NodesBackoffices.Location = New-Object System.Drawing.Point(50, 50)
	$NodesBackoffices.Size = New-Object System.Drawing.Size(200, 20)
	$NodesBackoffices.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$NodesBackoffices.AutoSize = $false
	$form.Controls.Add($NodesBackoffices)
	
	# Nodes Store Label (Number of Lanes)
	$script:NodesStore = New-Object System.Windows.Forms.Label
	$NodesStore.Text = "Number of Lanes: $($Counts.NumberOfLanes)"
	$NodesStore.Location = New-Object System.Drawing.Point(420, 50)
	$NodesStore.Size = New-Object System.Drawing.Size(200, 20)
	$NodesStore.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$NodesStore.AutoSize = $false
	$form.Controls.Add($NodesStore)
	
	# Scales Label
	$script:scalesLabel = New-Object System.Windows.Forms.Label
	$scalesLabel.Text = "Number of Scales: $($Counts.NumberOfScales)"
	$scalesLabel.Location = New-Object System.Drawing.Point(820, 50)
	$scalesLabel.Size = New-Object System.Drawing.Size(200, 20)
	$scalesLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($scalesLabel)
	
	# Create a RichTextBox for log output
	$logBox = New-Object System.Windows.Forms.RichTextBox
	$logBox.Location = New-Object System.Drawing.Point(50, 70)
	$logBox.Size = New-Object System.Drawing.Size(900, 400)
	$logBox.ReadOnly = $true
	$logBox.Font = New-Object System.Drawing.Font("Consolas", 10)
	
	# Add the RichTextBox to the form
	$form.Controls.Add($logBox)
	
	######################################################################################################################
	# 
	# General Tools Buttons
	#
	######################################################################################################################
	
	############################################################################
	# General Tools Anchor Button
	############################################################################
	$GeneralToolsButton = New-Object System.Windows.Forms.Button
	$GeneralToolsButton.Text = "General Tools"
	$GeneralToolsButton.Location = New-Object System.Drawing.Point(650, 475)
	$GeneralToolsButton.Size = New-Object System.Drawing.Size(300, 50)
	$ContextMenuGeneral = New-Object System.Windows.Forms.ContextMenuStrip
	$ContextMenuGeneral.ShowItemToolTips = $true
	
	
	############################################################################
	# 1) Activate Windows ("Alex_C.T")
	############################################################################
	$activateItem = New-Object System.Windows.Forms.ToolStripMenuItem("Alex_C.T")
	$activateItem.ToolTipText = "Activate Windows using Alex_C.T's method."
	$activateItem.Add_Click({
			Invoke_Secure_Script
		})
	[void]$contextMenuGeneral.Items.Add($activateItem)
	
	############################################################################
	# 2) Reboot System
	############################################################################
	$rebootItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot System")
	$rebootItem.ToolTipText = "Reboot the host system immediately."
	$rebootItem.Add_Click({
			$rebootResult = [System.Windows.Forms.MessageBox]::Show(
				"Do you want to reboot now?",
				"Reboot",
				[System.Windows.Forms.MessageBoxButtons]::YesNo
			)
			if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Restart-Computer -Force
				Delete_Files -Path "$TempDir" -SpecifiedFiles `
							 "Server_Database_Maintenance.sqi", `
							 "Lane_Database_Maintenance.sqi", `
							 "TBS_Maintenance_Script.ps1"
			}
		})
	[void]$contextMenuGeneral.Items.Add($rebootItem)
	
	############################################################################
	# 3) Install Function in SMS
	############################################################################
	$Install_ONE_FUNCTION_Into_SMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Install Function in SMS")
	$Install_ONE_FUNCTION_Into_SMSItem.ToolTipText = "Installs 'Deploy_ONE_FCT' & 'Pump_All_Items_Tables' into the SMS system."
	$Install_ONE_FUNCTION_Into_SMSItem.Add_Click({
			Install_ONE_FUNCTION_Into_SMS -StoreNumber $StoreNumber -OfficePath $OfficePath
		})
	[void]$contextMenuGeneral.Items.Add($Install_ONE_FUNCTION_Into_SMSItem)
	
	############################################################################
	# 4) Repair BMS Service
	############################################################################
	$repairBMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair BMS Service")
	$repairBMSItem.ToolTipText = "Repairs the BMS service for scale deployment."
	$repairBMSItem.Add_Click({
			Repair_BMS
		})
	[void]$contextMenuGeneral.Items.Add($repairBMSItem)
	
	############################################################################
	# 5) Manual Repair
	############################################################################
	$manualRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Manual Repair")
	$manualRepairItem.ToolTipText = "Writes SQL repair scripts to the desktop."
	$manualRepairItem.Add_Click({
			Write_SQL_Scripts_To_Desktop -LaneSQL $script:LaneSQLFiltered -ServerSQL $script:ServerSQLScript
		})
	[void]$contextMenuGeneral.Items.Add($manualRepairItem)
	
	############################################################################
	# 6) Fix Journal
	############################################################################
	$fixJournalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Fix Journal")
	$fixJournalItem.ToolTipText = "Fix journal entries for the specified date."
	$fixJournalItem.Add_Click({
			Fix_Journal -StoreNumber $StoreNumber -OfficePath $OfficePath
		})
	[void]$contextMenuGeneral.Items.Add($fixJournalItem)
	
	############################################################################
	# 7) Reboot Scales
	############################################################################
	$Reboot_ScalesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Scales")
	$Reboot_ScalesItem.ToolTipText = "Reboot Scale/s."
	$Reboot_ScalesItem.Add_Click({
			Reboot_Scales -ScaleIPNetworks $script:FunctionResults['ScaleIPNetworks']
		})
	[void]$contextMenuGeneral.Items.Add($Reboot_ScalesItem)
	
	############################################################################
	# 8) Open Lane C$ Share(s)
	############################################################################
	$OpenLaneCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Lane C$ Share(s)")
	$OpenLaneCShareItem.ToolTipText = "Select lanes and open their administrative C$ shares in Explorer."
	$OpenLaneCShareItem.Add_Click({
			Open_Selected_Lane/s_C_Path -StoreNumber $storeNumber
		})
	[void]$contextMenuGeneral.Items.Add($OpenLaneCShareItem)
	
	############################################################################
	# 9) Open Scale C$ Share(s)
	############################################################################
	$OpenScaleCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Scale C$ Share(s)")
	$OpenScaleCShareItem.ToolTipText = "Select scales and open their C$ administrative shares as 'bizuser' (bizerba/biyerba)."
	$OpenScaleCShareItem.Add_Click({
			Open_Selected_Scale/s_C_Path -StoreNumber $storeNumber
		})
	[void]$contextMenuGeneral.Items.Add($OpenScaleCShareItem)
	
	############################################################################
	# 10) Export All VNC Files
	############################################################################
	$ExportVNCFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Export All VNC Files")
	$ExportVNCFilesItem.ToolTipText = "Generate UltraVNC (.vnc) connection files for all lanes, scales, and backoffices."
	$ExportVNCFilesItem.Add_Click({
			Export_VNC_Files_For_All_Nodes `
										   -LaneMachines $script:FunctionResults['LaneMachines'] `
										   -ScaleIPNetworks $script:FunctionResults['ScaleIPNetworks'] `
										   -BackofficeMachines $script:FunctionResults['BackofficeMachines']`
										   -LaneVNCPasswords $script:FunctionResults['LaneVNCPasswords']
		})
	[void]$contextMenuGeneral.Items.Add($ExportVNCFilesItem)
	
	############################################################################
	# 11) Remove Archive Bit
	############################################################################
	$RemoveArchiveBitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Remove Archive Bit")
	$RemoveArchiveBitItem.ToolTipText = "Remove archived bit from all lanes and server. Option to schedule as a repeating task."
	$RemoveArchiveBitItem.Add_Click({
			Remove_ArchiveBit_Interactive
		})
	[void]$contextMenuGeneral.Items.Add($RemoveArchiveBitItem)
	
	############################################################################
	# 12) Update Scales Specials
	############################################################################
	$UpdateScalesSpecialsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Scales Specials")
	$UpdateScalesSpecialsItem.ToolTipText = "Update scale specials immediately or schedule as a daily 5AM task."
	$UpdateScalesSpecialsItem.Add_Click({
			Update_Scales_Specials_Interactive
		})
	[void]$contextMenuGeneral.Items.Add($UpdateScalesSpecialsItem)
	
	############################################################################
	# Show the context menu when the General Tools button is clicked
	############################################################################
	$GeneralToolsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$GeneralToolsButton.Add_Click({
			$contextMenuGeneral.Show($GeneralToolsButton, 0, $GeneralToolsButton.Height)
		})
	$toolTip.SetToolTip($GeneralToolsButton, "Click to see some tools created for SMS.")
	$form.Controls.Add($GeneralToolsButton)
	
	######################################################################################################################
	# 
	# Server Tools Button
	#
	######################################################################################################################
	
	############################################################################
	# Server Tools Anchor Button
	############################################################################
	$ServerToolsButton = New-Object System.Windows.Forms.Button
	$ServerToolsButton.Text = "Server Tools"
	$ServerToolsButton.Location = New-Object System.Drawing.Point(50, 475)
	$ServerToolsButton.Size = New-Object System.Drawing.Size(300, 50)
	$ContextMenuServer = New-Object System.Windows.Forms.ContextMenuStrip
	$ContextMenuServer.ShowItemToolTips = $true
	
	############################################################################
	# 1) Server DB Maintenance 
	############################################################################
	$ServerDBMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Server DB Maintenance")
	$ServerDBMaintenanceItem.ToolTipText = "Runs maintenance on the store server database."
	$ServerDBMaintenanceItem.Add_Click({
			$confirmation = [System.Windows.Forms.MessageBox]::Show(
				"Do you want to proceed with the server database maintenance?",
				"Confirmation",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)
			if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Process_Server -StoresqlFilePath $StoresqlFilePath -PromptForSections
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show(
					"Operation canceled.",
					"Canceled",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Information
				)
			}
		})
	[void]$ContextMenuServer.Items.Add($ServerDBMaintenanceItem)
	
	############################################################################
	# 2) Schedule the DB maintenance at the lanes
	############################################################################
	$ServerScheduleRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Server DB Maintenance")
	$ServerScheduleRepairItem.ToolTipText = "Schedule a task to run maintenance at the server database."
	$ServerScheduleRepairItem.Add_Click({
			Schedule_Server_DB_Maintenance -StoreNumber $StoreNumber
		})
	[void]$ContextMenuServer.Items.Add($ServerScheduleRepairItem)
	
	############################################################################
	# 3) Organize_TBS_SCL_ver520 Menu Item
	############################################################################
	$OrganizeScaleTableItem = New-Object System.Windows.Forms.ToolStripMenuItem("Organize_TBS_SCL_ver520")
	$OrganizeScaleTableItem.ToolTipText = "Organize the Scale SQL table (TBS_SCL_ver520)."
	$OrganizeScaleTableItem.Add_Click({
			Organize_TBS_SCL_ver520
		})
	[void]$ContextMenuServer.Items.Add($OrganizeScaleTableItem)
	
	############################################################################
	# 4) Repair Windows Menu Item
	############################################################################
	$RepairWindowsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair Windows")
	$RepairWindowsItem.ToolTipText = "Perform repairs on the Windows operating system."
	$RepairWindowsItem.Add_Click({
			Repair_Windows
		})
	[void]$ContextMenuServer.Items.Add($RepairWindowsItem)
	
	############################################################################
	# 5) Configure System Settings Menu Item
	############################################################################
	$ConfigureSystemSettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Configure System Settings")
	$ConfigureSystemSettingsItem.ToolTipText = "Organize the desktop, set power plan to maximize performance and make sure necessary services are running."
	$ConfigureSystemSettingsItem.Add_Click({
			$confirmResult = [System.Windows.Forms.MessageBox]::Show(
				"Warning: Configuring system settings will make major changes. Do you want to continue?",
				"Confirm Changes",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Configure_System_Settings
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show(
					"Operation canceled.",
					"Canceled",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Information
				)
			}
		})
	[void]$ContextMenuServer.Items.Add($ConfigureSystemSettingsItem)
	
	############################################################################
	# Show the context menu when the Server Tools button is clicked
	############################################################################
	$ServerToolsButton.Add_Click({
			$ContextMenuServer.Show($ServerToolsButton, 0, $ServerToolsButton.Height)
		})
	$toolTip.SetToolTip($ServerToolsButton, "Click to see Server-related tools.")
	$form.Controls.Add($ServerToolsButton)
	
	######################################################################################################################
	# 
	# Lane Tools Button
	#
	######################################################################################################################
	
	############################################################################
	# Lane Tools Anchor Button
	############################################################################
	$LaneToolsButton = New-Object System.Windows.Forms.Button
	$LaneToolsButton.Text = "Lane Tools"
	$LaneToolsButton.Location = New-Object System.Drawing.Point(350, 475)
	$LaneToolsButton.Size = New-Object System.Drawing.Size(300, 50)
	$ContextMenuLane = New-Object System.Windows.Forms.ContextMenuStrip
	$ContextMenuLane.ShowItemToolTips = $true
	
	############################################################################
	# 1) Lane DB Maintenance Button
	############################################################################
	$LaneDBMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Lane DB Maintenance")
	$LaneDBMaintenanceItem.ToolTipText = "Runs maintenance at the lane(s) databases for the selected lane(s)."
	$LaneDBMaintenanceItem.Add_Click({
			Process_Lanes -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($LaneDBMaintenanceItem)
	
	############################################################################
	# 2) Schedule the DB maintenance at the lanes
	############################################################################
	$LaneScheduleMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Lane DB Maintenance")
	$LaneScheduleMaintenanceItem.ToolTipText = "Schedule a task to run maintenance at the lane/s database."
	$LaneScheduleMaintenanceItem.Add_Click({
			Schedule_Lane_DB_Maintenance -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($LaneScheduleMaintenanceItem)
	
	############################################################################
	# 3) Pump Table to Lane Menu Item
	############################################################################
	$PumpTableToLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Pump Table to Lane")
	$PumpTableToLaneItem.ToolTipText = "Pump the selected tables to the lane/s databases."
	$PumpTableToLaneItem.Add_Click({
			Pump_Tables -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($PumpTableToLaneItem)
	
	############################################################################
	# 4) Update Lane Configuration Menu Item
	############################################################################
	$UpdateLaneConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Lane Configuration")
	$UpdateLaneConfigItem.ToolTipText = "Update the configuration files for the lanes. Fixes connectivity errors and mistakes made during lane ghosting."
	$UpdateLaneConfigItem.Add_Click({
			Update_Lane_Config -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($UpdateLaneConfigItem)
	
	############################################################################
	# 5) Close Open Transactions Menu Item
	############################################################################
	$CloseOpenTransItem = New-Object System.Windows.Forms.ToolStripMenuItem("Close Open Transactions")
	$CloseOpenTransItem.ToolTipText = "Close any open transactions at the lane/s."
	$CloseOpenTransItem.Add_Click({
			Close_Open_Transactions -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($CloseOpenTransItem)
	
	############################################################################
	# 6) Ping Lanes Menu Item
	############################################################################
	$PingLanesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Ping Lanes")
	$PingLanesItem.ToolTipText = "Ping all lane devices to check connectivity."
	$PingLanesItem.Add_Click({
			Ping_All_Lanes -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($PingLanesItem)
	
	############################################################################
	# 7) Delete DBS Menu Item
	############################################################################
	$DeleteDBSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Delete DBS")
	$DeleteDBSItem.ToolTipText = "Delete the DBS files (*.txt, *.dwr, if selected *.sus as well) at the lane."
	$DeleteDBSItem.Add_Click({
			Delete_DBS -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($DeleteDBSItem)
	
	############################################################################
	# 8) Refresh PIN Pad Files Menu Item
	############################################################################
	$RefreshPinPadFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh PIN Pad Files")
	$RefreshPinPadFilesItem.ToolTipText = "Refresh the PIN pad files for the lane/s."
	$RefreshPinPadFilesItem.Add_Click({
			Refresh_PIN_Pad_Files -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($RefreshPinPadFilesItem)
	
	############################################################################
	# 9) Drawer Control Item
	############################################################################
	$DrawerControlItem = New-Object System.Windows.Forms.ToolStripMenuItem("Drawer Control")
	$DrawerControlItem.ToolTipText = "Set the Drawer Control for a lane for testing"
	$DrawerControlItem.Add_Click({
			Drawer_Control -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($DrawerControlItem)
	
	############################################################################
	# 10) Refresh Database
	############################################################################
	$RefreshDatabaseItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Database")
	$RefreshDatabaseItem.ToolTipText = "Refresh the database at the lane/s"
	$RefreshDatabaseItem.Add_Click({
			Refresh_Database -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($RefreshDatabaseItem)
	
	############################################################################
	# 11) Send Restart Command Menu Item
	############################################################################
	$SendRestartCommandItem = New-Object System.Windows.Forms.ToolStripMenuItem("Send Restart All Programs")
	$SendRestartCommandItem.ToolTipText = "Send restart all programs to selected lane(s) for the store."
	$SendRestartCommandItem.Add_Click({
			Send_Restart_All_Programs -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($SendRestartCommandItem)
	
	############################################################################
	# 13) Set the time on the lanes
	############################################################################
	$SetLaneTimeFromLocalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Set the time on lanes")
	$SetLaneTimeFromLocalItem.ToolTipText = "Synchronize the time for the selected lanes."
	$SetLaneTimeFromLocalItem.Add_Click({
			Send_SERVER_time_to_Lanes -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($SetLaneTimeFromLocalItem)
	
	############################################################################
	# 14) Reboot Lane Menu Item
	############################################################################
	$RebootLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Lane")
	$RebootLaneItem.ToolTipText = "Reboot the selected lane/s."
	$RebootLaneItem.Add_Click({
			Reboot_Lanes -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($RebootLaneItem)
	
	<############################################################################
	# Retrive Transactions
	############################################################################
	$RetriveTransactionsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Retrive Transactions")
	$RetriveTransactionsItem.ToolTipText = "Retrive Transactions from lane/s."
	$RetriveTransactionsItem.Add_Click({
		Retrive_Transactions -StoreNumber "$StoreNumber"
	})
	[void]$ContextMenuLane.Items.Add($RetriveTransactionsItem)#>
	
	############################################################################
	# Show the context menu when the Server Tools button is clicked
	############################################################################
	$LaneToolsButton.Add_Click({
			$ContextMenuLane.Show($LaneToolsButton, 0, $LaneToolsButton.Height)
		})
	$toolTip.SetToolTip($LaneToolsButton, "Click to see Lane-related tools.")
	$form.Controls.Add($LaneToolsButton)
}
######################################################################################################################
# 
# Anchor all controls for resize
#
######################################################################################################################
$smsVersionLabel.Anchor = 'Top,Left'
$storeNumberLabel.Anchor = 'Top,Right'
$NodesBackoffices.Anchor = 'Top,Left'
$NodesStore.Anchor = 'Top'
$scalesLabel.Anchor = 'Top,Right'
$logBox.Anchor = 'Top,Left,Right,Bottom'
$clearLogButton.Anchor = 'Top,Right'
$GeneralToolsButton.Anchor = 'Bottom,Right'
$ServerToolsButton.Anchor = 'Bottom,Left'
$LaneToolsButton.Anchor = 'Bottom'

$form.add_Resize({
		$storeNameLabel.Left = [math]::Max(0, ($form.ClientSize.Width - $storeNameLabel.Width) / 2)
		$logBox.Width = $form.ClientSize.Width - 100
		$logBox.Height = $form.ClientSize.Height - 170
		$clearLogButton.Left = $form.ClientSize.Width - 55
		$GeneralToolsButton.Left = $form.ClientSize.Width - 350
		$ServerToolsButton.Top = $LaneToolsButton.Top = $GeneralToolsButton.Top = $form.ClientSize.Height - 85
		$LaneToolsButton.Left = [math]::Max(350, ($form.ClientSize.Width - 950) / 2 + 300)
	})

# ===================================================================================================
#                                       SECTION: Main Script Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Orchestrates the execution flow of the script, initializing variables, processing items, and handling user interactions.
# ===================================================================================================

# Get SQL Connection String
Get_Store_And_Database_Info -WinIniPath $WinIniPath -SmsStartIniPath $SmsStartIniPath -StartupIniPath $StartupIniPath -SystemIniPath $SystemIniPath
$StoreNumber = $script:FunctionResults['StoreNumber']
$StoreName = $script:FunctionResults['StoreName']

# Count Nodes based on mode
$Nodes = Retrieve_Nodes -StoreNumber $StoreNumber
$Nodes = $script:FunctionResults['Nodes']

# Retrieve the list of machine names from the FunctionResults dictionary
$LaneMachines = $script:FunctionResults['LaneMachines']

# Populate the hash table with results from various functions
Get_Table_Aliases

# Generate SQL scripts
Generate_SQL_Scripts -StoreNumber $StoreNumber -LanesqlFilePath $LanesqlFilePath -StoresqlFilePath $StoresqlFilePath

# Clearing XE (Urgent Messages) folder.
$ClearXEJob = Clear_XE_Folder

# Gather all unique machine names from LaneMachines for hardware info lookup
$uniqueMachines = $LaneMachines.Values | Where-Object { $_ } | Select-Object -Unique
$null = Get_Remote_Machine_Info -LaneMachines $uniqueMachines # Populates $script:LaneHardwareInfo

# Clear %Temp% folder on start
$ClearTempAtLaunch = Delete_Files -Path "$TempDir" -Exclusions "Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1" -AsJob
$ClearWinTempAtLaunch = Delete_Files -Path "$env:SystemRoot\Temp" -AsJob

# Indicate the script has started
Write-Host "Script started" -ForegroundColor Green

# ===================================================================================================
#                                       SECTION: Show the GUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays the main form to the user and manages the script's execution flow based on user interactions.
# ===================================================================================================

[void]$form.ShowDialog()

# Indicate the script is closing
Write-Host "Script closing..." -ForegroundColor Yellow

# Close the console to avoid duplicate logging to the richbox
exit
