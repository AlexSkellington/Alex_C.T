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
$VersionNumber = "2.3.8"
$VersionDate = "2025-08-04"

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
$script:LaneProtocols = @{ }
$script:LaneProtocolJobs = @{ }

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
# Passwords for Bizerba Scales
# ---------------------------------------------------------------------------------------------------
$bizuser = "bizuser"
$passwordBizerba = ConvertTo-SecureString "bizerba" -AsPlainText -Force
$passwordBiyerba = ConvertTo-SecureString "biyerba" -AsPlainText -Force

$script:credBizerba = New-Object System.Management.Automation.PSCredential ($bizuser, $passwordBizerba)
$script:credBiyerba = New-Object System.Management.Automation.PSCredential ($bizuser, $passwordBiyerba)

# ---------------------------------------------------------------------------------------------------
# Locate Base Path: Storeman Folder Detection (case-insensitive)
# ---------------------------------------------------------------------------------------------------
$BasePath = $null
$targetSubPathPattern = 'Office\Dbs\INFO_*901_WIN.INI'
$storemanDirs = @()
$fixedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 -and $_.Root -match '^[A-Z]:\\$' }

foreach ($drive in $fixedDrives)
{
	# Check every top-level directory (no name filter) for the required subpath
	$dirs = Get-ChildItem -Path "$($drive.Root)" -Directory -ErrorAction SilentlyContinue |
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
# Path where all script files will be saved
# ---------------------------------------------------------------------------------------------------
$script:ScriptsFolder = "C:\Tecnica_Systems\Alex_C.T\Scripts"

# ---------------------------------------------------------------------------------------------------
# Path where all tools will be saved
# ---------------------------------------------------------------------------------------------------
$script:ToolsDir = "C:\Tecnica_Systems\Alex_C.T\Tools"

# ===================================================================================================
#   Detect -ConnectionString support ONCE (run at top of script, before any SQL commands)
# ===================================================================================================
$script:SqlcmdSupportsConnectionString = $null
try
{
	$script:SqlcmdSupportsConnectionString = (Get-Command Invoke-Sqlcmd -ErrorAction Stop).Parameters.Keys -contains "ConnectionString"
}
catch { $script:SqlcmdSupportsConnectionString = $false }

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
#                                      FUNCTION: Get-PsExec
# ---------------------------------------------------------------------------------------------------
# Description:
#   Ensures Sysinternals PsExec.exe is present at the specified tools directory (default: C:\Tecnica_Systems\Alex_C.T\Tools).
#   If missing, downloads PSTools.zip from Microsoft, extracts PsExec.exe in the background,
#   and provides a visible progress indicator compatible with ISE and console hosts.
#   Cleans up temporary files and prints clear log output for all actions and errors.
#
# Improvements:
#   - Fully self-contained, no helpers/nested functions.
#   - Compatible with PowerShell ISE and Windows PowerShell 5+.
#   - Progress display uses Write-Host for log-friendly output (no cursor jumps).
#   - Extraction is robust: uses manual file copy for maximum compatibility.
#   - Handles concurrent job detection; never double-downloads.
#   - Detailed feedback on errors or success, always visible to the user.
#
# Author: Alex_C.T
# ===================================================================================================

function Get_PsExec
{
	param (
		[string]$ToolsDir = $script:ToolsDir
	)
	
	$psexecPath = Join-Path $ToolsDir "PsExec.exe"
	$pstoolsZip = Join-Path $ToolsDir "PSTools.zip"
	$pstoolsUrl = "https://download.sysinternals.com/files/PSTools.zip"
	$jobName = "Get_PsExec_Download_Job"
	
	# Check if PsExec.exe already exists
	if (Test-Path $psexecPath)
	{
		Write-Host "PsExec.exe is ready to be used at $psexecPath."
		return $psexecPath
	}
	
	# Check for existing running job (compatible method)
	$existingJob = Get-Job | Where-Object { $_.Name -eq $jobName -and $_.State -eq 'Running' }
	if ($existingJob)
	{
		Write-Host "Download job already running. Waiting for completion..."
		$job = $existingJob
	}
	else
	{
		Write-Host "PsExec.exe not found. Starting background download and extraction..."
		$job = Start-Job -Name $jobName -ScriptBlock {
			param ($pstoolsUrl,
				$pstoolsZip,
				$ToolsDir,
				$psexecPath)
			try
			{
				if (!(Test-Path $ToolsDir))
				{
					Write-Host "[Job] Creating directory: $ToolsDir"
					New-Item -Path $ToolsDir -ItemType Directory | Out-Null
				}
				Write-Host "[Job] Downloading PSTools.zip..."
				Invoke-WebRequest -Uri $pstoolsUrl -OutFile $pstoolsZip -UseBasicParsing -ErrorAction Stop
				
				Write-Host "[Job] Extracting PsExec.exe..."
				Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
				$zip = [System.IO.Compression.ZipFile]::OpenRead($pstoolsZip)
				$entry = $zip.Entries | Where-Object { $_.Name -ieq "PsExec.exe" }
				if ($entry)
				{
					# Manual extraction (compat with older PowerShell/ISE)
					$fs = $entry.Open()
					$bytes = New-Object byte[] $entry.Length
					[void]$fs.Read($bytes, 0, $entry.Length)
					$fs.Close()
					[System.IO.File]::WriteAllBytes($psexecPath, $bytes)
					Write-Host "[Job] PsExec.exe extracted to $psexecPath"
				}
				else
				{
					Write-Host "[Job] WARNING: PsExec.exe not found in ZIP!"
				}
				$zip.Dispose()
				
				if (Test-Path $pstoolsZip) { Remove-Item $pstoolsZip -ErrorAction SilentlyContinue }
				if (Test-Path $psexecPath)
				{
					Write-Host "[Job] PsExec is ready at $psexecPath"
				}
				else
				{
					Write-Host "[Job] WARNING: PsExec.exe not found after extraction."
				}
			}
			catch
			{
				Write-Host "[Job] ERROR: $($_.Exception.Message)"
				if (Test-Path $pstoolsZip) { Remove-Item $pstoolsZip -ErrorAction SilentlyContinue }
			}
		} -ArgumentList $pstoolsUrl, $pstoolsZip, $ToolsDir, $psexecPath
	}
	
	# Simple progress indicator (compatible with ISE)
	Write-Host -NoNewline "Downloading and extracting PsExec.exe"
	while ($job.State -eq "Running")
	{
		Write-Host -NoNewline "."
		Start-Sleep -Seconds 1
		$job = Get-Job | Where-Object { $_.Id -eq $job.Id }
	}
	Write-Host ""
	
	# Show job output
	Receive-Job -Id $job.Id | ForEach-Object { Write-Host $_ }
	Remove-Job -Id $job.Id -Force
	
	if (Test-Path $psexecPath)
	{
		Write-Host "All done! PsExec.exe is ready at $psexecPath"
		return $psexecPath
	}
	else
	{
		Write-Host "ERROR: PsExec.exe was not found after extraction. Check above for errors."
		return $null
	}
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
	
	# Prepare timestamp if needed
	#$timestamp = if ($IncludeTimestamp) { Get-Date -Format 'yyyy-MM-dd HH:mm:ss' } else { "" }
	
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
#   **Also checks for available SQL module (SqlServer or SQLPS) and stores result.**
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
	# Detect available SQL PowerShell module (SqlServer preferred, fallback to SQLPS)
	# ------------------------------------------------------------------------------------------------
	$availableSqlModule = $null
	if (Get-Module -ListAvailable -Name SqlServer) { $availableSqlModule = "SqlServer" }
	elseif (Get-Module -ListAvailable -Name SQLPS) { $availableSqlModule = "SQLPS" }
	else { $availableSqlModule = "None" }
	$script:FunctionResults['SqlModuleName'] = $availableSqlModule
	
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
#                            FUNCTION: Get_All_Lanes_Database_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the DB server names and database names for all lanes in TER_TAB or for a single lane.
#   If -LaneNumber is supplied, only looks up info for that lane.
#   Updates $script:FunctionResults['LaneDatabaseInfo'] and returns lane info if single.
# ===================================================================================================

function Get_All_Lanes_Database_Info
{
	param (
		[string]$LaneNumber
	)
	
	if (-not $script:FunctionResults.ContainsKey('LaneDatabaseInfo'))
	{
		$script:FunctionResults['LaneDatabaseInfo'] = @{ }
	}
	$LaneDatabaseInfo = $script:FunctionResults['LaneDatabaseInfo']
	
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName) { return $null }
	
	$lanesToProcess = if ($LaneNumber) { @($LaneNumber) }
	else { $LaneNumToMachineName.Keys }
	
	foreach ($laneNumber in $lanesToProcess)
	{
		if ($LaneDatabaseInfo.ContainsKey($laneNumber))
		{
			if ($LaneNumber) { return $LaneDatabaseInfo[$laneNumber] }
			continue
		}
		# Skip unwanted lanes for full mode
		if (-not $LaneNumber -and ($laneNumber -match '^(8|9)' -or $laneNumber -eq '901' -or $laneNumber -eq '999')) { continue }
		
		$machineName = $LaneNumToMachineName[$laneNumber]
		if (-not $machineName) { continue }
		
		# ---------- 1. Try LOCAL Startup file (INI\STARTUP.###) ----------
		$startupLocalPath = Join-Path $OfficePath "INI\STARTUP.$($laneNumber.PadLeft(3, '0'))"
		$startupIniPath = $null
		$content = $null
		$source = $null
		
		if (Test-Path $startupLocalPath)
		{
			try
			{
				$content = Get-Content -Path $startupLocalPath -ErrorAction Stop
				$source = "local"
			}
			catch { $content = $null }
		}
		
		# ---------- 2. Fallback: Try REMOTE Startup.ini ----------
		if (-not $content)
		{
			$startupIniPath = "\\$machineName\storeman\Startup.ini"
			if (Test-Path $startupIniPath)
			{
				try
				{
					$content = Get-Content -Path $startupIniPath -ErrorAction Stop
					$source = "remote"
				}
				catch { $content = $null }
			}
		}
		
		if (-not $content) { continue }
		
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
		
		# -------------- PATCH: Remove trailing backslash unless a named instance is present --------------
		# If it's just 'POS005\' (with no instance), make it 'POS005'
		if ($dbServerRaw -match '^[^\\]+\\$')
		{
			$dbServerRaw = $dbServerRaw.TrimEnd('\')
		}
		# PATCH: If (LOCAL) or localhost (any case), replace with actual machine name
		if ($dbServerRaw -match '^(?i)\(LOCAL\)$' -or $dbServerRaw -match '^(?i)localhost$' -or $dbServerRaw -eq "")
		{
			$dbServer = $machineName
		}
		else
		{
			$dbServer = $dbServerRaw
		}
		# END PATCH --------------------------------------------------------------------------------------
		
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
		
		$laneInfo = @{
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
			'Source'		    = $source
		}
		
		$LaneDatabaseInfo[$laneNumber] = $laneInfo
		
		if ($LaneNumber) { return $laneInfo }
	}
	if ($LaneNumber) { return $null }
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
#       - `$LaneMachineNames`: Array to hold lane identifiers.
#       - `$LaneNumToMachineName`: Hashtable to map lane numbers to machine names.
#       - `$ScaleCodeToIPInfo`: Hashtable to map scale identifiers to their IPNetwork values.
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
	
	# ====================================================================================
	# INITIALIZE ALL MAPPING STRUCTURES
	# ====================================================================================
	$HostPath = "$OfficePath"
	$LaneMachineNames = @() # List of all lane machine names (order not guaranteed)
	$LaneNumToMachineName = @{ } # Key: Lane num (008), Value: Machine name (POS008)
	$LaneMachineLabels = @{ } # Key: Machine name, Value: Label from DB/file
	$LaneMachinePath = @{ } # Key: Machine name, Value: UNC/physical path
	$LaneMachineToServerPath = @{ } # Key: Machine name, Value: Server path
	$ScaleCodes = @() # List of all scale codes
	$ScaleLabels = @{ } # Key: Scale code, Value: Label
	$ScaleExePaths = @{ } # Key: Scale code, Value: Path to scale EXE
	$ScaleCodeToIPInfo = @{ } # Key: Scale code, Value: Scale info object (Bizerba, etc)
	$BackofficeNumToMachineName = @{ } # Key: Backoffice terminal num (e.g. 902), Value: Name
	$BackofficeNumToLabel = @{ } # Key: Backoffice terminal num, Value: Label
	$BackofficeNumToPath = @{ } # Key: Backoffice terminal num, Value: Path
	$ServerMachineName = $null # Store server machine name
	$ServerLabel = $null # Server label (from DB/file)
	$ServerPath = $null # Server path (from DB/file)
	$TerLoadSqlPath = Join-Path $LoadPath 'Ter_Load.sql'
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$NodesFromDatabase = $false
	$SqlModule = $script:FunctionResults['SqlModuleName']
	$server = $script:FunctionResults['DBSERVER']
	$database = $script:FunctionResults['DBNAME']
	
	# ====================================================================================
	# DETECT SQL MODULE
	# ====================================================================================
	if (-not $SqlModule -or $SqlModule -eq 'None')
	{
		Write_Log "No SQL PowerShell module found! Cannot query database for node info." "red"
		$ConnectionString = $null
	}
	
	# ====================================================================================
	# 1. PRIMARY: LOAD FROM DATABASE (TER_TAB, TBS_SCL_ver520)
	# ====================================================================================
	if ($ConnectionString)
	{
		$invokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModule -ErrorAction SilentlyContinue
		$supportsConnStr = $false
		if ($invokeSqlCmd) { $supportsConnStr = $invokeSqlCmd.Parameters.ContainsKey('ConnectionString') }
		else { Write_Log "Invoke-Sqlcmd command not found in module $SqlModule." "yellow" }
		
		$NodesFromDatabase = $true
		try
		{
			# -------------------------------------------------------------------------
			# LOAD LANES, BACKOFFICES, SERVER FROM TER_TAB
			# -------------------------------------------------------------------------
			$queryTerTab = @"
SELECT F1057, F1058, F1125, F1169
FROM TER_TAB
WHERE F1056 = '$StoreNumber'
"@
			if ($supportsConnStr)
			{
				$terTabResult = & $invokeSqlCmd -ConnectionString $ConnectionString -Query $queryTerTab -ErrorAction Stop
			}
			else
			{
				$terTabResult = & $invokeSqlCmd -ServerInstance $server -Database $database -Query $queryTerTab -ErrorAction Stop
			}
			
			foreach ($row in $terTabResult)
			{
				$terminal = $row.F1057
				$label = $row.F1058
				$path = $row.F1125
				$hostPath = $row.F1169
				
				# -------------------------- LANES --------------------------
				# (terminal is 0xx or path is UNC, but not scale/backoffice)
				if (
					($terminal -match '^0\d\d$' -or $path -match '^\\\\[^\\]+\\') -and
					$terminal -notmatch '^8' -and $terminal -notmatch '^9'
				)
				{
					$machineName = $null
					if ($path -match '\\\\([^\\]+)\\') { $machineName = $matches[1] }
					else { $machineName = $terminal }
					$LaneMachineNames += $machineName
					if ($terminal -match '^0\d\d$') { $LaneNumToMachineName[$terminal] = $machineName }
					$LaneNumToMachineName[$machineName] = $machineName
					$LaneMachineLabels[$machineName] = $label
					$LaneMachinePath[$machineName] = $path
					$LaneMachineToServerPath[$machineName] = $hostPath
				}
				# -------------------------- SCALES --------------------------
				elseif ($terminal -match '^8\d\d$' -and $terminal -notmatch '^0' -and $terminal -notmatch '^9' -and $path -match '(?i)^[cC]:\\.*XchScale\\XchScale\.exe$')
				{
					$ScaleCodes += $terminal
					$ScaleLabels[$terminal] = $label
					$ScaleExePaths[$terminal] = $path
					$ScaleCodeToIPInfo[$terminal] = [PSCustomObject]@{
						Code	 = $terminal
						Label    = $label
						Path	 = $path
						HostPath = $hostPath
					}
				}
				# -------------------------- SERVER --------------------------
				elseif ($terminal -eq '901' -and $path -match '^@[^@]+$')
				{
					$ServerMachineName = $path -replace '^@', ''
					$ServerLabel = $label
					$ServerPath = $path
				}
				# -------------------------- BACKOFFICES --------------------------
				elseif ($terminal -match '^9(0[2-9]|[1-8]\d|9[0-8])$' -and $path -match '^@[^@]+$')
				{
					$machineName = $path -replace '^@', ''
					$BackofficeNumToMachineName[$terminal] = $machineName
					$BackofficeNumToLabel[$terminal] = $label
					$BackofficeNumToPath[$terminal] = $path
				}
			}
			
			$NumberOfLanes = $LaneMachineNames.Count
			$NumberOfScales = $ScaleCodes.Count
			$NumberOfServers = if ($ServerMachineName) { 1 }
			else { 0 }
			$NumberOfBackoffices = $BackofficeNumToMachineName.Count
			
			# -------------------------------------------------------------------------
			# LOAD SCALES FROM TBS_SCL_ver520 (preferred, for real scale data)
			# -------------------------------------------------------------------------
			$queryTbsSclScales = @"
SELECT ScaleCode, ScaleName, ScaleLocation, IPNetwork, IPDevice, Active, ScaleBrand, ScaleModel
FROM TBS_SCL_ver520
WHERE Active = 'Y'
"@
			try
			{
				if ($supportsConnStr)
				{
					$tbsSclScalesResult = & $invokeSqlCmd -ConnectionString $ConnectionString -Query $queryTbsSclScales -ErrorAction Stop
				}
				else
				{
					$tbsSclScalesResult = & $invokeSqlCmd -ServerInstance $server -Database $database -Query $queryTbsSclScales -ErrorAction Stop
				}
			}
			catch
			{
				if ($_.Exception.Message -match "Invalid object name 'TBS_SCL_ver520'")
				{
					$tbsSclScalesResult = $null
				}
				else
				{
					throw
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
					$ScaleCodeToIPInfo[$row.ScaleCode] = $scaleObj
				}
			}
		}
		catch
		{
			Write_Log "Failed to retrieve counts from the database: $_" "yellow"
			$NodesFromDatabase = $false
		}
	}
	
	# ====================================================================================
	# 2. FALLBACK: LOAD FROM Ter_Load.sql IF DB IS NOT AVAILABLE
	# ====================================================================================
	$TerLoadUsed = $false
	if ((-not $NodesFromDatabase) -and (Test-Path $TerLoadSqlPath))
	{
		Write_Log "Using Ter_Load.sql as backup for TER_TAB information." "yellow"
		$TerLoadUsed = $true
		$content = Get-Content $TerLoadSqlPath -Raw
		if ($content -match 'INSERT INTO Ter_Load VALUES\s*(.*);' -or $content -match 'INSERT INTO Ter_Load VALUES\s*(.*)')
		{
			$insertBlock = $matches[1]
			$values = $insertBlock -split '\),\s*\(' | ForEach-Object { $_.Trim('() ') }
			
			foreach ($value in $values)
			{
				$fields = $value -split "',\s*'" | ForEach-Object { $_.Trim("'") }
				if ($fields.Count -ge 5 -and $fields[0] -eq $StoreNumber)
				{
					$terminal = $fields[1]
					$label = $fields[2]
					$path = $fields[3]
					$hostPath = $fields[4]
					
					# Lanes: same rules as above
					if (
						($terminal -match '^0\d\d$' -or $path -match '^\\\\[^\\]+\\') -and
						$terminal -notmatch '^8' -and $terminal -notmatch '^9'
					)
					{
						$machineName = $null
						if ($path -match '\\\\([^\\]+)\\') { $machineName = $matches[1] }
						else { $machineName = $terminal }
						$LaneMachineNames += $machineName
						if ($terminal -match '^0\d\d$') { $LaneNumToMachineName[$terminal] = $machineName }
						$LaneNumToMachineName[$machineName] = $machineName
						$LaneMachineLabels[$machineName] = $label
						$LaneMachinePath[$machineName] = $path
						$LaneMachineToServerPath[$machineName] = $hostPath
					}
					elseif ($terminal -match '^8\d\d$' -and $terminal -notmatch '^0' -and $terminal -notmatch '^9' -and $path -match '(?i)^[cC]:\\.*XchScale\\XchScale\.exe$')
					{
						$ScaleCodes += $terminal
						$ScaleLabels[$terminal] = $label
						$ScaleExePaths[$terminal] = $path
						$ScaleCodeToIPInfo[$terminal] = [PSCustomObject]@{
							Code	 = $terminal
							Label    = $label
							Path	 = $path
							HostPath = $hostPath
						}
					}
					elseif ($terminal -eq '901' -and $path -match '^@[^@]+$')
					{
						$ServerMachineName = $path -replace '^@', ''
						$ServerLabel = $label
						$ServerPath = $path
					}
					elseif ($terminal -match '^9(0[2-9]|[1-8]\d|9[0-8])$' -and $path -match '^@[^@]+$')
					{
						$machineName = $path -replace '^@', ''
						$BackofficeNumToMachineName[$terminal] = $machineName
						$BackofficeNumToLabel[$terminal] = $label
						$BackofficeNumToPath[$terminal] = $path
					}
				}
			}
		}
		$NumberOfLanes = $LaneMachineNames.Count
		$NumberOfScales = $ScaleCodes.Count
		$NumberOfServers = if ($ServerMachineName) { 1 }
		else { 0 }
		$NumberOfBackoffices = $BackofficeNumToMachineName.Count
	}
	
	# ====================================================================================
	# 3. EXPAND ALL LOOKUPS: MAKE REVERSE AND FLEXIBLE KEYS FOR ALL NODES
	# ====================================================================================
	# ---- LANES ----
	$ExpandedLaneNumToMachineName = @{ }
	$ExpandedMachineNameToLaneNum = @{ }
	foreach ($kv in $LaneNumToMachineName.GetEnumerator())
	{
		$key = $kv.Key
		$machine = $kv.Value
		if ($key -match '(\d{3})')
		{
			$laneNum = $matches[1]
			$ExpandedLaneNumToMachineName[$laneNum] = $machine
			$ExpandedLaneNumToMachineName["POS$laneNum"] = $machine
			$ExpandedLaneNumToMachineName["pos$laneNum"] = $machine
			$ExpandedLaneNumToMachineName[$machine.ToUpper()] = $machine
			$ExpandedLaneNumToMachineName[$machine.ToLower()] = $machine
			$ExpandedMachineNameToLaneNum[$machine.ToUpper()] = $laneNum
			$ExpandedMachineNameToLaneNum[$machine.ToLower()] = $laneNum
			$ExpandedMachineNameToLaneNum["POS$laneNum"] = $laneNum
			$ExpandedMachineNameToLaneNum["pos$laneNum"] = $laneNum
		}
	}
	$LaneNumToMachineName = $ExpandedLaneNumToMachineName
	$MachineNameToLaneNum = $ExpandedMachineNameToLaneNum
	
	# ---- SCALES ----
	$ExpandedScaleCodeToInfo = @{ }
	$ExpandedScaleNameToCode = @{ }
	foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
	{
		$code = $kv.Key
		$scale = $kv.Value
		if ($code -match '(\d{1,3})')
		{
			$ExpandedScaleCodeToInfo[$code] = $scale
			$ExpandedScaleCodeToInfo["SC$code"] = $scale
			$ExpandedScaleCodeToInfo["sc$code"] = $scale
			if ($scale.ScaleName)
			{
				$ExpandedScaleCodeToInfo[$scale.ScaleName.ToUpper()] = $scale
				$ExpandedScaleCodeToInfo[$scale.ScaleName.ToLower()] = $scale
				$ExpandedScaleNameToCode[$scale.ScaleName.ToUpper()] = $code
				$ExpandedScaleNameToCode[$scale.ScaleName.ToLower()] = $code
			}
		}
	}
	$ScaleCodeToIPInfo = $ExpandedScaleCodeToInfo
	
	# ---- BACKOFFICES ----
	$ExpandedBackofficeNumToMachineName = @{ }
	$ExpandedMachineNameToBackofficeNum = @{ }
	foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
	{
		$key = $kv.Key
		$machine = $kv.Value
		if ($key -match '(\d{3})')
		{
			$boNum = $matches[1]
			$ExpandedBackofficeNumToMachineName[$boNum] = $machine
			$ExpandedBackofficeNumToMachineName["BO$boNum"] = $machine
			$ExpandedBackofficeNumToMachineName["bo$boNum"] = $machine
			$ExpandedBackofficeNumToMachineName[$machine.ToUpper()] = $machine
			$ExpandedBackofficeNumToMachineName[$machine.ToLower()] = $machine
			$ExpandedMachineNameToBackofficeNum[$machine.ToUpper()] = $boNum
			$ExpandedMachineNameToBackofficeNum[$machine.ToLower()] = $boNum
			$ExpandedMachineNameToBackofficeNum["BO$boNum"] = $boNum
			$ExpandedMachineNameToBackofficeNum["bo$boNum"] = $boNum
		}
	}
	$BackofficeNumToMachineName = $ExpandedBackofficeNumToMachineName
	$MachineNameToBackofficeNum = $ExpandedMachineNameToBackofficeNum
	
	# ====================================================================================
	# 4. BUILD RETURN OBJECT & STORE TO GLOBAL FUNCTIONRESULTS
	# ====================================================================================
	$Nodes = [PSCustomObject]@{
		NumberOfLanes			   = $NumberOfLanes
		NumberOfServers		       = $NumberOfServers
		NumberOfBackoffices	       = $NumberOfBackoffices
		NumberOfScales			   = $NumberOfScales
		LaneMachineNames		   = $LaneMachineNames
		LaneNumToMachineName	   = $LaneNumToMachineName
		LaneMachineLabels		   = $LaneMachineLabels
		LaneMachinePath		       = $LaneMachinePath
		LaneMachineToServerPath    = $LaneMachineToServerPath
		ScaleCodes				   = $ScaleCodes
		ScaleLabels			       = $ScaleLabels
		ScaleExePaths			   = $ScaleExePaths
		ScaleCodeToIPInfo		   = $ScaleCodeToIPInfo
		BackofficeNumToMachineName = $BackofficeNumToMachineName
		BackofficeNumToLabel	   = $BackofficeNumToLabel
		BackofficeNumToPath	       = $BackofficeNumToPath
		ServerMachineName		   = $ServerMachineName
		ServerLabel			       = $ServerLabel
		ServerPath				   = $ServerPath
	}
	
	# Write everything into the global FunctionResults for ALL script logic to use!
	$script:FunctionResults['NumberOfLanes'] = $NumberOfLanes
	$script:FunctionResults['NumberOfServers'] = $NumberOfServers
	$script:FunctionResults['NumberOfBackoffices'] = $NumberOfBackoffices
	$script:FunctionResults['NumberOfScales'] = $NumberOfScales
	$script:FunctionResults['LaneMachineNames'] = $LaneMachineNames
	$script:FunctionResults['LaneNumToMachineName'] = $LaneNumToMachineName
	$script:FunctionResults['MachineNameToLaneNum'] = $MachineNameToLaneNum
	$script:FunctionResults['LaneMachineLabels'] = $LaneMachineLabels
	$script:FunctionResults['LaneMachinePath'] = $LaneMachinePath
	$script:FunctionResults['LaneMachineToServerPath'] = $LaneMachineToServerPath
	$script:FunctionResults['ScaleCodes'] = $ScaleCodes
	$script:FunctionResults['ScaleLabels'] = $ScaleLabels
	$script:FunctionResults['ScaleExePaths'] = $ScaleExePaths
	$script:FunctionResults['ScaleCodeToIPInfo'] = $ScaleCodeToIPInfo
	$script:FunctionResults['ScaleNameToCode'] = $ExpandedScaleNameToCode
	$script:FunctionResults['BackofficeNumToMachineName'] = $BackofficeNumToMachineName
	$script:FunctionResults['MachineNameToBackofficeNum'] = $MachineNameToBackofficeNum
	$script:FunctionResults['BackofficeNumToLabel'] = $BackofficeNumToLabel
	$script:FunctionResults['BackofficeNumToPath'] = $BackofficeNumToPath
	$script:FunctionResults['ServerMachineName'] = $ServerMachineName
	$script:FunctionResults['ServerLabel'] = $ServerLabel
	$script:FunctionResults['ServerPath'] = $ServerPath
	$script:FunctionResults['Nodes'] = $Nodes
	
	# ====================================================================================
	# 5. BUILD WINDOWS SCALES ONLY (EX: Bizerba)
	# ====================================================================================
	$WindowsScales = @{ }
	foreach ($code in $ScaleCodeToIPInfo.Keys)
	{
		$scale = $ScaleCodeToIPInfo[$code]
		if ($scale.ScaleBrand -and $scale.ScaleBrand -match 'bizerba') { $WindowsScales[$code] = $scale }
	}
	$script:FunctionResults['WindowsScales'] = $WindowsScales
	
	# ====================================================================================
	# 6. OPTIONAL: UPDATE GUI LABELS IF PRESENT
	# ====================================================================================
	if ($NodesHost -ne $null) { $NodesHost.Text = "Number of Servers: $NumberOfServers" }
	if ($NodesBackoffices -ne $null) { $NodesBackoffices.Text = "Number of Backoffices: $NumberOfBackoffices" }
	if ($NodesStore -ne $null) { $NodesStore.Text = "Number of Lanes: $NumberOfLanes" }
	if ($scalesLabel -ne $null) { $scalesLabel.Text = "Number of Scales: $NumberOfScales" }
	if ($form -ne $null) { $form.Refresh() }
	
	# ====================================================================================
	# 7. RETURN THE NODES OBJECT FOR SCRIPT CALLERS
	# ====================================================================================
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
	
	<# Optionally write to file as fallback
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
#                        FUNCTION: Get_Table_Aliases
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns hashtable mapping table aliases to table names.
#   Uses hardcoded mapping if SMSVersion is 3.3.0.0 to 3.6.0.8 (inclusive),
#   else runs dynamic scan of *_Load.sql files.
#   Relies on $script:FunctionResults['SMSVersionFull'] (set by Get_Store_And_Database_Info).
# ===================================================================================================

function Get_Table_Aliases
{
	$MinSupportedSMSVersion = "3.3.0.0"
	$MaxSupportedSMSVersion = "3.6.0.8"
	
	$SMSVersion = $script:FunctionResults['SMSVersionFull']
	if ($SMSVersion -match '([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)')
	{
		$SMSVersion = $Matches[1]
	}
	else
	{
		$SMSVersion = "0.0.0.0"
	}
	
	# Version range check
	$versionInRange = $false
	try
	{
		$vTest = [version]$SMSVersion
		$vMin = [version]$MinSupportedSMSVersion
		$vMax = [version]$MaxSupportedSMSVersion
		if ($vTest -ge $vMin -and $vTest -le $vMax) { $versionInRange = $true }
	}
	catch { $versionInRange = $false }
	
	# Hardcoded alias → table mapping
	$AliasToTable = @{
		'ALT'	   = 'ALT_TAB'
		'BIO'	   = 'BIO_TAB'
		'BTL'	   = 'BTL_TAB'
		'CAT'	   = 'CAT_TAB'
		'CFG'	   = 'CFG_TAB'
		'CLF'	   = 'CLF_TAB'
		'CLF_SDP'  = 'CLF_SDP_TAB'
		'CLG'	   = 'CLG_TAB'
		'CLK'	   = 'CLK_TAB'
		'CLL'	   = 'CLL_TAB'
		'CLR'	   = 'CLR_TAB'
		'CLT'	   = 'CLT_TAB'
		'CLT_ITM'  = 'CLT_ITM_TAB'
		'COST'	   = 'COST_TAB'
		'CPN'	   = 'CPN_TAB'
		'DELV'	   = 'DELV_TAB'
		'DEPT'	   = 'DEPT_TAB'
		'DSD'	   = 'DSD_TAB'
		'ECL'	   = 'ECL_TAB'
		'FAM'	   = 'FAM_TAB'
		'KIT'	   = 'KIT_TAB'
		'LOC'	   = 'LOC_TAB'
		'LVL'	   = 'LVL_TAB'
		'MIX'	   = 'MIX_TAB'
		'MOD'	   = 'MOD_TAB'
		'OBJ'	   = 'OBJ_TAB'
		'POS'	   = 'POS_TAB'
		'PRICE'    = 'PRICE_TAB'
		'PUB'	   = 'PUB_TAB'
		'RES'	   = 'RES_TAB'
		'ROUTE'    = 'ROUTE_TAB'
		'RPC'	   = 'RPC_TAB'
		'SCAL_ITM' = 'SCL_TAB'
		'SCAL_NUT' = 'SCL_NUT_TAB'
		'SCAL_TXT' = 'SCL_TXT_TAB'
		'SDP'	   = 'SDP_TAB'
		'STD'	   = 'STD_TAB'
		'TAR'	   = 'TAR_TAB'
		'UNT'	   = 'UNT_TAB'
		'VENDOR'   = 'VENDOR_TAB'
	}
	
	# Static: build reverse (table → alias) as well
	$TableToAlias = @{ }
	foreach ($k in $AliasToTable.Keys)
	{
		$TableToAlias[$AliasToTable[$k]] = $k
	}
	
	if ($versionInRange)
	{
		# Build a detailed array of objects (as if they were parsed from files)
		$aliasResults = @()
		foreach ($alias in $AliasToTable.Keys)
		{
			$table = $AliasToTable[$alias]
			$aliasInfo = [PSCustomObject]@{
				File	   = ""
				Table	   = $table
				Alias	   = $alias
				LineNumber = 0
				Context    = "@CREATE('$table','$alias');"
			}
			$aliasResults += $aliasInfo
		}
		$script:FunctionResults['Get_Table_Aliases'] = @{
			Aliases   = $aliasResults
			AliasHash = $AliasToTable
			TableHash = $TableToAlias
		}
		return $AliasToTable
	}
	
	# Out of range: fallback to dynamic scan (and store full results)
	Write-Host "SMS Version $SMSVersion is outside supported range ($MinSupportedSMSVersion - $MaxSupportedSMSVersion). Scanning SQL files for table/alias map..." -ForegroundColor Yellow
	
	$BaseTables = @(
		'OBJ', 'POS', 'PRICE', 'COST', 'DSD', 'KIT', 'LOC', 'ALT', 'ECL',
		'SCL', 'SCL_TXT', 'SCL_NUT', 'DEPT', 'SDP', 'CAT', 'RPC', 'FAM',
		'CPN', 'PUB', 'BIO', 'CLK', 'LVL', 'MIX', 'BTL', 'TAR', 'UNT',
		'RES', 'ROUTE', 'VENDOR', 'DELV', 'CLT', 'CLG', 'CLF', 'CLR',
		'CLL', 'CLT_ITM', 'CLF_SDP', 'STD', 'CFG', 'MOD'
	)
	$escapedTables = $BaseTables | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }
	$tablesPattern = $escapedTables -join '|'
	$pattern = "^\s*@CREATE\s*\(\s*['""]?(?<Table>($tablesPattern)(_[A-Z]+))?['""]?\s*,\s*['""]?(?<Alias>[\w-]+)['""]?\s*\);"
	$regex = [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
	
	$allSqlFiles = Get-ChildItem -Path $LoadPath -Recurse -Filter '*_Load.sql' -ErrorAction SilentlyContinue
	
	$aliasResults = New-Object System.Collections.ArrayList
	$AliasToTableLive = @{ }
	$TableToAliasLive = @{ }
	foreach ($file in $allSqlFiles)
	{
		foreach ($baseTable in $BaseTables)
		{
			if ($file.Name -ieq "$baseTable`_Load.sql")
			{
				$content = Get-Content $file.FullName
				$lineNum = 0
				foreach ($line in $content)
				{
					$lineNum++
					$lineClean = $line -replace '--.*', '' -replace '/\*.*?\*/', ''
					if ($lineClean -match '@CREATE')
					{
						$match = $regex.Match($lineClean)
						if ($match.Success)
						{
							$table = $match.Groups['Table'].Value
							$alias = $match.Groups['Alias'].Value
							if ($table -and $alias)
							{
								# Add to array of objects
								$aliasInfo = [PSCustomObject]@{
									File	   = $file.FullName
									Table	   = $table
									Alias	   = $alias
									LineNumber = $lineNum
									Context    = $lineClean.Trim()
								}
								[void]$aliasResults.Add($aliasInfo)
								$AliasToTableLive[$alias] = $table
								$TableToAliasLive[$table] = $alias
							}
						}
					}
				}
				break
			}
		}
	}
	if ($AliasToTableLive.Count -eq 0)
	{
		Write-Warning "No table-alias pairs detected. Using empty hashtable."
	}
	
	$script:FunctionResults['Get_Table_Aliases'] = @{
		Aliases   = $aliasResults
		AliasHash = $AliasToTableLive
		TableHash = $TableToAliasLive
	}
	return $AliasToTableLive
}

# ===================================================================================================
#                            FUNCTION: Get_All_VNC_Passwords
# ---------------------------------------------------------------------------------------------------
# Description:
#   Scans all specified lanes for the UltraVNC configuration file (UltraVNC.ini, any case) in the
#   default install locations (Program Files and Program Files (x86)). Extracts the encrypted VNC
#   password and stores results in a hashtable keyed by machine name. Handles all filename case 
#   variations. Designed for remote auditing of VNC password status across all lanes.
#
# Parameters:
#   - LaneNumToMachineName   [hashtable]: LaneNumber => MachineName mapping.
#
# Details:
#   - Searches for UltraVNC.ini with any capitalization in both standard install folders.
#   - Uses PowerShell remoting (Invoke-Command) to access remote lane files.
#   - Finds the "passwd=" entry in the INI (case-insensitive, first match).
#   - Returns a hashtable: MachineName => Password (or $null if not found).
#   - Uses Write_Log for status, progress, and error messages.
#
# Usage:
#   $LanePasswords = Get-AllLaneVNCPasswords -LaneNumToMachineName $LaneNumToMachineName
#
# Author: Alex_C.T
# ===================================================================================================

function Get_All_VNC_Passwords
{
	param (
		[Parameter(Mandatory = $false)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$ScaleCodeToIPInfo,
		[Parameter(Mandatory = $false)]
		[hashtable]$BackofficeNumToMachineName
	)
	
	# Default VNC password for lanes, backoffices, and scales with fixed passwords (e.g., Ishida)
	$DefaultVNCPassword = "4330df922eb03b6e"
	$script:DefaultVNCPassword = $DefaultVNCPassword
	
	$uvncFolders = @(
		"C$\Program Files\uvnc bvba\UltraVNC",
		"C$\Program Files (x86)\uvnc bvba\UltraVNC"
	)
	$AllVNCPasswords = @{ }
	
	# 1. Build main node list and tag brands for scales
	$NodeList = @()
	$BizerbaScales = @()
	if ($LaneNumToMachineName) { $NodeList += $LaneNumToMachineName.Values | Where-Object { $_ } }
	if ($ScaleCodeToIPInfo)
	{
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			$ip = $null
			$isIshida = $false
			$isBizerba = $false
			if ($scaleObj -is [string]) { $ip = $scaleObj }
			elseif ($null -ne $scaleObj)
			{
				if (($scaleObj.PSObject.Properties.Name -contains "ScaleBrand") -and ($scaleObj.ScaleBrand -match "Ishida"))
				{
					$isIshida = $true
				}
				if (($scaleObj.PSObject.Properties.Name -contains "ScaleBrand") -and ($scaleObj.ScaleBrand -match "BIZERBA"))
				{
					$isBizerba = $true
				}
				if ($scaleObj.PSObject.Properties.Name -contains "FullIP" -and $scaleObj.FullIP)
				{
					$ip = $scaleObj.FullIP
				}
				elseif (($scaleObj.PSObject.Properties.Name -contains "IPNetwork") -and ($scaleObj.PSObject.Properties.Name -contains "IPDevice") -and $scaleObj.IPNetwork -and $scaleObj.IPDevice)
				{
					$ip = "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)"
				}
			}
			if ($ip)
			{
				if ($isIshida)
				{
					Write_Log "Ishida scale [$ip] will use default VNC password.`r`n" "yellow"
					$AllVNCPasswords[$ip] = $DefaultVNCPassword
					continue
				}
				if ($isBizerba) { $BizerbaScales += @{ Host = $ip; Obj = $scaleObj }; continue }
				$NodeList += $ip
			}
		}
	}
	if ($BackofficeNumToMachineName) { $NodeList += $BackofficeNumToMachineName.Values | Where-Object { $_ } }
	$NodeList = $NodeList | Sort-Object -Unique
	
	if (($NodeList.Count -eq 0) -and ($BizerbaScales.Count -eq 0)) { throw "No machines provided for password extraction." }
	
	# 2. Ping regular nodes first
	$OnlineNodes = @()
	foreach ($node in $NodeList)
	{
		if (Test-Connection -ComputerName $node -Count 1 -Quiet -ErrorAction SilentlyContinue)
		{
			$OnlineNodes += $node
		}
		else
		{
			$AllVNCPasswords[$node] = $null
		}
	}
	
	# 3. Start jobs for online regular nodes
	$jobs = @()
	$MaxConcurrentJobs = 20 # Increased for faster processing
	
	foreach ($machineName in $OnlineNodes)
	{
		while ($jobs.Count -ge $MaxConcurrentJobs)
		{
			$done = Wait-Job -Job $jobs -Any -Timeout 10
			foreach ($j in $done) { Remove-Job $j -Force }
			$jobs = $jobs | Where-Object { $_.State -eq "Running" }
		}
		$jobs += Start-Job -ArgumentList $machineName, $uvncFolders -ScriptBlock {
			param ($machineName,
				$uvncFolders)
			$password = $null
			foreach ($folder in $uvncFolders)
			{
				try
				{
					$iniFiles = Invoke-Command -ComputerName $machineName -ScriptBlock {
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
						$content = Invoke-Command -ComputerName $machineName -ScriptBlock {
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
								break
							}
						}
						if ($password) { break }
					}
				}
				catch
				{
					try
					{
						$remotePath = "\\$machineName\$folder"
						if (Test-Path $remotePath -ErrorAction SilentlyContinue)
						{
							$iniFiles = Get-ChildItem -Path $remotePath -Filter "*.ini" -File -ErrorAction SilentlyContinue | Where-Object {
								$_.Name.ToLower() -eq "ultravnc.ini"
							}
							foreach ($iniFile in $iniFiles)
							{
								$content = Get-Content $iniFile.FullName -ErrorAction Stop
								foreach ($line in $content)
								{
									if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
									{
										$password = $matches[1]
										break
									}
								}
								if ($password) { break }
							}
						}
					}
					catch { continue }
				}
				if ($password) { break }
			}
			return @{ Machine = $machineName; Password = $password }
		}
	}
	
	Wait-Job -Job $jobs | Out-Null
	
	foreach ($j in $jobs)
	{
		$result = Receive-Job $j
		if ($result -and $result.Machine)
		{
			$AllVNCPasswords[$result.Machine] = $result.Password
		}
		Remove-Job $j -Force
	}
	
	# 4. Handle Bizerba scales (using cmdkey/bizuser logic, serially)
	foreach ($b in $BizerbaScales)
	{
		$scaleHost = $b.Host
		$uvncIniRelativePaths = @(
			"Program Files\uvnc bvba\UltraVNC\ultravnc.ini",
			"Program Files (x86)\uvnc bvba\UltraVNC\ultravnc.ini"
		)
		$passwords = @("bizerba", "biyerba")
		$username = "bizuser"
		$password = $null
		$fullIniPath = $null
		
		# Ping check first
		if (-not (Test-Connection -ComputerName $scaleHost -Count 1 -Quiet -ErrorAction SilentlyContinue))
		{
			$AllVNCPasswords[$scaleHost] = $null
			continue
		}
		
		foreach ($uvncIniRel in $uvncIniRelativePaths)
		{
			foreach ($pw in $passwords)
			{
				# Remove any previous credential
				cmdkey /delete:$scaleHost | Out-Null
				cmdkey /add:$scaleHost /user:$username /pass:$pw | Out-Null
				$shareIniPath = "\\$scaleHost\c$\$uvncIniRel"
				if (Test-Path $shareIniPath -ErrorAction SilentlyContinue)
				{
					try
					{
						$content = Get-Content $shareIniPath -ErrorAction Stop
						foreach ($line in $content)
						{
							if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
							{
								$password = $matches[1]
								$fullIniPath = $shareIniPath
								break
							}
						}
					}
					catch { }
				}
				# Remove credential after attempt
				cmdkey /delete:$scaleHost | Out-Null
				if ($password) { break }
			}
			if ($password) { break }
		}
		$AllVNCPasswords[$scaleHost] = $password
	}
	$script:FunctionResults['AllVNCPasswords'] = $AllVNCPasswords
	return $AllVNCPasswords
}

# ===================================================================================================
#                              FUNCTION: Insert_Test_Item
# ---------------------------------------------------------------------------------------------------
# Description:
#   Inserts or updates a test item record (PLU '0020077700000') in the SCL_TAB, OBJ_TAB, POS_TAB, and PRICE_TAB tables
#   using the provided SQL connection string. If the record does not exist in a table, it inserts a new one with specified
#   non-null fields. If it exists, it updates the relevant non-null fields. For PRICE_TAB specifically, the price field (F30)
#   is only updated to 777.77 if its current value is exactly 777.77; otherwise, it remains unchanged. Null fields are ignored
#   in both insert and update operations to minimize unnecessary changes. Errors during execution are silently caught to
#   prevent script interruption.
#
# Improvements:
#   - Handles both insert and update scenarios with existence checks for each table.
#   - Selective update for PRICE_TAB price field to avoid overwriting custom values.
#   - Optimized queries to include only non-null fields, reducing query complexity.
#   - Silent error handling for robustness in production environments.
#   - Uses Invoke-Sqlcmd for efficient SQL execution.
#
# Author: Alex_C.T
# ===================================================================================================

function Insert_Test_Item
{
	param (
		[string]$ConnectionString = $script:FunctionResults['ConnectionString']
	)
	
	if (-not $ConnectionString) { return }
	
	Write_Log "`r`n==================== Starting Insert_Test_Item ====================`r`n" "blue"
	
	$now = Get-Date
	$nowFull = $now.ToString("yyyy-MM-dd HH:mm:ss.fff")
	$nowDate = $now.ToString("yyyy-MM-dd 00:00:00.000")
	
	$preferredPLU = '0020077700000'
	$alternativePLU = '0020777700000'
	$fallbackPLU = '0027777700000'
	$doInsert = $false
	$PLU = $null
	$TestF267 = 777
	
	# Check preferred PLU
	$isPreferredTest = $false
	$descPOS = ""
	$descOBJ = ""
	try
	{
		$posResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT F02 FROM POS_TAB WHERE F01 = '$preferredPLU'"
		if ($posResult) { $descPOS = $posResult.F02 }
		$objResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT F29 FROM OBJ_TAB WHERE F01 = '$preferredPLU'"
		if ($objResult) { $descOBJ = $objResult.F29 }
		if ($descPOS -match '(?i)test' -or $descPOS -match '(?i)tecnica' -or $descOBJ -match '(?i)test' -or $descOBJ -match '(?i)tecnica')
		{
			$isPreferredTest = $true
		}
	}
	catch { }
	
	if ($isPreferredTest -or ($descPOS -eq "" -and $descOBJ -eq ""))
	{
		# Preferred PLU is a test or does not exist, safe to use
		$PLU = $preferredPLU
		$TestF267 = 777
		$doInsert = $true
		Write_Log "Using preferred PLU: $PLU with F267: $TestF267" "green"
	}
	else
	{
		# Check alternate PLU
		$isAltTest = $false
		$descPOS2 = ""
		$descOBJ2 = ""
		try
		{
			$posResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT F02 FROM POS_TAB WHERE F01 = '$alternativePLU'"
			if ($posResult) { $descPOS2 = $posResult.F02 }
			$objResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT F29 FROM OBJ_TAB WHERE F01 = '$alternativePLU'"
			if ($objResult) { $descOBJ2 = $objResult.F29 }
			if ($descPOS2 -match '(?i)test' -or $descPOS2 -match '(?i)tecnica' -or $descOBJ2 -match '(?i)test' -or $descOBJ2 -match '(?i)tecnica')
			{
				$isAltTest = $true
			}
		}
		catch { }
		if ($isAltTest -or ($descPOS2 -eq "" -and $descOBJ2 -eq ""))
		{
			$PLU = $alternativePLU
			$TestF267 = 7777
			$doInsert = $true
			Write_Log "Using alternative PLU: $PLU with F267: $TestF267" "green"
		}
		else
		{
			# Check fallback PLU
			$isFallbackTest = $false
			$descPOS3 = ""
			$descOBJ3 = ""
			try
			{
				$posResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT F02 FROM POS_TAB WHERE F01 = '$fallbackPLU'"
				if ($posResult) { $descPOS3 = $posResult.F02 }
				$objResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT F29 FROM OBJ_TAB WHERE F01 = '$fallbackPLU'"
				if ($objResult) { $descOBJ3 = $objResult.F29 }
				if ($descPOS3 -match '(?i)test' -or $descPOS3 -match '(?i)tecnica' -or $descOBJ3 -match '(?i)test' -or $descOBJ3 -match '(?i)tecnica')
				{
					$isFallbackTest = $true
				}
			}
			catch { }
			if ($isFallbackTest -or ($descPOS3 -eq "" -and $descOBJ3 -eq ""))
			{
				$PLU = $fallbackPLU
				$TestF267 = 77777
				$doInsert = $true
				Write_Log "Using fallback PLU: $PLU with F267: $TestF267" "green"
			}
			else
			{
				Write_Log "No suitable PLU found for test item insertion" "red"
			}
		}
	}
	
	if ($doInsert -and $PLU)
	{
		Write_Log "Deleting existing records for PLU: $PLU and F267: $TestF267" "yellow"
		# Always delete old rows for the chosen PLU and F267 code
		$deleteQueries = @(
			"DELETE FROM SCL_TAB WHERE F01 = '$PLU'",
			"DELETE FROM OBJ_TAB WHERE F01 = '$PLU'",
			"DELETE FROM POS_TAB WHERE F01 = '$PLU'",
			"DELETE FROM PRICE_TAB WHERE F01 = '$PLU'",
			"DELETE FROM SCL_TXT_TAB WHERE F267 = $TestF267"
		)
		foreach ($query in $deleteQueries)
		{
			try { Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query }
			catch { Write_Log "Error during deletion: $_" "red" }
		}
		
		Write_Log "Inserting into SCL_TAB..." "yellow"
		# ===== Insert into SCL_TAB =====
		try
		{
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query @"
INSERT INTO SCL_TAB (F01, F1000, F902, F1001, F253, F258, F264, F267, F1952, F1964, F2581, F2582)
VALUES ('$PLU', 'PAL', 'MANUAL', 1, '$nowFull', 10, 7, $TestF267, 'Test Descriptor 2', '001', 'Test Descriptor 3', 'Test Descriptor 4')
"@
			Write_Log "SCL_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into SCL_TAB: $_" "red" }
		
		Write_Log "Inserting into OBJ_TAB..." "yellow"
		# ===== Insert into OBJ_TAB =====
		try
		{
			$F29 = 'Tecnica Test Item'
			if ($F29.Length -gt 60) { $F29 = $F29.Substring(0, 60) }
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query @"
INSERT INTO OBJ_TAB (F01, F902, F1001, F21, F29, F270, F1118, F1959)
VALUES ('$PLU', '00001153', 0, 1, '$F29', 123.45, '001', '001')
"@
			Write_Log "OBJ_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into OBJ_TAB: $_" "red" }
		
		Write_Log "Inserting into POS_TAB..." "yellow"
		# ===== Insert into POS_TAB =====
		try
		{
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query @"
INSERT INTO POS_TAB (F01, F1000, F902, F1001, F02, F09, F79, F80, F82, F104, F115, F176, F178, F217, F1964, F2119)
VALUES ('$PLU', 'PAL', 'MANUAL', 0, 'Tecnica Test Item', '$nowDate', '1', '1', '1', '0', '0', '1', '1', 1.0, '001', '1')
"@
			Write_Log "POS_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into POS_TAB: $_" "red" }
		
		Write_Log "Inserting into PRICE_TAB..." "yellow"
		# ===== Insert into PRICE_TAB =====
		try
		{
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query @"
INSERT INTO PRICE_TAB (F01, F1000, F126, F902, F1001, F21, F30, F31, F113, F1006, F1007, F1008, F1009, F1803)
VALUES ('$PLU', 'PAL', 1, 'MANUAL', 0, 1, 777.77, 1, 'REG', 1, 777.77, '$nowDate', '1858', 1.0)
"@
			Write_Log "PRICE_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into PRICE_TAB: $_" "red" }
		
		Write_Log "Inserting into SCL_TXT_TAB..." "yellow"
		# ===== Insert into SCL_TXT_TAB =====
		try
		{
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query @"
INSERT INTO SCL_TXT_TAB
(F267, F1000, F253, F297, F902, F1001, F1836)
VALUES
($TestF267, 'PAL', '$nowFull', 'Ingredients Test', 'MANUAL', 0, 'Tecnica Test Item')
"@
			Write_Log "SCL_TXT_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into SCL_TXT_TAB: $_" "red" }
		
		Write_Log "`r`n==================== Insert_Test_Item Completed ====================`r`n" "blue"
	}
	else
	{
		Write_Log "`r`n==================== Insert_Test_Item Completed (No Data) ====================`r`n" "blue"
	}
}

# ===================================================================================================
#                              FUNCTION: Get_Remote_Machine_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays a Windows Form with tabs for Lanes, Scales, and Backoffices, allowing selection via checkboxes.
#   For selected nodes, enables and starts Remote Registry if necessary, queries hardware info (manufacturer, model, CPU, RAM)
#   using WMI (preferred) or REG.exe (fallback), and restores Remote Registry state.
#   Writes sorted Info.txt files to Desktop\Lanes, Desktop\Scales, Desktop\BackOffices.
#   Stores results in $script:LaneHardwareInfo, $script:ScaleHardwareInfo, $script:BackofficeHardwareInfo.
#   (Assumes variables populated by Retrieve_Nodes; handles non-Windows devices gracefully.)
#
# Improvements:
#   - Added restoration of Remote Registry service state (query before, restore after).
#   - Improved error handling with specific messages (e.g., for non-Windows scales).
#   - Added progress feedback via Write_Log as jobs complete.
#   - Enhanced fallbacks: Added CIM for RAM/CPU if WMI fails but PSRemoting possible.
#   - Sorted output numerically if machine names are numeric (e.g., IPs for scales).
#   - Validation: Skip if no selections; better user feedback.
#   - Granular timeout handling: More detailed logging.
#   - Code cleanup: Reduced duplication; inlined all logic without helpers.
#   - Compatibility: Removed ternary operators for PS5 support; used if-else instead.
#   - For scales: Process Bizerba (Windows) with WMI/CIM/REG; skip Ishida (non-Windows) with message.
#   - For Backoffices: Added INI fallback method after all others fail.
#
# Author: Alex_C.T (original)
# ===================================================================================================

function Get_Remote_Machine_Info
{
	Write_Log "`r`n==================== Starting Get_Remote_Machine_Info ====================`r`n" "blue"
	
	# Set up concurrent job and timeout parameters
	$maxConcurrentJobs = 10
	$wmiTimeoutSeconds = 5
	$cimTimeoutSeconds = 10
	$regTimeoutSeconds = 30
	
	# Clear previous results
	$script:LaneHardwareInfo = $null
	$script:ScaleHardwareInfo = $null
	$script:BackofficeHardwareInfo = $null
	
	# ============================
	# MAPPING FROM FUNCTIONRESULTS
	# ============================
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$MachineNameToLaneNum = $script:FunctionResults['MachineNameToLaneNum']
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
	$MachineNameToBackofficeNum = $script:FunctionResults['MachineNameToBackofficeNum']
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$DbsPath = $script:DbsPath
	
	# ===========================
	# NODE SELECTION
	# ===========================
	$nodeSelection = Show_Node_Selection_Form -StoreNumber $StoreNumber `
											  -NodeTypes @("Lane", "Scale", "Backoffice") `
											  -Title "Select Nodes to Pull Hardware Info" `
											  -OnlyBizerbaScales
	
	if (-not $nodeSelection)
	{
		Write_Log "Get_Remote_Machine_Info cancelled by user." "yellow"
		return $false
	}
	
	$selectedLanes = $nodeSelection.Lanes
	$selectedScales = $nodeSelection.Scales
	$selectedBOs = $nodeSelection.Backoffices
	
	# Validation: No selections
	if ((-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		Write_Log "No nodes selected. Operation aborted." "yellow"
		return $false
	}
	
	# Prepare export directories
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"
	foreach ($dir in @($lanesDir, $scalesDir, $backofficesDir))
	{
		if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory | Out-Null }
	}
	
	# Collect all unique scale IPs/machine names only for TrustedHosts update
	$scaleRemotes = @()
	foreach ($sel in $selectedScales)
	{
		if ($sel.ScaleBrand -and $sel.ScaleBrand -match 'bizerba' -and $sel.FullIP)
		{
			$scaleRemotes += $sel.FullIP
		}
	}
	$scaleRemotes = $scaleRemotes | Sort-Object -Unique
	
	# Add scale remotes to TrustedHosts only
	if ($scaleRemotes.Count -gt 0)
	{
		$currentTrusted = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
		$currentList = if ([string]::IsNullOrEmpty($currentTrusted)) { @() }
		else { $currentTrusted.Split(',') | ForEach-Object { $_.Trim() } }
		$toAdd = $scaleRemotes | Where-Object { $_ -and ($currentList -notcontains $_) }
		if ($toAdd.Count -gt 0)
		{
			$newList = $currentList + $toAdd
			$newTrusted = ($newList | Sort-Object -Unique) -join ','
			Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newTrusted -Force
			Write_Log "Updated TrustedHosts with scales: $($toAdd -join ', ')" "cyan"
		}
	}
	
	# Deduplicate Backoffices by BONumber first
	if ($selectedBOs -and $selectedBOs.Count -gt 1)
	{
		$selectedBOs = $selectedBOs | Sort-Object BONumber -Unique
	}
	
	# Optional final deduplication by MachineName to prevent any target duplicates
	$uniqueTargets = @{ }
	$dedupedBackoffices = @()
	foreach ($bo in $selectedBOs)
	{
		if (-not $uniqueTargets.ContainsKey($bo.MachineName))
		{
			$uniqueTargets[$bo.MachineName] = $true
			$dedupedBackoffices += $bo
		}
	}
	$selectedBOs = $dedupedBackoffices
	
	foreach ($section in @(
			@{ Name = 'Lanes'; Selected = $selectedLanes; Dir = $lanesDir; ScriptVar = 'LaneHardwareInfo'; InfoLinesVar = 'LaneInfoLines'; ResultsVar = 'LaneResults'; FileName = 'Lanes_Info.txt'; IsWindows = $true },
			@{ Name = 'Scales'; Selected = $selectedScales; Dir = $scalesDir; ScriptVar = 'ScaleHardwareInfo'; InfoLinesVar = 'ScaleInfoLines'; ResultsVar = 'ScaleResults'; FileName = 'Scales_Info.txt'; IsWindows = $null },
			@{ Name = 'BackOffices'; Selected = $selectedBOs; Dir = $backofficesDir; ScriptVar = 'BackofficeHardwareInfo'; InfoLinesVar = 'BOInfoLines'; ResultsVar = 'BOResults'; FileName = 'Backoffices_Info.txt'; IsWindows = $true }
		))
	{
		if (-not $section.Selected -or $section.Selected.Count -eq 0) { continue }
		Write_Log "Processing $($section.Name) nodes..." "yellow"
		Set-Variable -Name $($section.ResultsVar) -Value @{ }
		Set-Variable -Name $($section.InfoLinesVar) -Value @()
		$jobs = @()
		$pending = @{ }
		
		foreach ($sel in ($section.Selected | Sort-Object {
					if ($section.Name -eq 'Lanes' -and $_.PSObject.Properties.Name -contains 'LaneNumber') { [int]$_.LaneNumber }
					elseif ($section.Name -eq 'Scales' -and $_.PSObject.Properties.Name -contains 'ScaleCode') { [int]$_.ScaleCode }
					elseif ($section.Name -eq 'BackOffices' -and $_.PSObject.Properties.Name -contains 'BONumber') { [int]$_.BONumber }
					else { $_ }
				}))
		{
			# Canonical mappings for all node types using Retrieve_Nodes logic!
			if ($section.Name -eq 'Lanes')
			{
				# $sel is a PSCustomObject with LaneNumber (3-digit) and MachineName (POSnnn)
				$laneNum = $sel.LaneNumber
				$machineName = $sel.MachineName
				$resolvedRemote = $machineName
				$nodeNumOrCode = $laneNum
			}
			elseif ($section.Name -eq 'Scales')
			{
				# $sel is a PSCustomObject with ScaleCode and FullIP
				$scaleCode = $sel.ScaleCode
				$scaleFullIP = $sel.FullIP
				$resolvedRemote = $scaleFullIP
				$nodeNumOrCode = $scaleCode
			}
			elseif ($section.Name -eq 'BackOffices')
			{
				# $sel is a PSCustomObject with BONumber (3-digit) and MachineName (BOxxx)
				$boNum = $sel.BONumber
				$boName = $sel.MachineName
				$resolvedRemote = $boName
				$nodeNumOrCode = $boNum
			}
			else
			{
				$resolvedRemote = $sel
				$nodeNumOrCode = $sel
			}
			
			# You can pass both $resolvedRemote (for connect/UNC) and $nodeNumOrCode (for INI fallback/logging)
			$job = Start-Job -ArgumentList $resolvedRemote, $nodeNumOrCode, $wmiTimeoutSeconds, $cimTimeoutSeconds, $regTimeoutSeconds, $DbsPath, $StoreNumber, $section, $BackofficeNumToMachineName, $credBizerba, $credBiyerba `
							 -ScriptBlock {
				param (
					$resolvedRemote,
					# For connect/UNC
					$nodeNumOrCode,
					# For INI fallback/logging
					$wmiTimeoutSeconds,
					$cimTimeoutSeconds,
					$regTimeoutSeconds,
					$DbsPath,
					$StoreNumber,
					$section,
					$boDict,
					$credBizerba,
					$credBiyerba
				)
				$info = [PSCustomObject]@{
					Success			    = $false
					SystemManufacturer  = $null
					SystemProductName   = $null
					CPU				    = $null
					RAM				    = $null
					OSInfo			    = $null
					Method			    = $null
					Error			    = $null
					MachineNameOverride = $null
				}
				$isBackoffice = ($section.Name -eq 'BackOffices')
				
				# Inline: Get original RemoteRegistry state
				$originalState = $null
				try
				{
					$stateOutput = sc.exe "\\$resolvedRemote" query RemoteRegistry 2>$null | Select-String "STATE" | ForEach-Object { $_.Line.Split(":")[1].Trim() }
					$startTypeOutput = sc.exe "\\$resolvedRemote" qc RemoteRegistry 2>$null | Select-String "START_TYPE" | ForEach-Object { $_.Line.Split(":")[1].Trim() }
					$originalState = @{ State = $stateOutput; StartType = $startTypeOutput }
				}
				catch { }
				
				# -- WMI Method -- (timeout enforced)
				if (-not $isBackoffice -and -not (Test-Connection -ComputerName $resolvedRemote -Count 1 -Quiet -ErrorAction SilentlyContinue))
				{
					$info.Error = "Offline or unreachable."
					$info.Method = "Offline"
				}
				else
				{
					$wmiJob = Start-Job -ScriptBlock {
						param ($resolvedRemote,
							$credBizerba,
							$credBiyerba,
							$isScale)
						try
						{
							# Use credentials if scale; otherwise default call
							if ($isScale)
							{
								$sys = $null
								try
								{
									$sys = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $resolvedRemote -Credential $credBizerba -ErrorAction SilentlyContinue
								}
								catch { }
								if (-not $sys)
								{
									$sys = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $resolvedRemote -Credential $credBiyerba -ErrorAction SilentlyContinue
								}
								if ($sys)
								{
									$cpu = Get-WmiObject -Class Win32_Processor -ComputerName $resolvedRemote -Credential $sys.PSComputerName -ErrorAction SilentlyContinue | Select-Object -First 1
									$os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $resolvedRemote -Credential $sys.PSComputerName -ErrorAction SilentlyContinue
								}
							}
							else
							{
								$sys = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $resolvedRemote -ErrorAction SilentlyContinue
								$cpu = Get-WmiObject -Class Win32_Processor -ComputerName $resolvedRemote -ErrorAction SilentlyContinue | Select-Object -First 1
								$os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $resolvedRemote -ErrorAction SilentlyContinue
							}
							if ($sys -and $sys.Manufacturer -and $sys.Model)
							{
								[PSCustomObject]@{
									SystemManufacturer = $sys.Manufacturer
									SystemProductName  = $sys.Model
									CPU			       = $cpu.Name
									RAM			       = [math]::Round($sys.TotalPhysicalMemory / 1GB, 1)
									OSInfo			   = "$($os.Caption) ($($os.Version))"
									Method			   = "WMI"
								}
							}
							else { $null }
						}
						catch { $null }
					} -ArgumentList $resolvedRemote, $credBizerba, $credBiyerba, ($section.Name -eq 'Scales')
					
					if (Wait-Job $wmiJob -Timeout $wmiTimeoutSeconds)
					{
						$wmiResult = Receive-Job $wmiJob 2>$null
						Remove-Job $wmiJob -Force -ErrorAction SilentlyContinue
					}
					else
					{
						Stop-Job $wmiJob -ErrorAction SilentlyContinue
						Remove-Job $wmiJob -Force -ErrorAction SilentlyContinue
						$wmiResult = $null
					}
					if ($wmiResult -and $wmiResult.SystemManufacturer -and $wmiResult.SystemProductName)
					{
						$info.SystemManufacturer = $wmiResult.SystemManufacturer
						$info.SystemProductName = $wmiResult.SystemProductName
						$info.CPU = $wmiResult.CPU
						$info.RAM = $wmiResult.RAM
						$info.OSInfo = $wmiResult.OSInfo
						$info.Method = "WMI"
						$info.Success = $true
					}
					else
					{
						$info.Error = "WMI failed (credential or access issue)"
					}
				}
				
				# -- CIM Method -- (timeout enforced)
				if (-not $info.Success)
				{
					$cimJob = Start-Job -ScriptBlock {
						param ($resolvedRemote,
							$credBizerba,
							$credBiyerba,
							$isScale)
						try
						{
							$session = $null
							if ($isScale)
							{
								try
								{
									$session = New-CimSession -ComputerName $resolvedRemote -Credential $credBizerba -ErrorAction SilentlyContinue
								}
								catch { }
								if (-not $session)
								{
									$session = New-CimSession -ComputerName $resolvedRemote -Credential $credBiyerba -ErrorAction SilentlyContinue
								}
							}
							else
							{
								$session = New-CimSession -ComputerName $resolvedRemote -ErrorAction SilentlyContinue
							}
							if ($session)
							{
								$sys = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem 2>$null
								$cpu = Get-CimInstance -CimSession $session -ClassName Win32_Processor 2>$null | Select-Object -First 1
								$os = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem 2>$null
								Remove-CimSession $session 2>$null
								if ($sys -and $sys.Manufacturer -and $sys.Model)
								{
									[PSCustomObject]@{
										SystemManufacturer = $sys.Manufacturer
										SystemProductName  = $sys.Model
										CPU			       = $cpu.Name
										RAM			       = [math]::Round($sys.TotalPhysicalMemory / 1GB, 1)
										OSInfo			   = "$($os.Caption) ($($os.Version))"
										Method			   = "CIM"
									}
								}
								else { $null }
							}
							else { $null }
						}
						catch { $null }
					} -ArgumentList $resolvedRemote, $credBizerba, $credBiyerba, ($section.Name -eq 'Scales')
					
					if (Wait-Job $cimJob -Timeout $cimTimeoutSeconds)
					{
						$cimResult = Receive-Job $cimJob 2>$null
						Remove-Job $cimJob -Force -ErrorAction SilentlyContinue
					}
					else
					{
						Stop-Job $cimJob -ErrorAction SilentlyContinue
						Remove-Job $cimJob -Force -ErrorAction SilentlyContinue
						$cimResult = $null
					}
					if ($cimResult -and $cimResult.SystemManufacturer -and $cimResult.SystemProductName)
					{
						$info.SystemManufacturer = $cimResult.SystemManufacturer
						$info.SystemProductName = $cimResult.SystemProductName
						$info.CPU = $cimResult.CPU
						$info.RAM = $cimResult.RAM
						$info.OSInfo = $cimResult.OSInfo
						$info.Method = "CIM"
						$info.Success = $true
					}
					else
					{
						$info.Error = "CIM failed (credential or access issue)"
					}
				}
				
				# -- REG Method (wrapped for timeout and try/catch) --
				if (-not $info.Success)
				{
					$regResult = $null
					try
					{
						$regJob = Start-Job -ScriptBlock {
							param ($resolvedRemote,
								$originalState,
								$credBizerba,
								$credBiyerba,
								$isScale)
							
							# Ensure RemoteRegistry service is running
							if ($originalState.StartType -ne "AUTO_START" -and $originalState.StartType -ne "DEMAND_START")
							{
								sc.exe "\\$resolvedRemote" config RemoteRegistry start= demand | Out-Null
							}
							if ($originalState.State -ne "RUNNING")
							{
								sc.exe "\\$resolvedRemote" start RemoteRegistry | Out-Null
								Start-Sleep -Milliseconds 500
							}
							
							try
							{
								if ($isScale)
								{
									try
									{
										$session = New-PSSession -ComputerName $resolvedRemote -Credential $credBizerba -ErrorAction SilentlyContinue
									}
									catch { }
									if (-not $session)
									{
										$session = New-PSSession -ComputerName $resolvedRemote -Credential $credBiyerba -ErrorAction SilentlyContinue
									}
								}
								else
								{
									$session = New-PSSession -ComputerName $resolvedRemote -ErrorAction SilentlyContinue
								}
								if ($session)
								{
									$result = Invoke-Command -Session $session -ScriptBlock {
										try
										{
											$manuf = Get-ItemProperty -Path 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemManufacturer -ErrorAction SilentlyContinue
											$prod = Get-ItemProperty -Path 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemProductName -ErrorAction SilentlyContinue
											$cpu = Get-ItemProperty -Path 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -Name ProcessorNameString -ErrorAction SilentlyContinue
											$osKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
											$osName = Get-ItemProperty -Path $osKey -Name ProductName -ErrorAction SilentlyContinue
											$osVer = Get-ItemProperty -Path $osKey -Name CurrentVersion -ErrorAction SilentlyContinue
											$osBuild = Get-ItemProperty -Path $osKey -Name DisplayVersion -ErrorAction SilentlyContinue
											
											[PSCustomObject]@{
												SystemManufacturer = $manuf.SystemManufacturer
												SystemProductName  = $prod.SystemProductName
												CPU			       = $cpu.ProcessorNameString
												OSInfo			   = "$($osName.ProductName) ($($osVer.CurrentVersion)" + (if ($osBuild) { ", $($osBuild.DisplayVersion))" } else { ")" })
												Method			   = "REG"
												Success		       = $true
												Error			   = $null
											}
										}
										catch
										{
											[PSCustomObject]@{
												SystemManufacturer = $null
												SystemProductName  = $null
												CPU			       = $null
												OSInfo			   = $null
												Method			   = "REG"
												Success		       = $false
												Error			   = $_.Exception.Message
											}
										}
									} 2>$null
									
									Remove-PSSession $session -ErrorAction SilentlyContinue
									return $result
								}
								else
								{
									[PSCustomObject]@{
										SystemManufacturer = $null
										SystemProductName  = $null
										CPU			       = $null
										OSInfo			   = $null
										Method			   = "REG"
										Success		       = $false
										Error			   = "Session creation failed (credential or access issue)"
									}
								}
							}
							catch
							{
								[PSCustomObject]@{
									SystemManufacturer = $null
									SystemProductName  = $null
									CPU			       = $null
									OSInfo			   = $null
									Method			   = "REG"
									Success		       = $false
									Error			   = "REG fallback failed silently"
								}
							}
						} -ArgumentList $resolvedRemote, $originalState, $credBizerba, $credBiyerba, ($section.Name -eq 'Scales')
						
						if (Wait-Job $regJob -Timeout $regTimeoutSeconds)
						{
							$regResult = Receive-Job $regJob 2>$null
							Remove-Job $regJob -Force -ErrorAction SilentlyContinue
						}
						else
						{
							Stop-Job $regJob -ErrorAction SilentlyContinue
							Remove-Job $regJob -Force -ErrorAction SilentlyContinue
							$regResult = [PSCustomObject]@{
								SystemManufacturer = $null
								SystemProductName  = $null
								CPU			       = $null
								OSInfo			   = $null
								Method			   = "REG"
								Success		       = $false
								Error			   = "REG query timed out after $regTimeoutSeconds seconds."
							}
						}
					}
					catch
					{
						$regResult = [PSCustomObject]@{
							SystemManufacturer = $null
							SystemProductName  = $null
							CPU			       = $null
							OSInfo			   = $null
							Method			   = "REG"
							Success		       = $false
							Error			   = "REG method failed silently"
						}
					}
					
					# Set info from regResult no matter what
					$info.SystemManufacturer = $regResult.SystemManufacturer
					$info.SystemProductName = $regResult.SystemProductName
					$info.CPU = $regResult.CPU
					$info.RAM = $null
					$info.OSInfo = $regResult.OSInfo
					$info.Method = "REG"
					$info.Error = $regResult.Error
					$info.Success = $regResult.Success
				}
				
				# --- INI Block (for Lanes/BackOffices), ALWAYS RUN ---
				$returnInfo = [PSCustomObject]@{
					Machine			    = $resolvedRemote
					MachineNameOverride = $info.MachineNameOverride
					Info			    = $info
					IniFound		    = $false
					IniPath			    = $null
				}
				
				try
				{
					if ($section.Name -in @('Lanes', 'BackOffices'))
					{
						if ($section.Name -eq 'Lanes')
						{
							$pattern = "INFO_${StoreNumber}${nodeNumOrCode}_SMSStart.ini"
							$iniPath = Join-Path "\\$resolvedRemote\storeman\office\dbs" $pattern
						}
						else
						{
							$pattern = "INFO_${StoreNumber}${nodeNumOrCode}_SMSStart.ini"
							$iniPath = Join-Path $DbsPath $pattern
						}
						
						$iniFile = Get-ChildItem -Path (Split-Path $iniPath) -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
						if ($iniFile -and (Test-Path $iniFile.FullName))
						{
							$returnInfo.IniFound = $true
							$returnInfo.IniPath = $iniFile.FullName
							
							$iniLines = Get-Content $iniFile.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
							$sections = @{ }
							$curSec = ""
							foreach ($line in $iniLines)
							{
								if ($line -match '^\[(.+)\]$')
								{
									$curSec = $matches[1]
									$sections[$curSec] = @{ }
								}
								elseif ($line -match '^\s*([^=]+?)\s*=\s*(.*)$' -and $curSec)
								{
									$sections[$curSec][$matches[1].Trim()] = $matches[2].Trim()
								}
							}
							
							if ($info.CPU -eq "Unknown" -and $sections.ContainsKey('PROCESSOR') -and $sections['PROCESSOR'].ContainsKey('Cores'))
							{
								$cores = $sections['PROCESSOR']['Cores']
								$arch = $sections['PROCESSOR']['Architecture']
								$info.CPU = "$cores cores" + ($(if ($arch) { " ($arch)" }
										else { "" }))
							}
							if ($info.RAM -eq "Unknown" -and $sections.ContainsKey('Memory') -and $sections['Memory'].ContainsKey('PhysicalMemory'))
							{
								$ramMb = $sections['Memory']['PhysicalMemory']
								$info.RAM = $(if ($ramMb -match '^\d+$') { "{0:N1}" -f ([double]$ramMb / 1024) }
									else { $ramMb })
							}
							if ($info.OSInfo -eq "Unknown" -and $sections.ContainsKey('OperatingSystem') -and $sections['OperatingSystem'].ContainsKey('ProductName'))
							{
								$info.OSInfo = $sections['OperatingSystem']['ProductName']
							}
							if ($sections.ContainsKey('ORIGIN') -and $sections['ORIGIN'].ContainsKey('ComputerName'))
							{
								$info.MachineNameOverride = $sections['ORIGIN']['ComputerName']
							}
						}
					}
				}
				catch { }
				
				return $returnInfo
			}
			$jobs += $job
			$pending[$job.Id] = $resolvedRemote
			
			# Throttle background jobs for max parallelism
			while ($jobs.Count -ge $maxConcurrentJobs)
			{
				$done = Wait-Job -Job $jobs -Any -Timeout 60
				if ($done)
				{
					foreach ($j in $done)
					{
						$result = Receive-Job $j 2>$null
						$remoteName = $pending[$j.Id]
						$info = $result.Info
						# --- INI Debug output ---
						if ($result.PSObject.Properties.Name -contains 'IniFound')
						{
							if ($section.Name -in @('Lanes', 'BackOffices'))
							{
								if ($result.IniFound)
								{
									Write_Log "[INI] Found INI for $($section.Name) $remoteName at $($result.IniPath)" "cyan"
								}
								else
								{
									Write_Log "[INI] No INI found for $($section.Name) $remoteName" "darkyellow"
								}
							}
						}
						# Inline: Restore if not done in job (edge case)
						$originalState = $result.OriginalState
						if ($originalState)
						{
							if ($originalState.StartType) { sc.exe "\\$remoteName" config RemoteRegistry start= $originalState.StartType | Out-Null 2>$null }
							if ($originalState.State -ne "RUNNING") { sc.exe "\\$remoteName" stop RemoteRegistry | Out-Null 2>$null }
						}
						# Determine display name (MachineNameOverride logic, supports serialization)
						if ($section.Name -eq 'BackOffices')
						{
							$boNum = $remoteName
							$boName = $null
							if ($result.PSObject.Properties.Name -contains "MachineNameOverride" -and $result.MachineNameOverride)
							{
								$boName = $result.MachineNameOverride
							}
							elseif ($info.PSObject.Properties.Name -contains "MachineNameOverride" -and $info.MachineNameOverride)
							{
								$boName = $info.MachineNameOverride
							}
							elseif ($boDict -and $boDict.ContainsKey($boNum) -and $boDict[$boNum])
							{
								$boName = $boDict[$boNum]
							}
							else
							{
								$boName = $boNum
							}
							$displayName = "$boName ($boNum)"
						}
						else
						{
							if ($result.PSObject.Properties.Name -contains "MachineNameOverride" -and $result.MachineNameOverride)
							{
								$displayName = $result.MachineNameOverride
							}
							elseif ($info.PSObject.Properties.Name -contains "MachineNameOverride" -and $info.MachineNameOverride)
							{
								$displayName = $info.MachineNameOverride
							}
							else
							{
								$displayName = $remoteName
							}
						}
						$line = "Machine Name: $displayName |"
						if ($info.Success)
						{
							$line += " Manufacturer: $($info.SystemManufacturer) | Model: $($info.SystemProductName) | CPU: $($info.CPU)"
							if ($info.RAM -ne $null) { $line += " | RAM: $($info.RAM) GB" }
							if ($info.OSInfo) { $line += " | OS: $($info.OSInfo)" }
							$line += " | Method: $($info.Method)"
							Write_Log "Processed $displayName ($($section.Name)): Success [$($info.Method)]" "green"
						}
						elseif ($info.Error)
						{
							$line += " [Hardware info unavailable] Error: $($info.Error)"
							$line += " | Method: $($info.Method)"
							Write_Log "Processed $displayName ($($section.Name)): Error [$($info.Method)] - $($info.Error)" "red"
						}
						else
						{
							$line += " [No hardware info found] | Method: $($info.Method)"
							Write_Log "Processed $displayName ($($section.Name)): No info [$($info.Method)]" "yellow"
						}
						$infolines = Get-Variable -Name $($section.InfoLinesVar) -ValueOnly
						$infolines += $line
						Set-Variable -Name $($section.InfoLinesVar) -Value $infolines
						Stop-Job $j -ErrorAction SilentlyContinue
						Remove-Job $j -Force -ErrorAction SilentlyContinue
						$jobs = $jobs | Where-Object { $_.Id -ne $j.Id }
						$pending.Remove($j.Id)
					}
				}
			}
		}
		
		# Wait for all jobs to finish and collect results
		if ($jobs.Count -gt 0)
		{
			Wait-Job -Job $jobs -Timeout 60 | Out-Null
			foreach ($j in $jobs)
			{
				$remoteName = $pending[$j.Id]
				$result = Receive-Job $j 2>$null
				$info = $result.Info
				# --- INI Debug output ---
				if ($result.PSObject.Properties.Name -contains 'IniFound')
				{
					if ($section.Name -in @('Lanes', 'BackOffices'))
					{
						if ($result.IniFound)
						{
							Write_Log "[INI] Found INI for $($section.Name) $remoteName at $($result.IniPath)" "cyan"
						}
						else
						{
							Write_Log "[INI] No INI found for $($section.Name) $remoteName" "darkyellow"
						}
					}
				}
				# --- Restore RemoteRegistry state if needed ---
				$originalState = $result.OriginalState
				if ($originalState)
				{
					if ($originalState.StartType) { sc.exe "\\$remoteName" config RemoteRegistry start= $originalState.StartType | Out-Null 2>$null }
					if ($originalState.State -ne "RUNNING") { sc.exe "\\$remoteName" stop RemoteRegistry | Out-Null 2>$null }
				}
				if ($section.Name -eq 'BackOffices')
				{
					$boNum = $remoteName
					$boName = $null
					if ($result.PSObject.Properties.Name -contains "MachineNameOverride" -and $result.MachineNameOverride)
					{
						$boName = $result.MachineNameOverride
					}
					elseif ($info.PSObject.Properties.Name -contains "MachineNameOverride" -and $info.MachineNameOverride)
					{
						$boName = $info.MachineNameOverride
					}
					elseif ($BackofficeNumToMachineName -and $BackofficeNumToMachineName.ContainsKey($boNum) -and $BackofficeNumToMachineName[$boNum])
					{
						$boName = $BackofficeNumToMachineName[$boNum]
					}
					else
					{
						$boName = $boNum
					}
					$displayName = "$boName ($boNum)"
				}
				else
				{
					if ($result.PSObject.Properties.Name -contains "MachineNameOverride" -and $result.MachineNameOverride)
					{
						$displayName = $result.MachineNameOverride
					}
					elseif ($info.PSObject.Properties.Name -contains "MachineNameOverride" -and $info.MachineNameOverride)
					{
						$displayName = $info.MachineNameOverride
					}
					else
					{
						$displayName = $remoteName
					}
				}
				$line = "Machine Name: $displayName |"
				if ($info.Success)
				{
					$line += " Manufacturer: $($info.SystemManufacturer) | Model: $($info.SystemProductName) | CPU: $($info.CPU)"
					if ($info.RAM -ne $null) { $line += " | RAM: $($info.RAM) GB" }
					if ($info.OSInfo) { $line += " | OS: $($info.OSInfo)" }
					$line += " | Method: $($info.Method)"
					Write_Log "Processed $displayName ($($section.Name)): Success [$($info.Method)]" "green"
				}
				elseif ($info.Error)
				{
					$line += " [Hardware info unavailable] Error: $($info.Error)"
					$line += " | Method: $($info.Method)"
					Write_Log "Processed $displayName ($($section.Name)): Error [$($info.Method)] - $($info.Error)" "red"
				}
				else
				{
					$line += " [No hardware info found] | Method: $($info.Method)"
					Write_Log "Processed $displayName ($($section.Name)): No info [$($info.Method)]" "yellow"
				}
				$infolines = Get-Variable -Name $($section.InfoLinesVar) -ValueOnly
				$infolines += $line
				Set-Variable -Name $($section.InfoLinesVar) -Value $infolines
				Stop-Job $j -ErrorAction SilentlyContinue
				Remove-Job $j -Force -ErrorAction SilentlyContinue
			}
		}
		
		Set-Variable -Name ("script:" + $section.ScriptVar) -Value (Get-Variable -Name $($section.ResultsVar) -ValueOnly) -Scope Script
		$infolines = Get-Variable -Name $($section.InfoLinesVar) -ValueOnly
		# Write sorted output
		if ($infolines.Count)
		{
			$sortedLines = $infolines | Sort-Object {
				if ($_ -match "^Machine Name: ([^|]+)")
				{
					$name = $matches[1]
					if ($name -match '^\d+$' -or $name -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') { [int]$name.Split('.')[-1] }
					else { $name }
				}
				else { $_ }
			}
			$filePath = Join-Path $section.Dir $section.FileName
			$sortedLines -join "`r`n" | Set-Content -Path $filePath -Encoding Default
			Write_Log "Exported $($section.Name) info to $filePath" "green"
		}
		Write_Log "Completed processing $($section.Name).`r`n" "green"
	}
	
	# Safely retrieve info lines
	if (Get-Variable -Name LaneInfoLines -Scope Local -ErrorAction SilentlyContinue) { $laneLines = Get-Variable -Name LaneInfoLines -ValueOnly }
	else { $laneLines = @() }
	if (Get-Variable -Name ScaleInfoLines -Scope Local -ErrorAction SilentlyContinue) { $scaleLines = Get-Variable -Name ScaleInfoLines -ValueOnly }
	else { $scaleLines = @() }
	if (Get-Variable -Name BOInfoLines -Scope Local -ErrorAction SilentlyContinue) { $boLines = Get-Variable -Name BOInfoLines -ValueOnly }
	else { $boLines = @() }
	
	Write_Log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
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
	
	# ------------------------------------------------------------------------------------------------
	# Banner: Start
	# ------------------------------------------------------------------------------------------------
	Write_Log "`r`n==================== Starting Server Database Maintenance ====================`r`n" "blue"
	
	if ($PromptForSections)
	{
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
	}
	
	# ------------------------------------------------------------------------------------------------
	# Variables and settings
	# ------------------------------------------------------------------------------------------------
	$sqlScript = $script:ServerSQLScript
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']
	
	# ------------------------------------------------------------------------------------------------
	# Load SQL script content (from variable or file)
	# ------------------------------------------------------------------------------------------------
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
	
	# ------------------------------------------------------------------------------------------------
	# Parse SQL sections (comment blocks) to execute
	# ------------------------------------------------------------------------------------------------
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$matches = [regex]::Matches($sqlScript, $sectionPattern)
	
	if ($matches.Count -eq 0)
	{
		Write_Log "No SQL sections found to execute." "red"
		return
	}
	
	# ------------------------------------------------------------------------------------------------
	# Prompt for section selection if requested
	# ------------------------------------------------------------------------------------------------
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
	
	# ------------------------------------------------------------------------------------------------
	# Get connection string and import detected SQL module
	# ------------------------------------------------------------------------------------------------
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	
	if (-not $ConnectionString -or -not $server -or -not $dbName)
	{
		Write_Log "DB server, DB name, or connection string not found. Cannot execute SQL script." "red"
		return
	}
	
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		Import-Module $SqlModuleName -ErrorAction Stop
		$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
	}
	else
	{
		Write_Log "No valid SQL module available for SQL operations!" "red"
		return
	}
	
	# ------------------------------------------------------------------------------------------------
	# Check if Invoke-Sqlcmd supports ConnectionString parameter
	# ------------------------------------------------------------------------------------------------
	$supportsConnectionString = $false
	if ($InvokeSqlCmd)
	{
		$supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString'
	}
	
	# ------------------------------------------------------------------------------------------------
	# Execute each SQL section, one by one (NO RETRIES, NO FILE COPY)
	# ------------------------------------------------------------------------------------------------
	$failedSections = @()
	foreach ($match in $matches)
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
				& $InvokeSqlCmd -ConnectionString $ConnectionString -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
			}
			else
			{
				& $InvokeSqlCmd -ServerInstance $server -Database $dbName -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
			}
			Write_Log "Section '$sectionName' executed successfully." "green"
		}
		catch
		{
			Write_Log "Error executing section '$sectionName': $_" "red"
			$failedSections += $sectionName
		}
	}
	
	# ------------------------------------------------------------------------------------------------
	# Completion summary
	# ------------------------------------------------------------------------------------------------
	if ($failedSections.Count -eq 0)
	{
		Write_Log "`r`nAll SQL sections executed successfully." "green"
	}
	else
	{
		Write_Log "The following sections failed: $($failedSections -join ', ')" "red"
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
					# Use -Include and -Exclude for efficient batch deletion
					$matchedItems = Get-ChildItem -Path $targetPath -Include $SpecifiedFiles -Exclude $Exclusions -Recurse -Force -ErrorAction SilentlyContinue
					
					if ($matchedItems)
					{
						$matchedItems | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
						$deletedCount = $matchedItems.Count # Approximate count (assumes most succeed)
					}
					else
					{
						# Write_Log "No items matched the specified patterns in '$targetPath'." "Yellow"
					}
				}
				else
				{
					# Delete all except exclusions
					$allItems = Get-ChildItem -Path $targetPath -Exclude $Exclusions -Recurse -Force -ErrorAction SilentlyContinue
					
					if ($allItems)
					{
						$allItems | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
						$deletedCount = $allItems.Count # Approximate
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
				# Use -Include and -Exclude for efficient batch deletion
				$matchedItems = Get-ChildItem -Path $targetPath -Include $SpecifiedFiles -Exclude $Exclusions -Recurse -Force -ErrorAction SilentlyContinue
				
				if ($matchedItems)
				{
					$matchedItems | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
					$deletedCount = $matchedItems.Count # Approximate count (assumes most succeed)
				}
				else
				{
					Write_Log "No items matched the specified patterns in '$targetPath'." "Yellow"
				}
			}
			else
			{
				# Delete all except exclusions
				$allItems = Get-ChildItem -Path $targetPath -Exclude $Exclusions -Recurse -Force -ErrorAction SilentlyContinue
				
				if ($allItems)
				{
					$allItems | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
					$deletedCount = $allItems.Count # Approximate
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
#   (with embedded fixed header/footer), and logs all progress and errors.
#   Tries protocol execution first (from pre-populated protocol table), falls back to file writing.
#   Protocol detection is NOT performed here; background jobs must fill $script:LaneProtocols.
# ===================================================================================================

function Process_Lanes
{
	param (
		[string]$StoreNumber,
		[switch]$ProcessAllLanes
	)
	
	# ----------------------------------------
	# Banner: Start
	# ----------------------------------------
	Write_Log "`r`n==================== Starting Process_Lanes Function ====================`r`n" "blue"
	
	# ----------------------------------------
	# Check for required OfficePath
	# ----------------------------------------
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}
	
	# ----------------------------------------
	# Import detected SQL module for Invoke-Sqlcmd usage
	# ----------------------------------------
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		Import-Module $SqlModuleName -ErrorAction Stop
		$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
	}
	else
	{
		Write_Log "No valid SQL module available for SQL operations!" "red"
		return
	}
	
	# ----------------------------------------
	# Check for available Lane Machines
	# ----------------------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "No lanes available. Please retrieve nodes first." "red"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}
	$MachineNameToLaneNum = $script:FunctionResults['MachineNameToLaneNum']
	if (-not $MachineNameToLaneNum)
	{
		$MachineNameToLaneNum = @{ }
	}
	
	# ----------------------------------------
	# Get user's lane selection
	# ----------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($selection -eq $null)
	{
		Write_Log "Lane processing canceled by user." "yellow"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}
	
	# Support both string and object selections for lanes (for future-proofing GUI node returns)
	if ($selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		# Detect if these are node objects or just string numbers
		if ($selection.Lanes[0] -is [PSCustomObject] -and $selection.Lanes[0].PSObject.Properties.Name -contains 'LaneNumber')
		{
			$Lanes = $selection.Lanes | ForEach-Object { $_.LaneNumber }
		}
		else
		{
			$Lanes = $selection.Lanes
		}
	}
	else
	{
		Write_Log "No lanes selected." "yellow"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}
	
	# ----------------------------------------
	# Parse Lane SQL script sections for processing
	# ----------------------------------------
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$matches = [regex]::Matches($script:LaneSQLScript, $sectionPattern)
	if ($matches.Count -eq 0)
	{
		Write_Log "No sections found in Lane SQL script." "red"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
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
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}
	
	# ----------------------------------------
	# Pre-build the SQI script for fallback (same for all lanes)
	# ----------------------------------------
	$topBlock = "/* Set a long timeout so the entire script runs */`r`n@WIZRPL(DBASE_TIMEOUT=E);`r`n" +
	"--------------------------------------------------------------------------------`r`n"
	$bottomBlock = "--------------------------------------------------------------------------------`r`n" +
	"/* Clear the long database timeout */`r`n@WIZCLR(DBASE_TIMEOUT);"
	$middleBlock = ($matches | Where-Object {
			$SectionsToSend -contains $_.Groups['SectionName'].Value.Trim()
		}) | ForEach-Object {
		"/* $($_.Groups['SectionName'].Value.Trim()) */`r`n$($_.Groups['SQLCommands'].Value.Trim())"
	} | Out-String
	$finalScript = $topBlock + $middleBlock + $bottomBlock
	
	# ----------------------------------------
	# MULTIPLE lanes: Always file fallback
	# ----------------------------------------
	if ($Lanes.Count -gt 1)
	{
		Write_Log "Multiple lanes selected, using file-based fallback for all lanes." "yellow"
		foreach ($LaneNumber in $Lanes)
		{
			$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
			if (-not $laneInfo)
			{
				Write_Log "Could not get DB info for lane $LaneNumber. Skipping." "yellow"
				continue
			}
			if ($MachineNameToLaneNum -and $MachineNameToLaneNum.ContainsKey($LaneNumber))
			{
				$laneNum = $MachineNameToLaneNum[$LaneNumber]
			}
			else
			{
				$laneNum = $LaneNumber
			}
			$LaneLocalPath = Join-Path $OfficePath ("XF" + $StoreNumber + $laneNum)
			$machineName = $laneInfo['MachineName']
			# Write_Log "Protocol not attempted (file-based fallback used for all lanes) on $machineName." "gray"
			if (Test-Path $LaneLocalPath)
			{
				Write_Log "Writing Lane_Database_Maintenance.sqi to Lane $LaneNumber ($machineName)..." "blue"
				try
				{
					Set-Content -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Value $finalScript -Encoding Ascii
					Set-ItemProperty -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
					Write_Log "Created and wrote to file at Lane #${LaneNumber} ($machineName) successfully. (file copy)" "green"
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
	}
	else
	{
		# SINGLE lane: Try protocol, fallback to file
		foreach ($LaneNumber in $Lanes)
		{
			$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
			if (-not $laneInfo)
			{
				Write_Log "Could not get DB info for lane $LaneNumber. Skipping." "yellow"
				continue
			}
			
			$machineName = $laneInfo['MachineName']
			$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
			$tcpConnStr = $laneInfo['TcpConnStr']
			if ($MachineNameToLaneNum -and $MachineNameToLaneNum.ContainsKey($LaneNumber))
			{
				$laneNum = $MachineNameToLaneNum[$LaneNumber]
			}
			else
			{
				$laneNum = $LaneNumber
			}
			$LaneLocalPath = Join-Path $OfficePath ("XF" + $StoreNumber + $laneNum)
			
			# Get protocol for this lane
			$laneKey = $LaneNumber.PadLeft(3, '0')
			$protocolType = $script:LaneProtocols[$laneKey]
			$workingConnStr = $null
			if ($protocolType -eq "Named Pipes") { $workingConnStr = $namedPipesConnStr }
			elseif ($protocolType -eq "TCP") { $workingConnStr = $tcpConnStr }
			
			# DEBUG: Show actual conn string to log for this lane
			if ([string]::IsNullOrEmpty($protocolType)) { $protocolType = "File" }
			if ([string]::IsNullOrEmpty($workingConnStr)) { $workingConnStr = "File" }
			Write_Log "Lane $LaneNumber uses protocol: $protocolType" "gray"
			Write_Log "Lane $LaneNumber connection string: $workingConnStr" "gray"
			
			# If protocol not ready, fallback to file
			if (-not $protocolType -or $protocolType -eq "File" -or -not $workingConnStr)
			{
				Write_Log "Protocol not ready or unavailable for $machineName. Skipping protocol and using file copy." "yellow"
				if (Test-Path $LaneLocalPath)
				{
					Write_Log "`r`nProcessing $machineName using file fallback..." "blue"
					Write_Log "Lane path found: $LaneLocalPath" "blue"
					Write_Log "Writing Lane_Database_Maintenance.sqi to Lane..." "blue"
					try
					{
						Set-Content -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Value $finalScript -Encoding Ascii
						Set-ItemProperty -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
						Write_Log "Created and wrote to file at Lane #${LaneNumber} ($machineName) successfully. (file copy)" "green"
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
				continue
			}
			
			# ----------------------------------------
			# Protocol execution: Try SQL via protocol
			# ----------------------------------------
			$protocolWorked = $false
			$server = $laneInfo['TcpServer']
			$database = $laneInfo['DBName']
			$currentLogin = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
			
			try
			{
				$matchesFiltered = [regex]::Matches($script:LaneSQLFiltered, $sectionPattern)
				$sections = ($matchesFiltered | Where-Object { $SectionsToSend -contains $_.Groups['SectionName'].Value.Trim() })
				foreach ($match in $sections)
				{
					$sectionName = $match.Groups['SectionName'].Value.Trim()
					$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()
					Write_Log "Executing section: '$sectionName' on $machineName" "blue"
					Write_Log "--------------------------------------------------------------------------------"
					Write_Log "$sqlCommands" "orange"
					Write_Log "--------------------------------------------------------------------------------"
					$querySucceeded = $false
					$retriedForMapping = $false
					for ($try = 1; $try -le 2; $try++)
					{
						try
						{
							# Always use the correct module and per-lane connection string here
							& $InvokeSqlCmd -ConnectionString $workingConnStr -Query $sqlCommands -QueryTimeout 0 -ErrorAction Stop
							Write_Log "Section '$sectionName' executed successfully on $machineName using ($protocolType)." "green"
							$protocolWorked = $true
							$querySucceeded = $true
							break
						}
						catch
						{
							$errorMsg = $_.Exception.Message
							if ($errorMsg -match 'Login failed for user' -and -not $retriedForMapping)
							{
								Write_Log "Login failed for $currentLogin on $machineName. Attempting to map user and retry..." "yellow"
								$checkUserQuery = "SELECT COUNT(*) AS UserExists FROM sys.database_principals WHERE name = '$currentLogin'"
								try
								{
									$userExists = (& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $checkUserQuery -ErrorAction Stop).UserExists
								}
								catch { $userExists = 0 }
								if ($userExists -eq 0)
								{
									$createUserSql = @"
USE [$database];
CREATE USER [$currentLogin] FOR LOGIN [$currentLogin];
ALTER ROLE db_owner ADD MEMBER [$currentLogin];
"@
									try
									{
										& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $createUserSql -ErrorAction Stop
										Write_Log "Mapped and granted db_owner to $currentLogin in [$database]." "blue"
									}
									catch
									{
										Write_Log "Failed to map $currentLogin in [$database]: $_" "yellow"
									}
								}
								else
								{
									try
									{
										& $InvokeSqlCmd -ServerInstance $server -Database $database -Query "ALTER ROLE db_owner ADD MEMBER [$currentLogin];" -ErrorAction Stop
									}
									catch { }
									Write_Log "$currentLogin already mapped in [$database]." "gray"
								}
								$retriedForMapping = $true
								continue # Retry after mapping
							}
							elseif ($errorMsg -match 'Login failed for user')
							{
								Write_Log "Login failed for $currentLogin on $machineName. ServerInstance fallback will NOT be attempted (would repeat login failure)..." "yellow"
								$protocolWorked = $false
								break
							}
							else
							{
								Write_Log "ConnectionString method failed for section '$sectionName' on [$machineName]: $_. Trying ServerInstance fallback..." "yellow"
								try
								{
									& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $sqlCommands -QueryTimeout 0 -ErrorAction Stop
									Write_Log "Section '$sectionName' executed successfully on $machineName using ($protocolType, ServerInstance fallback)." "green"
									$protocolWorked = $true
									$querySucceeded = $true
								}
								catch
								{
									Write_Log "ServerInstance method also failed for section '$sectionName' on [$machineName]: $_" "red"
									$protocolWorked = $false
								}
								break
							}
						}
					}
					if (-not $querySucceeded) { break }
				}
			}
			catch
			{
				Write_Log "Failed to execute a section on $machineName via protocol: $_. Falling back to file." "yellow"
				$protocolWorked = $false
			}
			if ($protocolWorked)
			{
				if (-not ($script:ProcessedLanes -contains $LaneNumber))
				{
					$script:ProcessedLanes += $LaneNumber
				}
				continue
			}
			
			# Fallback: Classic file-based method
			if (Test-Path $LaneLocalPath)
			{
				Write_Log "`r`nProcessing $machineName using file copy..." "blue"
				Write_Log "Lane path found: $LaneLocalPath" "blue"
				Write_Log "Writing Lane_Database_Maintenance.sqi to Lane..." "blue"
				try
				{
					Set-Content -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Value $finalScript -Encoding Ascii
					Set-ItemProperty -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
					Write_Log "Created and wrote to file at Lane #${LaneNumber} ($machineName) successfully. (file copy)" "green"
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
	}
	
	# Final: Report processed lanes and finish banner
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
			$updateOutput = & "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -SignatureUpdate 2>&1
			Write_Log "Windows Defender signatures update output: $updateOutput" "cyan"
			Write_Log "Windows Defender signatures updated successfully." "green"
			Write_Log "Running Windows Defender full scan..." "blue"
			$scanOutput = & "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -Scan -ScanType 2 2>&1
			Write_Log "Windows Defender full scan output: $scanOutput" "cyan"
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
			$dismCleanupOutput = DISM /Online /Cleanup-Image /StartComponentCleanup /NoRestart 2>&1
			Write_Log "DISM StartComponentCleanup output: $dismCleanupOutput" "cyan"
			Write_Log "DISM StartComponentCleanup completed." "green"
			Write_Log "Running DISM RestoreHealth..." "blue"
			$dismRestoreOutput = DISM /Online /Cleanup-Image /RestoreHealth /NoRestart 2>&1
			Write_Log "DISM RestoreHealth output: $dismRestoreOutput" "cyan"
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
			$sfcOutput = SFC /scannow 2>&1
			Write_Log "System File Checker output: $sfcOutput" "cyan"
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
				$optimizeOutput = Optimize-Volume -DriveLetter $_.DriveLetter -Verbose 2>&1
				Write_Log "Optimize drive output: $optimizeOutput" "cyan"
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
			$checkDiskOutput = Start-Process "cmd.exe" -ArgumentList "/c echo Y|chkdsk C: /f /r" -Verb RunAs -Wait -PassThru
			Write_Log "Check Disk scheduling output: $checkDiskOutput" "cyan"
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
	
	if (-not (Test-Path $LoadPath))
	{
		Write_Log "`r`nLoad Base Path not found: $LoadPath" "yellow"
		return
	}
	
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($selection -eq $null)
	{
		Write_Log "`r`nLane processing canceled by user." "yellow"
		return
	}
	
	$Type = $selection.Type
	$Lanes = $selection.Lanes
	$processAllLanes = $false
	if ($Type -eq "All") { $processAllLanes = $true }
	
	if ($processAllLanes)
	{
		try
		{
			$LaneMachineNames = $script:FunctionResults['LaneMachineNames']
			if ($LaneMachineNames -and $LaneMachineNames.Count -gt 0)
			{
				$Lanes = $LaneMachineNames
			}
			else
			{
				throw "LaneMachineNames is empty or not available."
			}
		}
		catch
		{
			Write_Log "Failed to retrieve LaneMachineNames: $_. Falling back to user-selected lanes." "yellow"
			$processAllLanes = $false
		}
	}
	
	$ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)
	$runLoadFilename = "run_load.sql"
	$lnkLoadFilename = "lnk_load.sql"
	$stoLoadFilename = "sto_load.sql"
	$terLoadFilename = "ter_load.sql"
	
	# Original SQL data (unchanged for file copy)
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
	
	foreach ($laneNumber in $Lanes)
	{
		$protocolWorked = @{ }
		$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNumber
		$MachineName = $null
		$namedPipesConnStr = $null
		if ($laneInfo)
		{
			$MachineName = $laneInfo['MachineName']
			$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
			Write_Log "Lane #${laneNumber}: Connection info found. Machine: $MachineName" "green"
		}
		else
		{
			$MachineName = $script:FunctionResults['LaneNumToMachineName'][$laneNumber]
			if (-not $MachineName) { $MachineName = "POS${laneNumber}" }
			Write_Log "Lane #${laneNumber}: Fallback machine name '$MachineName'" "yellow"
		}
		
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
		$fileCopyNeeded = @()
		
		# --- FIX: Parse run_load rows ahead of time ---
		$runRows =
		$runLoadScript -split "INSERT INTO Run_Load VALUES", 2 |
		Select-Object -Last 1 |
		ForEach-Object { ($_ -replace "(?ms);.*", "") -split "`r?`n" } |
		Where-Object { $_ -match "^\s*\(" }
		
		if ($namedPipesConnStr -and (Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue))
		{
			foreach ($tableJob in @(
					@{
						Table    = 'RUN_TAB'
						Filename = $runLoadFilename
						Rows	 = $runRows
					}
					@{
						Table																																																										     = 'LNK_TAB'
						Filename																																																										 = $lnkLoadFilename
						Rows																																																											 = @(
							"('${laneNumber}','${StoreNumber}','${laneNumber}')",
							"('DSM','${StoreNumber}','${laneNumber}')",
							"('PAL','${StoreNumber}','${laneNumber}')",
							"('RAL','${StoreNumber}','${laneNumber}')",
							"('XAL','${StoreNumber}','${laneNumber}')"
						)
					}
					@{
						Table																																																		   = 'STO_TAB'
						Filename																																																	   = $stoLoadFilename
						Rows																																																		   = @(
							"('${laneNumber}','Terminal ${laneNumber}',1,1,1,,,,)",
							"('DSM','Deploy SMS',1,1,1,,,,)",
							"('PAL','Program all',0,0,1,1,,,)",
							"('RAL','Report all',1,0,0,,,,)",
							"('XAL','Exchange all',0,1,0,,,,)"
						)
					}
					@{
						Table																																																				     = 'TER_TAB'
						Filename																																																				 = $terLoadFilename
						Rows																																																					 = @(
							"('${StoreNumber}','${laneNumber}','Terminal ${laneNumber}','\\${MachineName}\storeman\office\XF${StoreNumber}${laneNumber}\','\\${MachineName}\storeman\office\XF${StoreNumber}901\')",
							"('${StoreNumber}','901','Server','','')"
						)
					}
				))
			{
				$table = $tableJob.Table
				$rows = $tableJob.Rows
				
				try
				{
					$srcConn = $script:FunctionResults['ConnectionString']
					$srcSqlConn = New-Object System.Data.SqlClient.SqlConnection($srcConn)
					$srcSqlConn.Open()
					$cmd = $srcSqlConn.CreateCommand()
					$cmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '$table'
ORDER BY ORDINAL_POSITION
"@
					$rdr = $cmd.ExecuteReader()
					$colDefs = @()
					while ($rdr.Read())
					{
						$col = $rdr["COLUMN_NAME"]
						$type = $rdr["DATA_TYPE"]
						$nullText = if ($rdr["IS_NULLABLE"] -eq "NO") { "NOT NULL" }
						else { "NULL" }
						switch ($type)
						{
							'nvarchar' { $typeText = "nvarchar($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'varchar'  { $typeText = "varchar($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'char'     { $typeText = "char($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'nchar'    { $typeText = "nchar($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'decimal'  { $typeText = "decimal($($rdr["NUMERIC_PRECISION"]),$($rdr["NUMERIC_SCALE"]))" }
							'numeric'  { $typeText = "numeric($($rdr["NUMERIC_PRECISION"]),$($rdr["NUMERIC_SCALE"]))" }
							default    { $typeText = $type }
						}
						$colDefs += "[$col] $typeText $nullText"
					}
					$rdr.Close()
					$srcSqlConn.Close()
					$colDefsText = $colDefs -join ", "
					$dropCreate = "IF OBJECT_ID(N'[$table]', N'U') IS NOT NULL DROP TABLE [$table]; CREATE TABLE [$table] ($colDefsText);"
					Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $dropCreate -QueryTimeout 30 -ErrorAction Stop
					Write_Log "Dropped and recreated $table on lane $laneNumber via protocol." "green"
					
					# Get column names for the insert
					$schema = Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '$table' ORDER BY ORDINAL_POSITION"
					$colNames = $schema | ForEach-Object { $_.COLUMN_NAME }
					
					$successfulRows = 0
					foreach ($row in $rows)
					{
						$cleanRow = $row.Trim().TrimEnd(',', ';')
						if ($cleanRow.StartsWith("(") -and $cleanRow.EndsWith(")"))
						{
							$cleanRow = $cleanRow.Substring(1, $cleanRow.Length - 2)
						}
						$splitRegex = ',(?=(?:[^'']*''[^'']*'')*[^'']*$)'
						$values = [regex]::Split($cleanRow, $splitRegex) | ForEach-Object {
							$v = $_.Trim()
							if ($v -eq "" -or $v -eq "' '") { "NULL" }
							else
							{
								if ($v.StartsWith("'") -and $v.EndsWith("'"))
								{
									$unquoted = $v.Trim("'")
									$res = $unquoted
									$res = $res -replace "@STORE", $StoreNumber.PadLeft(3, "0")
									$res = $res -replace "@TER", $laneNumber.PadLeft(3, "0")
									$res = $res -replace "@USER", $env:USERNAME
									$res = $res -replace "@USERNAME", $env:USERNAME
									$res = $res -replace "@RUN", "C:\storeman"
									$res = $res -replace "@OFFICE", "C:\storeman\office"
									$now = Get-Date
									$res = $res -replace "@TIME", $now.ToString("HHmm")
									$res = $res -replace "@NOW", $now.ToString("HHmmss")
									$res = $res -replace "@DSSF", $now.ToString("yyyyMMdd")
									$res = $res -replace "@DSW-001", $now.AddDays(-1).ToString("yyyyMMdd")
									$res = $res -replace "@DSW-008", $now.AddDays(-8).ToString("yyyyMMdd")
									$res = $res -replace "@DSS\+001", $now.AddDays(1).ToString("yyyyMMdd")
									$res = $res -replace "@DSS", $now.ToString("yyyyMMdd")
									"'" + $res + "'"
								}
								else { $v }
							}
						}
						# Fill with NULLs or trim
						if ($values.Count -lt $colNames.Count)
						{
							$values += @("NULL") * ($colNames.Count - $values.Count)
						}
						elseif ($values.Count -gt $colNames.Count)
						{
							$values = $values[0 .. ($colNames.Count - 1)]
						}
						$sql = "INSERT INTO $table ([{0}]) VALUES ({1})" -f ($colNames -join "],["), ($values -join ", ")
						try
						{
							Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $sql -QueryTimeout 30 -ErrorAction Stop
							$successfulRows++
						}
						catch
						{
							Write_Log "$table row insert failed: $sql; Error: $_" "red"
						}
					}
					if ($successfulRows -gt 0)
					{
						Write_Log "Inserted $successfulRows rows into $table on lane $laneNumber via Named Pipes protocol." "green"
						$protocolWorked[$table] = $true
						$actionSummaries += "Protocol loaded $($tableJob.Filename)"
					}
					else
					{
						Write_Log "No data inserted to $table on lane $laneNumber." "yellow"
						$fileCopyNeeded += $tableJob.Filename
						$protocolWorked[$table] = $false
					}
				}
				catch
				{
					Write_Log "Protocol copy failed for $table on lane [$laneNumber]: $_" "red"
					$fileCopyNeeded += $tableJob.Filename
					$protocolWorked[$table] = $false
				}
			}
		}
		else
		{
			# No protocol possible at all, mark all as needing fallback
			$fileCopyNeeded += $runLoadFilename, $lnkLoadFilename, $stoLoadFilename, $terLoadFilename
		}
		
		# --- Fallback file copy for only failed tables (preserving original format) ---
		$tableFileData = @{
			$runLoadFilename = $runLoadScript
			$lnkLoadFilename = $lnkLoadHeader + "`r`n" +
			"('${laneNumber}','${StoreNumber}','${laneNumber}'),`r`n" +
			"('DSM','${StoreNumber}','${laneNumber}'),`r`n" +
			"('PAL','${StoreNumber}','${laneNumber}'),`r`n" +
			"('RAL','${StoreNumber}','${laneNumber}'),`r`n" +
			"('XAL','${StoreNumber}','${laneNumber}');`r`n" +
			"`r`n" + $lnkLoadFooter.TrimStart() + "`r`n"
			$stoLoadFilename = $stoLoadHeader + "`r`n" +
			"('${laneNumber}','Terminal ${laneNumber}',1,1,1,,,,),`r`n" +
			"('DSM','Deploy SMS',1,1,1,,,,),`r`n" +
			"('PAL','Program all',0,0,1,1,,,),`r`n" +
			"('RAL','Report all',1,0,0,,,,),`r`n" +
			"('XAL','Exchange all',0,1,0,,,,);`r`n" +
			"`r`n" + $stoLoadFooter.TrimStart() + "`r`n"
			$terLoadFilename = $terLoadHeader + "`r`n" +
			"('${StoreNumber}','${laneNumber}','Terminal ${laneNumber}','\\${MachineName}\storeman\office\XF${StoreNumber}${laneNumber}\','\\${MachineName}\storeman\office\XF${StoreNumber}901\'),`r`n" +
			"('${StoreNumber}','901','Server','','');`r`n" +
			"`r`n" + $terLoadFooter.TrimStart() + "`r`n"
		}
		foreach ($file in $fileCopyNeeded | Select-Object -Unique)
		{
			try
			{
				$dest = Join-Path -Path $laneFolder.FullName -ChildPath $file
				[System.IO.File]::WriteAllText($dest, $tableFileData[$file], $ansiPcEncoding)
				Set-ItemProperty -Path $dest -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
				$actionSummaries += "Copied $file"
			}
			catch
			{
				$actionSummaries += "Failed to copy $file"
			}
		}
		
		$summaryMessage = "Lane ${laneNumber} (Machine: ${MachineName}): " + ($actionSummaries -join "; ")
		Write_Log $summaryMessage "green"
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
#   and copy to the specified lanes or hosts. Uses cached protocol (from $script:LaneProtocols)
#   for each lane to determine protocol or file copy. Uses LaneNumToMachineName mapping.
# ===================================================================================================

function Pump_Tables
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	# Log function start
	Write_Log "`r`n==================== Starting Pump_Tables Function ====================`r`n" "blue"
	
	# Ensure OfficePath is present and valid
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# ------------------------------------------------------------------------------------------------
	# STEP 1: Load and validate lane mappings from Retrieve_Nodes (and allow for all flexible lookups)
	# ------------------------------------------------------------------------------------------------
	if (-not ($script:FunctionResults.ContainsKey('LaneNumToMachineName')))
	{
		Write_Log "Lane mappings not found. Please run Retrieve_Nodes first." "red"
		return
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$MachineNameToLaneNum = $script:FunctionResults['MachineNameToLaneNum']
	
	# ------------------------------------------------------------------------------------------------
	# STEP 2: Prompt user for lane selection via GUI
	# ------------------------------------------------------------------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($null -eq $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "Lane processing canceled by user." "yellow"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	
	# Accept all supported forms (lane number, machine name, POSxxx, etc)
	$Lanes = $selection.Lanes
	
	# Force all selections to 3-digit lane number for folder/DB logic using mappings from Retrieve_Nodes
	$Lanes = $Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($MachineNameToLaneNum.ContainsKey($_)) { $MachineNameToLaneNum[$_] }
		else { $_ }
	}
	$Lanes = $Lanes | Where-Object { $LaneNumToMachineName.ContainsKey($_) }
	
	if (-not $Lanes -or $Lanes.Count -eq 0)
	{
		Write_Log "No valid lanes to process. Exiting Pump_Tables." "yellow"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	
	# ------------------------------------------------------------------------------------------------
	# STEP 3: Retrieve table aliases and prompt for tables to process
	# ------------------------------------------------------------------------------------------------
	if ($script:FunctionResults.ContainsKey('Get_Table_Aliases'))
	{
		$aliasData = $script:FunctionResults['Get_Table_Aliases']
		$aliasResults = $aliasData.Aliases
		$aliasHash = $aliasData.AliasHash
	}
	else
	{
		Write_Log "Alias data not found. Ensure Get_Table_Aliases has been run." "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	if ($aliasResults.Count -eq 0)
	{
		Write_Log "No tables found to process. Exiting Pump_Tables." "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	
	# Prompt user to select which tables to pump (alias objects or just names)
	$selectedTables = Show_Table_Selection_Form -AliasResults $aliasResults
	if (-not $selectedTables -or $selectedTables.Count -eq 0)
	{
		Write_Log "No tables were selected. Exiting Pump_Tables." "yellow"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	
	# ------------------------------------------------------------------------------------------------
	# STEP 4: Validate SQL Connection, and load SQL module only once
	# ------------------------------------------------------------------------------------------------
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write_Log "Connection string not found. Cannot proceed with Pump_Tables." "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		Import-Module $SqlModuleName -ErrorAction Stop
		$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
	}
	else
	{
		Write_Log "No valid SQL module available for SQL operations!" "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	$supportsConnectionString = $false
	if ($InvokeSqlCmd) { $supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString' }
	
	# Open SQL connection to source DB (ADO.NET, for schema/data pull)
	$srcSqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$srcSqlConnection.ConnectionString = $ConnectionString
	$srcSqlConnection.Open()
	
	# Prepare for tracking protocol/file status
	$ProcessedLanes = @()
	$protocolLanes = @()
	$fileCopyLanes = @()
	
	# Filter alias entries that match user's selection
	$filteredAliasEntries = $aliasResults | Where-Object { $selectedTables -contains $_.Table }
	
	# ------------------------------------------------------------------------------------------------
	# STEP 5: Process each lane: Try protocol (Named Pipes or TCP), else fall back to file copy
	# ------------------------------------------------------------------------------------------------
	foreach ($laneNum in $Lanes)
	{
		$machineName = $LaneNumToMachineName[$laneNum]
		
		if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
		{
			Write_Log "Lane #${laneNum}: Machine name is invalid or unknown. Skipping." "yellow"
			continue
		}
		
		# Retrieve DB connection info for the lane (connection strings, etc)
		$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNum
		if (-not $laneInfo)
		{
			Write_Log "Could not get DB info for lane $laneNum. Skipping." "yellow"
			continue
		}
		$connString = $laneInfo['ConnectionString']
		$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
		$tcpConnStr = $laneInfo['TcpConnStr']
		
		# Always use numeric lane number for folder path!
		$LaneLocalPath = Join-Path $OfficePath "XF${StoreNumber}${laneNum}"
		
		# Use cached protocol type if available, else fallback to File
		$protocolType = if ($script:LaneProtocols.ContainsKey($laneNum)) { $script:LaneProtocols[$laneNum] }
		else { "File" }
		$laneSqlConn = $null
		$protocolWorked = $false
		
		# ----------------------------------------------------------------------------------------
		# Try direct SQL protocol copy (Named Pipes/TCP)
		# ----------------------------------------------------------------------------------------
		if ($protocolType -eq "Named Pipes" -or $protocolType -eq "TCP")
		{
			try
			{
				if ($protocolType -eq "Named Pipes")
				{
					$laneSqlConn = New-Object System.Data.SqlClient.SqlConnection $namedPipesConnStr
				}
				elseif ($protocolType -eq "TCP")
				{
					$laneSqlConn = New-Object System.Data.SqlClient.SqlConnection $tcpConnStr
				}
				$laneSqlConn.Open()
				if ($laneSqlConn.State -eq 'Open')
				{
					Write_Log "`r`nCopying data to Lane $laneNum ($machineName) via SQL protocol [$protocolType]..." "blue"
					foreach ($aliasEntry in $filteredAliasEntries)
					{
						$table = $aliasEntry.Table
						Write_Log "Pumping table '$table' to lane $laneNum ($machineName) via SQL..." "blue"
						try
						{
							# Build CREATE TABLE from source schema
							$schemaCmd = $srcSqlConnection.CreateCommand()
							$schemaCmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '$table'
ORDER BY ORDINAL_POSITION
"@
							$reader = $schemaCmd.ExecuteReader()
							$colDefs = @()
							while ($reader.Read())
							{
								$colName = $reader["COLUMN_NAME"]
								$dataType = $reader["DATA_TYPE"]
								$isNullable = $reader["IS_NULLABLE"]
								$nullText = if ($isNullable -eq "NO") { "NOT NULL" }
								else { "NULL" }
								switch ($dataType)
								{
									'nvarchar' { $length = $reader["CHARACTER_MAXIMUM_LENGTH"]; $typeText = "nvarchar($length)" }
									'varchar'  { $length = $reader["CHARACTER_MAXIMUM_LENGTH"]; $typeText = "varchar($length)" }
									'char'     { $length = $reader["CHARACTER_MAXIMUM_LENGTH"]; $typeText = "char($length)" }
									'nchar'    { $length = $reader["CHARACTER_MAXIMUM_LENGTH"]; $typeText = "nchar($length)" }
									'decimal'  { $prec = $reader["NUMERIC_PRECISION"]; $scale = $reader["NUMERIC_SCALE"]; $typeText = "decimal($prec,$scale)" }
									'numeric'  { $prec = $reader["NUMERIC_PRECISION"]; $scale = $reader["NUMERIC_SCALE"]; $typeText = "numeric($prec,$scale)" }
									default    { $typeText = $dataType }
								}
								$colDefs += "[$colName] $typeText $nullText"
							}
							$reader.Close()
							$colDefsText = $colDefs -join ", "
							$createTableSQL = "IF OBJECT_ID(N'[$table]', N'U') IS NOT NULL DROP TABLE [$table]; CREATE TABLE [$table] ($colDefsText);"
							
							# Drop/recreate table structure
							$cmdCreate = $laneSqlConn.CreateCommand()
							$cmdCreate.CommandText = $createTableSQL
							$cmdCreate.ExecuteNonQuery() | Out-Null
							Write_Log "Recreated table structure for '$table' on $machineName" "green"
							
							# Select data from source and insert into target lane
							$dataQuery = "SELECT * FROM [$table]"
							$cmdSource = $srcSqlConnection.CreateCommand()
							$cmdSource.CommandText = $dataQuery
							$readerSource = $cmdSource.ExecuteReader()
							$schemaTable = $readerSource.GetSchemaTable()
							$colNames = $schemaTable | ForEach-Object { $_["ColumnName"] }
							$insertPrefix = "INSERT INTO [$table] ([$($colNames -join '],[')]) VALUES "
							$rowCountCopied = 0
							while ($readerSource.Read())
							{
								$values = @()
								foreach ($col in $colNames)
								{
									$val = $readerSource[$col]
									if ($val -eq $null -or $val -is [System.DBNull])
									{
										$values += "NULL"
									}
									elseif ($val -is [string])
									{
										$escaped = $val.Replace("'", "''")
										$values += "'$escaped'"
									}
									elseif ($val -is [datetime])
									{
										$values += "'" + $val.ToString("yyyy-MM-dd HH:mm:ss") + "'"
									}
									else
									{
										$values += $val.ToString()
									}
								}
								$insertCmd = $laneSqlConn.CreateCommand()
								$insertCmd.CommandText = $insertPrefix + "(" + ($values -join ",") + ")"
								$insertCmd.ExecuteNonQuery() | Out-Null
								$rowCountCopied++
							}
							$readerSource.Close()
							Write_Log "Copied $rowCountCopied rows to $table on lane $laneNum ($machineName) (SQL protocol)." "green"
						}
						catch
						{
							Write_Log "Failed to copy table '$table' to lane $laneNum ($machineName) via SQL: $_" "red"
						}
					}
					$laneSqlConn.Close()
					$protocolWorked = $true
					$protocolLanes += $laneNum
					$ProcessedLanes += $laneNum
				}
			}
			catch
			{
				Write_Log "SQL protocol copy failed for lane [$laneNum] ($machineName) ($protocolType): $_" "yellow"
				$protocolWorked = $false
			}
		}
		
		# --------------------------------------------------------------------------------------------
		# File fallback (if protocol copy not possible)
		# --------------------------------------------------------------------------------------------
		if (-not $protocolWorked)
		{
			if (Test-Path $LaneLocalPath)
			{
				Write_Log "`r`nCopying via FILE fallback for Lane #$laneNum ($machineName)..." "blue"
				foreach ($aliasEntry in $filteredAliasEntries)
				{
					$table = $aliasEntry.Table
					$baseTable = $table -replace '_TAB$', ''
					$sqlFileName = "${baseTable}_Load.sql"
					$localTempPath = Join-Path $env:TEMP $sqlFileName
					
					# Optionally reuse an existing file if < 1 hour old (for efficiency)
					$useExistingFile = $false
					if (Test-Path $localTempPath)
					{
						$fileInfo = Get-Item $localTempPath
						$fileAge = (Get-Date) - $fileInfo.LastWriteTime
						if ($fileAge.TotalHours -le 1) { $useExistingFile = $true }
					}
					
					if (-not $useExistingFile)
					{
						try
						{
							$ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)
							$streamWriter = New-Object System.IO.StreamWriter($localTempPath, $false, $ansiPcEncoding)
							$streamWriter.NewLine = "`r`n"
							
							# Get columns
							$columnDataTypesQuery = @"
SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '$table'
ORDER BY ORDINAL_POSITION
"@
							$cmdColumnTypes = $srcSqlConnection.CreateCommand()
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
							
							# Primary key
							$pkQuery = @"
SELECT c.COLUMN_NAME
FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE c
    ON c.CONSTRAINT_NAME = tc.CONSTRAINT_NAME
    AND c.TABLE_NAME = tc.TABLE_NAME
WHERE tc.TABLE_NAME = '$table' AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
ORDER BY c.ORDINAL_POSITION
"@
							$cmdPK = $srcSqlConnection.CreateCommand()
							$cmdPK.CommandText = $pkQuery
							$readerPK = $cmdPK.ExecuteReader()
							$pkColumns = @()
							while ($readerPK.Read()) { $pkColumns += $readerPK["COLUMN_NAME"] }
							$readerPK.Close()
							
							if ($pkColumns.Count -eq 0)
							{
								$primaryKeyColumns = @()
								$cmdFirstColumn = $srcSqlConnection.CreateCommand()
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
							
							$keyString = ($primaryKeyColumns | ForEach-Object { "$_=:$_" }) -join " AND "
							$viewName = $baseTable.Substring(0, 1).ToUpper() + $baseTable.Substring(1).ToLower() + '_Load'
							$columnList = ($columnDataTypes.Keys) -join ','
							
							# Write header
							$header = @"
@WIZRPL(DBASE_TIMEOUT=E);

CREATE VIEW $viewName AS SELECT $columnList FROM $table;

INSERT INTO $viewName VALUES
"@
							$header = $header -replace "(\r\n|\n|\r)", "`r`n"
							$streamWriter.WriteLine($header.TrimEnd())
							
							# Dump data as INSERTs (row by row)
							$dataQuery = "SELECT * FROM [$table]"
							$cmdData = $srcSqlConnection.CreateCommand()
							$cmdData.CommandText = $dataQuery
							$readerData = $cmdData.ExecuteReader()
							$firstRow = $true
							while ($readerData.Read())
							{
								if ($firstRow) { $firstRow = $false }
								else { $streamWriter.WriteLine(",") }
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
											if ([math]::Floor($val) -eq $val) { $values += $val.ToString() }
											else { $values += $val.ToString("0.00") }
											break
										}
										{ $_ -in @('tinyint', 'smallint', 'int', 'bigint') } {
											$values += $val.ToString()
											break
										}
										default {
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
						}
						catch
						{
							Write_Log "Error generating SQL for table '$table' (file fallback): $_" "red"
							continue
						}
					}
					
					# Copy file to the lane's folder
					try
					{
						$destinationPath = Join-Path $LaneLocalPath $sqlFileName
						Copy-Item -Path $localTempPath -Destination $destinationPath -Force -ErrorAction Stop
						$fileItem = Get-Item $destinationPath
						if ($fileItem.Attributes -band [System.IO.FileAttributes]::Archive)
						{
							$fileItem.Attributes -= [System.IO.FileAttributes]::Archive
						}
						Write_Log "Copied $sqlFileName to Lane #$laneNum ($machineName) (file fallback)." "green"
						$fileCopyLanes += $laneNum
						$ProcessedLanes += $laneNum
					}
					catch
					{
						Write_Log "Error copying $sqlFileName to Lane #[$laneNum] ($machineName): $_" "red"
					}
				}
			}
			else
			{
				Write_Log "Lane #$laneNum ($machineName) not found at path: $LaneLocalPath (file fallback failed)" "yellow"
			}
		}
	}
	
	# ------------------------------------------------------------------------------------------------
	# STEP 6: Clean up and final logging
	# ------------------------------------------------------------------------------------------------
	$srcSqlConnection.Close()
	
	$uniqueProcessedLanes = $ProcessedLanes | Select-Object -Unique
	Write_Log "`r`nTotal Lanes processed: $($uniqueProcessedLanes.Count)" "green"
	if ($protocolLanes.Count -gt 0)
	{
		Write_Log "Lanes processed via SQL protocol: $((($protocolLanes | Select-Object -Unique) -join ', '))" "green"
	}
	if ($fileCopyLanes.Count -gt 0)
	{
		Write_Log "Lanes processed via FILE fallback: $((($fileCopyLanes | Select-Object -Unique) -join ', '))" "yellow"
	}
	Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
}

# ===================================================================================================
#                                   FUNCTION: Close_Open_Transactions
# ---------------------------------------------------------------------------------------------------
# Description:
#   Monitors the XE folder for error files, extracts data, and closes open transactions on specified
#   lanes for a given store. Uses the correct SQL module and protocol. Falls back to file if needed.
#   All lookups use the current FunctionResults mappings. Logging is handled via Write_Log.
# ===================================================================================================

function Close_Open_Transactions
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	# ==================== Start and validate XE path ====================
	Write_Log "`r`n==================== Starting Close_Open_Transactions ====================`r`n" "blue"
	
	$XEFolderPath = "$OfficePath\XE${StoreNumber}901"
	if (-not (Test-Path $XEFolderPath))
	{
		Write_Log "XE folder not found: $XEFolderPath" "red"
		return
	}
	
	# ==================== Validate SQL module availability ====================
	$SqlModule = $script:FunctionResults['SqlModuleName']
	if (-not $SqlModule -or $SqlModule -eq "None")
	{
		Write_Log "No SQL Server module available (SqlServer or SQLPS). Cannot close transactions." "red"
		return
	}
	
	# SQL string for closing all open transactions (manual fallback)
	$CloseTransactionManual = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE')"
	
	# Paths for logs (keep a persistent transaction log)
	$LogFolderPath = "$BasePath\Scripts_by_Alex_C.T"
	$LogFilePath = Join-Path -Path $LogFolderPath -ChildPath "Closed_Transactions_LOG.txt"
	if (-not (Test-Path $LogFolderPath))
	{
		try { New-Item -Path $LogFolderPath -ItemType Directory -Force | Out-Null }
		catch { Write_Log "Failed to create log directory '$LogFolderPath'. Error: $_" "red"; return }
	}
	
	$MatchedTransactions = $false
	
	try
	{
		$currentTime = Get-Date
		# Only process XE files less than 30 days old and starting with S (SMSStart)
		$files = Get-ChildItem -Path $XEFolderPath -Filter "S*.???" | Where-Object {
			($currentTime - $_.LastWriteTime).TotalDays -le 30
		}
		
		# ==================== PROCESS EACH XE ERROR FILE FOUND ====================
		if ($files -and $files.Count -gt 0)
		{
			foreach ($file in $files)
			{
				try
				{
					# Attempt to extract the lane number from the filename
					if ($file.Name -match '^S.*\.(\d{3})$') { $LaneNumber = $Matches[1] }
					else { continue }
					
					# Read contents for parsing
					$content = Get-Content -Path $file.FullName
					$fromLine = $content | Where-Object { $_ -like 'From:*' }
					$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
					$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
					$lastRecordedStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
					
					# Parse out store/lane numbers from the "From:" line
					if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
					{
						$fileStoreNumber = $Matches[1]
						$fileLaneNumber = $Matches[2]
						if ($fileStoreNumber -eq $StoreNumber -and $fileLaneNumber -eq $LaneNumber)
						{
							# Get subject/MSG and check if it matches open transaction/health check
							if ($subjectLine -match 'Subject:\s*(.*)')
							{
								$subject = $Matches[1].Trim()
								if ($subject -eq 'Health' -and $msgLine -match 'MSG:\s*(.*)' -and $Matches[1].Trim() -eq 'This application is not running.')
								{
									# Get transaction number from status line
									if ($lastRecordedStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
									{
										$transactionNumber = $Matches[1]
										$laneKey = $LaneNumber.PadLeft(3, '0')
										$protocolType = if ($script:LaneProtocols.ContainsKey($laneKey)) { $script:LaneProtocols[$laneKey] }
										else { "File" }
										
										# Get full DB connection info for this lane
										$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
										$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
										$tcpConnStr = $laneInfo['TcpConnStr']
										$tcpServer = $laneInfo['TcpServer']
										$dbName = $laneInfo['DBName']
										$closeSQL = "UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1032 = $transactionNumber"
										$protocolWorked = $false
										
										# Try protocol (Named Pipes/TCP), otherwise fallback to file
										Import-Module $SqlModule -ErrorAction Stop
										$supportsConnStr = (Get-Command Invoke-Sqlcmd).Parameters.ContainsKey('ConnectionString')
										
										# ----------- Try Named Pipes protocol first -----------
										if ($protocolType -eq "Named Pipes")
										{
											try
											{
												if ($supportsConnStr)
												{
													Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												else
												{
													Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												$protocolWorked = $true
												Write_Log "Closed transaction $transactionNumber via SQL protocol (Named Pipes) on lane $LaneNumber." "green"
											}
											catch { Write_Log "Named Pipes failed for lane $LaneNumber." "yellow" }
										}
										# ----------- Try TCP protocol -----------
										elseif ($protocolType -eq "TCP")
										{
											try
											{
												if ($supportsConnStr)
												{
													Invoke-Sqlcmd -ConnectionString $tcpConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												else
												{
													Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												$protocolWorked = $true
												Write_Log "Closed transaction $transactionNumber via SQL protocol (TCP) on lane $LaneNumber." "green"
											}
											catch { Write_Log "TCP failed for lane $LaneNumber." "yellow" }
										}
										
										# ----------- Fallback: File-based SQI if protocol failed -----------
										if (-not $protocolWorked)
										{
											$CloseTransactionAuto = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1032 = $transactionNumber)"
											$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
											if (Test-Path $LaneDirectory)
											{
												$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
												Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionAuto -Encoding ASCII
												Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
												Write_Log "Wrote Close_Transaction.sqi file for lane $LaneNumber (file fallback)." "yellow"
											}
											else
											{
												Write_Log "Lane directory $LaneDirectory not found (file fallback)." "yellow"
											}
										}
										
										# Log the close action
										$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Closed transaction $transactionNumber on lane $LaneNumber"
										Add-Content -Path $LogFilePath -Value $logMessage
										
										# Remove processed XE file
										Remove-Item -Path $file.FullName -Force
										
										Write_Log "Processed file $($file.Name) for lane $LaneNumber and closed transaction $transactionNumber" "green"
										$MatchedTransactions = $true
										
										# ----------- Optionally restart all programs on the lane -----------
										Start-Sleep -Seconds 3
										$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
										if ($nodes)
										{
											$machineName = $nodes.LaneNumToMachineName[$LaneNumber]
											if ($machineName)
											{
												$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
												$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
												$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
												if ($result) { Write_Log "Restart command sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after deployment." "green" }
												else { Write_Log "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red" }
											}
										}
									}
									else
									{
										Write_Log "Could not extract transaction number from Last recorded status in file $($file.Name)" "red"
									}
								}
							}
						}
					}
				}
				catch
				{
					Write_Log "Error processing file $($file.Name): $_" "red"
				}
			}
		}
		
		# ==================== If no matching files: Manual prompt for lane ====================
		if (-not $MatchedTransactions)
		{
			Write_Log "No files or no matching transactions found. Prompting for lane number." "yellow"
			$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
			if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
			{
				Write_Log "Lane selection cancelled or returned no selection." "yellow"
				Write_Log "`r`n==================== Close_Open_Transactions Function Completed ====================" "blue"
				return
			}
			
			foreach ($LaneNumber in $selection.Lanes)
			{
				$LaneNumber = if ($LaneNumber -is [pscustomobject] -and $LaneNumber.LaneNumber) { $LaneNumber.LaneNumber }
				else { $LaneNumber }
				$LaneNumber = $LaneNumber.PadLeft(3, '0')
				$laneKey = $LaneNumber
				$protocolType = if ($script:LaneProtocols.ContainsKey($laneKey)) { $script:LaneProtocols[$laneKey] }
				else { "File" }
				$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
				$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
				$tcpConnStr = $laneInfo['TcpConnStr']
				$tcpServer = $laneInfo['TcpServer']
				$dbName = $laneInfo['DBName']
				$closeSQL = "UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE'"
				$protocolWorked = $false
				
				Import-Module $SqlModule -ErrorAction Stop
				$supportsConnStr = (Get-Command Invoke-Sqlcmd).Parameters.ContainsKey('ConnectionString')
				
				if ($protocolType -eq "Named Pipes")
				{
					try
					{
						if ($supportsConnStr)
						{
							Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						else
						{
							Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						$protocolWorked = $true
						Write_Log "Closed all open transactions via SQL protocol (Named Pipes) on lane $LaneNumber." "green"
					}
					catch { Write_Log "Named Pipes failed for lane $LaneNumber." "yellow" }
				}
				elseif ($protocolType -eq "TCP")
				{
					try
					{
						if ($supportsConnStr)
						{
							Invoke-Sqlcmd -ConnectionString $tcpConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						else
						{
							Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						$protocolWorked = $true
						Write_Log "Closed all open transactions via SQL protocol (TCP) on lane $LaneNumber." "green"
					}
					catch { Write_Log "TCP failed for lane $LaneNumber." "yellow" }
				}
				
				if (-not $protocolWorked)
				{
					$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
					if (Test-Path $LaneDirectory)
					{
						$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
						Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionManual -Encoding ASCII
						Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
						Write_Log "Deployed Close_Transaction.sqi to lane $LaneNumber (fallback)." "yellow"
					}
					else
					{
						Write_Log "Lane directory $LaneDirectory not found" "yellow"
					}
				}
				$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - User deployed Close_Transaction to lane $LaneNumber"
				Add-Content -Path $LogFilePath -Value $logMessage
				# Clean up XE folder, except FATALs
				Get-ChildItem -Path $XEFolderPath -File | Where-Object { $_.Name -notlike "*FATAL*" } | Remove-Item -Force
				$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
				if ($nodes)
				{
					$machineName = $nodes.LaneNumToMachineName[$LaneNumber]
					if ($machineName)
					{
						$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
						$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
						$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
						if ($result) { Write_Log "Restart All Programs sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after user deployment." "green" }
						else { Write_Log "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red" }
					}
				}
				Write_Log "Prompt deployment process completed." "yellow"
			}
		}
	}
	catch
	{
		Write_Log "An error occurred during monitoring: $_" "red"
	}
	Write_Log "No further matching files were found after processing." "yellow"
	Write_Log "`r`n==================== Close_Open_Transactions Function Completed ====================" "blue"
}

# ===================================================================================================
#                                  FUNCTION: Ping_All_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#    Ping all nodes of a given type (Lanes, Scales, or Backoffices) for a store.
#    Usage:
#        Ping_All_Nodes -NodeType "Lane"       -StoreNumber "001"
#        Ping_All_Nodes -NodeType "Scale"
#        Ping_All_Nodes -NodeType "Backoffice" -StoreNumber "001"
#    All context (Lane/Scale/Backoffice lists) are sourced from $script:FunctionResults.
# ===================================================================================================

function Ping_All_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string]$NodeType,
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Ping_All_Nodes ($NodeType) Function ====================`r`n" "blue"
	
	$nodesToPing = @()
	$nodeLabel = ""
	$nodeSummary = ""
	
	switch ($NodeType)
	{
		"Lane" {
			$nodeLabel = "Lane"
			$nodeSummary = "Lanes"
			if (-not $script:FunctionResults.ContainsKey('LaneNumToMachineName'))
			{
				Write_Log "Lane mappings not available. Please run Retrieve_Nodes first." "red"
				return
			}
			$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
			foreach ($laneNum in $LaneNumToMachineName.Keys | Where-Object { $_ -match '^\d{3}$' })
			{
				$machineName = $LaneNumToMachineName[$laneNum]
				if ($machineName -and $machineName -notin @("Unknown", "Not Found"))
				{
					$nodesToPing += [PSCustomObject]@{
						Key    = $laneNum
						Target = $machineName
						Label  = "" # Do not use $null or $null string
					}
				}
			}
			$nodesToPing = $nodesToPing | Sort-Object { [int]$_.Key }
		}
		
		"Scale" {
			$nodeLabel = "Scale"
			$nodeSummary = "Scales"
			if (-not $script:FunctionResults.ContainsKey('ScaleCodeToIPInfo') -or $script:FunctionResults['ScaleCodeToIPInfo'].Count -eq 0)
			{
				Write_Log "No scales found to ping." "yellow"
				return
			}
			$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
			foreach ($code in $ScaleCodeToIPInfo.Keys | Where-Object { $_ -match '^\d{1,3}$' })
			{
				$scaleObj = $ScaleCodeToIPInfo[$code]
				$ip = $null
				if ($scaleObj.PSObject.Properties['FullIP'])
				{
					$ip = $scaleObj.FullIP
				}
				elseif ($scaleObj.PSObject.Properties['IPNetwork'] -and $scaleObj.PSObject.Properties['IPDevice'])
				{
					$ip = "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)"
				}
				if ($ip -and $ip -notin @("Unknown", "Not Found", ""))
				{
					$nodesToPing += [PSCustomObject]@{
						Key    = $code
						Target = $ip
						Label  = $scaleObj.ScaleName
					}
				}
			}
			$nodesToPing = $nodesToPing | Sort-Object { [int]$_.Key }
			# Remove duplicate IPs
			$uniqueTargets = @{ }
			$finalNodesToPing = @()
			foreach ($node in $nodesToPing)
			{
				if (-not $uniqueTargets.ContainsKey($node.Target))
				{
					$uniqueTargets[$node.Target] = $true
					$finalNodesToPing += $node
				}
			}
			$nodesToPing = $finalNodesToPing
		}
		
		"Backoffice" {
			$nodeLabel = "Backoffice"
			$nodeSummary = "Backoffices"
			if (-not $script:FunctionResults.ContainsKey('BackofficeNumToMachineName'))
			{
				Write_Log "Backoffice information is not available. Please run Retrieve_Nodes first." "red"
				return
			}
			$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
			foreach ($boNum in $BackofficeNumToMachineName.Keys | Where-Object { $_ -match '^\d{3}$' })
			{
				$machineName = $BackofficeNumToMachineName[$boNum]
				if ($machineName -and $machineName -notin @("Unknown", "Not Found"))
				{
					$nodesToPing += [PSCustomObject]@{
						Key    = $boNum
						Target = $machineName
						Label  = "" # Do not use $null or $null string
					}
				}
			}
			$nodesToPing = $nodesToPing | Sort-Object { [int]$_.Key }
		}
	}
	
	# Final deduplication by Target (just to be sure)
	$uniqueTargets = @{ }
	$finalList = @()
	foreach ($node in $nodesToPing)
	{
		if (-not $uniqueTargets.ContainsKey($node.Target))
		{
			$uniqueTargets[$node.Target] = $node
			$finalList += $node
		}
	}
	$nodesToPing = $finalList
	
	if ($nodesToPing.Count -eq 0)
	{
		Write_Log "No valid $nodeSummary found to ping." "yellow"
		return
	}
	
	Write_Log "All $nodeSummary will be pinged." "green"
	
	$successCount = 0
	$failureCount = 0
	
	foreach ($node in $nodesToPing)
	{
		$primary = $node.Key
		$target = $node.Target
		
		# Compose labelInfo only for scales if present and non-empty
		$labelInfo = ""
		if ($NodeType -eq "Scale" -and $node.PSObject.Properties.Name -contains "Label" -and $node.Label -and $node.Label -ne "")
		{
			$labelInfo = " [($($node.Label))]"
		}
		
		if ([string]::IsNullOrWhiteSpace($target) -or $target -in @("Unknown", "Not Found"))
		{
			# For lanes/backoffices, $labelInfo is always ""
			Write_Log "$nodeLabel #[$primary]: Target '$target'. Status: Skipped." "yellow"
			continue
		}
		try
		{
			$pingResult = Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction Stop
			if ($pingResult)
			{
				Write_Log ("$nodeLabel #[$primary]${labelInfo}: Target '$target' is reachable. Status: Success.") "green"
				$successCount++
			}
			else
			{
				Write_Log ("$nodeLabel #[$primary]${labelInfo}: Target '$target' is not reachable. Status: Failed.") "red"
				$failureCount++
			}
		}
		catch
		{
			Write_Log ("$nodeLabel #[$primary]${labelInfo}: Failed to ping target '$target'. Error: $($_.Exception.Message)") "red"
			$failureCount++
		}
	}
	
	$summaryMsg = "Ping Summary for $nodeSummary"
	if ($StoreNumber) { $summaryMsg += " (Store Number: $StoreNumber)" }
	$summaryMsg += " - Success: $successCount, Failed: $failureCount."
	Write_Log $summaryMsg "blue"
	Write_Log "`r`n==================== Ping_All_Nodes ($NodeType) Function Completed ====================" "blue"
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
	if (-not $script:FunctionResults.ContainsKey('LaneMachineNames') -or
		-not $script:FunctionResults.ContainsKey('LaneNumToMachineName'))
	{
		Write_Log "Lane information is not available. Please run Retrieve_Nodes first." "Red"
		return
	}
	
	# Retrieve lane information
	$LaneMachineNames = $script:FunctionResults['LaneMachineNames']
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	
	if ($LaneMachineNames.Count -eq 0)
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
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	
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
		if ($LaneNumToMachineName.ContainsKey($lane))
		{
			$machineName = $LaneNumToMachineName[$lane]
			
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
#     4. **Configures Visual Settings**:
#        - Enables "Show thumbnails instead of icons" in Explorer.
#        - Enables "Smooth edges of screen fonts" (font smoothing with ClearType).
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
		[string]$UnorganizedFolderName = "Unorganized Items"
	)
	
	Write_Log "`r`n==================== Starting Configure_System_Settings Function ====================`r`n" "blue"
	
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
		$excludedFolders = @("Lanes", "Scales", "BackOffices", "Unorganized Items")
		
		# Create Unorganized Items folder if it doesn't exist
		$folderPath = Join-Path -Path $DesktopPath -ChildPath "Unorganized Items"
		if (-not (Test-Path -Path $folderPath))
		{
			New-Item -Path $folderPath -ItemType Directory | Out-Null
			Write_Log "Created folder: $folderPath" "green"
		}
		else
		{
			Write_Log "Folder already exists: $folderPath" "Cyan"
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
		
		# ===========================================
		# 4. Configure Visual Settings
		# ===========================================
		Write_Log "`r`nConfiguring visual settings..." "Blue"
		
		try
		{
			# Enable "Show thumbnails instead of icons" (disable "Always show icons, never thumbnails")
			Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "IconsOnly" -Value 0 -Type DWord -ErrorAction Stop
			Write_Log "Enabled 'Show thumbnails instead of icons'." "Green"
		}
		catch
		{
			Write_Log "Failed to enable thumbnails. Error: $_" "Red"
		}
		
		try
		{
			# Enable "Smooth edges of screen fonts" (font smoothing with ClearType)
			Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothing" -Value "2" -Type String -ErrorAction Stop
			Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothingType" -Value 2 -Type DWord -ErrorAction Stop
			Write_Log "Enabled 'Smooth edges of screen fonts'." "Green"
		}
		catch
		{
			Write_Log "Failed to enable font smoothing. Error: $_" "Red"
		}
		
		Write_Log "Visual settings configuration complete." "Green"
		
		Write_Log "Restarting Explorer to apply changes..." "Yellow"
		Stop-Process -Name explorer -Force
		Write_Log "Explorer restarted." "Green"
		
		Write_Log "All system configurations have been applied successfully." "Green"
		Write_Log "`r`n==================== Configure_System_Settings Function Completed ====================`r`n" "blue"
	}
	catch
	{
		Write_Log "An unexpected error occurred: $_" "Red"
	}
}

# ===================================================================================================
#                            FUNCTION: Refresh_PIN_Pad_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Refreshes critical EMV PIN pad configuration files (.ini) on selected lanes for a specified store.
#   It "touches" (updates the LastWriteTime) on each target file, ensuring the system sees them as
#   modified and refreshes them accordingly.
#   Mappings are sourced from LaneNumToMachineName (from Retrieve_Nodes).
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: (Mandatory) The store number in 3-digit string format (e.g., "003")
# Requirements:
#   - Retrieve_Nodes must have been executed to populate FunctionResults.
#   - LaneNumToMachineName mapping must be present in FunctionResults.
#   - Write_Log and Show_Node_Selection_Form functions must be available.
# ===================================================================================================

function Refresh_PIN_Pad_Files
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Refresh_PIN_Pad_Files Function ====================`r`n" "blue"
	
	# ============================= Validate Environment =============================
	# Validate OfficePath base directory
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Ensure required functions are loaded
	foreach ($func in @('Show_Node_Selection_Form', 'Write_Log', 'Retrieve_Nodes'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Ensure lane mapping is available from Retrieve_Nodes
	if (-not ($script:FunctionResults.ContainsKey('LaneNumToMachineName')))
	{
		Write_Log "No lane information found. Please ensure Retrieve_Nodes has been executed." "red"
		return
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	
	# ============================= Lane Selection UI =============================
	# Use the node selector form to allow the user to choose which lanes to refresh
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($null -eq $selection)
	{
		Write_Log "Operation canceled by user." "yellow"
		return
	}
	$Lanes = $selection.Lanes
	if (-not $Lanes -or $Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for processing." "yellow"
		return
	}
	
	# ============================= File Types =============================
	# List of PIN pad config files that should be "touched" (refreshed)
	$fileExtensions = @("PreferredAIDs.ini", "EMVCAPKey.ini", "EMVAID.ini")
	Write_Log "Starting refresh of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "green"
	
	# ============================= Main Processing Loop =============================
	$totalRefreshed = 0
	$totalFailed = 0
	
	foreach ($laneObj in $Lanes)
	{
		# Accept both plain lane number string or PSCustomObject with .LaneNumber property
		$laneNum = if ($laneObj -is [pscustomobject] -and $laneObj.LaneNumber) { $laneObj.LaneNumber }
		else { $laneObj }
		$laneNum = $laneNum.PadLeft(3, '0')
		
		# Ensure this lane exists in LaneNumToMachineName
		if ($LaneNumToMachineName.ContainsKey($laneNum))
		{
			$machineName = $LaneNumToMachineName[$laneNum]
			
			if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
			{
				Write_Log "Lane #${laneNum}: Machine name is invalid or unknown. Skipping refresh." "yellow"
				continue
			}
			
			# Build UNC path to lane's EMV config directory
			$targetPath = "\\$machineName\Storeman\XchDev\EMVConfig\"
			
			if (-not (Test-Path -Path $targetPath))
			{
				Write_Log "Lane #${laneNum}: Target path '$targetPath' does not exist. Skipping." "yellow"
				continue
			}
			
			Write_Log "Processing Lane #$laneNum ($machineName) at '$targetPath'." "blue"
			
			foreach ($file in $fileExtensions)
			{
				$filePath = Join-Path -Path $targetPath -ChildPath $file
				if (Test-Path -Path $filePath)
				{
					try
					{
						# "Touch" the file (set LastWriteTime to current time)
						(Get-Item -Path $filePath).LastWriteTime = Get-Date
						Write_Log "Lane #${laneNum}: Refreshed file '$filePath'." "green"
						$totalRefreshed++
					}
					catch
					{
						Write_Log "Lane #${laneNum}: Failed to refresh file '$filePath'. Error: $_" "red"
						$totalFailed++
					}
				}
				else
				{
					Write_Log "Lane #${laneNum}: File '$filePath' does not exist. Skipping." "yellow"
				}
			}
		}
		else
		{
			Write_Log "Lane #${laneNum}: Machine information not found in LaneNumToMachineName. Skipping." "yellow"
			continue
		}
	}
	
	Write_Log "Refresh Summary for Store Number: $StoreNumber - Total Files Refreshed: $totalRefreshed, Total Failures: $totalFailed." "green"
	Write_Log "`r`n==================== Refresh_PIN_Pad_Files Function Completed ====================" "blue"
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
	
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes)
	{
		Write_Log "No lanes selected. Cancelling operation." "yellow"
		return
	}
	# --- ENSURE NODE MAPPING IS PRESENT ---
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName)
	{
		$null = Retrieve_Nodes -StoreNumber $StoreNumber
		$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
		if (-not $LaneNumToMachineName)
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
		$LaneMachineName = $LaneNumToMachineName[$LaneNumber]
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
#   BIZERBA and ISHIDA records. 
#   - Sets BufferTime to 1 for first BIZERBA, 5 for others.
#   - Updates ScaleName for BIZERBA with IPDevice.
#   - Reassigns ScaleCode to ensure BIZERBA are first, then ISHIDA.
#   - Updates ScaleName and BufferTime for ISHIDA WMAI records.
#   Displays the organized table at the end.
# ===================================================================================================

function Organize_TBS_SCL_ver520
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Organize_TBS_SCL_ver520 Function ====================`r`n" "blue"
	
	# Get connection string and module
	$connectionString = $script:FunctionResults['ConnectionString']
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	if (-not $connectionString)
	{
		Write_Log "Connection string not found in `$script:FunctionResults['ConnectionString']`." "red"
		return
	}
	if (-not $SqlModuleName -or $SqlModuleName -eq "None")
	{
		Write_Log "No valid SQL module detected for SQL operations." "red"
		return
	}
	Import-Module $SqlModuleName -ErrorAction Stop
	$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
	$supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString'
	
	# Define the SQL commands
	$updateQueries = @"
-------------------------------------------------------------------------------
-- 1) Update ISHIDA WMAI ScaleName and BufferTime based on the record count
-------------------------------------------------------------------------------
DECLARE @IshidaWMAICount INT;
SELECT @IshidaWMAICount = COUNT(*) FROM [TBS_SCL_ver520] WHERE ScaleBrand = 'ISHIDA' AND ScaleModel = 'WMAI';
IF @IshidaWMAICount > 1
BEGIN
    UPDATE [TBS_SCL_ver520] SET ScaleName = CONCAT('Ishida Wrapper ', IPDevice), BufferTime = '1' WHERE ScaleBrand = 'ISHIDA' AND ScaleModel = 'WMAI';
END
ELSE
BEGIN
    UPDATE [TBS_SCL_ver520] SET ScaleName = 'Ishida Wrapper', BufferTime = '1' WHERE ScaleBrand = 'ISHIDA' AND ScaleModel = 'WMAI';
END;

-------------------------------------------------------------------------------
-- 2) Update BIZERBA ScaleName
-------------------------------------------------------------------------------
UPDATE [TBS_SCL_ver520] SET ScaleName = CONCAT('Scale ', IPDevice) WHERE ScaleBrand = 'BIZERBA';

-------------------------------------------------------------------------------
-- 3) Update ScaleCode for BIZERBA, starting at 10, in IPDevice ascending order
-------------------------------------------------------------------------------
WITH BIZERBA_CTE AS (
    SELECT ScaleCode, IPDevice, rn = ROW_NUMBER() OVER (ORDER BY TRY_CAST(IPDevice AS INT)) 
    FROM [TBS_SCL_ver520] WHERE ScaleBrand = 'BIZERBA'
)
UPDATE T
SET T.ScaleCode = 10 + B.rn - 1
FROM [TBS_SCL_ver520] AS T
JOIN BIZERBA_CTE AS B ON T.ScaleCode = B.ScaleCode
WHERE T.ScaleBrand = 'BIZERBA';

-------------------------------------------------------------------------------
-- 4) Update ScaleCode for ISHIDA, starting after the new max BIZERBA ScaleCode.
-------------------------------------------------------------------------------
;WITH MaxBizerba AS (
    SELECT MAX(ScaleCode) AS MaxCode FROM [TBS_SCL_ver520] WHERE ScaleBrand = 'BIZERBA'
),
ISHIDA_CTE AS (
    SELECT ScaleCode, IPDevice, rn = ROW_NUMBER() OVER (ORDER BY TRY_CAST(IPDevice AS INT))
    FROM [TBS_SCL_ver520] WHERE ScaleBrand = 'ISHIDA'
)
UPDATE T
SET T.ScaleCode = (SELECT MaxCode FROM MaxBizerba) + 10 + I.rn - 1
FROM [TBS_SCL_ver520] AS T
JOIN ISHIDA_CTE AS I ON T.ScaleCode = I.ScaleCode
WHERE T.ScaleBrand = 'ISHIDA';

-------------------------------------------------------------------------------
-- 5) Now set BufferTime for BIZERBA records:
-------------------------------------------------------------------------------
WITH BIZ_ORDER AS (
    SELECT ScaleCode, RN = ROW_NUMBER() OVER (ORDER BY ScaleCode ASC)
    FROM [TBS_SCL_ver520] WHERE ScaleBrand = 'BIZERBA'
)
UPDATE T
SET T.BufferTime = CASE WHEN B.RN = 1 THEN '1' ELSE '5' END
FROM [TBS_SCL_ver520] T
INNER JOIN BIZ_ORDER B ON T.ScaleCode = B.ScaleCode
WHERE T.ScaleBrand = 'BIZERBA';
"@
	
	$selectQuery = @"
SELECT 
    ScaleCode, ScaleName, ScaleLocation, IPNetwork, IPDevice, Active,
    SystemLocalTime, AutoStart, AutoTransmit, BufferTime, ScaleBrand, ScaleModel
FROM [TBS_SCL_ver520]
ORDER BY ScaleCode ASC;
"@
	
	# -------------------------------------------------------------------------
	# Execute update and select queries using detected module and connection string
	# -------------------------------------------------------------------------
	try
	{
		Write_Log "Executing update queries to modify ScaleName, BufferTime, and ScaleCode..." "blue"
		if ($supportsConnectionString)
		{
			Invoke-Sqlcmd -ConnectionString $connectionString -Query $updateQueries -ErrorAction Stop
		}
		else
		{
			$server = $script:FunctionResults['DBSERVER']
			$database = $script:FunctionResults['DBNAME']
			if (-not $server -or -not $database)
			{
				Write_Log "Invalid ConnectionString. Missing Server or Database information." "red"
				throw "Invalid ConnectionString. Cannot parse Server or Database."
			}
			Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $updateQueries -ErrorAction Stop
		}
		Write_Log "Update queries executed successfully." "green"
	}
	catch
	{
		Write_Log "An error occurred while executing update queries: $_" "red"
		return
	}
	
	# -------------------------------------------------------------------------
	# Retrieve and display organized data
	# -------------------------------------------------------------------------
	try
	{
		Write_Log "Retrieving organized data..." "blue"
		if ($supportsConnectionString)
		{
			$data = Invoke-Sqlcmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
		}
		else
		{
			$server = $script:FunctionResults['DBSERVER']
			$database = $script:FunctionResults['DBNAME']
			if (-not $server -or -not $database)
			{
				Write_Log "Invalid ConnectionString. Missing Server or Database information." "red"
				throw "Invalid ConnectionString. Cannot parse Server or Database."
			}
			$data = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
		}
		Write_Log "Data retrieval successful." "green"
	}
	catch
	{
		Write_Log "An error occurred while retrieving data: $_" "red"
		return
	}
	
	# -------------------------------------------------------------------------
	# Display organized data
	# -------------------------------------------------------------------------
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
		[array]$LaneNumbers,
		[Parameter(Mandatory = $false)]
		[switch]$Silent
	)
	
	if (-not $Silent)
	{
		Write_Log "`r`n==================== Starting Send_Restart_All_Programs Function ====================`r`n" "blue"
	}
	
	# Retrieve lane mappings using Retrieve_Nodes (guaranteed current)
	$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
	if (-not $nodes -or -not $nodes.LaneNumToMachineName)
	{
		Write_Log "Failed to retrieve node information for store $StoreNumber." "red"
		return
	}
	
	# Lane selection: use passed-in list or picker
	$lanes =
	if ($LaneNumbers -and $LaneNumbers.Count -gt 0)
	{
		$LaneNumbers | ForEach-Object {
			if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
			elseif ($_ -is [int]) { "{0:D3}" -f $_ }
			elseif ($_.Length -lt 3) { $_.PadLeft(3, '0') }
			else { $_ }
		}
	}
	else
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
		if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
		{
			Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
			return
		}
		$selection.Lanes | ForEach-Object {
			if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
			elseif ($_ -is [int]) { "{0:D3}" -f $_ }
			elseif ($_.Length -lt 3) { $_.PadLeft(3, '0') }
			else { $_ }
		}
	}
	
	foreach ($lane in $lanes)
	{
		$machineName = $nodes.LaneNumToMachineName[$lane]
		if (-not $machineName)
		{
			Write_Log "No machine found for lane $lane. Skipping." "yellow"
			continue
		}
		
		$mailslotAddress = "\\$machineName\Mailslot\SMSStart_${StoreNumber}${lane}"
		$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
		
		try
		{
			$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
			if ($result)
			{
				if (-not $Silent)
				{
					Write_Log "Command sent successfully to Machine $machineName (Store $StoreNumber, Lane $lane)." "green"
				}
			}
			else
			{
				if (-not $Silent)
				{
					Write_Log "Failed to send command to Machine $machineName (Store $StoreNumber, Lane $lane)." "red"
				}
			}
		}
		catch
		{
			if (-not $Silent)
			{
				Write_Log "Exception sending to $machineName (Lane $lane): $_" "red"
			}
		}
	}
	
	if (-not $Silent)
	{
		Write_Log "`r`n==================== Send_Restart_All_Programs Function Completed ====================" "blue"
	}
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
		[string]$StoreNumber,
		[switch]$Schedule # Optional: enables scheduling mode (prompt for interval)
	)
	
	Write_Log "`r`n==================== Starting Time Sync ====================`r`n" "blue"
	
	# 1. Unified Lane Picker
	$laneSelection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $laneSelection -or -not $laneSelection.Lanes -or $laneSelection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for time sync." "yellow"
		Write_Log "`r`n==================== Time Sync Aborted ====================`r`n" "blue"
		return
	}
	$selectedLanes = $laneSelection.Lanes
	
	# 2. Scheduling Mode Prompt (Optional)
	$isScheduling = $false
	$intervalMinutes = 0
	if ($Schedule)
	{
		Add-Type -AssemblyName System.Windows.Forms
		$formInterval = New-Object System.Windows.Forms.Form
		$formInterval.Text = "Set Sync Interval"
		$formInterval.Size = New-Object System.Drawing.Size(300, 150)
		$formInterval.StartPosition = "CenterScreen"
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "Enter interval in minutes:"
		$label.Location = New-Object System.Drawing.Point(20, 20)
		$formInterval.Controls.Add($label)
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(20, 50)
		$textBox.Size = New-Object System.Drawing.Size(240, 20)
		$formInterval.Controls.Add($textBox)
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(100, 80)
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$formInterval.Controls.Add($okButton)
		$result = $formInterval.ShowDialog()
		if ($result -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "Scheduling canceled." "yellow"
			return
		}
		if (-not [int]::TryParse($textBox.Text, [ref]$intervalMinutes) -or $intervalMinutes -le 0)
		{
			Write_Log "Invalid interval. Must be a positive integer." "red"
			return
		}
		$isScheduling = $true
	}
	
	# 3. Build Time String
	$now = Get-Date
	$dateStr = $now.ToString("MM/dd/yyyy")
	$timeStr = $now.ToString("HHmmss")
	$wizCommand = "@WIZRPL(DATE=$dateStr)@WIZRPL(TIME=$timeStr)"
	
	foreach ($lane in $selectedLanes)
	{
		# Lane number normalization
		$laneNum = if ($lane -is [pscustomobject] -and $lane.LaneNumber) { $lane.LaneNumber }
		elseif ($lane -is [int]) { "{0:D3}" -f $lane }
		elseif ($lane.Length -lt 3) { $lane.PadLeft(3, '0') }
		else { $lane }
		
		$machine = $script:FunctionResults['LaneNumToMachineName'][$laneNum]
		if (-not $machine)
		{
			Write_Log "No machine name found for lane $laneNum." "red"
			continue
		}
		
		$success = $false
		
		# -------- PRIMARY: REMOTE SCHTASKS --------
		try
		{
			# Use the server's IP for net time
			try
			{
				$serverIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -ErrorAction Stop).IPv4Address.IPAddressToString
			}
			catch
			{
				$serverIP = $env:COMPUTERNAME
			}
			$command = "net time \\$serverIP /set /yes"
			
			$taskName = if ($isScheduling) { "ScheduledTimeSync" }
			else { "SyncTime" }
			$scheduleSc = if ($isScheduling) { "/sc minute" }
			else { "/sc once" }
			$scheduleMo = if ($isScheduling) { "/mo $intervalMinutes" }
			else { "" }
			$scheduleSt = if ($isScheduling) { "" }
			else { "" } # Omit /st for immediate run
			$logMessage = if ($isScheduling) { "Scheduled time sync every $intervalMinutes minutes on lane" }
			else { "Executed time sync on lane" }
			
			$createCmd = "schtasks /create /s $machine /tn $taskName /tr `"$command`" $scheduleSc $scheduleMo $scheduleSt /ru SYSTEM /f /rl HIGHEST"
			$createOutput = Invoke-Expression $createCmd 2>&1
			if ($LASTEXITCODE -eq 0)
			{
				if (-not $isScheduling)
				{
					$runCmd = "schtasks /run /s $machine /tn $taskName"
					$runOutput = Invoke-Expression $runCmd 2>&1
					if ($LASTEXITCODE -eq 0)
					{
						Write_Log "$logMessage $laneNum." "green"
						$success = $true
					}
					else
					{
						Write_Log "Run output for [$machine]: $runOutput" "yellow"
					}
				}
				else
				{
					Write_Log "$logMessage $laneNum." "green"
					$success = $true
				}
				# Clean up (delete task after one-shot run)
				if ($success -and -not $isScheduling)
				{
					Start-Sleep -Seconds 5
					$deleteCmd = "schtasks /delete /s $machine /tn $taskName /f"
					$deleteOutput = Invoke-Expression $deleteCmd 2>&1
					if ($LASTEXITCODE -ne 0)
					{
						Write_Log "Delete output for [$machine]: $deleteOutput" "yellow"
					}
				}
			}
			else
			{
				Write_Log "Create output for [$machine]: $createOutput" "yellow"
			}
		}
		catch
		{
			Write_Log "Exception during schtasks create/run: $_" "yellow"
		}
		
		# -------- SECONDARY: FILE-DROP WITH @EXEC --------
		if (-not $success)
		{
			try
			{
				$TempDir = $env:TEMP
				$xfFolder = Join-Path $OfficePath "XF${StoreNumber}${laneNum}"
				if (-not (Test-Path $xfFolder))
				{
					Write_Log "XF folder for lane $laneNum does not exist: $xfFolder" "red"
					continue
				}
				try
				{
					$serverIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -ErrorAction Stop).IPv4Address.IPAddressToString
				}
				catch
				{
					$serverIP = $env:COMPUTERNAME
				}
				$remoteTempPath = "\\$machine\C$\Windows\Temp\sync_time.bat"
				$batContent = "@echo off`r`nnet time \\$serverIP /set /yes"
				$localBatPath = Join-Path $TempDir "sync_time_$laneNum.bat"
				Set-Content -Path $localBatPath -Value $batContent -Encoding Ascii
				Copy-Item -Path $localBatPath -Destination $remoteTempPath -Force -ErrorAction Stop
				Write_Log "Copied sync_time.bat to temp folder on lane $laneNum." "green"
				Remove-Item -Path $localBatPath -Force
				$execFilePath = Join-Path $xfFolder "exec_time_sync.txt"
				$execContent = "@EXEC(Run='C:\Windows\Temp\sync_time.bat')"
				Set-Content -Path $execFilePath -Value $execContent -Encoding Ascii -ErrorAction Stop
				# Clear the archive bit
				$attr = (Get-Item $execFilePath).Attributes
				Set-ItemProperty -Path $execFilePath -Name Attributes -Value ($attr -band -bnot [System.IO.FileAttributes]::Archive)
				Write_Log "Created @EXEC trigger in XF folder for lane $laneNum and cleared archive bit." "green"
				# Wait for execution (adjust as needed)
				Start-Sleep -Seconds 10
				Remove-Item -Path $remoteTempPath -Force -ErrorAction SilentlyContinue
				$success = $true
			}
			catch
			{
				Write_Log "File copy/@EXEC fallback failed for lane ${laneNum}: $_" "red"
			}
		}
		
		# -------- TERTIARY: MAILSLOT --------
		if (-not $success)
		{
			try
			{
				$mailslot = "\\$machine\mailslot\SMSStart_${StoreNumber}${laneNum}"
				$result = [MailslotSender]::SendMailslotCommand($mailslot, $wizCommand)
				if ($result)
				{
					Write_Log "Time sync sent to $machine (Lane $laneNum) via mailslot (final fallback)." "green"
					$success = $true
				}
				else
				{
					Write_Log "Mailslot failed for $machine (Lane $laneNum) [final fallback]." "yellow"
				}
			}
			catch
			{
				Write_Log "Exception during mailslot send (final fallback): $_" "yellow"
			}
		}
		
		if (-not $success)
		{
			Write_Log "All methods failed for lane $laneNum!" "red"
		}
	}
	
	Write_Log "`r`n==================== Ending Time Sync ====================`r`n" "blue"
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
	
	# Prompt for drawer state (Enable/Disable)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$stateForm = New-Object System.Windows.Forms.Form
	$stateForm.Text = "Select Drawer State"
	$stateForm.Size = New-Object System.Drawing.Size(400, 180)
	$stateForm.StartPosition = "CenterScreen"
	$stateLabel = New-Object System.Windows.Forms.Label
	$stateLabel.Text = "Select Drawer State:"
	$stateLabel.Location = New-Object System.Drawing.Point(10, 20)
	$stateLabel.AutoSize = $true
	$stateForm.Controls.Add($stateLabel)
	$radioEnable = New-Object System.Windows.Forms.RadioButton
	$radioEnable.Text = "Enable"
	$radioEnable.Location = New-Object System.Drawing.Point(10, 50)
	$radioEnable.AutoSize = $true
	$radioEnable.Checked = $true
	$stateForm.Controls.Add($radioEnable)
	$radioDisable = New-Object System.Windows.Forms.RadioButton
	$radioDisable.Text = "Disable"
	$radioDisable.Location = New-Object System.Drawing.Point(10, 80)
	$radioDisable.AutoSize = $true
	$stateForm.Controls.Add($radioDisable)
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(80, 120)
	$okButton.Add_Click({
			$stateForm.Tag = if ($radioEnable.Checked) { "1" }
			else { "0" }
			$stateForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$stateForm.Close()
		})
	$stateForm.Controls.Add($okButton)
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
		Write_Log "User cancelled at drawer state selection." "yellow"
		Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	$DrawerState = $stateForm.Tag
	Write_Log "Drawer state selected: $DrawerState" "green"
	
	# Select lanes with unified picker
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected or selection cancelled." "yellow"
		Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	
	# Node mapping for machine names if needed
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	
	$lanes = $selection.Lanes
	$successCount = 0
	
	foreach ($laneObj in $lanes)
	{
		# Support all output shapes from picker (object/string/int)
		$laneNum = $null; $machineName = $null
		if ($laneObj -is [pscustomobject])
		{
			$laneNum = if ($laneObj.PSObject.Properties.Name -contains "LaneNumber") { $laneObj.LaneNumber }
			else { $laneObj.ToString() }
			$machineName = if ($laneObj.PSObject.Properties.Name -contains "MachineName") { $laneObj.MachineName }
		}
		else { $laneNum = $laneObj.ToString() }
		$laneNum = $laneNum.PadLeft(3, '0')
		if (-not $machineName -and $LaneNumToMachineName.ContainsKey($laneNum)) { $machineName = $LaneNumToMachineName[$laneNum] }
		
		$LaneDirectory = "$OfficePath\XF${StoreNumber}${laneNum}"
		if (-not (Test-Path $LaneDirectory))
		{
			Write_Log "Lane directory $LaneDirectory not found. Skipping lane $laneNum." "yellow"
			continue
		}
		
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
		
		$SQIContent = $SQIContent -replace "`n", "`r`n"
		$SQIFilePath = Join-Path -Path $LaneDirectory -ChildPath "DrawerControl.sqi"
		Set-Content -Path $SQIFilePath -Value $SQIContent -Encoding ASCII
		Set-ItemProperty -Path $SQIFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Deployed DrawerControl.sqi to lane $laneNum ($machineName) with state '$DrawerState' in $LaneDirectory." "green"
		
		# Send Restart All Programs
		try
		{
			if ($machineName)
			{
				$mailslot = "\\$machineName\mailslot\SMSStart_${StoreNumber}${laneNum}"
				$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
				$result = $null
				if ($global:MailslotSender -or ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetTypes() -match 'MailslotSender' }))
				{
					$result = [MailslotSender]::SendMailslotCommand($mailslot, $commandMessage)
				}
				if ($result)
				{
					Write_Log "Restart All Programs sent to $machineName (Lane $laneNum)" "cyan"
				}
				else
				{
					Write_Log "Could not send restart command to $machineName (Lane $laneNum)" "yellow"
				}
			}
		}
		catch
		{
			Write_Log "Error sending restart command to $machineName (Lane $laneNum): $_" "red"
		}
		$successCount++
	}
	
	Write_Log "$successCount lanes processed for Drawer_Control." "blue"
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
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Refresh_Database Function ====================`r`n" "blue"
	
	if (-not $OfficePath)
	{
		Write_Log "OfficePath variable not defined. Aborting." "red"
		return
	}
	if (-not $script:FunctionResults.ContainsKey('LaneNumToMachineName'))
	{
		Write_Log "Lane mapping not available. Please run Retrieve_Nodes first." "red"
		return
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	
	# Lane picker (compatible with objects or string output)
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected. Exiting." "yellow"
		return
	}
	
	$lanes = $selection.Lanes
	$refreshedCount = 0
	
	foreach ($laneObj in $lanes)
	{
		# Support either [pscustomobject] or plain lane number/string
		$laneNum = $null
		$machineName = $null
		
		if ($laneObj -is [pscustomobject])
		{
			if ($laneObj.PSObject.Properties.Name -contains "LaneNumber") { $laneNum = $laneObj.LaneNumber }
			elseif ($laneObj.PSObject.Properties.Name -contains "Key") { $laneNum = $laneObj.Key }
			else { $laneNum = $laneObj.ToString() }
			if ($laneObj.PSObject.Properties.Name -contains "MachineName") { $machineName = $laneObj.MachineName }
		}
		else
		{
			$laneNum = $laneObj.ToString()
		}
		
		# Always 3-digit padded
		$laneNum = $laneNum.PadLeft(3, '0')
		if (-not $machineName -and $LaneNumToMachineName.ContainsKey($laneNum))
		{
			$machineName = $LaneNumToMachineName[$laneNum]
		}
		if (-not $machineName)
		{
			Write_Log "No machine mapping for lane $laneNum. Skipping." "yellow"
			continue
		}
		
		$regDir = Join-Path $OfficePath "XF${StoreNumber}${laneNum}"
		if (-not (Test-Path $regDir))
		{
			Write_Log "Directory not found for lane ${laneNum}: $regDir. Skipping." "yellow"
			continue
		}
		
		$filePath = Join-Path $regDir "db_refresh.sqi"
		$content = "@WIZSET(TARGET=@TER);`r`n@EXEC(SQM=exe_activate_accept_sys);`r`n"
		if (-not $script:ansiPcEncoding) { $script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252) }
		[System.IO.File]::WriteAllText($filePath, $content, $script:ansiPcEncoding)
		Write_Log "Deployed db_refresh.sqi to $regDir (Lane $laneNum, $machineName)" "green"
		$refreshedCount++
		
		# --- Restart All Programs after file drop ---
		try
		{
			$mailslot = "\\$machineName\mailslot\SMSStart_${StoreNumber}${laneNum}"
			$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
			$result = $null
			if ($global:MailslotSender -or ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetTypes() -match 'MailslotSender' }))
			{
				$result = [MailslotSender]::SendMailslotCommand($mailslot, $commandMessage)
			}
			if ($result)
			{
				Write_Log "Restart All Programs sent to $machineName (Lane $laneNum)" "cyan"
			}
			else
			{
				Write_Log "Could not send restart command to $machineName (Lane $laneNum)" "yellow"
			}
		}
		catch
		{
			Write_Log "Error sending restart command to $machineName (Lane $laneNum): $_" "red"
		}
	}
	
	Write_Log "$refreshedCount registers refreshed for Store $StoreNumber." "blue"
	Write_Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
}

# ===================================================================================================
#                                  FUNCTION: Reboot_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Presents a tabbed dialog for rebooting Lanes (with mailslot), Scales, and/or Backoffices.
#   - For Lanes: tries SMSStart mailslot, then shutdown.exe, then Restart-Computer.
#   - For Scales and Backoffices: tries shutdown.exe, then Restart-Computer.
#   - Tabs are dynamically shown depending on -NodeTypes.
# Usage:
#   Reboot_Nodes -StoreNumber "001" -NodeTypes Lane
#   Reboot_Nodes -StoreNumber "001" -NodeTypes Scale
#   Reboot_Nodes -StoreNumber "001" -NodeTypes Lane,Scale
#   Reboot_Nodes -StoreNumber "001" -NodeTypes Backoffice
#   Reboot_Nodes -StoreNumber "001" -NodeTypes Lane,Scale,Backoffice
# ===================================================================================================

function Reboot_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter()]
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string[]]$NodeTypes = @("Lane", "Scale", "Backoffice")
	)
	Write_Log "`r`n==================== Starting Reboot_Nodes Function ====================`r`n" "blue"
	
	# Load global node data
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
	
	# --- Use the shared selection dialog ---
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes $NodeTypes
	if (-not $selection)
	{
		Write_Log "No nodes selected or dialog cancelled." "yellow"
		Write_Log "`r`n==================== Reboot_Nodes Function Completed ====================" "blue"
		return
	}
	
	# ---- Lanes ----
	if ($NodeTypes -contains "Lane" -and $selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		foreach ($lane in $selection.Lanes)
		{
			$laneNum = $lane
			if ($LaneNumToMachineName.ContainsKey($laneNum))
			{
				$machine = $LaneNumToMachineName[$laneNum]
				Write_Log "Lane $laneNum on [$machine]: attempting mailslot reboot" "Yellow"
				$msResult = $false
				try
				{
					$mailslot = "\\$machine\mailslot\SMSStart_${StoreNumber}${laneNum}"
					$msResult = [MailslotSender]::SendMailslotCommand($mailslot, '@exec(REBOOT=1).')
				}
				catch { $msResult = $false }
				if ($msResult)
				{
					Write_Log "Mailslot reboot sent to $machine (Lane $laneNum)" "Green"
					continue
				}
				Write_Log "Mailslot reboot failed for $machine. Falling back to shutdown.exe" "Yellow"
				cmd.exe /c "shutdown /r /m \\$machine /t 0 /f" | Out-Null
				if ($LASTEXITCODE -eq 0)
				{
					Write_Log "shutdown.exe reboot succeeded for $machine" "Green"
					continue
				}
				Write_Log "shutdown.exe exit code $LASTEXITCODE. Now trying Restart-Computer" "Yellow"
				try
				{
					Restart-Computer -ComputerName $machine -Force -ErrorAction SilentlyContinue
					if ($?)
					{
						Write_Log "Restart-Computer succeeded for $machine" "Green"
					}
					else
					{
						Write_Log "All reboot methods failed for $machine (Lane $laneNum)" "Red"
					}
				}
				catch
				{
					Write_Log "All reboot methods failed for $machine (Lane $laneNum)" "Red"
				}
			}
			else
			{
				Write_Log "Lane mapping not found for $laneNum." "yellow"
			}
		}
	}
	
	# ---- Scales ----
	if ($NodeTypes -contains "Scale" -and $selection.Scales -and $selection.Scales.Count -gt 0)
	{
		foreach ($code in $selection.Scales)
		{
			if ($ScaleCodeToIPInfo.ContainsKey($code))
			{
				$scale = $ScaleCodeToIPInfo[$code]
				$ip = $scale.FullIP
				$name = $scale.ScaleName
				$display = if ($name) { "$name [$ip]" }
				else { "$code [$ip]" }
				Write_Log "Attempting to reboot scale: $display at $ip" "Yellow"
				try
				{
					$shutdownArgs = "/r /m \\$ip /t 0 /f"
					$proc = Start-Process -FilePath "shutdown.exe" -ArgumentList $shutdownArgs -Wait -PassThru -ErrorAction Stop
					if ($proc.ExitCode -ne 0) { throw "Shutdown command exited with code $($proc.ExitCode)" }
					Write_Log "Shutdown command executed successfully for $ip." "Green"
				}
				catch
				{
					Write_Log "Shutdown command failed for $ip. Falling back to Restart-Computer." "Red"
					try
					{
						Restart-Computer -ComputerName $ip -Force -ErrorAction Stop
						Write_Log "Restart-Computer command executed successfully for $ip." "Green"
					}
					catch
					{
						Write_Log "Failed to reboot scale $ip using both methods: $_" "Red"
					}
				}
			}
			else
			{
				Write_Log "Scale mapping not found for $code." "yellow"
			}
		}
	}
	
	# ---- Backoffices ----
	if ($NodeTypes -contains "Backoffice" -and $selection.Backoffices -and $selection.Backoffices.Count -gt 0)
	{
		foreach ($boNum in $selection.Backoffices)
		{
			if ($BackofficeNumToMachineName.ContainsKey($boNum))
			{
				$machine = $BackofficeNumToMachineName[$boNum]
				Write_Log "Attempting to reboot backoffice $boNum [$machine]" "Yellow"
				try
				{
					$shutdownArgs = "/r /m \\$machine /t 0 /f"
					$proc = Start-Process -FilePath "shutdown.exe" -ArgumentList $shutdownArgs -Wait -PassThru -ErrorAction Stop
					if ($proc.ExitCode -ne 0) { throw "Shutdown command exited with code $($proc.ExitCode)" }
					Write_Log "Shutdown command executed successfully for $machine." "Green"
				}
				catch
				{
					Write_Log "Shutdown command failed for $machine. Falling back to Restart-Computer." "Red"
					try
					{
						Restart-Computer -ComputerName $machine -Force -ErrorAction Stop
						Write_Log "Restart-Computer command executed successfully for $machine." "Green"
					}
					catch
					{
						Write_Log "Failed to reboot backoffice $boNum ($machine) using both methods: $_" "Red"
					}
				}
			}
			else
			{
				Write_Log "Backoffice mapping not found for $boNum." "yellow"
			}
		}
	}
	
	Write_Log "`r`n==================== Reboot_Nodes Function Completed ====================" "blue"
	[System.Windows.Forms.MessageBox]::Show("Reboot commands issued for selected nodes.", "Reboot",
	[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
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
	$scriptFolder = $script:ScriptsFolder
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
	if (-not $script:FunctionResults.ContainsKey('LaneNumToMachineName') -or -not $script:FunctionResults['LaneNumToMachineName'])
	{
		Write_Log "No lane machine paths found. Did you run Retrieve_Nodes?" "red"
	}
	else
	{
		foreach ($laneNum in $script:FunctionResults['LaneNumToMachineName'].Keys | Sort-Object { [int]$_ })
		{
			$machine = $script:FunctionResults['LaneNumToMachineName'][$laneNum]
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
#                      FUNCTION: Enable_SQL_Protocols_On_Selected_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Enables TCP/IP, Named Pipes, and Shared Memory protocols for all SQL Server instances
#   on the selected remote lane machines. Sets static TCP port (default: 1433), disables dynamic ports,
#   and restarts each SQL Server service. Handles both 64-bit and 32-bit registry locations.
#   Logs detailed status for each lane and each protocol.
#
# Parameters:
#   - [string]$StoreNumber
#         A 3-digit identifier for the store (SSS). Used to retrieve lanes and machine mappings.
#   - [string]$tcpPort
#         The static TCP port to set for all SQL Server instances (default: "1433")
#
# Workflow:
#   1. Prompt user for lanes using Show_Lane_Selection_Form.
#   2. For each selected lane, find the machine and enumerate all SQL Server instances.
#   3. For each instance:
#         - Enable TCP/IP, Named Pipes, Shared Memory protocols (in both 64- and 32-bit registry if found)
#         - Set static TCP port and clear dynamic port
#         - Restart the SQL Server service for the instance
#         - Log all actions and verification results
#
# Returns:
#   None. Outputs results and errors via Write_Log.
#
# Example Usage:
#   Enable_SQL_Protocols_On_Selected_Lanes -StoreNumber "123"
#
# Notes:
#   - RemoteRegistry and service control must be allowed on the remote machine.
#   - $script:FunctionResults['LaneNumToMachineName'] must be defined and populated.
#   - Show_Node_Selection_Form and Write_Log must be available.
# ===================================================================================================

function Enable_SQL_Protocols_On_Selected_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$tcpPort = "1433"
	)
	
	Write_Log "`r`n==================== Starting Enable_SQL_Protocols_On_Selected_Lanes Function ====================`r`n" "blue"
	
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "No lanes available. Please retrieve nodes first." "red"
		Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Function Completed ====================" "blue"
		return
	}
	
	# --- Always use the new, universal selection form ---
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
		Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Function Completed ====================" "blue"
		return
	}
	
	# Extract lanes: always convert to zero-padded 3-digit string, deduplicate
	$lanes = $selection.Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		else { $_ }
	} | ForEach-Object { $_.ToString().PadLeft(3, '0') } | Sort-Object -Unique
	
	Write_Log "Selected lanes: $($lanes -join ', ')" "green"
	
	$isSingle = ($lanes.Count -eq 1)
	
	# Ensure protocol tracking tables exist
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
	if (-not $script:ProtocolResults) { $script:ProtocolResults = @() }
	
	$jobs = @()
	
	foreach ($lane in $lanes)
	{
		$machine = $LaneNumToMachineName[$lane]
		if (-not $machine)
		{
			Write_Log "Machine name not found for lane $lane. Skipping." "yellow"
			continue
		}
		
		# --------- Inline for single lane, job for multi ----------
		if ($isSingle)
		{
			Write_Log "`r`n--- Processing Machine: $machine (Store $StoreNumber, Lane $lane) ---" "blue"
			try
			{				Write_Log "Ensuring RemoteRegistry is running on $machine..." "gray"
				sc.exe "\\$machine" config RemoteRegistry start= demand | Out-Null
				sc.exe "\\$machine" start RemoteRegistry | Out-Null
				$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $machine)
				$instanceRootPaths = @(
					"SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL",
					"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL"
				)
				$allInstances = @{ }
				foreach ($rootPath in $instanceRootPaths)
				{
					$instKey = $reg.OpenSubKey($rootPath)
					if ($instKey)
					{
						foreach ($name in $instKey.GetValueNames())
						{
							$id = $instKey.GetValue($name)
							if ($id -and !$allInstances.ContainsKey($name))
							{
								$allInstances[$name] = $id
							}
						}
						$instKey.Close()
					}
				}
				if ($allInstances.Count -eq 0)
				{
					Write_Log "No SQL instances found on $machine." "yellow"
					$reg.Close()
				}
				else
				{
					$laneNeedsRestart = $false
					foreach ($instanceName in $allInstances.Keys)
					{
						$instanceID = $allInstances[$instanceName]
						Write_Log "Processing SQL Instance: $instanceName (ID: $instanceID)" "blue"
						$needsRestart = $false
						
						# --- Mixed Mode ---
						$authPaths = @(
							"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer",
							"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer"
						)
						$authSet = $false
						foreach ($authPath in $authPaths)
						{
							$authKey = $reg.OpenSubKey($authPath, $true)
							if ($authKey)
							{
								$loginMode = $authKey.GetValue("LoginMode", 1)
								if ($loginMode -ne 2)
								{
									$authKey.SetValue("LoginMode", 2, [Microsoft.Win32.RegistryValueKind]::DWord)
									Write_Log "Mixed Mode Authentication enabled at $authPath." "green"
									$needsRestart = $true
								}
								else
								{
									Write_Log "Mixed Mode Authentication already enabled at $authPath." "gray"
								}
								$authKey.Close()
								$authSet = $true
								break
							}
						}
						if (-not $authSet)
						{
							Write_Log "LoginMode registry path not found for $instanceName." "yellow"
						}
						
						# --- TCP/IP Protocol ---
						$basePaths = @(
							"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp",
							"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp"
						)
						foreach ($basePath in $basePaths)
						{
							$regKey = $reg.OpenSubKey($basePath, $true)
							if ($regKey)
							{
								$tcpEnabled = $regKey.GetValue('Enabled', 0)
								if ($tcpEnabled -eq 1)
								{
									Write_Log "TCP/IP already enabled at $basePath." "gray"
								}
								else
								{
									$regKey.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
									Write_Log "TCP/IP protocol enabled at $basePath." "green"
									$needsRestart = $true
								}
								$regKey.Close()
								break
							}
						}
						# --- TCP Port ---
						$ipAllPaths = @(
							"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll",
							"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll"
						)
						foreach ($ipAllPath in $ipAllPaths)
						{
							$regKey = $reg.OpenSubKey($ipAllPath, $true)
							if ($regKey)
							{
								$curPort = $regKey.GetValue('TcpPort', "")
								$curDyn = $regKey.GetValue('TcpDynamicPorts', "")
								if ($curPort -eq $tcpPort -and $curDyn -eq "")
								{
									Write_Log "TCP port already set to $tcpPort at $ipAllPath." "gray"
								}
								else
								{
									$regKey.SetValue('TcpPort', $tcpPort, [Microsoft.Win32.RegistryValueKind]::String)
									$regKey.SetValue('TcpDynamicPorts', '', [Microsoft.Win32.RegistryValueKind]::String)
									Write_Log "Registry port set to $tcpPort at $ipAllPath." "green"
									$needsRestart = $true
								}
								$regKey.Close()
								break
							}
						}
						# --- Named Pipes ---
						$npBasePaths = @(
							"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np",
							"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np"
						)
						foreach ($npBasePath in $npBasePaths)
						{
							$regKey = $reg.OpenSubKey($npBasePath, $true)
							if ($regKey)
							{
								$npEnabled = $regKey.GetValue('Enabled', 0)
								if ($npEnabled -eq 1)
								{
									Write_Log "Named Pipes already enabled at $npBasePath." "gray"
								}
								else
								{
									$regKey.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
									Write_Log "Named Pipes protocol enabled at $npBasePath." "green"
									$needsRestart = $true
								}
								$regKey.Close()
								break
							}
						}
						# --- Shared Memory ---
						$smBasePaths = @(
							"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm",
							"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm"
						)
						foreach ($smBasePath in $smBasePaths)
						{
							$regKey = $reg.OpenSubKey($smBasePath, $true)
							if ($regKey)
							{
								$smEnabled = $regKey.GetValue('Enabled', 0)
								if ($smEnabled -eq 1)
								{
									Write_Log "Shared Memory already enabled at $smBasePath." "gray"
								}
								else
								{
									$regKey.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
									Write_Log "Shared Memory protocol enabled at $smBasePath." "green"
									$needsRestart = $true
								}
								$regKey.Close()
								break
							}
						}
						# --- Service restart if needed ---
						if ($needsRestart)
						{
							$svcName = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' }
							else { "MSSQL`$$instanceName" }
							Write_Log "Restarting SQL Service $svcName on $machine..." "gray"
							sc.exe "\\$machine" stop $svcName | Out-Null
							Start-Sleep -Seconds 10
							sc.exe "\\$machine" start $svcName | Out-Null
							Start-Sleep -Seconds 3
							Write_Log "SQL Service $svcName restarted successfully on $machine." "green"
							$laneNeedsRestart = $true
						}
						else
						{
							Write_Log "No protocol or auth changes required for $instanceName on $machine. No restart needed." "green"
						}
					}
					$reg.Close()
					
					if ($laneNeedsRestart)
					{
						Send_Restart_All_Programs -StoreNumber $StoreNumber -LaneNumbers @($lane) -Silent
						Write_Log "Restart All Programs sent to $machine (Lane $lane) after protocol update." "green"
					}
					
					# --- Test actual protocol ---
					$protocol = "File"
					try
					{
						$tcpClient = New-Object System.Net.Sockets.TcpClient
						$connectTask = $tcpClient.ConnectAsync($machine, 1433)
						if ($connectTask.Wait(600) -and $tcpClient.Connected)
						{
							$tcpClient.Close()
							$protocol = "TCP"
						}
						else
						{
							try
							{
								Import-Module SqlServer -ErrorAction Stop
								$npConn = "Server=$machine;Database=master;Integrated Security=True;Network Library=dbnmpntw"
								Invoke-Sqlcmd -ConnectionString $npConn -Query "SELECT 1" -QueryTimeout 1 -ErrorAction Stop | Out-Null
								$protocol = "Named Pipes"
							}
							catch { }
						}
					}
					catch { }
					$script:LaneProtocols[$lane] = $protocol
					$script:ProtocolResults = $script:ProtocolResults | Where-Object { $_.Lane -ne $lane }
					if ($null -eq $script:ProtocolResults -or $script:ProtocolResults -isnot [System.Collections.IEnumerable]) { $script:ProtocolResults = @() }
					$script:ProtocolResults += [PSCustomObject]@{ Lane = $lane; Protocol = $protocol }
					Write_Log "Protocol detected for $machine (Lane $lane): $protocol" "magenta"
				}
			}
			catch
			{
				Write_Log "Failed to process [$machine]: $_" "red"
				$script:LaneProtocols[$lane] = "File"
				$script:ProtocolResults = $script:ProtocolResults | Where-Object { $_.Lane -ne $lane }
				$script:ProtocolResults += [PSCustomObject]@{ Lane = $lane; Protocol = "File" }
			}
		}
		else
		{			# Multi-lane: job logic
			$job = Start-Job -ArgumentList $machine, $lane, $StoreNumber, $tcpPort -ScriptBlock {
				param ($machine,
					$lane,
					$StoreNumber,
					$tcpPort)
				$output = @()
				$laneNeedsRestart = $false
				try
				{
					$output += @{ Text = "`r`n--- Processing Machine: $machine (Store $StoreNumber, Lane $lane) ---"; Color = "blue" }
					$output += @{ Text = "Ensuring RemoteRegistry is running on $machine..."; Color = "gray" }
					$null = sc.exe "\\$machine" config RemoteRegistry start= demand
					$null = sc.exe "\\$machine" start RemoteRegistry
					$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $machine)
					$instanceRootPaths = @(
						"SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL",
						"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL"
					)
					$allInstances = @{ }
					foreach ($rootPath in $instanceRootPaths)
					{
						$instKey = $reg.OpenSubKey($rootPath)
						if ($instKey)
						{
							foreach ($name in $instKey.GetValueNames())
							{
								$id = $instKey.GetValue($name)
								if ($id -and !$allInstances.ContainsKey($name))
								{
									$allInstances[$name] = $id
								}
							}
							$instKey.Close()
						}
					}
					if ($allInstances.Count -eq 0)
					{
						$output += @{ Text = "No SQL instances found on $machine."; Color = "yellow" }
						$reg.Close()
					}
					else
					{
						foreach ($instanceName in $allInstances.Keys)
						{
							$instanceID = $allInstances[$instanceName]
							$output += @{ Text = "Processing SQL Instance: $instanceName (ID: $instanceID)"; Color = "blue" }
							$needsRestart = $false
							
							# --- Mixed Mode ---
							$authPaths = @(
								"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer",
								"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer"
							)
							$authSet = $false
							foreach ($authPath in $authPaths)
							{
								$authKey = $reg.OpenSubKey($authPath, $true)
								if ($authKey)
								{
									$loginMode = $authKey.GetValue("LoginMode", 1)
									if ($loginMode -ne 2)
									{
										$authKey.SetValue("LoginMode", 2, [Microsoft.Win32.RegistryValueKind]::DWord)
										$output += @{ Text = "Mixed Mode Authentication enabled at $authPath."; Color = "green" }
										$needsRestart = $true
									}
									else
									{
										$output += @{ Text = "Mixed Mode Authentication already enabled at $authPath."; Color = "gray" }
									}
									$authKey.Close()
									$authSet = $true
									break
								}
							}
							if (-not $authSet)
							{
								$output += @{ Text = "LoginMode registry path not found for $instanceName."; Color = "yellow" }
							}
							
							# --- TCP/IP Protocol ---
							$basePaths = @(
								"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp",
								"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp"
							)
							foreach ($basePath in $basePaths)
							{
								$regKey = $reg.OpenSubKey($basePath, $true)
								if ($regKey)
								{
									$tcpEnabled = $regKey.GetValue('Enabled', 0)
									if ($tcpEnabled -eq 1)
									{
										$output += @{ Text = "TCP/IP already enabled at $basePath."; Color = "gray" }
									}
									else
									{
										$regKey.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
										$output += @{ Text = "TCP/IP protocol enabled at $basePath."; Color = "green" }
										$needsRestart = $true
									}
									$regKey.Close()
									break
								}
							}
							# --- TCP Port ---
							$ipAllPaths = @(
								"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll",
								"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll"
							)
							foreach ($ipAllPath in $ipAllPaths)
							{
								$regKey = $reg.OpenSubKey($ipAllPath, $true)
								if ($regKey)
								{
									$curPort = $regKey.GetValue('TcpPort', "")
									$curDyn = $regKey.GetValue('TcpDynamicPorts', "")
									if ($curPort -eq $tcpPort -and $curDyn -eq "")
									{
										$output += @{ Text = "TCP port already set to $tcpPort at $ipAllPath."; Color = "gray" }
									}
									else
									{
										$regKey.SetValue('TcpPort', $tcpPort, [Microsoft.Win32.RegistryValueKind]::String)
										$regKey.SetValue('TcpDynamicPorts', '', [Microsoft.Win32.RegistryValueKind]::String)
										$output += @{ Text = "Registry port set to $tcpPort at $ipAllPath."; Color = "green" }
										$needsRestart = $true
									}
									$regKey.Close()
									break
								}
							}
							# --- Named Pipes ---
							$npBasePaths = @(
								"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np",
								"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np"
							)
							foreach ($npBasePath in $npBasePaths)
							{
								$regKey = $reg.OpenSubKey($npBasePath, $true)
								if ($regKey)
								{
									$npEnabled = $regKey.GetValue('Enabled', 0)
									if ($npEnabled -eq 1)
									{
										$output += @{ Text = "Named Pipes already enabled at $npBasePath."; Color = "gray" }
									}
									else
									{
										$regKey.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
										$output += @{ Text = "Named Pipes protocol enabled at $npBasePath."; Color = "green" }
										$needsRestart = $true
									}
									$regKey.Close()
									break
								}
							}
							# --- Shared Memory ---
							$smBasePaths = @(
								"SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm",
								"SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm"
							)
							foreach ($smBasePath in $smBasePaths)
							{
								$regKey = $reg.OpenSubKey($smBasePath, $true)
								if ($regKey)
								{
									$smEnabled = $regKey.GetValue('Enabled', 0)
									if ($smEnabled -eq 1)
									{
										$output += @{ Text = "Shared Memory already enabled at $smBasePath."; Color = "gray" }
									}
									else
									{
										$regKey.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
										$output += @{ Text = "Shared Memory protocol enabled at $smBasePath."; Color = "green" }
										$needsRestart = $true
									}
									$regKey.Close()
									break
								}
							}
							# --- Service restart if needed ---
							if ($needsRestart)
							{
								$svcName = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' }
								else { "MSSQL`$$instanceName" }
								$output += @{ Text = "Restarting SQL Service $svcName on $machine..."; Color = "gray" }
								sc.exe "\\$machine" stop $svcName | Out-Null
								Start-Sleep -Seconds 10
								sc.exe "\\$machine" start $svcName | Out-Null
								Start-Sleep -Seconds 3
								$output += @{ Text = "SQL Service $svcName restarted successfully on $machine."; Color = "green" }
								$laneNeedsRestart = $true
							}
							else
							{
								$output += @{ Text = "No protocol or auth changes required for $instanceName on $machine. No restart needed."; Color = "green" }
							}
						}
						$reg.Close()
					}
					
					# --- Test actual protocol ---
					$protocol = "File"
					try
					{
						$tcpClient = New-Object System.Net.Sockets.TcpClient
						$connectTask = $tcpClient.ConnectAsync($machine, 1433)
						if ($connectTask.Wait(600) -and $tcpClient.Connected)
						{
							$tcpClient.Close()
							$protocol = "TCP"
						}
						else
						{
							try
							{
								Import-Module SqlServer -ErrorAction Stop
								$npConn = "Server=$machine;Database=master;Integrated Security=True;Network Library=dbnmpntw"
								Invoke-Sqlcmd -ConnectionString $npConn -Query "SELECT 1" -QueryTimeout 2 -ErrorAction Stop | Out-Null
								$protocol = "Named Pipes"
							}
							catch { }
						}
					}
					catch { }
					$output += @{ Text = "Protocol detected for $machine (Lane $lane): $protocol"; Color = "magenta" }
				}
				catch
				{
					$output += @{ Text = "Failed to process [$machine]: $_"; Color = "red" }
					$protocol = "File"
				}
				[PSCustomObject]@{
					Output    = $output
					Protocol  = $protocol
					Lane	  = $lane
					Restarted = $laneNeedsRestart
				}
			}
			$jobs += @{ Lane = $lane; Job = $job }
		}
	}
	
	if (-not $isSingle)
	{
		$laneOrder = $lanes | Sort-Object
		$jobMap = @{ }
		foreach ($j in $jobs) { $jobMap[$j.Lane] = $j.Job }
		$restartedLanes = @()
		foreach ($lane in $laneOrder)
		{
			$job = $jobMap[$lane]
			Wait-Job $job | Out-Null
			$result = Receive-Job $job
			Remove-Job $job
			foreach ($line in $result.Output) { Write_Log $line.Text $line.Color }
			$script:LaneProtocols[$result.Lane] = $result.Protocol
			$script:ProtocolResults = $script:ProtocolResults | Where-Object { $_.Lane -ne $result.Lane }
			$script:ProtocolResults += [PSCustomObject]@{ Lane = $result.Lane; Protocol = $result.Protocol }
			if ($result.Restarted)
			{
				$restartedLanes += $lane
			}
		}
		if ($restartedLanes.Count -gt 0)
		{
			Send_Restart_All_Programs -StoreNumber $StoreNumber -LaneNumbers $restartedLanes -Silent
			Write_Log "Restart All Programs sent to restarted lanes: $($restartedLanes -join ', ')" "green"
		}
	}
	Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Function Completed ====================" "blue"
}

# ===================================================================================================
#                 FUNCTION: Open_Selected_Node_C_Path (Unified, uses tabbed picker)
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts user with Show_Node_Selection_Form for Lanes and/or Bizerba Scales,
#   then opens \\MACHINE\c$ or \\SCALEIP\c$ accordingly.
#   - For scales: tries bizuser/bizerba, then bizuser/biyerba if needed.
# Usage:
#   Open_Selected_Node_C_Path -StoreNumber "001" -NodeTypes Lane
#   Open_Selected_Node_C_Path -StoreNumber "001" -NodeTypes Scale
#   Open_Selected_Node_C_Path -StoreNumber "001" -NodeTypes Lane,Scale
# ===================================================================================================

function Open_Selected_Node_C_Path
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter()]
		[ValidateSet("Lane", "Scale")]
		[string[]]$NodeTypes = @("Lane", "Scale")
	)
	
	Write_Log "`r`n==================== Starting Open_Selected_Node_C_Path Function ====================`r`n" "blue"
	
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes $NodeTypes
	if (-not $selection)
	{
		Write_Log "No selection made. Exiting." "Yellow"
		return
	}
	
	# Normalize lane selection to 3-digit and get machine name
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$lanes = @()
	if ($selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		foreach ($laneSel in $selection.Lanes)
		{
			$laneNum = if ($laneSel -is [PSCustomObject] -and $laneSel.LaneNumber) { $laneSel.LaneNumber }
			elseif ($laneSel -match '^\d{3}$') { $laneSel }
			elseif ($LaneNumToMachineName.ContainsKey($laneSel)) { $laneSel }
			else { $null }
			if ($laneNum -and $LaneNumToMachineName.ContainsKey($laneNum))
			{
				$machine = $LaneNumToMachineName[$laneNum]
				$sharePath = "\\$machine\c$"
				Write_Log "Opened $sharePath ..." "Green"
				Start-Process "explorer.exe" $sharePath
			}
			else
			{
				Write_Log "Machine not found for lane '$laneSel'." "Red"
			}
		}
	}
	
	# Normalize scale selection (scale objects or just codes)
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	if ($selection.Scales -and $selection.Scales.Count -gt 0)
	{
		foreach ($scaleObj in $selection.Scales)
		{
			$scale = $null
			if ($scaleObj -is [PSCustomObject])
			{
				$scale = $scaleObj
			}
			elseif ($ScaleCodeToIPInfo.ContainsKey($scaleObj))
			{
				$scale = $ScaleCodeToIPInfo[$scaleObj]
			}
			if (-not $scale) { Write_Log "Could not resolve scale info for $scaleObj." "Red"; continue }
			
			$scaleHost = $null
			if ($scale.PSObject.Properties['FullIP'] -and $scale.FullIP)
			{
				$scaleHost = $scale.FullIP
			}
			elseif ($scale.PSObject.Properties['IPNetwork'] -and $scale.PSObject.Properties['IPDevice'])
			{
				$scaleHost = "$($scale.IPNetwork)$($scale.IPDevice)"
			}
			elseif ($scale.PSObject.Properties['ScaleName'])
			{
				$scaleHost = $scale.ScaleName
			}
			if (-not $scaleHost) { Write_Log "Could not determine host for $($scale.ScaleCode)." "Red"; continue }
			
			$sharePath = "\\$scaleHost\c$"
			$opened = $false
			
			# Try bizuser/bizerba, then bizuser/biyerba
			cmdkey /add:$scaleHost /user:bizuser /pass:bizerba | Out-Null
			Start-Process "explorer.exe" $sharePath
			Start-Sleep -Seconds 2
			if (Test-Path $sharePath)
			{
				Write_Log "Opened $sharePath as bizuser using password 'bizerba'." "Green"
				$opened = $true
			}
			else
			{
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
				}
			}
			# Clean up credentials (optional)
			# cmdkey /delete:$scaleHost | Out-Null
		}
	}
	Write_Log "`r`n==================== Open_Selected_Node_C_Path Function Completed ====================" "blue"
}

# ===================================================================================================
#                      FUNCTION: Remove_Duplicate_Files_From_toBizerba
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user to either run the Duplicate File Monitor now (as a background job, controllable via the GUI) or
#   schedule it as a Windows scheduled task (runs at logon, hidden, always-on).
#   Monitors the folder 'C:\Bizerba\RetailConnect\BMS\toBizerba' for duplicate files by content
#   (using hash), and deletes all but the oldest file (by CreationTime).
#   Writes PowerShell script to disk and manages the Windows scheduled task.
#   For "Run Now", starts as a job tied to the session, keeps the GUI open with a "Stop" button, and stops the job on close/stop.
#   Author: Alex_C.T
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   (none - uses script context)
# ===================================================================================================

function Remove_Duplicate_Files_From_toBizerba
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Remove_Duplicate_Files_From_toBizerba Function ====================`r`n" "blue"
	
	$scriptFolder = $script:ScriptsFolder
	$psScriptName = "Remove_Duplicate_Files_From_toBizerba.ps1"
	$psScriptPath = Join-Path $scriptFolder $psScriptName
	$TargetPath = "C:\Bizerba\RetailConnect\BMS\toBizerba"
	$taskName = "Remove_Duplicate_Files_From_toBizerba"
	$logPath = Join-Path $scriptFolder "Remove_Duplicates.log"
	
	if (-not (Test-Path $scriptFolder)) { New-Item -Path $scriptFolder -ItemType Directory | Out-Null }
	
	# Write the PowerShell watcher script (with param and logging)
	$psScriptContent = @"
param([int]`$IntervalSeconds = 5)
`$Path = '$TargetPath'
`$LogPath = '$logPath'
function Remove-DuplicateFilesByContent {
    param([string]`$Path)
    `$files = Get-ChildItem -Path `$Path -File -ErrorAction SilentlyContinue
    `$hashTable = @{}
    foreach (`$file in `$files) {
        try { `$hash = (Get-FileHash -Path `$file.FullName -Algorithm SHA256).Hash }
        catch { continue }
        if (-not `$hashTable.ContainsKey(`$hash)) { `$hashTable[`$hash] = @() }
        `$hashTable[`$hash] += `$file
    }
    foreach (`$entry in `$hashTable.GetEnumerator()) {
        `$fileList = `$entry.Value
        if (`$fileList.Count -gt 1) {
            Add-Content -Path `$LogPath -Value "`$(Get-Date): Found duplicates for hash `$($entry.Key): `$(`$fileList.Count) files"
            `$fileList = `$fileList | Sort-Object CreationTime
            `$original = `$fileList[0]
            `$duplicates = `$fileList[1..(`$fileList.Count - 1)]
            foreach (`$dup in `$duplicates) {
                try { 
                    Remove-Item `$dup.FullName -Force 
                    Add-Content -Path `$LogPath -Value "`$(Get-Date): Removed duplicate `$($dup.FullName), kept `$($original.FullName)"
                } catch {
                    Add-Content -Path `$LogPath -Value "`$(Get-Date): Failed to remove `$($dup.FullName): `$_"
                }
            }
        }
    }
}
Add-Content -Path `$LogPath -Value "`$(Get-Date): Monitor started with interval `$IntervalSeconds seconds"
while (`$true) {
    Remove-DuplicateFilesByContent -Path `$Path
    Start-Sleep -Seconds `$IntervalSeconds
}
"@
	Set-Content -Path $psScriptPath -Value $psScriptContent -Encoding UTF8
	
	# ---- Check if task exists ----
	$hasTask = [bool](schtasks /Query /TN "$taskName" 2>$null)
	
	# --- Build the GUI ---
	Add-Type -AssemblyName System.Windows.Forms
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Duplicate File Monitor"
	$form.Size = New-Object System.Drawing.Size(440, 230)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "How do you want to run the Duplicate File Monitor for:`r`n$TargetPath"
	$label.Location = New-Object System.Drawing.Point(16, 15)
	$label.Size = New-Object System.Drawing.Size(390, 40)
	$label.TextAlign = 'MiddleCenter'
	$form.Controls.Add($label)
	
	$btnRunNow = New-Object System.Windows.Forms.Button
	$btnRunNow.Text = "Run Now (as Job)"
	$btnRunNow.Location = New-Object System.Drawing.Point(10, 70)
	$btnRunNow.Size = New-Object System.Drawing.Size(110, 35)
	$form.Controls.Add($btnRunNow)
	
	$btnSchedule = New-Object System.Windows.Forms.Button
	$btnSchedule.Text = "Schedule (background)"
	$btnSchedule.Location = New-Object System.Drawing.Point(125, 70)
	$btnSchedule.Size = New-Object System.Drawing.Size(140, 35)
	$form.Controls.Add($btnSchedule)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(10, 148)
	$btnCancel.Size = New-Object System.Drawing.Size(400, 35)
	$form.Controls.Add($btnCancel)
	
	$lblSec = New-Object System.Windows.Forms.Label
	$lblSec.Text = "Interval (seconds):"
	$lblSec.Location = New-Object System.Drawing.Point(70, 120)
	$lblSec.Size = New-Object System.Drawing.Size(110, 22)
	$form.Controls.Add($lblSec)
	
	$numSec = New-Object System.Windows.Forms.NumericUpDown
	$numSec.Location = New-Object System.Drawing.Point(185, 118)
	$numSec.Size = New-Object System.Drawing.Size(50, 24)
	$numSec.Minimum = 1
	$numSec.Maximum = 3600
	$numSec.Value = 5
	$form.Controls.Add($numSec)
	
	# --- Delete Scheduled Task button, only enabled if task exists ---
	$btnDeleteTask = New-Object System.Windows.Forms.Button
	$btnDeleteTask.Text = "Delete Scheduled Task"
	$btnDeleteTask.Location = New-Object System.Drawing.Point(270, 70)
	$btnDeleteTask.Size = New-Object System.Drawing.Size(140, 35)
	$btnDeleteTask.Enabled = $hasTask
	$form.Controls.Add($btnDeleteTask)
	
	$selectedAction = [ref] ""
	$intervalSeconds = [ref] 5
	$deleteScheduledTask = [ref]$false
	$monitorJob = $null
	
	# Handle form closing: stop job and child processes if running
	$form.Add_FormClosing({
			if ($monitorJob -and $monitorJob.State -eq 'Running')
			{
				try
				{
					Stop-Job -Job $monitorJob -Force
					Remove-Job -Job $monitorJob -Force
					Write_Log "Monitor job stopped on form close." "yellow"
				}
				catch
				{
					Write_Log "Failed to stop monitor job on close: $_" "red"
				}
			}
			# Stop any child powershell processes
			Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
				$_.CommandLine -and $_.CommandLine -match [Regex]::Escape($psScriptPath)
			} | ForEach-Object {
				try
				{
					Stop-Process -Id $_.ProcessId -Force
					Write_Log "Stopped powershell.exe PID $($_.ProcessId) on form close" "yellow"
				}
				catch
				{
					Write_Log "Failed to stop powershell.exe PID $($_.ProcessId) on close" "red"
				}
			}
		})
	
	$btnRunNow.Add_Click({
			if ($btnRunNow.Text -eq "Run Now (as Job)")
			{
				$intervalSeconds.Value = $numSec.Value
				$monitorJob = Start-Job -ScriptBlock {
					param ($scriptPath,
						$interval)
					& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -IntervalSeconds $interval
				} -ArgumentList $psScriptPath, $intervalSeconds.Value
				
				Write_Log "Started duplicate file monitor as job (ID: $($monitorJob.Id))." "green"
				
				$btnRunNow.Text = "Stop Monitor"
				$btnSchedule.Enabled = $false
				$btnDeleteTask.Enabled = $false
				$btnCancel.Text = "Close (Monitor Running)"
			}
			else
			{
				if ($monitorJob -and $monitorJob.State -eq 'Running')
				{
					try
					{
						Stop-Job -Job $monitorJob -Force
						Remove-Job -Job $monitorJob -Force
						Write_Log "Monitor job stopped by user." "yellow"
					}
					catch
					{
						Write_Log "Failed to stop monitor job: $_" "red"
					}
				}
				# Stop any child powershell processes
				Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
					$_.CommandLine -and $_.CommandLine -match [Regex]::Escape($psScriptPath)
				} | ForEach-Object {
					try
					{
						Stop-Process -Id $_.ProcessId -Force
						Write_Log "Stopped powershell.exe PID $($_.ProcessId) by user stop" "yellow"
					}
					catch
					{
						Write_Log "Failed to stop powershell.exe PID $($_.ProcessId)" "red"
					}
				}
				$btnRunNow.Text = "Run Now (as Job)"
				$btnSchedule.Enabled = $true
				$btnDeleteTask.Enabled = $hasTask
				$btnCancel.Text = "Cancel"
			}
		})
	
	$btnSchedule.Add_Click({
			$intervalSeconds.Value = $numSec.Value
			$selectedAction.Value = "schedule"
			$form.Close()
		})
	
	$btnCancel.Add_Click({
			$form.Close()
		})
	
	$btnDeleteTask.Add_Click({
			$deleteScheduledTask.Value = $true
			$form.Close()
		})
	
	$form.ShowDialog() | Out-Null
	
	# --- Post-dialog actions ---
	
	# Stop existing processes/jobs if deleting or scheduling
	if ($deleteScheduledTask.Value -or $selectedAction.Value -eq "schedule")
	{
		Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
			$_.CommandLine -and $_.CommandLine -match [Regex]::Escape($psScriptPath)
		} | ForEach-Object {
			try
			{
				Stop-Process -Id $_.ProcessId -Force
				Write_Log "Stopped powershell.exe PID $($_.ProcessId) running $psScriptPath" "yellow"
			}
			catch
			{
				Write_Log "Failed to stop powershell.exe PID $($_.ProcessId)" "red"
			}
		}
		Get-Job | Where-Object {
			$_.Command -like "*$psScriptPath*"
		} | ForEach-Object {
			try { Stop-Job $_ -Force }
			catch { }
			try { Remove-Job $_ -Force }
			catch { }
		}
	}
	
	if ($deleteScheduledTask.Value)
	{
		$deleteOut = schtasks /Delete /TN "$taskName" /F 2>&1
		if ($LASTEXITCODE -ne 0 -and -not $deleteOut -match "cannot find the file specified")
		{
			Write_Log "Failed to delete scheduled task: $deleteOut" "red"
		}
		else
		{
			Write_Log "Scheduled task deleted." "green"
		}
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}
	
	if (-not $selectedAction.Value)
	{
		Write_Log "User cancelled." "yellow"
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}
	
	if ($selectedAction.Value -eq "schedule")
	{
		if (-not (Test-Path $psScriptPath))
		{
			Write_Log "Script file missing: $psScriptPath" "red"
			return
		}
		
		$escapedPath = $psScriptPath -replace '"', '""'
		$action = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$escapedPath`" -IntervalSeconds $($intervalSeconds.Value)"
		
		# Delete if exists
		if (schtasks /Query /TN "$taskName" 2>$null)
		{
			schtasks /Delete /TN "$taskName" /F | Out-Null
		}
		
		$createArgs = @(
			'/Create',
			'/TN', $taskName,
			'/TR', $action,
			'/SC', 'ONLOGON',
			'/RL', 'HIGHEST',
			'/RU', 'SYSTEM',
			'/F'
		)
		$createOut = schtasks @createArgs 2>&1
		if ($LASTEXITCODE -ne 0)
		{
			Write_Log "Failed to schedule task: $createOut" "red"
			return
		}
		Write_Log "Task scheduled successfully." "green"
		
		# Start immediately
		try
		{
			Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$psScriptPath`" -IntervalSeconds $($intervalSeconds.Value)"
			Write_Log "Started monitor immediately." "green"
		}
		catch
		{
			Write_Log "Failed to start immediately: $_" "yellow"
		}
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
	}
}

# ===================================================================================================
#                                FUNCTION: Update_Scales_Specials_Interactive
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
	Write_Log "`r`n==================== Starting Update_Scales_Specials_Interactive Function ====================`r`n" "blue"
	
	$DeployRestored = [ref]$false
	$scriptFolder = $script:ScriptsFolder
	$batchName_Daily = "Update_Scales_Specials.bat"
	$batchPath_Daily = Join-Path $scriptFolder $batchName_Daily
	$batchName_Minutes = "Update_Scales_Specials_Minutes.bat"
	$batchPath_Minutes = Join-Path $scriptFolder $batchName_Minutes
	
	$deployChgFile = Join-Path $OfficePath "DEPLOY_CHG.sql"
	
	# Check for all executables
	$hasDeployChg = Test-Path $deployChgFile -ErrorAction SilentlyContinue
	$hasFastDeploy = Test-Path "C:\ScaleCommApp\ScaleManagementApp_FastDEPLOY.exe" -ErrorAction SilentlyContinue
	$hasRegularDeploy = Test-Path "C:\ScaleCommApp\ScaleManagementApp.exe" -ErrorAction SilentlyContinue
	$hasUpdateSpecials = Test-Path "C:\ScaleCommApp\ScaleManagementAppUpdateSpecials.exe" -ErrorAction SilentlyContinue
	
	# Enable/disable schedule buttons based on file existence
	$enableDaily = $hasDeployChg -and $hasUpdateSpecials
	$enableMinutes = $hasDeployChg -and ($hasFastDeploy -or $hasRegularDeploy)
	
	if (-not $hasDeployChg)
	{
		Write_Log "DEPLOY_CHG.sql not found in $OfficePath" "yellow"
	}
	if (-not $hasUpdateSpecials)
	{
		Write_Log "ScaleManagementAppUpdateSpecials.exe not found for 5AM task in C:\ScaleCommApp!" "yellow"
	}
	if (-not ($hasFastDeploy -or $hasRegularDeploy))
	{
		Write_Log "Neither FastDEPLOY nor regular ScaleManagementApp.exe found in C:\ScaleCommApp for minutes task!" "yellow"
	}
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Update Scales Specials"
	$form.Size = New-Object System.Drawing.Size(495, 220)
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
	$btnSchedule.Location = New-Object System.Drawing.Point(5, 90)
	$btnSchedule.Size = New-Object System.Drawing.Size(150, 40)
	$btnSchedule.Enabled = $enableDaily
	$form.Controls.Add($btnSchedule)
	
	$btnRestoreDeployLine = New-Object System.Windows.Forms.Button
	$btnRestoreDeployLine.Text = "Restore DEPLOY_CHG.sql Line"
	$btnRestoreDeployLine.Location = New-Object System.Drawing.Point(155, 90)
	$btnRestoreDeployLine.Size = New-Object System.Drawing.Size(170, 40)
	$btnRestoreDeployLine.Enabled = $false
	$form.Controls.Add($btnRestoreDeployLine)
	
	$btnScheduleMinutes = New-Object System.Windows.Forms.Button
	$btnScheduleMinutes.Text = "Schedule Task (Minutes)"
	$btnScheduleMinutes.Location = New-Object System.Drawing.Point(325, 90)
	$btnScheduleMinutes.Size = New-Object System.Drawing.Size(150, 40)
	$btnScheduleMinutes.Enabled = $enableMinutes
	$form.Controls.Add($btnScheduleMinutes)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(200, 145)
	$btnCancel.Size = New-Object System.Drawing.Size(100, 30)
	$form.Controls.Add($btnCancel)
	
	# --- If missing any necessary files, label displays ---
	if (-not $hasDeployChg)
	{
		$label.Text += "`r`n(DEPLOY_CHG.sql missing - scheduling is disabled)"
	}
	if (-not $hasUpdateSpecials)
	{
		$label.Text += "`r`n(ScaleManagementAppUpdateSpecials.exe missing - 5AM scheduling is disabled)"
	}
	if (-not ($hasFastDeploy -or $hasRegularDeploy))
	{
		$label.Text += "`r`n(Neither FastDEPLOY nor regular ScaleManagementApp.exe present - minutes scheduling is disabled)"
	}
	
	# Determine the correct @EXEC line for deploy file (minutes schedule always prefers FAST if present)
	$correctExeLine = ""
	if ($hasFastDeploy)
	{
		$correctExeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='C:\ScaleCommApp\ScaleManagementApp_FastDEPLOY.exe');"
	}
	elseif ($hasRegularDeploy)
	{
		$correctExeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='C:\ScaleCommApp\ScaleManagementApp.exe');"
	}
	
	# Enable restore button if the exact correct line is missing (exists but wrong format will enable it)
	$deployLineCorrect = $false
	if (Test-Path $deployChgFile)
	{
		$deployContent = Get-Content $deployChgFile -Raw
		if ($deployContent -match [regex]::Escape($correctExeLine))
		{
			$deployLineCorrect = $true
		}
	}
	if (-not $deployLineCorrect) { $btnRestoreDeployLine.Enabled = $true }
	
	$btnRestoreDeployLine.Add_Click({
			# --- Restore DEPLOY_CHG.sql ---
			if (Test-Path $deployChgFile)
			{
				try
				{
					$content = Get-Content $deployChgFile -Raw
					$newContent = [System.Collections.Generic.List[string]]@(
						($content -split "`r?`n") | Where-Object {
							$_ -notmatch '^\s*/\* Deploy price changes to the scales \*/' -and
							$_ -notmatch '(?i)@EXEC\(RUN=''C:\\ScaleCommApp\\ScaleManagementApp(_FastDEPLOY)?\.exe''\);'
						}
					)
					while ($newContent.Count -gt 0 -and ($newContent[-1] -match '^\s*$'))
					{
						$null = $newContent.RemoveAt($newContent.Count - 1)
					}
					$newContent += ""
					$newContent += $correctExeLine
					$newContent -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default
					Write_Log "Restored line to DEPLOY_CHG.sql: $correctExeLine" "green"
					$btnRestoreDeployLine.Enabled = $false
				}
				catch
				{
					Write_Log "Failed to restore line to DEPLOY_CHG.sql: $_" "red"
				}
			}
			else
			{
				Write_Log "DEPLOY_CHG.sql not found for restore." "yellow"
			}
			
			# --- Delete the minutes task if it exists ---
			$minutesTaskName = "Update_Scales_Specials_Task_Minutes"
			$taskExists = schtasks /Query /TN "$minutesTaskName" 2>&1 | Select-String -Quiet -Pattern "$minutesTaskName"
			if ($taskExists)
			{
				$deleteOut = schtasks /Delete /TN "$minutesTaskName" /F 2>&1
				if ($LASTEXITCODE -eq 0)
				{
					Write_Log "Deleted scheduled task '$minutesTaskName' after DEPLOY_CHG.sql restore." "yellow"
				}
				else
				{
					Write_Log "Attempted to delete '$minutesTaskName' after restore but schtasks said: $deleteOut" "yellow"
				}
			}
			
			Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
			$DeployRestored.Value = $true
			$form.Close()
			return
		})
	
	$form.add_FormClosing({
			if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK)
			{
				$script:selectedAction = "cancel"
			}
		})
	
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
	if ($DeployRestored.Value) { return }
	if ($script:selectedAction -eq "cancel" -or -not $script:selectedAction)
	{
		Write_Log "User cancelled Update_Scales_Specials_Interactive." "yellow"
		Write_Log "`r`n==================== Update_Scales_Specials_Interactive Function Completed ====================" "blue"
		return
	}
	
	if (-not (Test-Path $scriptFolder -ErrorAction SilentlyContinue)) { New-Item -Path $scriptFolder -ItemType Directory | Out-Null }
	
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
		if (Test-Path $deployChgFile -ErrorAction SilentlyContinue)
		{
			try
			{
				$lines = [System.Collections.Generic.List[string]]@(Get-Content $deployChgFile)
				$toRemoveIdx = @()
				for ($i = 0; $i -lt $lines.Count; $i++)
				{
					if (
						$lines[$i] -match '^\s*/\* Deploy price changes to the scales \*/' -or
						$lines[$i] -match '(?i)@EXEC\(RUN=''C:\\ScaleCommApp\\ScaleManagementApp(_FastDEPLOY)?\.exe''\);'
					)
					{
						if ($i -gt 0 -and $lines[$i - 1] -match '^\s*$') { $toRemoveIdx += ($i - 1) }
						$toRemoveIdx += $i
					}
				}
				$toRemoveIdx = $toRemoveIdx | Sort-Object -Unique -Descending
				foreach ($idx in $toRemoveIdx) { $lines.RemoveAt($idx) }
				while ($lines.Count -gt 0 -and $lines[-1] -match '^\s*$') { $null = $lines.RemoveAt($lines.Count - 1) }
				$lines -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default
				if ($toRemoveIdx.Count -gt 0)
				{
					Write_Log "Removed banner, @EXEC line, and any blank line above from $deployChgFile" "green"
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
if exist C:\ScaleCommApp\ScaleManagementApp_FastDEPLOY.exe (
    start C:\ScaleCommApp\ScaleManagementApp_FastDEPLOY.exe
) else (
    start C:\ScaleCommApp\ScaleManagementApp.exe
)
exit
"@
		Set-Content -Path $batchPath_Minutes -Value $batchContent -Encoding ASCII
		
		$taskName = "Update_Scales_Specials_Task_Minutes"
		$interval = [int]$script:minutesValue
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
#                                   FUNCTION: Fix_Deploy_CHG
# ---------------------------------------------------------------------------------------------------
# Description:
#   Restores the deploy line to DEPLOY_CHG.sql for scale management.
#   Checks for FastDEPLOY or regular ScaleManagementApp.exe and constructs the appropriate @EXEC line.
#   Removes any existing matching deploy lines from the file, then appends the new line.
#   Deletes the minutes scheduled task if it exists.
#   Uses Write_Log for progress and error reporting.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - OfficePath: Path to the Office folder containing DEPLOY_CHG.sql (defaults to script's $OfficePath).
# ===================================================================================================

function Fix_Deploy_CHG
{
	param (
		[string]$OfficePath = $script:OfficePath
	)
	
	Write_Log "`r`n==================== Starting Fix_Deploy_CHG Function ====================`r`n" "blue"
	
	$deployChgFile = Join-Path $OfficePath "DEPLOY_CHG.sql"
	
	if (-not (Test-Path $deployChgFile))
	{
		Write_Log "DEPLOY_CHG.sql not found at $deployChgFile" "red"
		Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
		return
	}
	
	$exeLine = ""
	if (Test-Path "C:\ScaleCommApp\ScaleManagementApp_FastDEPLOY.exe")
	{
		$exeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='C:\ScaleCommApp\ScaleManagementApp_FastDEPLOY.exe');"
	}
	elseif (Test-Path "C:\ScaleCommApp\ScaleManagementApp.exe")
	{
		$exeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='C:\ScaleCommApp\ScaleManagementApp.exe');"
	}
	else
	{
		Write_Log "Neither FastDEPLOY nor regular ScaleManagementApp.exe found in C:\ScaleCommApp!" "red"
		Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
		return
	}
	
	# --- Restore DEPLOY_CHG.sql ---
	try
	{
		$content = Get-Content $deployChgFile -Raw
		$newContent = [System.Collections.Generic.List[string]]@(
			($content -split "`r?`n") | Where-Object {
				$_ -notmatch '^\s*/\* Deploy price changes to the scales \*/' -and
				$_ -notmatch '(?i)@EXEC\(RUN=''C:\\ScaleCommApp\\ScaleManagementApp(_FastDEPLOY)?\.exe''\);'
			}
		)
		while ($newContent.Count -gt 0 -and ($newContent[-1] -match '^\s*$'))
		{
			$null = $newContent.RemoveAt($newContent.Count - 1)
		}
		$newContent += ""
		$newContent += $exeLine
		$newContent -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default
		Write_Log "Restored line to DEPLOY_CHG.sql: $exeLine" "green"
	}
	catch
	{
		Write_Log "Failed to restore line to DEPLOY_CHG.sql: $_" "red"
		Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
		return
	}
	
	# --- Delete the minutes task if it exists ---
	$minutesTaskName = "Update_Scales_Specials_Task_Minutes"
	$taskExists = schtasks /Query /TN "$minutesTaskName" 2>&1 | Select-String -Quiet -Pattern "$minutesTaskName"
	if ($taskExists)
	{
		$deleteOut = schtasks /Delete /TN "$minutesTaskName" /F 2>&1
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Deleted scheduled task '$minutesTaskName' after DEPLOY_CHG.sql restore." "yellow"
		}
		else
		{
			Write_Log "Attempted to delete '$minutesTaskName' after restore but schtasks said: $deleteOut" "yellow"
		}
	}
	
	Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
}

# ===================================================================================================
#                                 FUNCTION: Manage_Sa_Account
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays a Windows Form with buttons to enable or disable the 'sa' account on the local SQL Server.
#   Sets the password to 'TB$upp0rT' when enabling. Buttons are enabled/disabled based on the current
#   state of the 'sa' account. Uses integrated security for connection. Assumes default SQL instance.
#   Logs actions and errors using Write_Log. Closes the form after a successful enable or disable action.
#
# Assumptions:
#   - Script runs with sufficient privileges (sysadmin on SQL).
#   - Local SQL Server default instance ('.').
#   - Write_Log function is available for logging.
#
# Author: Alex_C.T
# ===================================================================================================

function Manage_Sa_Account
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Manage_Sa_Account Function ====================`r`n" "blue"
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Use already detected SQL module and connection info
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$serverInstance = $script:FunctionResults['DBSERVER']
	$database = "master"
	
	if (-not $SqlModuleName -or $SqlModuleName -eq "None")
	{
		Write_Log "No valid SQL module detected for SQL operations!" "red"
		return
	}
	Import-Module $SqlModuleName -ErrorAction Stop
	$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
	
	$supportsConnectionString = $false
	if ($InvokeSqlCmd)
	{
		$supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString'
	}
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Manage SQL 'sa' Account"
	$form.Size = New-Object System.Drawing.Size(300, 150)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$btnEnable = New-Object System.Windows.Forms.Button
	$btnEnable.Text = "Enable 'sa'"
	$btnEnable.Location = New-Object System.Drawing.Point(50, 30)
	$btnEnable.Size = New-Object System.Drawing.Size(200, 30)
	$form.Controls.Add($btnEnable)
	
	$btnDisable = New-Object System.Windows.Forms.Button
	$btnDisable.Text = "Disable 'sa'"
	$btnDisable.Location = New-Object System.Drawing.Point(50, 70)
	$btnDisable.Size = New-Object System.Drawing.Size(200, 30)
	$form.Controls.Add($btnDisable)
	
	# --- Initial state update ---
	try
	{
		$query = "SELECT is_disabled FROM sys.sql_logins WHERE name = 'sa'"
		if ($supportsConnectionString)
		{
			$result = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query -ErrorAction Stop
		}
		else
		{
			$result = Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $query -ErrorAction Stop
		}
		if ($result)
		{
			$isEnabled = ($result.is_disabled -eq 0)
		}
		else
		{
			Write_Log "'sa' account not found." "red"
			$isEnabled = $false
		}
	}
	catch
	{
		Write_Log "Error checking 'sa' status: $_" "red"
		$isEnabled = $false
	}
	$btnEnable.Enabled = -not $isEnabled
	$btnDisable.Enabled = $isEnabled
	if ($isEnabled)
	{
		Write_Log "'sa' is currently enabled. Enable button greyed out." "yellow"
	}
	else
	{
		Write_Log "'sa' is currently disabled. Disable button greyed out." "yellow"
	}
	
	# --- Enable 'sa' button click event ---
	$btnEnable.Add_Click({
			Write_Log "Enable button clicked. Attempting to enable 'sa'..." "blue"
			try
			{
				$enableQuery = "ALTER LOGIN sa ENABLE; ALTER LOGIN sa WITH PASSWORD = 'TB`$upp0rT';"
				if ($supportsConnectionString)
				{
					Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $enableQuery -ErrorAction Stop
				}
				else
				{
					Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $enableQuery -ErrorAction Stop
				}
				Write_Log "'sa' account enabled and password set successfully." "green"
				# Refresh state and close
				try
				{
					$query = "SELECT is_disabled FROM sys.sql_logins WHERE name = 'sa'"
					if ($supportsConnectionString)
					{
						$result = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query -ErrorAction Stop
					}
					else
					{
						$result = Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $query -ErrorAction Stop
					}
					if ($result) { $isEnabled = ($result.is_disabled -eq 0) }
					else { $isEnabled = $false }
				}
				catch { $isEnabled = $false }
				$btnEnable.Enabled = -not $isEnabled
				$btnDisable.Enabled = $isEnabled
				$form.Close()
			}
			catch
			{
				Write_Log "Error enabling 'sa' account: $_" "red"
			}
		})
	
	# --- Disable 'sa' button click event ---
	$btnDisable.Add_Click({
			Write_Log "Disable button clicked. Attempting to disable 'sa'..." "blue"
			try
			{
				$disableQuery = "ALTER LOGIN sa DISABLE;"
				if ($supportsConnectionString)
				{
					Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $disableQuery -ErrorAction Stop
				}
				else
				{
					Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $disableQuery -ErrorAction Stop
				}
				Write_Log "'sa' account disabled successfully." "green"
				# Refresh state and close
				try
				{
					$query = "SELECT is_disabled FROM sys.sql_logins WHERE name = 'sa'"
					if ($supportsConnectionString)
					{
						$result = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query -ErrorAction Stop
					}
					else
					{
						$result = Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $query -ErrorAction Stop
					}
					if ($result) { $isEnabled = ($result.is_disabled -eq 0) }
					else { $isEnabled = $false }
				}
				catch { $isEnabled = $false }
				$btnEnable.Enabled = -not $isEnabled
				$btnDisable.Enabled = $isEnabled
				$form.Close()
			}
			catch
			{
				Write_Log "Error disabling 'sa' account: $_" "red"
			}
		})
	
	$form.ShowDialog() | Out-Null
	Write_Log "`r`n==================== Manage_Sa_Account Function Completed ====================" "blue"
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
#   - LaneNumToMachineName      [hashtable]: LaneNumber => MachineName mapping.
#   - ScaleCodeToIPInfo   [hashtable]: ScaleCode => Scale Object (includes IP).
#   - BackofficeNumToMachineName[hashtable]: BackofficeTerminal => MachineName mapping.
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
#   Export-VNCFiles-ForAllNodes -LaneNumToMachineName $... -ScaleCodeToIPInfo $... -BackofficeNumToMachineName $... [-LaneVNCPasswords $...]
#
# Author: Alex_C.T
# ===================================================================================================

function Export_VNC_Files_For_All_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory = $true)]
		[hashtable]$ScaleCodeToIPInfo,
		[Parameter(Mandatory = $true)]
		[hashtable]$BackofficeNumToMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$AllVNCPasswords
	)
	
	Write_Log "`r`n==================== Starting Export_VNCFiles_ForAllNodes ====================`r`n" "blue"
	$DefaultVNCPassword = "4330df922eb03b6e"
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"
	
	# ---- If passwords not provided, scan them ----
	if (-not $AllVNCPasswords -or $AllVNCPasswords.Count -eq 0)
	{
		Write_Log "Gathering VNC passwords for all machines`r`n" "magenta"
		$AllVNCPasswords = Get_All_VNC_Passwords -LaneNumToMachineName $LaneNumToMachineName -ScaleCodeToIPInfo $ScaleCodeToIPInfo -BackofficeNumToMachineName $BackofficeNumToMachineName
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
	Write_Log "-------------------- Exporting Lane VNC Files --------------------" "blue"
	$laneCount = 0
	$laneInfoLines = @()
	
	foreach ($lane in $LaneNumToMachineName.GetEnumerator())
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
		if ($AllVNCPasswords -and $AllVNCPasswords.ContainsKey($machineName) -and $AllVNCPasswords[$machineName])
		{
			$VNCPassword = $AllVNCPasswords[$machineName]
		}
		$content = $vncTemplate.Replace('%%HOST%%', $machineName).Replace('%%PASSWORD%%', $VNCPassword)
		[System.IO.File]::WriteAllText($filePath, $content, $script:ansiPcEncoding)
		Write_Log "Created: $filePath" "green"
		$laneCount++
	}
	Write_Log "$laneCount lane VNC files written to $lanesDir`r`n" "blue"
	
	# ---- Scales ---- #
	Write_Log "-------------------- Exporting Scale VNC Files --------------------" "blue"
	$scaleCount = 0
	foreach ($scale in $ScaleCodeToIPInfo.GetEnumerator())
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
			# Set AllowUntrustedServers=1 for Ishida, else use template as-is
			$VNCPassword = $DefaultVNCPassword
			if ($AllVNCPasswords -and $AllVNCPasswords.ContainsKey($ip) -and $AllVNCPasswords[$ip])
			{
				$VNCPassword = $AllVNCPasswords[$ip]
			}
			if ($brand -like '*Ishida*')
			{
				$content = $vncTemplate.Replace('%%HOST%%', $ip).Replace('%%PASSWORD%%', $VNCPassword)
				$content = $content -replace 'AllowUntrustedServers=0', 'AllowUntrustedServers=1'
			}
			else
			{
				$content = $vncTemplate.Replace('%%HOST%%', $ip).Replace('%%PASSWORD%%', $VNCPassword)
			}
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
	Write_Log "-------------------- Exporting Backoffice VNC Files --------------------" "blue"
	$boCount = 0
	foreach ($bo in $BackofficeNumToMachineName.GetEnumerator())
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
#                               FUNCTION: Update_ScaleConfig_And_DB
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user to choose the weighted item marking logic (SCL or POS).
#   Updates all ScaleCommApp XML config files (ProdType_F272) in-place using robust XML logic.
#   If the setting doesn't exist, it is inserted; if it exists, it is updated.
#   Runs the correct SQL update/merge using your item key (default F01) for SCL_TAB and POS_TAB.
#   Uses Write_Log for all progress, result, and error reporting.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   Folder              Directory containing ScaleCommApp config files.           (default: C:\ScaleCommApp)
#   MinItem             Lower bound of item code range.                           (default: 0020000100000)
#   MaxItem             Upper bound of item code range.                           (default: 0029999900000)
#   POS_Field           Field in POS_TAB for weighted flag.                       (default: F82)
#   POS_Value           Value in POS_TAB.$POS_Field to trigger update.            (default: 1)
#   SCL_Field           Field in SCL_TAB to set (usually F272).                   (default: F272)
#   SCL_Value           Value to set in SCL_TAB when in SCL mode.                 (default: 3)
#   POS_Default_F272_Value  Value to set in SCL_TAB when reverting to POS mode.   (default: 0)
#   ItemKey             Column name for unique item code in both tables.          (default: F01)
# ---------------------------------------------------------------------------------------------------

function Update_ScaleConfig_And_DB
{
	Write_Log "`r`n==================== Starting Update_ScaleConfig_And_DB Function ====================`r`n" "blue"
	
	# Set all defaults up front (edit as needed)
	$Folder = 'C:\ScaleCommApp'
	$MinItem = '0020000100000'
	$MaxItem = '0029999900000'
	$POS_Field = 'F82'
	$POS_Value = 0
	$SCL_Field = 'F272'
	$SCL_Value = 3
	$POS_Default_F272_Value = 1
	$ItemKey = 'F01'
	
	# -----------------------------------------------------------------------------
	# Prompt the user for which mode to use (GUI)
	# -----------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Weighted Item Marking Method"
	$form.Size = New-Object System.Drawing.Size(420, 170)
	$form.StartPosition = "CenterScreen"
	$form.Topmost = $true
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Which marking method do you want to implement for weighted items?"
	$label.Size = New-Object System.Drawing.Size(390, 40)
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$form.Controls.Add($label)
	
	$radioSCL = New-Object System.Windows.Forms.RadioButton
	$radioSCL.Text = "New Way (SCL.F272 = 3, config to SCL)"
	$radioSCL.Location = New-Object System.Drawing.Point(30, 50)
	$radioSCL.Size = New-Object System.Drawing.Size(350, 20)
	$radioSCL.Checked = $true
	$form.Controls.Add($radioSCL)
	
	$radioPOS = New-Object System.Windows.Forms.RadioButton
	$radioPOS.Text = "Default / Old Way (POS.F82, config to POS, F272 = 0)"
	$radioPOS.Location = New-Object System.Drawing.Point(30, 75)
	$radioPOS.Size = New-Object System.Drawing.Size(350, 20)
	$form.Controls.Add($radioPOS)
	
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(220, 110)
	$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
	$form.Controls.Add($okButton)
	
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(310, 110)
	$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
	$form.Controls.Add($cancelButton)
	
	$form.AcceptButton = $okButton
	$form.CancelButton = $cancelButton
	$form.Tag = $null
	$form.ShowDialog() | Out-Null
	
	if ($form.Tag -ne "OK")
	{
		Write_Log "User canceled Update_ScaleConfig_And_DB." "yellow"
		Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
		return
	}
	$Mode = if ($radioSCL.Checked) { "SCL" }
	else { "POS" }
	
	# -----------------------------------------------------------------------------
	# XML Config Update (robust, always update or insert key)
	# -----------------------------------------------------------------------------
	$Files = @(
		'ScaleManagementApp.exe.config',
		'ScaleManagementAppUpdateSpecials.exe.config',
		'ScaleManagementAppSetup.exe.config',
		'ScaleManagementApp_FastDEPLOY.exe.config'
	)
	$ProdTypeKey = "ProdType_F272"
	$ProdTypeValue = if ($Mode -eq "SCL") { "SCL.F272" }
	else { "POS.F82" }
	
	foreach ($file in $Files)
	{
		$FullPath = Join-Path $Folder $file
		if (-not (Test-Path $FullPath))
		{
			Write_Log "[$file] Not found, skipped" "yellow"
			continue
		}
		try
		{
			[xml]$xml = Get-Content $FullPath -Raw
			$settings = $xml.configuration.appSettings
			$existing = $settings.add | Where-Object { $_.key -eq $ProdTypeKey }
			if ($existing)
			{
				if ($existing.value -ne $ProdTypeValue)
				{
					$existing.value = $ProdTypeValue
					Write_Log "[$file] Updated key '$ProdTypeKey' to '$ProdTypeValue'" "green"
				}
				else
				{
					Write_Log "[$file] '$ProdTypeKey' already set to '$ProdTypeValue', no change needed" "gray"
				}
			}
			else
			{
				$addElem = $xml.CreateElement("add")
				$addElem.SetAttribute("key", $ProdTypeKey)
				$addElem.SetAttribute("value", $ProdTypeValue)
				$settings.AppendChild($addElem) | Out-Null
				Write_Log "[$file] Inserted key '$ProdTypeKey' = '$ProdTypeValue'" "green"
			}
			$xml.Save($FullPath)
		}
		catch
		{
			Write_Log "[$file] XML update error: $_" "red"
		}
	}
	
	# -----------------------------------------------------------------------------
	# SQL Logic: run the appropriate update/merge
	# -----------------------------------------------------------------------------
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	
	if (-not $ConnectionString -or -not $server -or -not $dbName)
	{
		Write_Log "DB server, DB name, or connection string not found. Cannot execute SQL update." "red"
		Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
		return
	}
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		Import-Module $SqlModuleName -ErrorAction Stop
		$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
	}
	else
	{
		Write_Log "No valid SQL module available for SQL operations!" "red"
		Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
		return
	}
	$supportsConnectionString = $false
	if ($InvokeSqlCmd)
	{
		$supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString'
	}
	
	if ($Mode -eq "SCL")
	{
		$SqlQuery = @"
MERGE INTO SCL_TAB AS Target
USING (
    SELECT $ItemKey
    FROM POS_TAB
    WHERE $ItemKey BETWEEN '$MinItem' AND '$MaxItem'
      AND $POS_Field = $POS_Value
) AS Source
ON Target.$ItemKey = Source.$ItemKey
WHEN MATCHED THEN
    UPDATE SET $SCL_Field = $SCL_Value
WHEN NOT MATCHED BY TARGET THEN
    INSERT ($ItemKey, $SCL_Field) VALUES (Source.$ItemKey, $SCL_Value);
"@
	}
	else
	{
		$SqlQuery = @"
UPDATE SCL_TAB
SET $SCL_Field = $POS_Default_F272_Value
WHERE $ItemKey BETWEEN '$MinItem' AND '$MaxItem'
  AND EXISTS (
        SELECT 1 FROM POS_TAB
        WHERE POS_TAB.$ItemKey = SCL_TAB.$ItemKey
          AND POS_TAB.$POS_Field = $POS_Value
  );
"@
	}
	
	try
	{
		if ($supportsConnectionString)
		{
			& $InvokeSqlCmd -ConnectionString $ConnectionString -Query $SqlQuery -ErrorAction Stop -QueryTimeout 0
		}
		else
		{
			& $InvokeSqlCmd -ServerInstance $server -Database $dbName -Query $SqlQuery -ErrorAction Stop -QueryTimeout 0
		}
		Write_Log "Database updated for $Mode mode." "green"
	}
	catch
	{
		Write_Log "Error executing SQL update for $Mode mode: $_" "red"
	}
	
	Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
}

# ===================================================================================================
#                             FUNCTION: Show_Node_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Shows a node selector dialog for Lanes, Scales, Backoffices, or any combination.
#   Usage:
#      $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
#      $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Scale"
#      $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes @("Lane","Scale","Backoffice")
#      $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Scale" -OnlyBizerbaScales
#   Returns: hashtable with keys matching selected node types (Lanes, Scales, Backoffices)
# ===================================================================================================

function Show_Node_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter()]
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string[]]$NodeTypes = @("Lane", "Scale", "Backoffice"),
		[switch]$OnlyBizerbaScales,
		[string]$Title = "Select Nodes to Process"
	)
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = $Title
	$form.Size = New-Object System.Drawing.Size(430, 450)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$tabs = New-Object System.Windows.Forms.TabControl
	$tabs.Location = New-Object System.Drawing.Point(10, 10)
	$tabs.Size = New-Object System.Drawing.Size(390, 320)
	$form.Controls.Add($tabs)
	
	$tabControls = @{ }
	
	# ----- Lanes Tab -----
	if ("Lane" -in $NodeTypes)
	{
		$tabLanes = New-Object System.Windows.Forms.TabPage
		$tabLanes.Text = "Lanes"
		$clbLanes = New-Object System.Windows.Forms.CheckedListBox
		$clbLanes.Location = New-Object System.Drawing.Point(10, 10)
		$clbLanes.Size = New-Object System.Drawing.Size(350, 270)
		$clbLanes.CheckOnClick = $true
		$tabLanes.Controls.Add($clbLanes)
		$tabs.TabPages.Add($tabLanes)
		$tabControls["Lanes"] = $clbLanes
		
		# Load lane node data from FunctionResults
		$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
		$LaneMachineLabels = $script:FunctionResults['LaneMachineLabels']
		$LaneMachinePath = $script:FunctionResults['LaneMachinePath']
		$LaneMachineToServerPath = $script:FunctionResults['LaneMachineToServerPath']
		$allLanes = @()
		if ($script:FunctionResults.ContainsKey('LaneMachineNames') -and $script:FunctionResults['LaneMachineNames'].Count -gt 0)
		{
			$allLanes = $script:FunctionResults['LaneMachineNames']
		}
		elseif (Test-Path -Path $OfficePath)
		{
			$laneFolders = Get-ChildItem -Path $OfficePath -Directory -Filter "XF${StoreNumber}0*"
			if ($laneFolders)
			{
				$allLanes = $laneFolders | ForEach-Object { $_.Name.Substring($_.Name.Length - 3, 3) }
			}
		}
		$sortedLanes = $allLanes | Sort-Object
		foreach ($laneMachine in $sortedLanes)
		{
			# Find numeric lane (if possible)
			$laneNum = $null
			foreach ($key in $LaneNumToMachineName.Keys)
			{
				if ($LaneNumToMachineName[$key] -eq $laneMachine -and $key -match '^\d{3}$') { $laneNum = $key; break }
			}
			# Compose object for each lane, store all mapping info
			$obj = [PSCustomObject]@{
				LaneNumber  = $laneNum
				MachineName = $laneMachine
				Label	    = $LaneMachineLabels[$laneMachine]
				Path	    = $LaneMachinePath[$laneMachine]
				ServerPath  = $LaneMachineToServerPath[$laneMachine]
				DisplayName = if ($laneNum) { "$laneMachine ($laneNum)" } else { $laneMachine }
			}
			$obj | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.DisplayName } -Force
			$clbLanes.Items.Add($obj) | Out-Null
		}
	}
	
	# ----- Scales Tab -----
	if ("Scale" -in $NodeTypes)
	{
		$tabScales = New-Object System.Windows.Forms.TabPage
		$tabScales.Text = "Scales"
		$clbScales = New-Object System.Windows.Forms.CheckedListBox
		$clbScales.Location = New-Object System.Drawing.Point(10, 10)
		$clbScales.Size = New-Object System.Drawing.Size(350, 270)
		$clbScales.CheckOnClick = $true
		$tabScales.Controls.Add($clbScales)
		$tabs.TabPages.Add($tabScales)
		$tabControls["Scales"] = $clbScales
		
		$allScales = @()
		if ($script:FunctionResults.ContainsKey('ScaleCodeToIPInfo'))
		{
			$allScales = $script:FunctionResults['ScaleCodeToIPInfo'].Values
			if ($OnlyBizerbaScales)
			{
				$allScales = $allScales | Where-Object { $_.ScaleBrand -match 'bizerba' }
			}
		}
		$sortedScales = $allScales | Sort-Object { [int]($_.ScaleCode) }
		
		# --- Deduplicate by FullIP (show only one entry per unique scale IP) ---
		$uniqueScaleIPs = @{ }
		$dedupedScales = @()
		foreach ($scale in $sortedScales)
		{
			# Use .FullIP, fallback to .IPAddress if not present
			$ipKey = if ($scale.FullIP) { $scale.FullIP }
			elseif ($scale.IPNetwork -and $scale.IPDevice) { "$($scale.IPNetwork)$($scale.IPDevice)" }
			else { $null }
			if ($ipKey -and -not $uniqueScaleIPs.ContainsKey($ipKey))
			{
				$uniqueScaleIPs[$ipKey] = $true
				$dedupedScales += $scale
			}
		}
		
		foreach ($scale in $dedupedScales)
		{
			$ip = if ($scale.IPNetwork -and $scale.IPDevice) { "$($scale.IPNetwork)$($scale.IPDevice)" }
			else { "" }
			$displayName = "$($scale.ScaleName) [$ip]"
			$scaleObj = [PSCustomObject]@{
				ScaleCode   = $scale.ScaleCode
				ScaleName   = $scale.ScaleName
				IPAddress   = $ip
				Vendor	    = $scale.Vendor
				FullIP	    = $scale.FullIP
				IPNetwork   = $scale.IPNetwork
				IPDevice    = $scale.IPDevice
				ScaleBrand  = $scale.ScaleBrand
				DisplayName = $displayName
			}
			$scaleObj | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.DisplayName } -Force
			$clbScales.Items.Add($scaleObj) | Out-Null
		}
	}
	
	# ----- Backoffices Tab -----
	if ("Backoffice" -in $NodeTypes)
	{
		$tabBO = New-Object System.Windows.Forms.TabPage
		$tabBO.Text = "Backoffices"
		$clbBO = New-Object System.Windows.Forms.CheckedListBox
		$clbBO.Location = New-Object System.Drawing.Point(10, 10)
		$clbBO.Size = New-Object System.Drawing.Size(350, 270)
		$clbBO.CheckOnClick = $true
		$tabBO.Controls.Add($clbBO)
		$tabs.TabPages.Add($tabBO)
		$tabControls["Backoffices"] = $clbBO
		
		$boDict = if ($script:FunctionResults.ContainsKey('BackofficeNumToMachineName'))
		{
			$script:FunctionResults['BackofficeNumToMachineName']
		}
		else { @{ } }
		$boLabels = $script:FunctionResults['BackofficeNumToLabel']
		$boPaths = $script:FunctionResults['BackofficeNumToPath']
		
		# Deduplicate Backoffices by numeric BONumber only (extract 3-digit number)
		$seenBonumbers = @{ }
		foreach ($boNumKey in $boDict.Keys | Sort-Object)
		{
			# Extract numeric part only (3 digits)
			if ($boNumKey -match '(\d{3})')
			{
				$bonum = $matches[1]
				if (-not $seenBonumbers.ContainsKey($bonum))
				{
					$seenBonumbers[$bonum] = $true
					
					# Compose display name from MachineName & bonum
					$machineName = $boDict[$boNumKey]
					$label = $boLabels[$boNumKey]
					$path = $boPaths[$boNumKey]
					
					$obj = [PSCustomObject]@{
						BONumber    = $bonum
						MachineName = $machineName
						Label	    = $label
						Path	    = $path
						DisplayName = if ($machineName) { "$machineName ($bonum)" } else { "Unknown ($bonum)" }
					}
					$obj | Add-Member ScriptMethod ToString { $this.DisplayName } -Force
					$clbBO.Items.Add($obj) | Out-Null
				}
			}
		}
	}
	
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Location = New-Object System.Drawing.Point(20, 340)
	$btnSelectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.BackColor = [System.Drawing.SystemColors]::Control
	
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Location = New-Object System.Drawing.Point(220, 340)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.BackColor = [System.Drawing.SystemColors]::Control
	
	# Utility: Set button color (no nested function)
	$setBtnColor = {
		param ($btn,
			$state)
		switch ($state)
		{
			1 { $btn.BackColor = [System.Drawing.Color]::Yellow }
			2 { $btn.BackColor = [System.Drawing.Color]::LightGreen }
			Default { $btn.BackColor = [System.Drawing.SystemColors]::Control }
		}
	}
	&$setBtnColor $btnSelectAll 0
	&$setBtnColor $btnDeselectAll 0
	
	# --- SINGLE TAB SELECT ALL LOGIC ---
	if ($tabs.TabPages.Count -eq 1)
	{
		$clb = $tabControls[$tabs.TabPages[0].Text]
		$isAllChecked = {
			for ($i = 0; $i -lt $clb.Items.Count; $i++)
			{
				if (-not $clb.GetItemChecked($i)) { return $false }
			}
			return $clb.Items.Count -gt 0
		}
		$isAnyChecked = {
			for ($i = 0; $i -lt $clb.Items.Count; $i++)
			{
				if ($clb.GetItemChecked($i)) { return $true }
			}
			return $false
		}
		$btnSelectAll.Add_Click({
				$allChecked = & $isAllChecked
				if (-not $allChecked)
				{
					for ($i = 0; $i -lt $clb.Items.Count; $i++)
					{
						$clb.SetItemChecked($i, $true)
					}
					& $setBtnColor $btnSelectAll 2
				}
				# If already all checked, do nothing
			})
		$btnDeselectAll.Add_Click({
				for ($i = 0; $i -lt $clb.Items.Count; $i++)
				{
					$clb.SetItemChecked($i, $false)
				}
				& $setBtnColor $btnSelectAll 0
				& $setBtnColor $btnDeselectAll 0
			})
		foreach ($event in @("Add_ItemCheck", "Add_MouseUp", "Add_KeyUp"))
		{
			$clb.$event.Invoke({
					Start-Sleep -Milliseconds 30
					$allChecked = & $isAllChecked
					$anyChecked = & $isAnyChecked
					if ($allChecked)
					{
						& $setBtnColor $btnSelectAll 2
					}
					elseif ($anyChecked)
					{
						& $setBtnColor $btnSelectAll 1
					}
					else
					{
						& $setBtnColor $btnSelectAll 0
					}
				})
		}
	}
	else
	{
		# ---- NORMAL MULTI-TAB LOGIC ----
		$tabSelectState = @{ }
		$lastSelectTabIndex = $null
		$selectAllYellowTabIndex = $null
		
		$btnSelectAll.Add_Click({
				$tabName = $tabs.SelectedTab.Text
				$clb = $tabControls[$tabName]
				$tabIndex = $tabs.SelectedIndex
				$currentTabAllChecked = $true
				for ($i = 0; $i -lt $clb.Items.Count; $i++)
				{
					if (-not $clb.GetItemChecked($i)) { $currentTabAllChecked = $false; break }
				}
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++)
					{
						if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break }
					}
				}
				if ($allTabsChecked)
				{
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++)
						{
							if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break }
						}
						if ($tabIndex2 -eq -1) { continue }
						$tabSelectState[$tabIndex2] = 0
					}
					$tabSelectState[$tabIndex] = 0
					$selectAllYellowTabIndex = $null
					&$setBtnColor $btnSelectAll 0
					return
				}
				if ($currentTabAllChecked -and -not $allTabsChecked)
				{
					foreach ($k in $tabControls.Keys)
					{
						$list = $tabControls[$k]
						for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) }
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++)
						{
							if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break }
						}
						if ($tabIndex2 -eq -1) { continue }
						$tabSelectState[$tabIndex2] = 2
					}
					&$setBtnColor $btnSelectAll 2
					$selectAllYellowTabIndex = $null
					return
				}
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
				$tabSelectState[$tabIndex] = 1
				$lastSelectTabIndex = $tabIndex
				$selectAllYellowTabIndex = $tabIndex
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++)
					{
						if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break }
					}
				}
				if ($allTabsChecked)
				{
					&$setBtnColor $btnSelectAll 2
					$selectAllYellowTabIndex = $null
				}
				else
				{
					&$setBtnColor $btnSelectAll 1
				}
			})
		$btnDeselectAll.Add_Click({
				$tabName = $tabs.SelectedTab.Text
				$clb = $tabControls[$tabName]
				$tabIndex = $tabs.SelectedIndex
				$noneChecked = $true
				for ($i = 0; $i -lt $clb.Items.Count; $i++)
				{
					if ($clb.GetItemChecked($i)) { $noneChecked = $false; break }
				}
				if ($noneChecked)
				{
					$originalTab = $tabs.SelectedTab
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++)
						{
							if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break }
						}
						if ($tabIndex2 -eq -1) { continue }
						$tabs.SelectedTab = $tabs.TabPages[$tabIndex2]
						$list = $tabControls[$k]
						for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) }
						$tabSelectState[$tabIndex2] = 0
						$list.Refresh()
					}
					$tabs.SelectedTab = $originalTab
				}
				else
				{
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
					$tabSelectState[$tabIndex] = 0
				}
				&$setBtnColor $btnDeselectAll 0
				&$setBtnColor $btnSelectAll 0
				$selectAllYellowTabIndex = $null
			})
		foreach ($clb in $tabControls.Values)
		{
			$clb.Add_ItemCheck({
					Start-Sleep -Milliseconds 50
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++)
						{
							if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break }
						}
					}
					if ($allTabsChecked)
					{
						&$setBtnColor $btnSelectAll 2
						$selectAllYellowTabIndex = $null
					}
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++)
						{
							if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break }
						}
						if ($allChecked -and $clbLocal.Items.Count -gt 0)
						{
							&$setBtnColor $btnSelectAll 1
							$selectAllYellowTabIndex = $tabIndex
						}
						else
						{
							&$setBtnColor $btnSelectAll 0
							$selectAllYellowTabIndex = $null
						}
					}
				})
			$clb.Add_MouseUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++)
						{
							if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break }
						}
					}
					if ($allTabsChecked)
					{
						&$setBtnColor $btnSelectAll 2
						$selectAllYellowTabIndex = $null
					}
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++)
						{
							if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break }
						}
						if ($allChecked -and $clbLocal.Items.Count -gt 0)
						{
							&$setBtnColor $btnSelectAll 1
							$selectAllYellowTabIndex = $tabIndex
						}
						else
						{
							&$setBtnColor $btnSelectAll 0
							$selectAllYellowTabIndex = $null
						}
					}
				})
			$clb.Add_KeyUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++)
						{
							if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break }
						}
					}
					if ($allTabsChecked)
					{
						&$setBtnColor $btnSelectAll 2
						$selectAllYellowTabIndex = $null
					}
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++)
						{
							if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break }
						}
						if ($allChecked -and $clbLocal.Items.Count -gt 0)
						{
							&$setBtnColor $btnSelectAll 1
							$selectAllYellowTabIndex = $tabIndex
						}
						else
						{
							&$setBtnColor $btnSelectAll 0
							$selectAllYellowTabIndex = $null
						}
					}
				})
		}
		$tabs.add_SelectedIndexChanged({
				$tabIndex = $tabs.SelectedIndex
				$clb = $tabControls[$tabs.SelectedTab.Text]
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++)
					{
						if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break }
					}
				}
				if ($allTabsChecked)
				{
					&$setBtnColor $btnSelectAll 2
					$selectAllYellowTabIndex = $null
				}
				else
				{
					$allChecked = $true
					for ($i = 0; $i -lt $clb.Items.Count; $i++)
					{
						if (-not $clb.GetItemChecked($i)) { $allChecked = $false; break }
					}
					if ($allChecked -and $clb.Items.Count -gt 0)
					{
						&$setBtnColor $btnSelectAll 1
						$selectAllYellowTabIndex = $tabIndex
					}
					else
					{
						&$setBtnColor $btnSelectAll 0
						$selectAllYellowTabIndex = $null
					}
				}
				&$setBtnColor $btnDeselectAll 0
			})
	}
	
	$form.Controls.Add($btnSelectAll)
	$form.Controls.Add($btnDeselectAll)
	
	# OK Button
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(60, 380)
	$btnOK.Size = New-Object System.Drawing.Size(140, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	# Cancel Button
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(220, 380)
	$btnCancel.Size = New-Object System.Drawing.Size(140, 32)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	# ----- Show form & collect selections -----
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
	
	$result = @{ }
	if ("Lane" -in $NodeTypes)
	{
		$clb = $tabControls["Lanes"]
		$selectedLaneNums = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['LaneNumber'] -and $item.LaneNumber)
				{
					$selectedLaneNums += $item.LaneNumber
				}
			}
		}
		$result.Lanes = $selectedLaneNums
	}
	if ("Scale" -in $NodeTypes)
	{
		$clb = $tabControls["Scales"]
		$selectedScaleCodes = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['ScaleCode'] -and $item.ScaleCode)
				{
					$selectedScaleCodes += "$($item.ScaleCode)"
				}
			}
		}
		$result.Scales = $selectedScaleCodes
	}
	if ("Backoffice" -in $NodeTypes)
	{
		$clb = $tabControls["Backoffices"]
		$selectedBONums = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['BONumber'] -and $item.BONumber)
				{
					$selectedBONums += $item.BONumber
				}
			}
		}
		$result.Backoffices = $selectedBONums
	}
	
	# Basic validation
	if ($result.Keys.Count -eq 0 -or ($result.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum -eq 0)
	{
		[System.Windows.Forms.MessageBox]::Show("No nodes selected.", "Information",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		return $null
	}
	return $result
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
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Tables to Process"
	$form.Size = New-Object System.Drawing.Size(450, 570)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Please select the tables you want to pump:"
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$label.AutoSize = $true
	$form.Controls.Add($label)
	
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 40)
	$checkedListBox.Size = New-Object System.Drawing.Size(410, 400)
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)
	
	$distinctTables = $AliasResults | Select-Object -ExpandProperty Table -Unique | Sort-Object
	foreach ($tableName in $distinctTables)
	{
		[void]$checkedListBox.Items.Add($tableName, $false)
	}
	
	# Buttons styled/positioned like Node form (Y=460 and 500 for a 550-height window)
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Location = New-Object System.Drawing.Point(20, 460)
	$btnSelectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnSelectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnSelectAll)
	
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Location = New-Object System.Drawing.Point(220, 460)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnDeselectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnDeselectAll)
	
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(60, 500)
	$btnOK.Size = New-Object System.Drawing.Size(140, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(220, 500)
	$btnCancel.Size = New-Object System.Drawing.Size(140, 32)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	$setBtnColor = {
		param ($btn,
			$state)
		switch ($state)
		{
			1 { $btn.BackColor = [System.Drawing.Color]::Yellow }
			2 { $btn.BackColor = [System.Drawing.Color]::LightGreen }
			Default { $btn.BackColor = [System.Drawing.SystemColors]::Control }
		}
	}
	
	$isAllChecked = {
		for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
		{
			if (-not $checkedListBox.GetItemChecked($i)) { return $false }
		}
		return $checkedListBox.Items.Count -gt 0
	}
	$isAnyChecked = {
		for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
		{
			if ($checkedListBox.GetItemChecked($i)) { return $true }
		}
		return $false
	}
	
	$btnSelectAll.Add_Click({
			$allChecked = & $isAllChecked
			if (-not $allChecked)
			{
				for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
				{
					$checkedListBox.SetItemChecked($i, $true)
				}
				& $setBtnColor $btnSelectAll 2
			}
		})
	$btnDeselectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
			& $setBtnColor $btnSelectAll 0
		})
	
	$checkedListBox.Add_ItemCheck({
			Start-Sleep -Milliseconds 30
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	$checkedListBox.Add_MouseUp({
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	$checkedListBox.Add_KeyUp({
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
	{
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
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select SQL Sections"
	$form.StartPosition = "CenterScreen"
	$form.Size = New-Object System.Drawing.Size(550, 440)
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Check the sections you want to run, then click OK."
	$label.AutoSize = $true
	$label.Left = 20
	$label.Top = 10
	$form.Controls.Add($label)
	
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Width = 500
	$checkedListBox.Height = 280
	$checkedListBox.Left = 20
	$checkedListBox.Top = 40
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)
	
	foreach ($name in $SectionNames)
	{
		[void]$checkedListBox.Items.Add($name, $false)
	}
	
	# Place Select All and Deselect All like the Node dialog (Y=310)
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Location = New-Object System.Drawing.Point(75, 320)
	$btnSelectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnSelectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnSelectAll)
	
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Location = New-Object System.Drawing.Point(275, 320)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnDeselectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnDeselectAll)
	
	# Place OK/Cancel like Node dialog (Y=350)
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(115, 360)
	$btnOK.Size = New-Object System.Drawing.Size(140, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(275, 360)
	$btnCancel.Size = New-Object System.Drawing.Size(140, 32)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	$setBtnColor = {
		param ($btn,
			$state)
		switch ($state)
		{
			1 { $btn.BackColor = [System.Drawing.Color]::Yellow }
			2 { $btn.BackColor = [System.Drawing.Color]::LightGreen }
			Default { $btn.BackColor = [System.Drawing.SystemColors]::Control }
		}
	}
	
	$isAllChecked = {
		for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
		{
			if (-not $checkedListBox.GetItemChecked($i)) { return $false }
		}
		return $checkedListBox.Items.Count -gt 0
	}
	$isAnyChecked = {
		for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
		{
			if ($checkedListBox.GetItemChecked($i)) { return $true }
		}
		return $false
	}
	
	$btnSelectAll.Add_Click({
			$allChecked = & $isAllChecked
			if (-not $allChecked)
			{
				for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
				{
					$checkedListBox.SetItemChecked($i, $true)
				}
				& $setBtnColor $btnSelectAll 2
			}
		})
	$btnDeselectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
			& $setBtnColor $btnSelectAll 0
		})
	
	$checkedListBox.Add_ItemCheck({
			Start-Sleep -Milliseconds 30
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	$checkedListBox.Add_MouseUp({
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	$checkedListBox.Add_KeyUp({
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
	{
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
				# Existing cleanup...
				if ($global:ProtocolFormTimer) { $global:ProtocolFormTimer.Stop(); $global:ProtocolFormTimer.Dispose() }
				foreach ($job in $script:LaneProtocolJobs.Values)
				{
					try { Stop-Job $job -Force }
					catch { }
				}
				$script:LaneProtocolJobs.Clear()
				Write_Log "Form is closing. Performing cleanup." "green"
				Delete_Files -Path "$TempDir" -SpecifiedFiles "*.sqi", "*.sql"
			}
		})
	
	# Create a Clear Log button
	$clearLogButton = New-Object System.Windows.Forms.Button
	$clearLogButton.Text = "Clear Log"
	$clearLogButton.Location = New-Object System.Drawing.Point(950, 70)
	$clearLogButton.Size = New-Object System.Drawing.Size(39, 34)
	$clearLogButton.add_Click({
			$logBox.Clear()
			Write_Log "Log Cleared"
		})
	$form.Controls.Add($clearLogButton)
	$toolTip.SetToolTip($clearLogButton, "Clears the log display area.")
	
	# Protocol Table Popup Form and Timer	
	$rowHeight = 19
	$rowCount = 25
	$gridHeight = ($rowCount * $rowHeight) + 28
	
	if (-not $global:ProtocolForm)
	{
		$global:ProtocolForm = New-Object System.Windows.Forms.Form
		$global:ProtocolForm.Text = "Lane PS"
		$global:ProtocolForm.Size = New-Object System.Drawing.Size(257, 500)
		$global:ProtocolForm.StartPosition = "CenterScreen"
		$global:ProtocolForm.Topmost = $true
		
		$global:ProtocolGrid = New-Object System.Windows.Forms.DataGridView
		$global:ProtocolGrid.Location = New-Object System.Drawing.Point(10, 10)
		$global:ProtocolGrid.Size = New-Object System.Drawing.Size(222, 400)
		$global:ProtocolGrid.ColumnCount = 2
		$global:ProtocolGrid.Columns[0].Name = "Lane"
		$global:ProtocolGrid.Columns[1].Name = "Protocol"
		$global:ProtocolGrid.ReadOnly = $true
		$global:ProtocolGrid.RowHeadersVisible = $false
		$global:ProtocolGrid.AllowUserToAddRows = $false
		$global:ProtocolGrid.AllowUserToDeleteRows = $false
		$global:ProtocolGrid.AllowUserToResizeRows = $false
		$global:ProtocolGrid.AllowUserToResizeColumns = $false
		$global:ProtocolGrid.SelectionMode = "FullRowSelect"
		$global:ProtocolGrid.Font = New-Object System.Drawing.Font("Consolas", 10)
		$global:ProtocolForm.Controls.Add($global:ProtocolGrid)
		
		$closeBtn = New-Object System.Windows.Forms.Button
		$closeBtn.Text = "Hide"
		$closeBtn.Location = New-Object System.Drawing.Point(60, 420)
		$closeBtn.Size = New-Object System.Drawing.Size(120, 30)
		$closeBtn.Add_Click({ $global:ProtocolForm.Hide() })
		$global:ProtocolForm.Controls.Add($closeBtn)
		$global:ProtocolForm.add_FormClosing({
				$_.Cancel = $true
				$global:ProtocolForm.Hide()
			})
	}
	if (-not $global:ProtocolFormTimer)
	{
		$global:ProtocolFormTimer = New-Object System.Windows.Forms.Timer
		$global:ProtocolFormTimer.Interval = 500
		$global:ProtocolFormTimer.add_Tick({
				$global:ProtocolGrid.Rows.Clear()
				$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
				
				# Before populating, store old scroll and selection index if available
				$scrollIndex = 0
				if ($global:ProtocolGrid.Rows.Count -gt 0)
				{
					$scrollIndex = $global:ProtocolGrid.FirstDisplayedScrollingRowIndex
				}
				$selIndex = $null
				if ($global:ProtocolGrid.SelectedRows.Count -gt 0)
				{
					$selIndex = $global:ProtocolGrid.SelectedRows[0].Index
				}
				
				$script:ProtocolResults | Sort-Object Lane | ForEach-Object {
					$row = $global:ProtocolGrid.Rows.Add()
					$machineName = $LaneNumToMachineName[$_.Lane]
					if (-not $machineName) { $machineName = $_.Lane } # Fallback if mapping is missing
					$global:ProtocolGrid.Rows[$row].Cells[0].Value = $machineName
					$global:ProtocolGrid.Rows[$row].Cells[1].Value = $_.Protocol
				}
				
				# --- Auto-resize Protocol column based on scroll bar presence ---
				$visibleRowCount = [math]::Floor($global:ProtocolGrid.DisplayRectangle.Height / $global:ProtocolGrid.RowTemplate.Height)
				$scrollBarVisible = $global:ProtocolGrid.Rows.Count -gt $visibleRowCount
				
				# Set Lane column width (fixed)
				$global:ProtocolGrid.Columns[0].Width = 60
				
				# Set Protocol column width (auto)
				if ($scrollBarVisible)
				{
					$global:ProtocolGrid.Columns[1].Width = $global:ProtocolGrid.Width - 60 - 4 - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
				}
				else
				{
					$global:ProtocolGrid.Columns[1].Width = $global:ProtocolGrid.Width - 60 - 4
				}
				
				# Restore scroll and selection safely
				$rowCount = $global:ProtocolGrid.Rows.Count
				if ($rowCount -gt 0)
				{
					if ($scrollIndex -lt 0) { $scrollIndex = 0 }
					if ($scrollIndex -ge $rowCount) { $scrollIndex = $rowCount - 1 }
					$global:ProtocolGrid.FirstDisplayedScrollingRowIndex = $scrollIndex
				}
				if ($selIndex -ne $null -and $global:ProtocolGrid.Rows.Count -gt $selIndex -and $selIndex -ge 0)
				{
					$global:ProtocolGrid.Rows[$selIndex].Selected = $true
				}
			})
		$global:ProtocolFormTimer.Start()
	}
	
	# Show/Hide Protocol Table Button
	$showProtocolBtn = New-Object System.Windows.Forms.Button
	$showProtocolBtn.Text = "Protocols"
	$showProtocolBtn.Size = New-Object System.Drawing.Size(40, 34)
	$showProtocolBtn.Location = New-Object System.Drawing.Point(950, 105) # top-right, next to Clear Log
	$showProtocolBtn.Add_Click({
			$global:ProtocolForm.Show()
			$global:ProtocolForm.BringToFront()
		})
	$form.Controls.Add($showProtocolBtn)
	$toolTip.SetToolTip($showProtocolBtn, "Show/hide the live lane protocol status table.")
	
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
	# Server Tools Button
	#
	######################################################################################################################
	
	############################################################################
	# Server Tools Anchor Button
	############################################################################
	$ServerToolsButton = New-Object System.Windows.Forms.Button
	$ServerToolsButton.Text = "Server Tools"
	$ServerToolsButton.Location = New-Object System.Drawing.Point(50, 475)
	$ServerToolsButton.Size = New-Object System.Drawing.Size(200, 50)
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
	# 4) Manage SQL 'sa' Account Menu Item
	############################################################################
	$ManageSaAccountItem = New-Object System.Windows.Forms.ToolStripMenuItem("Manage SQL 'sa' Account")
	$ManageSaAccountItem.ToolTipText = "Enable or disable the 'sa' account on the local SQL Server with a predefined password."
	$ManageSaAccountItem.Add_Click({
			Manage_Sa_Account
		})
	[void]$ContextMenuServer.Items.Add($ManageSaAccountItem)
	
	############################################################################
	# 5) Repair Windows Menu Item
	############################################################################
	$RepairWindowsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair Windows")
	$RepairWindowsItem.ToolTipText = "Perform repairs on the Windows operating system."
	$RepairWindowsItem.Add_Click({
			Repair_Windows
		})
	[void]$ContextMenuServer.Items.Add($RepairWindowsItem)
	
	############################################################################
	# 6) Configure System Settings Menu Item
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
	$LaneToolsButton.Location = New-Object System.Drawing.Point(275, 475)
	$LaneToolsButton.Size = New-Object System.Drawing.Size(200, 50)
	$ContextMenuLane = New-Object System.Windows.Forms.ContextMenuStrip
	$ContextMenuLane.ShowItemToolTips = $true
	
	############################################################################
	# 1) Lane DB Maintenance Button
	############################################################################
	$LaneDBMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Lane DB Maintenance")
	$LaneDBMaintenanceItem.ToolTipText = "Runs maintenance at the lane(s) databases for the selected lane(s)."
	$LaneDBMaintenanceItem.Add_Click({
			Process_Lanes -StoreNumber $StoreNumber
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
			Ping_All_Nodes -NodeType "Lane" -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($PingLanesItem)
	
	############################################################################
	# 7) Open Lane C$ Share(s)
	############################################################################
	$OpenLaneCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Lane C$ Share(s)")
	$OpenLaneCShareItem.ToolTipText = "Select lanes and open their administrative C$ shares in Explorer."
	$OpenLaneCShareItem.Add_Click({
			Open_Selected_Node_C_Path -StoreNumber $StoreNumber -NodeTypes Lane
		})
	[void]$ContextMenuLane.Items.Add($OpenLaneCShareItem)
	
	############################################################################
	# 8) Delete DBS Menu Item
	############################################################################
	$DeleteDBSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Delete DBS")
	$DeleteDBSItem.ToolTipText = "Delete the DBS files (*.txt, *.dwr, if selected *.sus as well) at the lane."
	$DeleteDBSItem.Add_Click({
			Delete_DBS -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($DeleteDBSItem)
	
	############################################################################
	# 9) Refresh PIN Pad Files Menu Item
	############################################################################
	$RefreshPinPadFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh PIN Pad Files")
	$RefreshPinPadFilesItem.ToolTipText = "Refresh the PIN pad files for the lane/s."
	$RefreshPinPadFilesItem.Add_Click({
			Refresh_PIN_Pad_Files -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($RefreshPinPadFilesItem)
	
	############################################################################
	# 10) Drawer Control Item
	############################################################################
	$DrawerControlItem = New-Object System.Windows.Forms.ToolStripMenuItem("Drawer Control")
	$DrawerControlItem.ToolTipText = "Set the Drawer Control for a lane for testing"
	$DrawerControlItem.Add_Click({
			Drawer_Control -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($DrawerControlItem)
	
	############################################################################
	# 11) Refresh Database
	############################################################################
	$RefreshDatabaseItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Database")
	$RefreshDatabaseItem.ToolTipText = "Refresh the database at the lane/s"
	$RefreshDatabaseItem.Add_Click({
			Refresh_Database -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($RefreshDatabaseItem)
	
	############################################################################
	# 12) Send Restart Command Menu Item
	############################################################################
	$SendRestartCommandItem = New-Object System.Windows.Forms.ToolStripMenuItem("Send Restart All Programs")
	$SendRestartCommandItem.ToolTipText = "Send restart all programs to selected lane(s) for the store."
	$SendRestartCommandItem.Add_Click({
			Send_Restart_All_Programs -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($SendRestartCommandItem)
	
	############################################################################
	# 13) Enable SQL Protocols Menu Item
	############################################################################
	$EnableSQLProtocolsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Enable SQL Protocols")
	$EnableSQLProtocolsItem.ToolTipText = "Enable TCP/IP, Named Pipes, and set static port for SQL Server on selected lane(s)."
	$EnableSQLProtocolsItem.Add_Click({
			Enable_SQL_Protocols_On_Selected_Lanes -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($EnableSQLProtocolsItem)
	
	############################################################################
	# 14) Set the time on the lanes
	############################################################################
	$SetLaneTimeFromLocalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Set/Schedule Time on Lanes")
	$SetLaneTimeFromLocalItem.ToolTipText = "Synchronize or schedule time sync for selected lanes."
	$SetLaneTimeFromLocalItem.Add_Click({
			# Prompt for mode: one-time or schedule
			$modeResult = [System.Windows.Forms.MessageBox]::Show(
				"Do you want to schedule recurring sync? (Yes for schedule, No for one-time)",
				"Choose Mode",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)
			$isSchedule = ($modeResult -eq [System.Windows.Forms.DialogResult]::Yes)
			Send_SERVER_time_to_Lanes -StoreNumber $StoreNumber -Schedule:$isSchedule
		})
	[void]$ContextMenuLane.Items.Add($SetLaneTimeFromLocalItem)
	
	############################################################################
	# 15) Reboot Lane Menu Item
	############################################################################
	$RebootLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Lane")
	$RebootLaneItem.ToolTipText = "Reboot the selected lane/s."
	$RebootLaneItem.Add_Click({
			Reboot_Nodes -StoreNumber "$StoreNumber" -NodeTypes Lane
		})
	[void]$ContextMenuLane.Items.Add($RebootLaneItem)
	
	############################################################################
	# Show the context menu when the Server Tools button is clicked
	############################################################################
	$LaneToolsButton.Add_Click({
			$ContextMenuLane.Show($LaneToolsButton, 0, $LaneToolsButton.Height)
		})
	$toolTip.SetToolTip($LaneToolsButton, "Click to see Lane-related tools.")
	$form.Controls.Add($LaneToolsButton)
	
	######################################################################################################################
	# 
	# Scales Tools Button
	#
	######################################################################################################################
	
	############################################################################
	# Scales Tools Anchor Button
	############################################################################
	$ScaleToolsButton = New-Object System.Windows.Forms.Button
	$ScaleToolsButton.Text = "Scale Tools"
	$ScaleToolsButton.Location = New-Object System.Drawing.Point(525, 475)
	$ScaleToolsButton.Size = New-Object System.Drawing.Size(200, 50)
	$ContextMenuScale = New-Object System.Windows.Forms.ContextMenuStrip
	$ContextMenuScale.ShowItemToolTips = $true
	
	############################################################################
	# 1) Ping Scales Menu Item
	############################################################################
	$PingScalesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Ping Scales")
	$PingScalesItem.ToolTipText = "Ping all scale devices to check connectivity."
	$PingScalesItem.Add_Click({
			Ping_All_Nodes -NodeType "Scale" -StoreNumber $StoreNumber
		})
	[void]$ContextMenuScale.Items.Add($PingScalesItem)
	
	############################################################################
	# 2) Repair BMS Service
	############################################################################
	$repairBMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair BMS Service")
	$repairBMSItem.ToolTipText = "Repairs the BMS service for scale deployment."
	$repairBMSItem.Add_Click({
			Repair_BMS
		})
	[void]$ContextMenuScale.Items.Add($repairBMSItem)
	
	############################################################################
	# 3) Reboot Scales
	############################################################################
	$Reboot_ScalesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Scales")
	$Reboot_ScalesItem.ToolTipText = "Reboot Scale/s."
	$Reboot_ScalesItem.Add_Click({
			Reboot_Nodes -StoreNumber "$StoreNumber" -NodeTypes Scale
		})
	[void]$ContextMenuScale.Items.Add($Reboot_ScalesItem)
	
	############################################################################
	# 4) Open Scale C$ Share(s)
	############################################################################
	$OpenScaleCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Scale C$ Share(s)")
	$OpenScaleCShareItem.ToolTipText = "Select scales and open their C$ administrative shares as 'bizuser' (bizerba/biyerba)."
	$OpenScaleCShareItem.Add_Click({
			Open_Selected_Node_C_Path -StoreNumber $StoreNumber -NodeTypes Scale
		})
	[void]$ContextMenuScale.Items.Add($OpenScaleCShareItem)
	
	############################################################################
	# 5) Update Scales Specials
	############################################################################
	$UpdateScalesSpecialsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Scales Specials")
	$UpdateScalesSpecialsItem.ToolTipText = "Update scale specials immediately or schedule as a daily 5AM task."
	$UpdateScalesSpecialsItem.Add_Click({
			Update_Scales_Specials_Interactive
		})
	[void]$ContextMenuScale.Items.Add($UpdateScalesSpecialsItem)
	
	############################################################################
	# 6) Update Scale Config and DB (F272 Upsert)
	############################################################################
	$UpdateScaleConfigAndDBItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Scale Config && DB (F272 Upsert)")
	$UpdateScaleConfigAndDBItem.ToolTipText = "Updates ScaleCommApp configs and upserts F272 in SCL_TAB for POS_TAB F82=1 in item range."
	$UpdateScaleConfigAndDBItem.Add_Click({
			Update_ScaleConfig_And_DB
		})
	[void]$ContextMenuScale.Items.Add($UpdateScaleConfigAndDBItem)
	
	############################################################################
	# 7) Schedule Duplicate File Monitor
	############################################################################
	$ScheduleRemoveDupesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Remove duplicate files from (toBizerba)")
	$ScheduleRemoveDupesItem.ToolTipText = "Monitor for and auto-delete duplicate files in (toBizerba). Run now or schedule as SYSTEM."
	$ScheduleRemoveDupesItem.Add_Click({
			Remove_Duplicate_Files_From_toBizerba
		})
	[void]$ContextMenuScale.Items.Add($ScheduleRemoveDupesItem)
	
	############################################################################
	# Show the context menu when the Server Tools button is clicked
	############################################################################
	$ScaleToolsButton.Add_Click({
			$ContextMenuScale.Show($ScaleToolsButton, 0, $ScaleToolsButton.Height)
		})
	$toolTip.SetToolTip($ScaleToolsButton, "Click to see Scale-related tools.")
	$form.Controls.Add($ScaleToolsButton)
	
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
	$GeneralToolsButton.Location = New-Object System.Drawing.Point(750, 475)
	$GeneralToolsButton.Size = New-Object System.Drawing.Size(200, 50)
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
	# 7) Ping Backoffices Menu Item
	############################################################################
	$PingBackofficesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Ping Backoffices")
	$PingBackofficesItem.ToolTipText = "Ping all backoffice devices to check connectivity."
	$PingBackofficesItem.Add_Click({
			Ping_All_Nodes -NodeType "Backoffice" -StoreNumber $StoreNumber
		})
	[void]$contextMenuGeneral.Items.Add($PingBackofficesItem)
	
	############################################################################
	# 8) Reboot selected Backoffices
	############################################################################
	$RebootBackofficesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Backoffices")
	$RebootBackofficesItem.ToolTipText = "Reboot the selected Backoffice/s."
	$RebootBackofficesItem.Add_Click({
			Reboot_Nodes -StoreNumber $StoreNumber -NodeTypes Backoffice
		})
	[void]$contextMenuGeneral.Items.Add($RebootBackofficesItem)
	
	############################################################################
	# 9) Export All VNC Files
	############################################################################
	$ExportVNCFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Export All VNC Files")
	$ExportVNCFilesItem.ToolTipText = "Generate UltraVNC (.vnc) connection files for all lanes, scales, and backoffices."
	$ExportVNCFilesItem.Add_Click({
			Export_VNC_Files_For_All_Nodes `
										   -LaneNumToMachineName $script:FunctionResults['LaneNumToMachineName'] `
										   -ScaleCodeToIPInfo $script:FunctionResults['ScaleCodeToIPInfo'] `
										   -BackofficeNumToMachineName $script:FunctionResults['BackofficeNumToMachineName']`
										   -AllVNCPasswords $script:FunctionResults['AllVNCPasswords']
		})
	[void]$contextMenuGeneral.Items.Add($ExportVNCFilesItem)
	
	############################################################################
	# 10) Export Machines Hardware Info
	############################################################################
	$ExportMachineHardwareInfoItem = New-Object System.Windows.Forms.ToolStripMenuItem("Export Machines Hardware Info")
	$ExportMachineHardwareInfoItem.ToolTipText = "Collect and export manufacturer/model for all machines"
	$ExportMachineHardwareInfoItem.Add_Click({
			$didExport = Get_Remote_Machine_Info
			if ($didExport)
			{
				[System.Windows.Forms.MessageBox]::Show("Machine hardware info exported", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
		})
	[void]$contextMenuGeneral.Items.Add($ExportMachineHardwareInfoItem)
	
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
	# 12) Insert Test Item
	############################################################################
	$InsertTestItem = New-Object System.Windows.Forms.ToolStripMenuItem("Insert Test Item")
	$InsertTestItem.ToolTipText = "Inserts or updates a test item (PLU 0020077700000 or alternatives) in the database."
	$InsertTestItem.Add_Click({
			Insert_Test_Item
		})
	[void]$ContextMenuGeneral.Items.Add($InsertTestItem)
	
	############################################################################
	# 13) Fix Deploy CHG
	############################################################################
	$FixDeployCHGItem = New-Object System.Windows.Forms.ToolStripMenuItem("Fix Deploy_CHG")
	$FixDeployCHGItem.ToolTipText = "Restores the deploy line to DEPLOY_CHG.sql for scale management."
	$FixDeployCHGItem.Add_Click({
			Fix_Deploy_CHG
		})
	[void]$ContextMenuGeneral.Items.Add($FixDeployCHGItem)
	
	############################################################################
	# Show the context menu when the General Tools button is clicked
	############################################################################
	$GeneralToolsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
	$GeneralToolsButton.Add_Click({
			$contextMenuGeneral.Show($GeneralToolsButton, 0, $GeneralToolsButton.Height)
		})
	$toolTip.SetToolTip($GeneralToolsButton, "Click to see some tools created for SMS.")
	$form.Controls.Add($GeneralToolsButton)
}


######################################################################################################################
# 
# Anchor all controls for resize (PowerShell WinForms)
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
$ScaleToolsButton.Anchor = 'Bottom'

$form.add_Resize({
		# Margin between logBox and Clear Log button
		$buttonMargin = 10
		
		# Position Clear Log button in top-right
		$clearLogButton.Left = $form.ClientSize.Width - $clearLogButton.Width - 15
		$clearLogButton.Top = 70 # or wherever you want it (70 is your original)
		
		# Calculate the rightmost edge the logBox should go to
		$logBoxRightEdge = $clearLogButton.Left - $buttonMargin
		
		# Make logBox fill to just before the Clear Log button
		$logBox.Left = 50
		$logBox.Top = 70
		$logBox.Width = [math]::Max(100, $logBoxRightEdge - $logBox.Left)
		$logBox.Height = $form.ClientSize.Height - 170
		
		# Center store name label
		$storeNameLabel.Left = [math]::Max(0, ($form.ClientSize.Width - $storeNameLabel.Width) / 2)
		$NodesStore.Left = [math]::Max(0, ($form.ClientSize.Width - $NodesStore.Width) / 2)
		
		# Space the bottom buttons evenly
		$buttonWidth = 200
		$buttonHeight = 50
		$numButtons = 4
		$availableWidth = $form.ClientSize.Width
		$gap = [math]::Max(10, ($availableWidth - ($numButtons * $buttonWidth)) / ($numButtons + 1))
		$ServerToolsButton.Left = $gap
		$LaneToolsButton.Left = $ServerToolsButton.Left + $buttonWidth + $gap
		$ScaleToolsButton.Left = $LaneToolsButton.Left + $buttonWidth + $gap
		$GeneralToolsButton.Left = $ScaleToolsButton.Left + $buttonWidth + $gap
		$ServerToolsButton.Top = $LaneToolsButton.Top = $ScaleToolsButton.Top = $GeneralToolsButton.Top = $form.ClientSize.Height - ($buttonHeight + $buttonHeight)
	})

# ===================================================================================================
#                                       SECTION: Main Script Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Orchestrates the execution flow of the script, initializing variables, processing items, and handling user interactions.
# ===================================================================================================

# Check for the precense of PsExec for later use
# $GetPsExec = Get_PsExec

# Get SQL Connection String
Get_Store_And_Database_Info -WinIniPath $WinIniPath -SmsStartIniPath $SmsStartIniPath -StartupIniPath $StartupIniPath -SystemIniPath $SystemIniPath
$StoreNumber = $script:FunctionResults['StoreNumber']
$StoreName = $script:FunctionResults['StoreName']
$SqlModuleName = $script:FunctionResults['SqlModuleName']

# Count Nodes based on mode
$Nodes = Retrieve_Nodes -StoreNumber $StoreNumber
$Nodes = $script:FunctionResults['Nodes']

# Retrieve the list of machine names from the FunctionResults dictionary
$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']

# Get the SQL connection string for all machines
Get_All_Lanes_Database_Info | Out-Null

# Start per-lane jobs for protocol checks (PS5-compatible parallelism via multiple jobs)
$script:LaneProtocolJobs = @{ }
$script:LaneProtocols = @{ }
$script:ProtocolResults = @()

foreach ($laneNum in $LaneNumToMachineName.Keys | Where-Object { $_ -match '^\d{3}$' } | Sort-Object)
{
	$machine = $LaneNumToMachineName[$laneNum]
	# Launch job with ($machine, $laneNum, ...) and use $machine for connections, $laneNum for reference
	$script:LaneProtocolJobs[$laneNum] = Start-Job -ArgumentList $machine, $laneNum, $SqlModuleName -ScriptBlock {
		param ($machine,
			$laneNum,
			$SqlModuleName)
		
		$protocol = "File"
		$tcpConn = "Server=$machine;Database=master;Integrated Security=True;Network Library=DBMSSOCN"
		$npConn = "Server=$machine;Database=master;Integrated Security=True;Network Library=dbnmpntw"
		
		# Step 1: Fast TCP port check
		$tcpPortOpen = $false
		try
		{
			$tcpClient = New-Object System.Net.Sockets.TcpClient
			$connectTask = $tcpClient.ConnectAsync($machine, 1433)
			if ($connectTask.Wait(500) -and $tcpClient.Connected)
			{
				$tcpPortOpen = $true
				$tcpClient.Close()
			}
		}
		catch { }
		
		# Step 2: If TCP port is open, try SQL query over TCP
		if ($tcpPortOpen)
		{
			try
			{
				Import-Module $SqlModuleName -ErrorAction Stop
				Invoke-Sqlcmd -ConnectionString $tcpConn -Query "SELECT 1" -QueryTimeout 1 -ErrorAction Stop | Out-Null
				$protocol = "TCP"
			}
			catch { }
		}
		
		# Step 3: If not TCP, try Named Pipes SQL query
		if ($protocol -eq "File")
		{
			try
			{
				Import-Module $SqlModuleName -ErrorAction Stop
				Invoke-Sqlcmd -ConnectionString $npConn -Query "SELECT 1" -QueryTimeout 1 -ErrorAction Stop | Out-Null
				$protocol = "Named Pipes"
			}
			catch { }
		}
		
		# Step 4: If neither works, protocol remains "File"
		[PSCustomObject]@{ Lane = $laneNum; Protocol = $protocol }
	}
}

# Live-poll table view (keeps running, shows table as long as PowerShell window is open)
$protocolTimer = New-Object System.Windows.Forms.Timer
$protocolTimer.Interval = 1000
$protocolTimer.add_Tick({
		$keysCopy = @($script:LaneProtocolJobs.Keys)
		foreach ($lane in $keysCopy)
		{
			$job = $script:LaneProtocolJobs[$lane]
			if ($job.State -eq 'Completed')
			{
				$result = Receive-Job $job -ErrorAction SilentlyContinue
				if ($result -and $result.Lane -and $result.Protocol)
				{
					$rawLane = $result.Lane
					$numericLane = ($rawLane -replace '[^\d]', '').PadLeft(3, '0')
					# Add all possible key forms
					$script:LaneProtocols[$numericLane] = $result.Protocol
					$script:LaneProtocols[$rawLane] = $result.Protocol					
					# If you have LaneNumToMachineName mapping loaded, add machine name as key
					if ($script:FunctionResults -and $script:FunctionResults['LaneNumToMachineName'])
					{
						$machineName = $script:FunctionResults['LaneNumToMachineName'][$numericLane]
						if ($machineName) { $script:LaneProtocols[$machineName] = $result.Protocol }
					}
					
					# Also cache with lowercased machine name just in case
					if ($machineName) { $script:LaneProtocols[$machineName.ToLower()] = $result.Protocol }
					$already = $script:ProtocolResults | Where-Object { $_.Lane -eq $rawLane }
					if (-not $already) { $script:ProtocolResults += $result }
				}
				Remove-Job $job -Force
				$script:LaneProtocolJobs.Remove($lane)
			}
			elseif ($job.State -eq 'Failed')
			{
				Write-Host "`r`nJob for Lane $lane failed: $($job.ChildJobs[0].JobStateInfo.Reason)" -ForegroundColor Red
				
				# --- Same as above, add all key forms ---
				$rawLane = $lane
				$numericLane = ($rawLane -replace '[^\d]', '').PadLeft(3, '0')
				$script:LaneProtocols[$numericLane] = "File"
				$script:LaneProtocols[$rawLane] = "File"
				if ($script:FunctionResults -and $script:FunctionResults['LaneNumToMachineName'])
				{
					$machineName = $script:FunctionResults['LaneNumToMachineName'][$numericLane]
					if ($machineName) { $script:LaneProtocols[$machineName] = "File" }
				}
				if ($machineName) { $script:LaneProtocols[$machineName.ToLower()] = "File" }
				
				$already = $script:ProtocolResults | Where-Object { $_.Lane -eq $lane }
				if (-not $already)
				{
					$script:ProtocolResults += [PSCustomObject]@{ Lane = $lane; Protocol = "File" }
				}
				Remove-Job $job -Force
				$script:LaneProtocolJobs.Remove($lane)
			}
		}
	})
$protocolTimer.Start()

# Populate the hash table with results from various functions
$AliasToTable = Get_Table_Aliases

# Generate SQL scripts
Generate_SQL_Scripts -StoreNumber $StoreNumber -LanesqlFilePath $LanesqlFilePath -StoresqlFilePath $StoresqlFilePath

# Clearing XE (Urgent Messages) folder.
$ClearXEJob = Clear_XE_Folder

# Clear %Temp% folder on start
# $ClearTempAtLaunch = Delete_Files -Path "$TempDir" -Exclusions "Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1" -AsJob
# $ClearWinTempAtLaunch = Delete_Files -Path "$env:SystemRoot\Temp" -AsJob

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
