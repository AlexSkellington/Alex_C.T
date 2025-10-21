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
$VersionNumber = "2.4.9"
$VersionDate = "2025-10-21"

# Retrieve Major, Minor, Build, and Revision version numbers of PowerShell
$major = $PSVersionTable.PSVersion.Major
$minor = $PSVersionTable.PSVersion.Minor
$build = $PSVersionTable.PSVersion.Build
$revision = $PSVersionTable.PSVersion.Revision

# Idle timeout for the whole script
$script:IdleMinutesAllowed = 15 # <<< adjust as needed
$script:SuppressClosePrompt = $true

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

# === Directories for Backups and Scripts ===
$script:BackupRoot = "C:\Tecnica_Systems\Alex_C.T\Backups\"
$script:ScriptsFolder = "C:\Tecnica_Systems\Alex_C.T\Scripts\"

# === SQL Backup/Automation Credentials ===
$script:BackupSqlUser = "Tecnica"
$script:BackupSqlPass = "TB`$upp0rT"

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

# Local machine name
$script:LocalHost = $env:COMPUTERNAME

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
#                               FUNCTION: Repair_LOC_Databases_On_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Pick lane(s), then choose repair level (Audit / Quick / Deep) in a dialog.
#   Uses Get_All_Lanes_Database_Info to resolve DB server/name per lane.
#   Only runs on lanes with cached protocol in $script:LaneProtocols (TCP or Named Pipes).
#   Connects to master; repairs ONLY the lane DB from Startup.ini.
#   NEW: If a system DB (master/model/msdb/tempdb) is NOT ONLINE:
#        • tempdb -> offer to restart SQL service on the lane, then re-probe.
#        • model/msdb -> offer deep repair with explicit confirmation; else skip with guidance.
#        • master -> log guidance (restore/rebuild), skip lane (not automated).
#   Logging matches Process_Lanes; PS 5.1; no nested functions; no ternary; only Write_Log.
# ===================================================================================================

function Repair_LOC_Databases_On_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[int]$CommandTimeout = 900 # per-statement timeout (seconds)
	)
	
	# ----------------------------------------
	# Banner: Start
	# ----------------------------------------
	Write_Log "`r`n==================== Starting Repair_LOC_Databases_On_Lanes Function ====================`r`n" "blue"
	
	# ----------------------------------------
	# Import detected SQL module for Invoke-Sqlcmd usage (same pattern as Process_Lanes)
	# ----------------------------------------
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch
		{
			Write_Log "Failed to import SQL module or find Invoke-Sqlcmd: $_" "red"
			Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
			return
		}
	}
	else
	{
		Write_Log "No valid SQL module available for SQL operations!" "red"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	
	# ----------------------------------------
	# Check for available Lane Machines map
	# ----------------------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "No lanes available. Please retrieve nodes first." "red"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	
	# ----------------------------------------
	# Get user's lane selection (same UX as Process_Lanes)
	# ----------------------------------------
	$selection = $null
	try
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Lanes for DB Repair"
	}
	catch
	{
		Write_Log "Lane selection failed: $_" "red"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	if ($selection -eq $null)
	{
		Write_Log "Lane DB repair canceled by user." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	if (-not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	
	# Support both string and object selections for lanes
	$Lanes = @()
	if ($selection.Lanes[0] -is [PSCustomObject] -and $selection.Lanes[0].PSObject.Properties.Name -contains 'LaneNumber')
	{
		foreach ($item in $selection.Lanes) { $Lanes += $item.LaneNumber }
	}
	else
	{
		$Lanes = $selection.Lanes
	}
	
	# ----------------------------------------
	# In-function dialog to pick repair level (no external switches)
	# ----------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Database Repair Level"
	$form.Size = New-Object System.Drawing.Size(520, 240)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$lbl = New-Object System.Windows.Forms.Label
	$lbl.Text = "Choose how aggressive the repair should be:"
	$lbl.AutoSize = $true
	$lbl.Location = New-Object System.Drawing.Point(12, 12)
	$form.Controls.Add($lbl)
	
	$grp = New-Object System.Windows.Forms.GroupBox
	$grp.Text = "Repair Level"
	$grp.Location = New-Object System.Drawing.Point(12, 36)
	$grp.Size = New-Object System.Drawing.Size(490, 110)
	$form.Controls.Add($grp)
	
	$rbAudit = New-Object System.Windows.Forms.RadioButton
	$rbAudit.Text = "Audit only (DBCC CHECKDB, no changes)"
	$rbAudit.AutoSize = $true
	$rbAudit.Location = New-Object System.Drawing.Point(16, 24)
	$rbAudit.Checked = $true
	$grp.Controls.Add($rbAudit)
	
	$rbQuick = New-Object System.Windows.Forms.RadioButton
	$rbQuick.Text = "Quick repair (REPAIR_REBUILD)"
	$rbQuick.AutoSize = $true
	$rbQuick.Location = New-Object System.Drawing.Point(16, 48)
	$grp.Controls.Add($rbQuick)
	
	$rbDeep = New-Object System.Windows.Forms.RadioButton
	$rbDeep.Text = "Deep repair (REPAIR_ALLOW_DATA_LOSS)"
	$rbDeep.AutoSize = $true
	$rbDeep.Location = New-Object System.Drawing.Point(16, 72)
	$grp.Controls.Add($rbDeep)
	
	$chkConfirm = New-Object System.Windows.Forms.CheckBox
	$chkConfirm.Text = "I understand deep repair can cause data loss."
	$chkConfirm.AutoSize = $true
	$chkConfirm.Location = New-Object System.Drawing.Point(28, 152)
	$chkConfirm.Enabled = $false
	$form.Controls.Add($chkConfirm)
	
	[void]$rbDeep.Add_CheckedChanged({
			if ($rbDeep.Checked) { $chkConfirm.Enabled = $true }
			else { $chkConfirm.Enabled = $false; $chkConfirm.Checked = $false }
		})
	[void]$rbAudit.Add_CheckedChanged({ if ($rbAudit.Checked) { $chkConfirm.Enabled = $false; $chkConfirm.Checked = $false } })
	[void]$rbQuick.Add_CheckedChanged({ if ($rbQuick.Checked) { $chkConfirm.Enabled = $false; $chkConfirm.Checked = $false } })
	
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(314, 172)
	$btnOK.Size = New-Object System.Drawing.Size(85, 28)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(417, 172)
	$btnCancel.Size = New-Object System.Drawing.Size(85, 28)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	$dlg = $form.ShowDialog()
	if ($dlg -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Repair level selection canceled." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	
	$Level = 'Audit'
	if ($rbQuick.Checked) { $Level = 'Quick' }
	if ($rbDeep.Checked) { $Level = 'Deep' }
	if ($Level -eq 'Deep' -and -not $chkConfirm.Checked)
	{
		Write_Log "Deep repair selected but confirmation not checked. Aborting." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	
	Write_Log ("Selected repair level: {0}" -f $Level) "gray"
	
	# ----------------------------------------
	# Ensure protocol cache exists; we will NOT populate it here
	# ----------------------------------------
	if (-not $script:LaneProtocols -or $script:LaneProtocols.Keys.Count -eq 0)
	{
		Write_Log "Protocol cache is empty. Please run Start_Lane_Protocol_Jobs first. No lanes will be processed." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	
	# ----------------------------------------
	# Process lanes one by one (summary like Process_Lanes)
	# ----------------------------------------
	$laneSummary = New-Object System.Collections.Generic.List[pscustomobject]
	
	foreach ($LaneNumber in ($Lanes | Sort-Object))
	{
		$laneKey = ($LaneNumber -replace '[^\d]', '')
		$laneKeyP = $laneKey.PadLeft(3, '0')
		
		# Lookup machine
		$machineName = $null
		if ($LaneNumToMachineName.ContainsKey($LaneNumber)) { $machineName = $LaneNumToMachineName[$LaneNumber] }
		if (-not $machineName -and $LaneNumToMachineName.ContainsKey($laneKeyP)) { $machineName = $LaneNumToMachineName[$laneKeyP] }
		
		if (-not $machineName)
		{
			Write_Log ("Lane {0}: No machine mapping found. Skipping." -f $LaneNumber) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = ''; Protocol = ''; DBServer = ''; DBName = ''; Level = $Level; Result = 'NoMachine' })
			continue
		}
		
		# Resolve protocol from cache ONLY (do not detect here)
		$proto = $null
		if ($script:LaneProtocols.ContainsKey($laneKeyP)) { $proto = $script:LaneProtocols[$laneKeyP] }
		elseif ($script:LaneProtocols.ContainsKey($LaneNumber)) { $proto = $script:LaneProtocols[$LaneNumber] }
		elseif ($script:LaneProtocols.ContainsKey($machineName)) { $proto = $script:LaneProtocols[$machineName] }
		else
		{
			$lower = $machineName.ToLower()
			if ($script:LaneProtocols.ContainsKey($lower)) { $proto = $script:LaneProtocols[$lower] }
		}
		
		if (-not $proto)
		{
			Write_Log ("Lane {0} ({1}): No protocol found in cache. Run Start_Lane_Protocol_Jobs first. Skipping." -f $laneKeyP, $machineName) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = ''; DBServer = ''; DBName = ''; Level = $Level; Result = 'NoProtocolCached' })
			continue
		}
		
		if ($proto -ne 'TCP' -and $proto -ne 'Named Pipes')
		{
			Write_Log ("Lane {0} ({1}): Cached protocol is '{2}'. Only TCP or Named Pipes are supported. Skipping." -f $laneKeyP, $machineName, $proto) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = ''; DBName = ''; Level = $Level; Result = 'UnsupportedProtocol' })
			continue
		}
		
		# Resolve DB connection info from your helper (Startup.ini)
		$dbInfo = $null
		try { $dbInfo = Get_All_Lanes_Database_Info -LaneNumber $laneKeyP }
		catch { $dbInfo = $null }
		if (-not $dbInfo)
		{
			Write_Log ("Lane {0} ({1}): Could not get DB info. Skipping." -f $laneKeyP, $machineName) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = ''; DBName = ''; Level = $Level; Result = 'NoDbInfo' })
			continue
		}
		
		$dbName = $dbInfo['DBName']
		$dbServer = $dbInfo['DBServer']
		$instanceName = $dbInfo['InstanceName']
		$csNamedPipes = $dbInfo['NamedPipesConnStr']
		$csTcp = $dbInfo['TcpConnStr']
		
		Write_Log ("`r`n--- Lane {0} ({1}) | Protocol={2} | DB: {3} on {4} ---" -f $laneKeyP, $machineName, $proto, $dbName, $dbServer) "blue"
		
		if ([string]::IsNullOrWhiteSpace($dbName) -or [string]::IsNullOrWhiteSpace($dbServer))
		{
			Write_Log "Incomplete DB info (DBName or DBServer missing). Skipping lane." "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'IncompleteDbInfo' })
			continue
		}
		
		# Build a connection string to MASTER based on cached protocol (no fallbacks)
		$rawCs = $null
		if ($proto -eq 'TCP') { $rawCs = $csTcp }
		if ($proto -eq 'Named Pipes') { $rawCs = $csNamedPipes }
		
		if ([string]::IsNullOrWhiteSpace($rawCs))
		{
			Write_Log ("Lane {0} ({1}): Missing {2} connection string. Skipping." -f $laneKeyP, $machineName, $proto) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'NoConnStr' })
			continue
		}
		
		$connStr = $rawCs
		if ($connStr -match 'Database\s*=\s*[^;]+;')
		{
			$connStr = [System.Text.RegularExpressions.Regex]::Replace($connStr, 'Database\s*=\s*[^;]+;', 'Database=master;', 'IgnoreCase')
		}
		else
		{
			if ($connStr.EndsWith(';')) { $connStr = $connStr + 'Database=master;' }
			else { $connStr = $connStr + ';Database=master;' }
		}
		if ($connStr -notmatch 'TrustServerCertificate\s*=') { $connStr = $connStr + 'TrustServerCertificate=True;' }
		if ($connStr -notmatch 'Application Name\s*=') { $connStr = $connStr + 'Application Name=TBS_DBRepair;' }
		if ($connStr -notmatch 'Integrated Security\s*=') { $connStr = $connStr + 'Integrated Security=True;' }
		
		# Quick probe using the chosen method ONLY
		$probeOK = $false
		try
		{
			& $InvokeSqlCmd -ConnectionString $connStr -Query "SELECT 1 AS ok;" -QueryTimeout 8 -ErrorAction Stop | Out-Null
			$probeOK = $true
		}
		catch { $probeOK = $false }
		
		if (-not $probeOK)
		{
			Write_Log ("Lane {0} ({1}): {2} probe failed. Skipping." -f $laneKeyP, $machineName, $proto) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'ConnProbeFailed' })
			continue
		}
		
		# ----------------------------------------
		# System DB health gate with ACTIONS
		# ----------------------------------------
		$sysRows = $null
		try
		{
			$sysRows = & $InvokeSqlCmd -ConnectionString $connStr `
									   -Query "SELECT name, state_desc FROM sys.databases WHERE name IN ('master','model','msdb','tempdb');" `
									   -QueryTimeout 10 -ErrorAction Stop
		}
		catch
		{
			Write_Log ("Lane {0} ({1}): WARN: could not query system DB states: {2}" -f $laneKeyP, $machineName, $_.Exception.Message) "yellow"
			$laneSummary.Add([pscustomobject]@{
					Lane	    = $laneKeyP
					Machine	    = $machineName
					Protocol    = $proto
					DBServer    = $dbServer
					DBName	    = $dbName
					Level	    = $Level
					QuickRepair = $false
					DeepRepair  = $false
					Result	    = 'SystemDbStateUnknown'
				})
			continue
		}
		
		$badSystemDbs = @()
		foreach ($r in $sysRows)
		{
			if ($r -and $r.state_desc -and ($r.state_desc.ToString() -ne 'ONLINE'))
			{
				$badSystemDbs += [pscustomobject]@{ Name = [string]$r.name; State = [string]$r.state_desc }
			}
		}
		
		if ($badSystemDbs.Count -gt 0)
		{
			Write_Log ("Lane {0} ({1}): System DBs not ONLINE detected." -f $laneKeyP, $machineName) "yellow"
			foreach ($b in $badSystemDbs) { Write_Log ("  - {0}: {1}" -f $b.Name, $b.State) "yellow" }
			
			# Determine SQL service name on the remote machine
			$svcName = "MSSQLSERVER"
			if ($instanceName -and ($instanceName.ToUpper() -ne "MSSQLSERVER")) { $svcName = "MSSQL`$$instanceName" }
			
			# Handle tempdb specifically: offer restart of the SQL service on the lane
			$hasTempdbIssue = $false
			foreach ($b in $badSystemDbs) { if ($b.Name -eq 'tempdb') { $hasTempdbIssue = $true } }
			
			if ($hasTempdbIssue)
			{
				$msg = "Lane $laneKeyP ($machineName): tempdb is not ONLINE (`"$($badSystemDbs | Where-Object { $_.Name -eq 'tempdb' } | Select-Object -First 1).State`").`r`n`r`n" +
				"Would you like to restart the SQL Server service ($svcName) on $machineName now? This will drop connections and recreate tempdb."
				$resp = [System.Windows.Forms.MessageBox]::Show($msg, "Restart SQL Service for tempdb", `
					[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
				
				if ($resp -eq [System.Windows.Forms.DialogResult]::Yes)
				{
					Write_Log ("Attempting to restart service {0} on {1}..." -f $svcName, $machineName) "gray"
					
					# stop
					try
					{
						& sc.exe "\\$machineName" stop $svcName | Out-Null
					}
					catch { }
					# wait for stop up to ~60s
					$stopped = $false
					for ($i = 0; $i -lt 30; $i++)
					{
						Start-Sleep -Milliseconds 2000
						try
						{
							$chk = sc.exe "\\$machineName" query $svcName 2>$null
							$txt = if ($chk) { $chk | Out-String }
							else { "" }
							if ($txt -match 'STATE\s*:\s*\d+\s*STOPPED') { $stopped = $true; break }
						}
						catch { }
					}
					if ($stopped) { Write_Log "Service stopped." "gray" }
					else { Write_Log "WARN: Service did not report STOPPED in time." "yellow" }
					
					# start
					try
					{
						& sc.exe "\\$machineName" start $svcName | Out-Null
					}
					catch { }
					# wait for running up to ~90s
					$running = $false
					for ($i = 0; $i -lt 45; $i++)
					{
						Start-Sleep -Milliseconds 2000
						try
						{
							$chk = sc.exe "\\$machineName" query $svcName 2>$null
							$txt = if ($chk) { $chk | Out-String }
							else { "" }
							if ($txt -match 'STATE\s*:\s*\d+\s*RUNNING') { $running = $true; break }
						}
						catch { }
					}
					if ($running) { Write_Log "Service started." "green" }
					else { Write_Log "ERROR: Service did not report RUNNING in time." "red" }
					
					# re-probe connection (service restart broke previous connection context)
					$reprobeOK = $false
					for ($t = 1; $t -le 60; $t++)
					{
						try
						{
							& $InvokeSqlCmd -ConnectionString $connStr -Query "SELECT 1 AS ok;" -QueryTimeout 5 -ErrorAction Stop | Out-Null
							$reprobeOK = $true
							break
						}
						catch { Start-Sleep -Milliseconds 2000 }
					}
					if ($reprobeOK)
					{
						# check system DBs again
						try
						{
							$sysRows = & $InvokeSqlCmd -ConnectionString $connStr `
													   -Query "SELECT name, state_desc FROM sys.databases WHERE name IN ('master','model','msdb','tempdb');" `
													   -QueryTimeout 10 -ErrorAction Stop
							$badSystemDbs = @()
							foreach ($r in $sysRows) { if ($r -and $r.state_desc -and ($r.state_desc.ToString() -ne 'ONLINE')) { $badSystemDbs += [pscustomobject]@{ Name = [string]$r.name; State = [string]$r.state_desc } } }
						}
						catch
						{
							Write_Log "WARN: Re-check of system DB states failed after restart." "yellow"
						}
					}
					else
					{
						Write_Log "ERROR: Could not reconnect after service restart; skipping lane." "red"
						$laneSummary.Add([pscustomobject]@{
								Lane	    = $laneKeyP
								Machine	    = $machineName
								Protocol    = $proto
								DBServer    = $dbServer
								DBName	    = $dbName
								Level	    = $Level
								QuickRepair = $false
								DeepRepair  = $false
								Result	    = 'ServiceRestartFailed'
							})
						continue
					}
				}
				else
				{
					Write_Log "User declined SQL service restart for tempdb. Skipping lane with guidance: restart SQL service on the lane to recreate tempdb." "yellow"
					$laneSummary.Add([pscustomobject]@{
							Lane	    = $laneKeyP
							Machine	    = $machineName
							Protocol    = $proto
							DBServer    = $dbServer
							DBName	    = $dbName
							Level	    = $Level
							QuickRepair = $false
							DeepRepair  = $false
							Result	    = 'TempdbNotOnline_Skipped'
						})
					continue
				}
			}
			
			# After potential tempdb handling, if any system DB still not ONLINE, handle model/msdb or master
			$stillBad = $false
			foreach ($r in $badSystemDbs) { if ($r.Name -ne 'tempdb') { $stillBad = $true } }
			if ($badSystemDbs.Count -gt 0 -and -not $stillBad)
			{
				# only tempdb was bad and we handled it; re-check one more time
				$stillBad = $false
				foreach ($r in $badSystemDbs) { if ($r.state -ne 'ONLINE') { $stillBad = $true } }
			}
			
			if ($stillBad)
			{
				# master/model/msdb handling
				$hasMaster = $false
				$hasModel = $false
				$hasMsdb = $false
				foreach ($r in $badSystemDbs)
				{
					if ($r.Name -eq 'master') { $hasMaster = $true }
					if ($r.Name -eq 'model') { $hasModel = $true }
					if ($r.Name -eq 'msdb') { $hasMsdb = $true }
				}
				
				if ($hasMaster)
				{
					Write_Log "CRITICAL: master is not ONLINE. Automated repair is not supported here." "red"
					Write_Log "Action: Restore master from a known-good backup OR rebuild system DBs (setup.exe /ACTION=REBUILDDATABASE), then restore." "gray"
					$laneSummary.Add([pscustomobject]@{
							Lane	    = $laneKeyP
							Machine	    = $machineName
							Protocol    = $proto
							DBServer    = $dbServer
							DBName	    = $dbName
							Level	    = $Level
							QuickRepair = $false
							DeepRepair  = $false
							Result	    = 'MasterNotOnline'
						})
					continue
				}
				
				# Offer EMERGENCY deep-repair for model/msdb (with explicit confirmation)
				$targets = @()
				if ($hasModel) { $targets += 'model' }
				if ($hasMsdb) { $targets += 'msdb' }
				
				if ($targets.Count -gt 0)
				{
					$list = ($targets -join ', ')
					$warn = "Lane $laneKeyP ($machineName): $list is not ONLINE.`r`n`r`n" +
					"Attempt DEEP REPAIR (EMERGENCY + REPAIR_ALLOW_DATA_LOSS) for these system DB(s)?`r`n" +
					"This can cause data loss. Recommended alternative is RESTORE from backup."
					$resp2 = [System.Windows.Forms.MessageBox]::Show($warn, "Deep Repair system DB(s)?", `
						[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
					if ($resp2 -eq [System.Windows.Forms.DialogResult]::Yes)
					{
						foreach ($sysDb in $targets)
						{
							Write_Log ("Attempting deep repair on system DB: {0}" -f $sysDb) "yellow"
							$qSys = "[" + ($sysDb -replace "]", "]]") + "]"
							try
							{
								& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qSys SET EMERGENCY;" -QueryTimeout 30 -ErrorAction Stop | Out-Null
							}
							catch { Write_Log ("WARN: EMERGENCY failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "yellow" }
							try
							{
								& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qSys SET SINGLE_USER WITH ROLLBACK IMMEDIATE;" -QueryTimeout 30 -ErrorAction Stop | Out-Null
							}
							catch { Write_Log ("WARN: SINGLE_USER failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "yellow" }
							$okSys = $false
							try
							{
								& $InvokeSqlCmd -ConnectionString $connStr -Query "DBCC CHECKDB ($qSys, REPAIR_ALLOW_DATA_LOSS) WITH ALL_ERRORMSGS;" -QueryTimeout $CommandTimeout -ErrorAction Stop | Out-Null
								$okSys = $true
								Write_Log ("Deep repair completed on {0}." -f $sysDb) "green"
							}
							catch
							{
								$okSys = $false
								Write_Log ("Deep repair failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "red"
							}
							try
							{
								& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qSys SET MULTI_USER;" -QueryTimeout 30 -ErrorAction Stop | Out-Null
							}
							catch { Write_Log ("WARN: MULTI_USER restore failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "yellow" }
						}
						
						# Re-check system DBs after deep repair attempts
						try
						{
							$sysRows = & $InvokeSqlCmd -ConnectionString $connStr `
													   -Query "SELECT name, state_desc FROM sys.databases WHERE name IN ('master','model','msdb','tempdb');" `
													   -QueryTimeout 10 -ErrorAction Stop
							$badSystemDbs = @()
							foreach ($r in $sysRows) { if ($r -and $r.state_desc -and ($r.state_desc.ToString() -ne 'ONLINE')) { $badSystemDbs += [pscustomobject]@{ Name = [string]$r.name; State = [string]$r.state_desc } } }
						}
						catch
						{
							Write_Log "WARN: Re-check of system DB states failed after deep repair attempt." "yellow"
						}
					}
					else
					{
						Write_Log ("User declined deep repair for system DB(s): {0}. Skipping lane with guidance to RESTORE from backup." -f $list) "yellow"
						$laneSummary.Add([pscustomobject]@{
								Lane	    = $laneKeyP
								Machine	    = $machineName
								Protocol    = $proto
								DBServer    = $dbServer
								DBName	    = $dbName
								Level	    = $Level
								QuickRepair = $false
								DeepRepair  = $false
								Result	    = 'SystemDbNotOnline_Skipped'
							})
						continue
					}
				}
			}
			
			# Final gate after any actions: if any system DB still not ONLINE, skip.
			$finalBad = $false
			foreach ($r in $badSystemDbs)
			{
				if ($r -and $r.State -and ($r.State.ToString() -ne 'ONLINE')) { $finalBad = $true }
			}
			if ($finalBad)
			{
				Write_Log "System DBs are still not ONLINE after attempted actions. Skipping lane." "yellow"
				$laneSummary.Add([pscustomobject]@{
						Lane	    = $laneKeyP
						Machine	    = $machineName
						Protocol    = $proto
						DBServer    = $dbServer
						DBName	    = $dbName
						Level	    = $Level
						QuickRepair = $false
						DeepRepair  = $false
						Result	    = 'SystemDbStillNotOnline'
					})
				continue
			}
			else
			{
				Write_Log "All system DBs are ONLINE after remediation. Proceeding with lane DB repair." "green"
			}
		}
		
		# ----------------------------------------
		# Snapshot lane DB state BEFORE (aliased column + fallback)
		# ----------------------------------------
		$stateBefore = 'Unknown'
		try
		{
			$dbNameQuoted = $dbName.Replace("'", "''")
			$tsqlStateBefore = @"
IF EXISTS (SELECT 1 FROM sys.databases WHERE name = N'$dbNameQuoted')
    SELECT state_desc AS state_desc FROM sys.databases WHERE name = N'$dbNameQuoted';
ELSE
    SELECT CAST(DATABASEPROPERTYEX(N'$dbNameQuoted', N'Status') AS nvarchar(60)) AS state_desc;
"@
			$rows = & $InvokeSqlCmd -ConnectionString $connStr -Query $tsqlStateBefore -QueryTimeout 10 -ErrorAction Stop
			if ($rows)
			{
				$stateBefore = ($rows | Select-Object -First 1 -ExpandProperty state_desc)
				if ([string]::IsNullOrWhiteSpace($stateBefore)) { $stateBefore = 'Unknown' }
			}
		}
		catch { }
		Write_Log ("DB state before: {0}" -f $stateBefore) "gray"
		
		# ----------------------------------------
		# Repair the lane's application DB according to chosen level
		# ----------------------------------------
		$qDb = "[" + ($dbName -replace "]", "]]") + "]"
		$ok = $false
		$usedQuick = $false
		$usedDeep = $false
		
		if ($Level -eq 'Audit')
		{
			Write_Log ("{0}: Running DBCC CHECKDB (Audit)..." -f $dbName) "gray"
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "DBCC CHECKDB (N'$dbName') WITH NO_INFOMSGS, ALL_ERRORMSGS, TABLERESULTS;" -QueryTimeout $CommandTimeout -ErrorAction Stop | Out-Null
				$ok = $true
				Write_Log "Audit completed." "green"
			}
			catch
			{
				$ok = $false
				Write_Log ("Audit reported errors or failed: $_") "yellow"
			}
		}
		elseif ($Level -eq 'Quick')
		{
			Write_Log ("{0}: Quick repair (REPAIR_REBUILD)..." -f $dbName) "gray"
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qDb SET SINGLE_USER WITH ROLLBACK IMMEDIATE;" -QueryTimeout $CommandTimeout -ErrorAction Stop | Out-Null
			}
			catch { Write_Log ("WARN: SINGLE_USER failed: $_") "yellow" }
			
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "DBCC CHECKDB ($qDb, REPAIR_REBUILD) WITH ALL_ERRORMSGS, TABLERESULTS;" -QueryTimeout $CommandTimeout -ErrorAction Stop | Out-Null
				$ok = $true
				$usedQuick = $true
				Write_Log "Quick repair completed." "green"
			}
			catch
			{
				$ok = $false
				Write_Log ("Quick repair failed: $_") "yellow"
			}
			
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qDb SET MULTI_USER;" -QueryTimeout 60 -ErrorAction Stop | Out-Null
			}
			catch { Write_Log ("WARN: MULTI_USER restore failed: $_") "yellow" }
		}
		elseif ($Level -eq 'Deep')
		{
			Write_Log ("{0}: Deep repair (EMERGENCY + REPAIR_ALLOW_DATA_LOSS)..." -f $dbName) "gray"
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qDb SET EMERGENCY;" -QueryTimeout 60 -ErrorAction Stop | Out-Null
			}
			catch { Write_Log ("WARN: EMERGENCY failed: $_") "yellow" }
			
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qDb SET SINGLE_USER WITH ROLLBACK IMMEDIATE;" -QueryTimeout 60 -ErrorAction Stop | Out-Null
			}
			catch { Write_Log ("WARN: SINGLE_USER failed: $_") "yellow" }
			
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "DBCC CHECKDB ($qDb, REPAIR_ALLOW_DATA_LOSS) WITH ALL_ERRORMSGS, TABLERESULTS;" -QueryTimeout $CommandTimeout -ErrorAction Stop | Out-Null
				$ok = $true
				$usedDeep = $true
				Write_Log "Deep repair completed." "green"
			}
			catch
			{
				$ok = $false
				Write_Log ("Deep repair failed: $_") "red"
			}
			
			try
			{
				& $InvokeSqlCmd -ConnectionString $connStr -Query "ALTER DATABASE $qDb SET MULTI_USER;" -QueryTimeout 60 -ErrorAction Stop | Out-Null
			}
			catch { Write_Log ("WARN: MULTI_USER restore failed: $_") "yellow" }
		}
		
		# ----------------------------------------
		# Snapshot state AFTER (aliased column + fallback)
		# ----------------------------------------
		$stateAfter = 'Unknown'
		try
		{
			$dbNameQuoted2 = $dbName.Replace("'", "''")
			$tsqlStateAfter = @"
IF EXISTS (SELECT 1 FROM sys.databases WHERE name = N'$dbNameQuoted2')
    SELECT state_desc AS state_desc FROM sys.databases WHERE name = N'$dbNameQuoted2';
ELSE
    SELECT CAST(DATABASEPROPERTYEX(N'$dbNameQuoted2', N'Status') AS nvarchar(60)) AS state_desc;
"@
			$rows2 = & $InvokeSqlCmd -ConnectionString $connStr -Query $tsqlStateAfter -QueryTimeout 10 -ErrorAction Stop
			if ($rows2)
			{
				$stateAfter = ($rows2 | Select-Object -First 1 -ExpandProperty state_desc)
				if ([string]::IsNullOrWhiteSpace($stateAfter)) { $stateAfter = 'Unknown' }
			}
		}
		catch { }
		Write_Log ("DB state after:  {0}" -f $stateAfter) "gray"
		
		# Summarize this lane
		$result = 'OK'
		if (-not $ok -and $Level -eq 'Audit') { $result = 'AuditFoundErrorsOrFailed' }
		if (-not $ok -and $Level -ne 'Audit') { $result = 'Failed' }
		
		$laneSummary.Add([pscustomobject]@{
				Lane	    = $laneKeyP
				Machine	    = $machineName
				Protocol    = $proto
				DBServer    = $dbServer
				DBName	    = $dbName
				Level	    = $Level
				QuickRepair = $usedQuick
				DeepRepair  = $usedDeep
				Result	    = $result
			})
		
		Write_Log ("Lane {0} | DB {1} -> Level={2}, Protocol={3}, Quick={4}, Deep={5}, Result={6}" -f $laneKeyP, $dbName, $Level, $proto, $usedQuick, $usedDeep, $result) "gray"
	}
	
	# ----------------------------------------
	# Final: Summary table + finish banner
	# ----------------------------------------
	Write_Log "`r`n================ Lane DB Repair Summary ================" "blue"
	if ($laneSummary.Count -gt 0)
	{
		Write_Log ((
				$laneSummary |
				Sort-Object Lane |
				Format-Table -AutoSize Lane, Machine, Protocol, DBServer, DBName, Level, QuickRepair, DeepRepair, Result |
				Out-String
			)) "gray"
	}
	else
	{
		Write_Log "No lanes processed." "yellow"
	}
	
	Write_Log "==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
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
	# 3. CLEAN 1:1 MAPPINGS FOR ALL NODES (NO ALIASES, INCLUDE PATHS)
	# ====================================================================================
	
	# ---- LANES ----
	$CleanLaneNumToMachineName = @{ }
	$CleanMachineNameToLaneNum = @{ }
	$CleanLaneNumToPath = @{ }
	$CleanMachineNameToPath = @{ }
	$CleanLaneNumToServerPath = @{ }
	$CleanMachineNameToServerPath = @{ }
	
	foreach ($kv in $LaneNumToMachineName.GetEnumerator())
	{
		$laneNum = $kv.Key
		$machine = $kv.Value
		$CleanLaneNumToMachineName[$laneNum] = $machine
		$CleanMachineNameToLaneNum[$machine] = $laneNum
		# Paths
		if ($LaneMachinePath.ContainsKey($machine))
		{
			$CleanLaneNumToPath[$laneNum] = $LaneMachinePath[$machine]
			$CleanMachineNameToPath[$machine] = $LaneMachinePath[$machine]
		}
		if ($LaneMachineToServerPath.ContainsKey($machine))
		{
			$CleanLaneNumToServerPath[$laneNum] = $LaneMachineToServerPath[$machine]
			$CleanMachineNameToServerPath[$machine] = $LaneMachineToServerPath[$machine]
		}
	}
	$LaneNumToMachineName = $CleanLaneNumToMachineName
	$MachineNameToLaneNum = $CleanMachineNameToLaneNum
	$LaneNumToPath = $CleanLaneNumToPath
	$MachineNameToPath = $CleanMachineNameToPath
	$LaneNumToServerPath = $CleanLaneNumToServerPath
	$MachineNameToServerPath = $CleanMachineNameToServerPath
	
	# ---- SCALES ----
	$CleanScaleCodeToIPInfo = @{ }
	$CleanScaleNameToCode = @{ }
	$CleanScaleCodeToPath = @{ }
	$CleanScaleNameToPath = @{ }
	
	foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
	{
		$scaleCode = $kv.Key
		$scale = $kv.Value
		$CleanScaleCodeToIPInfo[$scaleCode] = $scale
		if ($scale.ScaleName)
		{
			$CleanScaleNameToCode[$scale.ScaleName] = $scaleCode
			$CleanScaleNameToPath[$scale.ScaleName] = $scale.Path
		}
		$CleanScaleCodeToPath[$scaleCode] = $scale.Path
	}
	$ScaleCodeToIPInfo = $CleanScaleCodeToIPInfo
	$ScaleNameToCode = $CleanScaleNameToCode
	$ScaleCodeToPath = $CleanScaleCodeToPath
	$ScaleNameToPath = $CleanScaleNameToPath
	
	# ---- BACKOFFICES ----
	$CleanBackofficeNumToMachineName = @{ }
	$CleanMachineNameToBackofficeNum = @{ }
	$CleanBackofficeNumToPath = @{ }
	$CleanMachineNameToBOPath = @{ }
	
	foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
	{
		$boNum = $kv.Key
		$machine = $kv.Value
		$CleanBackofficeNumToMachineName[$boNum] = $machine
		$CleanMachineNameToBackofficeNum[$machine] = $boNum
		# Paths
		if ($BackofficeNumToPath.ContainsKey($boNum))
		{
			$CleanBackofficeNumToPath[$boNum] = $BackofficeNumToPath[$boNum]
			$CleanMachineNameToBOPath[$machine] = $BackofficeNumToPath[$boNum]
		}
	}
	$BackofficeNumToMachineName = $CleanBackofficeNumToMachineName
	$MachineNameToBackofficeNum = $CleanMachineNameToBackofficeNum
	$BackofficeNumToPath = $CleanBackofficeNumToPath
	$MachineNameToBackofficePath = $CleanMachineNameToBOPath
	
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
	$dbServer = $script:FunctionResults['DBSERVER']
	
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
	
	#Always Enforce the Slash
	$BackupRoot = ($BackupRoot.TrimEnd('\') + '\')
	$ScriptsFolder = ($ScriptsFolder.TrimEnd('\') + '\')
	
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

/* Create "$BackupSqlUser" user in the database */
-- Create SQL Login if not exists
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = '$BackupSqlUser')
    CREATE LOGIN [$BackupSqlUser] WITH PASSWORD = '$BackupSqlPass', CHECK_POLICY = OFF;
-- Create User in DB if not exists, grant backup rights
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '$BackupSqlUser')
    CREATE USER [$BackupSqlUser] FOR LOGIN [$BackupSqlUser];
IF IS_ROLEMEMBER('db_owner', '$BackupSqlUser') = 0
    EXEC sp_addrolemember 'db_owner', '$BackupSqlUser';

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
		$uniqueScaleObjs = @{ }
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			# Compute a unique identifier, e.g., IP or some combination
			$ip = $null
			if ($scaleObj.FullIP) { $ip = $scaleObj.FullIP }
			elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { $ip = "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)" }
			# Add only once!
			if ($ip -and -not $uniqueScaleObjs.ContainsKey($ip))
			{
				$uniqueScaleObjs[$ip] = $scaleObj
			}
		}
		foreach ($ip in $uniqueScaleObjs.Keys)
		{
			$scaleObj = $uniqueScaleObjs[$ip]
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
#   Inserts a test item (PLU is chosen from 0020077700000, 0020777700000, 0027777700000) into:
#   - SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB, and SCL_TXT_TAB
#   Logic:
#     * We DELETE existing rows for the chosen PLU/F267 first to keep the test rows deterministic.
#     * Then we INSERT fixed values (same semantics as your last working version).
#
# Module/Connectivity Handling (integrated with Get_Store_And_Database_Info):
#   - Uses $script:FunctionResults['SqlModuleName'] to decide the best path.
#   - Tries Invoke-Sqlcmd with -ConnectionString if SqlServer module is available and probe succeeds.
#   - If unsupported or probe fails, FALLS BACK to -ServerInstance/-Database with
#     $script:FunctionResults['DBSERVER']/['DBNAME'] (older Windows / SQLPS path).
#
# Notes:
#   - Write_Log is used for colored status lines.
#   - If no module provides Invoke-Sqlcmd, we log and abort (keeps behavior predictable).
#
# Author: Alex_C.T
# ===================================================================================================

function Insert_Test_Item
{
	[CmdletBinding()]
	param (
		# If omitted, we use the ConnectionString assembled by Get_Store_And_Database_Info.
		[string]$ConnectionString = $script:FunctionResults['ConnectionString']
	)
	
	# ----------------------------------------------------------------------------------------------
	# Guard: Need connection context from Get_Store_And_Database_Info
	# (DBSERVER/DBNAME and/or ConnectionString should be available)
	# ----------------------------------------------------------------------------------------------
	if (-not $ConnectionString -and
		(-not $script:FunctionResults['DBSERVER'] -or -not $script:FunctionResults['DBNAME']))
	{
		Write_Log "Insert_Test_Item: No ConnectionString nor DBSERVER/DBNAME available. Run Get_Store_And_Database_Info first." "red"
		return
	}
	
	Write_Log "`r`n==================== Starting Insert_Test_Item ====================`r`n" "blue"
	
	# ----------------------------------------------------------------------------------------------
	# Ensure Invoke-Sqlcmd is available (prefer SqlServer, fallback to SQLPS)
	# We honor the detection already done in Get_Store_And_Database_Info.
	# ----------------------------------------------------------------------------------------------
	$sqlModule = $script:FunctionResults['SqlModuleName']
	if (-not (Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue))
	{
		try
		{
			if ($sqlModule -eq 'SqlServer')
			{
				Import-Module SqlServer -ErrorAction Stop
			}
			elseif ($sqlModule -eq 'SQLPS')
			{
				Import-Module SQLPS -DisableNameChecking -ErrorAction Stop
			}
		}
		catch
		{
			# If import fails, we'll check again below; if still missing, we abort.
		}
	}
	
	if (-not (Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue))
	{
		Write_Log "Insert_Test_Item: Invoke-Sqlcmd not available (SqlServer/SQLPS modules not loaded/installed)." "red"
		Write_Log "`r`n==================== Insert_Test_Item Aborted (No SQL Module) ====================`r`n" "blue"
		return
	}
	
	# ----------------------------------------------------------------------------------------------
	# Decide primary vs. fallback path:
	#   - If SqlServer module is detected, we ATTEMPT -ConnectionString (modern path).
	#   - Otherwise (SQLPS/None), we go straight to fallback (-ServerInstance/-Database).
	#   - Even with SqlServer, if the probe SELECT 1 fails, we flip to fallback.
	# ----------------------------------------------------------------------------------------------
	$useConnectionString = $false # Default to fallback unless we validate modern path
	$fallbackParams = @{ } # Will be used with -ServerInstance/-Database
	
	# Pull server/database from FunctionResults (preferred over parsing)
	$serverInstanceFromInfo = $script:FunctionResults['DBSERVER']
	$databaseFromInfo = $script:FunctionResults['DBNAME']
	
	# Compose fallback param set from known values
	if ($serverInstanceFromInfo) { $fallbackParams['ServerInstance'] = $serverInstanceFromInfo }
	if ($databaseFromInfo) { $fallbackParams['Database'] = $databaseFromInfo }
	
	# If we have SqlServer module, try modern path first.
	if ($sqlModule -eq 'SqlServer' -and $ConnectionString)
	{
		try
		{
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT 1" -QueryTimeout 5 -ErrorAction Stop | Out-Null
			$useConnectionString = $true
			Write_Log "SQL connectivity via -ConnectionString verified (SqlServer module)." "green"
		}
		catch
		{
			Write_Log "Modern path (-ConnectionString) failed: $($_.Exception.Message)`r`nFalling back to -ServerInstance/-Database..." "yellow"
			$useConnectionString = $false
		}
	}
	else
	{
		# Either SQLPS or no module indicated from detection - use fallback path.
		$useConnectionString = $false
	}
	
	# If planning to use fallback, validate it can connect
	if (-not $useConnectionString)
	{
		# If DBSERVER/DBNAME were missing, try to salvage from the connection string
		if (-not $fallbackParams['ServerInstance'] -or -not $fallbackParams['Database'])
		{
			# Last-resort: parse minimal fields from ConnectionString (only if needed)
			if ($ConnectionString -match '(?i)(?:Data\s*Source|Server)\s*=\s*([^;]+)')
			{
				$fallbackParams['ServerInstance'] = $matches[1].Trim()
			}
			if ($ConnectionString -match '(?i)(?:Initial\s*Catalog|Database)\s*=\s*([^;]+)')
			{
				$fallbackParams['Database'] = $matches[1].Trim()
			}
		}
		
		try
		{
			Invoke-Sqlcmd @fallbackParams -Query "SELECT 1" -QueryTimeout 5 -ErrorAction Stop | Out-Null
			$dbTxt = if ($fallbackParams['Database']) { " (DB=$($fallbackParams['Database']))" }
			else { "" }
			Write_Log "SQL connectivity via fallback -ServerInstance verified: $($fallbackParams['ServerInstance'])$dbTxt" "green"
		}
		catch
		{
			Write_Log "Fallback (-ServerInstance/-Database) failed: $($_.Exception.Message)" "red"
			Write_Log "`r`n==================== Insert_Test_Item Aborted (No Connectivity) ====================`r`n" "blue"
			return
		}
	}
	
	# ----------------------------------------------------------------------------------------------
	# Centralized SQL runner - ALWAYS use this for queries below so fallback is automatic.
	# (CHANGE: kept as a scriptblock variable - not a nested *function* - to honor 'no nested functions')
	# ----------------------------------------------------------------------------------------------
	$RunSql = {
		param ([Parameter(Mandatory = $true)]
			[string]$Query)
		if ($useConnectionString)
		{
			return Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $Query -ErrorAction Stop
		}
		else
		{
			$p = $fallbackParams.Clone()
			$p['Query'] = $Query
			$p['ErrorAction'] = 'Stop'
			return Invoke-Sqlcmd @p
		}
	}
	
	# ----------------------------------------------------------------------------------------------
	# Business logic: choose appropriate PLU (prefer the "test" one), then delete + insert rows
	# (CHANGE: removed nested helper function; inlined the PLU test logic 3x for clarity/compliance)
	# ----------------------------------------------------------------------------------------------
	$now = Get-Date
	$nowFull = $now.ToString("yyyy-MM-dd HH:mm:ss.fff")
	$nowDate = $now.ToString("yyyy-MM-dd 00:00:00.000")
	
	$preferredPLU = '0020077700000'
	$alternativePLU = '0020777700000'
	$fallbackPLU = '0027777700000'
	
	$PLU = $null
	$TestF267 = 777
	$doInsert = $false
	
	# ---- Check preferred PLU (INLINE)
	$okUse = $false
	try
	{
		$pos = & $RunSql "SELECT F02 FROM POS_TAB WHERE F01 = '$preferredPLU'"
		$obj = & $RunSql "SELECT F29 FROM OBJ_TAB WHERE F01 = '$preferredPLU'"
		$posDesc = if ($pos) { [string]$pos.F02 }
		else { "" }
		$objDesc = if ($obj) { [string]$obj.F29 }
		else { "" }
		$okUse = ($posDesc -match '(?i)test|tecnica') -or
		($objDesc -match '(?i)test|tecnica') -or
		([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
	}
	catch { $okUse = $true } # conservative default if lookups fail
	if ($okUse)
	{
		$PLU = $preferredPLU
		$TestF267 = 777
		$doInsert = $true
		Write_Log "Using preferred PLU: $PLU with F267: $TestF267" "green"
	}
	
	# ---- If not chosen yet, check alternative PLU (INLINE)
	if (-not $doInsert)
	{
		$okUse = $false
		try
		{
			$pos = & $RunSql "SELECT F02 FROM POS_TAB WHERE F01 = '$alternativePLU'"
			$obj = & $RunSql "SELECT F29 FROM OBJ_TAB WHERE F01 = '$alternativePLU'"
			$posDesc = if ($pos) { [string]$pos.F02 }
			else { "" }
			$objDesc = if ($obj) { [string]$obj.F29 }
			else { "" }
			$okUse = ($posDesc -match '(?i)test|tecnica') -or
			($objDesc -match '(?i)test|tecnica') -or
			([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
		}
		catch { $okUse = $true }
		if ($okUse)
		{
			$PLU = $alternativePLU
			$TestF267 = 7777
			$doInsert = $true
			Write_Log "Using alternative PLU: $PLU with F267: $TestF267" "green"
		}
	}
	
	# ---- If still not chosen, check fallback PLU (INLINE)
	if (-not $doInsert)
	{
		$okUse = $false
		try
		{
			$pos = & $RunSql "SELECT F02 FROM POS_TAB WHERE F01 = '$fallbackPLU'"
			$obj = & $RunSql "SELECT F29 FROM OBJ_TAB WHERE F01 = '$fallbackPLU'"
			$posDesc = if ($pos) { [string]$pos.F02 }
			else { "" }
			$objDesc = if ($obj) { [string]$obj.F29 }
			else { "" }
			$okUse = ($posDesc -match '(?i)test|tecnica') -or
			($objDesc -match '(?i)test|tecnica') -or
			([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
		}
		catch { $okUse = $true }
		if ($okUse)
		{
			$PLU = $fallbackPLU
			$TestF267 = 77777
			$doInsert = $true
			Write_Log "Using fallback PLU: $PLU with F267: $TestF267" "green"
		}
	}
	
	if ($doInsert -and $PLU)
	{
		Write_Log "Deleting existing records for PLU: $PLU and F267: $TestF267" "yellow"
		
		# --- Always delete old rows for the chosen PLU/F267 to keep inserts deterministic
		$deleteQueries = @(
			"DELETE FROM SCL_TAB     WHERE F01 = '$PLU'",
			"DELETE FROM OBJ_TAB     WHERE F01 = '$PLU'",
			"DELETE FROM POS_TAB     WHERE F01 = '$PLU'",
			"DELETE FROM PRICE_TAB   WHERE F01 = '$PLU'",
			"DELETE FROM SCL_TXT_TAB WHERE F267 = $TestF267"
		)
		foreach ($q in $deleteQueries)
		{
			try { & $RunSql $q | Out-Null }
			catch { Write_Log "Error during deletion: $($_.Exception.Message)" "red" }
		}
		
		# --- SCL_TAB
		Write_Log "Inserting into SCL_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO SCL_TAB (F01, F1000, F902, F1001, F253, F258, F264, F267, F1952, F1964, F2581, F2582)
VALUES
('$PLU', 'PAL', 'MANUAL', 1, '$nowFull', 10, 7, $TestF267, 'Test Descriptor 2', '001', 'Test Descriptor 3', 'Test Descriptor 4')
"@ | Out-Null
			Write_Log "SCL_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into SCL_TAB: $($_.Exception.Message)" "red" }
		
		# --- OBJ_TAB
		Write_Log "Inserting into OBJ_TAB..." "yellow"
		try
		{
			$F29 = 'Tecnica Test Item'
			if ($F29.Length -gt 60) { $F29 = $F29.Substring(0, 60) }
			& $RunSql @"
INSERT INTO OBJ_TAB (F01, F902, F1001, F21, F29, F270, F1118, F1959)
VALUES
('$PLU', '00001153', 0, 1, '$F29', 123.45, '001', '001')
"@ | Out-Null
			Write_Log "OBJ_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into OBJ_TAB: $($_.Exception.Message)" "red" }
		
		# --- POS_TAB
		Write_Log "Inserting into POS_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO POS_TAB (F01, F1000, F902, F1001, F02, F09, F79, F80, F82, F104, F115, F176, F178, F217, F1964, F2119)
VALUES
('$PLU', 'PAL', 'MANUAL', 0, 'Tecnica Test Item', '$nowDate', '1', '1', '1', '0', '0', '1', '1', 1.0, '001', '1')
"@ | Out-Null
			Write_Log "POS_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into POS_TAB: $($_.Exception.Message)" "red" }
		
		# --- PRICE_TAB
		Write_Log "Inserting into PRICE_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO PRICE_TAB (F01, F1000, F126, F902, F1001, F21, F30, F31, F113, F1006, F1007, F1008, F1009, F1803)
VALUES
('$PLU', 'PAL', 1, 'MANUAL', 0, 1, 777.77, 1, 'REG', 1, 777.77, '$nowDate', '1858', 1.0)
"@ | Out-Null
			Write_Log "PRICE_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into PRICE_TAB: $($_.Exception.Message)" "red" }
		
		# --- SCL_TXT_TAB
		Write_Log "Inserting into SCL_TXT_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO SCL_TXT_TAB (F267, F1000, F253, F297, F902, F1001, F1836)
VALUES
($TestF267, 'PAL', '$nowFull', 'Ingredients Test', 'MANUAL', 0, 'Tecnica Test Item')
"@ | Out-Null
			Write_Log "SCL_TXT_TAB insertion successful" "green"
		}
		catch { Write_Log "Error inserting into SCL_TXT_TAB: $($_.Exception.Message)" "red" }
		
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
#   Windows Form picker -> run concurrent probes -> write per-group text files and expose results.
#
# Key behavior (PS 5.1-safe):
#   • Selection UI is provided by your Show_Node_Selection_Form (not modified here).
#   • For each selected node:
#       WMI  →  CIM  →  (optional) PS-Session/Registry  →  REG.EXE (starts RemoteRegistry, then restores it)
#   • INI-based fallback augments fields when present (for Lanes/BackOffices).
#   • Writes Info files to Desktop\Lanes|Scales|BackOffices; logs incremental progress.
#   • Populates:
#         $script:LaneHardwareInfo, $script:ScaleHardwareInfo, $script:BackofficeHardwareInfo
#   • Throttles concurrency via $maxConcurrentJobs.
#
# Notes:
#   - Assumes these script vars already exist elsewhere in your toolchain:
#       $script:FunctionResults (LaneNumToMachineName, ScaleCodeToIPInfo, BackofficeNumToMachineName, StoreNumber)
#       $script:DbsPath  (server Office\Dbs path for BackOffice INIs)
#   - Requires admin rights on remotes for WMI/CIM/REG/SC operations.
# ===================================================================================================

function Get_Remote_Machine_Info
{
	Write_Log "`r`n==================== Starting Get_Remote_Machine_Info ====================`r`n" "blue"
	
	# --------------------------- Tunables ---------------------------
	$maxConcurrentJobs = 10
	$wmiTimeoutSeconds = 5
	$cimTimeoutSeconds = 10
	$regTimeoutSeconds = 30
	$usePSRemotingFallback = $false # set $true to try PS-Session registry read before REG.EXE
	
	# --------------------------- Outputs (reset) ---------------------------
	$script:LaneHardwareInfo = $null
	$script:ScaleHardwareInfo = $null
	$script:BackofficeHardwareInfo = $null
	
	# --------------------------- Inputs from your environment ---------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$DbsPath = $script:DbsPath
	
	# --------------------------- Let user pick nodes ---------------------------
	$nodeSelection = Show_Node_Selection_Form -StoreNumber $StoreNumber `
											  -NodeTypes @("Lane", "Scale", "Backoffice") `
											  -Title "Select Nodes to Pull Hardware Info"
	
	if (-not $nodeSelection)
	{
		Write_Log "Get_Remote_Machine_Info cancelled by user." "yellow"
		Write_Log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
		return $false
	}
	
	$selectedLanes = $nodeSelection.Lanes
	$selectedScales = $nodeSelection.Scales
	$selectedBOs = $nodeSelection.Backoffices
	
	if ((-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		Write_Log "No nodes selected. Operation aborted." "yellow"
		Write_Log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
		return $false
	}
	
	# --------------------------- Output folders ---------------------------
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"
	foreach ($dir in @($lanesDir, $scalesDir, $backofficesDir))
	{
		if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory | Out-Null }
	}
	
	# --------------------------- Process 3 categories uniformly ---------------------------
	foreach ($section in @(
			@{ Name = 'Lanes'; Selected = $selectedLanes; Dir = $lanesDir; ScriptVar = 'LaneHardwareInfo'; InfoLinesVar = 'LaneInfoLines'; ResultsVar = 'LaneResults'; FileName = 'Lanes_Info.txt' },
			@{ Name = 'Scales'; Selected = $selectedScales; Dir = $scalesDir; ScriptVar = 'ScaleHardwareInfo'; InfoLinesVar = 'ScaleInfoLines'; ResultsVar = 'ScaleResults'; FileName = 'Scales_Info.txt' },
			@{ Name = 'BackOffices'; Selected = $selectedBOs; Dir = $backofficesDir; ScriptVar = 'BackofficeHardwareInfo'; InfoLinesVar = 'BOInfoLines'; ResultsVar = 'BOResults'; FileName = 'Backoffices_Info.txt' }
		))
	{
		if (-not $section.Selected -or $section.Selected.Count -eq 0) { continue }
		
		Write_Log "Processing $($section.Name) nodes..." "yellow"
		Set-Variable -Name $($section.ResultsVar) -Value @{ }
		Set-Variable -Name $($section.InfoLinesVar) -Value @()
		
		$jobs = @()
		$pending = @{ }
		
		foreach ($sel in $section.Selected)
		{
			# -------- Canonical per-node identity: $NodeNumber (code) + $NodeName (machine/IP) --------
			$NodeNumber = $null
			$NodeName = $null
			
			if ($section.Name -eq 'Lanes')
			{
				if ($sel.PSObject.Properties.Name -contains 'LaneNumber' -and $sel.LaneNumber)
				{
					$NodeNumber = "{0:D3}" -f [int]$sel.LaneNumber
				}
				else { $NodeNumber = "$sel" }
				$NodeName = $LaneNumToMachineName[$NodeNumber]
			}
			elseif ($section.Name -eq 'Scales')
			{
				$NodeNumber = "$sel"
				if ($ScaleCodeToIPInfo.ContainsKey($NodeNumber) -and
					$ScaleCodeToIPInfo[$NodeNumber].PSObject.Properties.Name -contains 'FullIP')
				{
					$NodeName = $ScaleCodeToIPInfo[$NodeNumber].FullIP
				}
			}
			elseif ($section.Name -eq 'BackOffices')
			{
				if ($sel.PSObject.Properties.Name -contains 'BONumber' -and $sel.BONumber)
				{
					$NodeNumber = "{0:D3}" -f [int]$sel.BONumber
				}
				else { $NodeNumber = "$sel" }
				$NodeName = $BackofficeNumToMachineName[$NodeNumber]
			}
			
			if (-not $NodeName)
			{
				Write_Log "Skipping $($section.Name) $NodeNumber (no NodeName/mapping found)" "red"
				continue
			}
			
			# -------- Optional INI path (used to augment fields) --------
			$iniPattern = "INFO_${StoreNumber}${NodeNumber}_SMSStart.ini"
			if ($section.Name -eq 'Lanes') { $iniPath = Join-Path "\\$NodeName\storeman\office\dbs" $iniPattern }
			elseif ($section.Name -eq 'BackOffices') { $iniPath = Join-Path $DbsPath                         $iniPattern }
			else { $iniPath = $null }
			
			# -------- Per-node job (throttled outside) --------
			$job = Start-Job -ArgumentList $NodeName, $NodeNumber, $iniPath, $wmiTimeoutSeconds, $cimTimeoutSeconds, $regTimeoutSeconds, $usePSRemotingFallback `
							 -ScriptBlock {
				param (
					$NodeName,
					$NodeNumber,
					$iniPath,
					$wmiTimeoutSeconds,
					$cimTimeoutSeconds,
					$regTimeoutSeconds,
					$usePSRemotingFallback
				)
				
				# --- 0) Result object scaffold ---
				$info = [PSCustomObject]@{
					Success			    = $false
					SystemManufacturer  = $null
					SystemProductName   = $null
					CPU				    = $null
					RAM				    = $null
					OSInfo			    = $null
					BIOS			    = $null
					Method			    = $null
					Error			    = $null
					MachineNameOverride = $null
				}
				
				# --- 0.1) Remember RemoteRegistry state up-front (so we can restore) ---
				$originalState = $null
				try
				{
					$stateLine = sc.exe "\\$NodeName" query RemoteRegistry 2>$null | Select-String "STATE" | ForEach-Object { $_.Line }
					$startTypeLine = sc.exe "\\$NodeName" qc    RemoteRegistry 2>$null | Select-String "START_TYPE" | ForEach-Object { $_.Line }
					$originalState = @{ StateLine = $stateLine; StartTypeLine = $startTypeLine }
				}
				catch { }
				
				# --- 1) WMI (fast) ---
				if (-not (Test-Connection -ComputerName $NodeName -Count 1 -Quiet -ErrorAction SilentlyContinue))
				{
					$info.Error = "Offline or unreachable."
					$info.Method = "Offline"
				}
				else
				{
					$wmiJob = Start-Job -ScriptBlock {
						param ($NodeName)
						try
						{
							$sys = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $NodeName -ErrorAction SilentlyContinue
							$cpu = Get-WmiObject -Class Win32_Processor -ComputerName $NodeName -ErrorAction SilentlyContinue | Select-Object -First 1
							$os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $NodeName -ErrorAction SilentlyContinue
							
							# BIOS
							$biosVerOut = $null
							try
							{
								$bios = Get-WmiObject -Class Win32_BIOS -ComputerName $NodeName -ErrorAction SilentlyContinue
								if ($bios)
								{
									if ($bios.SMBIOSBIOSVersion) { $biosVerOut = $bios.SMBIOSBIOSVersion }
									elseif ($bios.BIOSVersion) { $biosVerOut = ($bios.BIOSVersion | Where-Object { $_ } | Select-Object -First 1) }
									elseif ($bios.Version) { $biosVerOut = $bios.Version }
									if ($bios.ReleaseDate)
									{
										try
										{
											$dt = [System.Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
											if ($dt)
											{
												if ($biosVerOut) { $biosVerOut = "$biosVerOut ($($dt.ToString('yyyy-MM-dd')))" }
												else { $biosVerOut = $dt.ToString('yyyy-MM-dd') }
											}
										}
										catch { }
									}
								}
							}
							catch { }
							
							if ($sys -and $sys.Manufacturer -and $sys.Model)
							{
								[PSCustomObject]@{
									SystemManufacturer = $sys.Manufacturer
									SystemProductName  = $sys.Model
									CPU			       = $cpu.Name
									RAM			       = [math]::Round($sys.TotalPhysicalMemory / 1GB, 1)
									OSInfo			   = "$($os.Caption) ($($os.Version))"
									BIOS			   = $biosVerOut
									Method			   = "WMI"
								}
							}
							else { $null }
						}
						catch { $null }
					} -ArgumentList $NodeName
					
					if (Wait-Job $wmiJob -Timeout $wmiTimeoutSeconds)
					{
						$wmiResult = Receive-Job $wmiJob 2>$null
						Remove-Job  $wmiJob -Force -ErrorAction SilentlyContinue
					}
					else
					{
						Stop-Job    $wmiJob -ErrorAction SilentlyContinue
						Remove-Job  $wmiJob -Force -ErrorAction SilentlyContinue
						$wmiResult = $null
					}
					
					if ($wmiResult -and $wmiResult.SystemManufacturer -and $wmiResult.SystemProductName)
					{
						$info.SystemManufacturer = $wmiResult.SystemManufacturer
						$info.SystemProductName = $wmiResult.SystemProductName
						$info.CPU = $wmiResult.CPU
						$info.RAM = $wmiResult.RAM
						$info.OSInfo = $wmiResult.OSInfo
						if ($wmiResult.BIOS) { $info.BIOS = $wmiResult.BIOS }
						$info.Method = "WMI"
						$info.Success = $true
					}
					else
					{
						$info.Error = "WMI failed (access issue or null)"
					}
				}
				
				# --- 2) CIM (if WMI failed) ---
				if (-not $info.Success)
				{
					$cimJob = Start-Job -ScriptBlock {
						param ($NodeName)
						try
						{
							$session = New-CimSession -ComputerName $NodeName -ErrorAction SilentlyContinue
							if ($session)
							{
								$sys = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem  2>$null
								$cpu = Get-CimInstance -CimSession $session -ClassName Win32_Processor       2>$null | Select-Object -First 1
								$os = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem  2>$null
								
								# BIOS
								$biosVerOut = $null
								try
								{
									$bios = Get-CimInstance -CimSession $session -ClassName Win32_BIOS 2>$null
									if ($bios)
									{
										if ($bios.SMBIOSBIOSVersion) { $biosVerOut = $bios.SMBIOSBIOSVersion }
										elseif ($bios.BIOSVersion) { $biosVerOut = ($bios.BIOSVersion | Where-Object { $_ } | Select-Object -First 1) }
										elseif ($bios.Version) { $biosVerOut = $bios.Version }
										if ($bios.ReleaseDate)
										{
											try
											{
												$dt = [System.Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
												if ($dt)
												{
													if ($biosVerOut) { $biosVerOut = "$biosVerOut ($($dt.ToString('yyyy-MM-dd')))" }
													else { $biosVerOut = $dt.ToString('yyyy-MM-dd') }
												}
											}
											catch { }
										}
									}
								}
								catch { }
								
								Remove-CimSession $session 2>$null
								if ($sys -and $sys.Manufacturer -and $sys.Model)
								{
									[PSCustomObject]@{
										SystemManufacturer = $sys.Manufacturer
										SystemProductName  = $sys.Model
										CPU			       = $cpu.Name
										RAM			       = [math]::Round($sys.TotalPhysicalMemory / 1GB, 1)
										OSInfo			   = "$($os.Caption) ($($os.Version))"
										BIOS			   = $biosVerOut
										Method			   = "CIM"
									}
								}
								else { $null }
							}
							else { $null }
						}
						catch { $null }
					} -ArgumentList $NodeName
					
					if (Wait-Job $cimJob -Timeout $cimTimeoutSeconds)
					{
						$cimResult = Receive-Job $cimJob 2>$null
						Remove-Job  $cimJob -Force -ErrorAction SilentlyContinue
					}
					else
					{
						Stop-Job    $cimJob -ErrorAction SilentlyContinue
						Remove-Job  $cimJob -Force -ErrorAction SilentlyContinue
						$cimResult = $null
					}
					
					if ($cimResult -and $cimResult.SystemManufacturer -and $cimResult.SystemProductName)
					{
						$info.SystemManufacturer = $cimResult.SystemManufacturer
						$info.SystemProductName = $cimResult.SystemProductName
						$info.CPU = $cimResult.CPU
						$info.RAM = $cimResult.RAM
						$info.OSInfo = $cimResult.OSInfo
						if ($cimResult.BIOS) { $info.BIOS = $cimResult.BIOS }
						$info.Method = "CIM"
						$info.Success = $true
					}
					else
					{
						$info.Error = "CIM failed (access issue or null)"
					}
				}
				
				# --- 3) PS-Remoting Registry (optional middle step) ---
				if (-not $info.Success -and $usePSRemotingFallback)
				{
					$regResult = $null
					try
					{
						$session = New-PSSession -ComputerName $NodeName -ErrorAction SilentlyContinue
						if ($session)
						{
							$regResult = Invoke-Command -Session $session -ScriptBlock {
								$out = [PSCustomObject]@{
									SystemManufacturer = $null
									SystemProductName  = $null
									CPU			       = $null
									OSInfo			   = $null
									BIOS			   = $null
									Success		       = $false
									Error			   = $null
									Method			   = 'REG(PSRemoting)'
								}
								try
								{
									$manuf = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemManufacturer -ErrorAction SilentlyContinue
									$prod = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemProductName -ErrorAction SilentlyContinue
									$cpu = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -Name ProcessorNameString -ErrorAction SilentlyContinue
									$os = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
									
									# BIOS pieces
									$biosVendor = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSVendor -ErrorAction SilentlyContinue).BIOSVendor
									$biosVerStr = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSVersion -ErrorAction SilentlyContinue).BIOSVersion
									if (-not $biosVerStr)
									{
										$biosVerStr = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemBiosVersion -ErrorAction SilentlyContinue).SystemBiosVersion
									}
									$biosDate = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSReleaseDate -ErrorAction SilentlyContinue).BIOSReleaseDate
									
									# Build OSInfo
									$osInfo = $null
									if ($os -and $os.ProductName) { $osInfo = $os.ProductName }
									if ($os -and $os.DisplayVersion)
									{
										if ($osInfo) { $osInfo = "$osInfo ($($os.DisplayVersion))" }
										else { $osInfo = $os.DisplayVersion }
									}
									elseif ($os -and $os.CurrentBuild)
									{
										if ($osInfo) { $osInfo = "$osInfo (Build $($os.CurrentBuild))" }
										else { $osInfo = "Build $($os.CurrentBuild)" }
									}
									
									# BIOS assemble
									$biosOut = $null
									if ($biosVendor -and $biosVerStr) { $biosOut = "$biosVendor $biosVerStr" }
									elseif ($biosVerStr) { $biosOut = $biosVerStr }
									elseif ($biosVendor) { $biosOut = $biosVendor }
									if ($biosDate)
									{
										if ($biosOut) { $biosOut = "$biosOut ($biosDate)" }
										else { $biosOut = $biosDate }
									}
									
									$out.SystemManufacturer = $manuf.SystemManufacturer
									$out.SystemProductName = $prod.SystemProductName
									$out.CPU = $cpu.ProcessorNameString
									$out.OSInfo = $osInfo
									$out.BIOS = $biosOut
									$out.Success = $true
								}
								catch { $out.Error = $_.Exception.Message }
								return $out
							} 2>$null
							Remove-PSSession $session -ErrorAction SilentlyContinue
						}
					}
					catch
					{
						$regResult = [PSCustomObject]@{
							SystemManufacturer = $null
							SystemProductName  = $null
							CPU			       = $null
							OSInfo			   = $null
							BIOS			   = $null
							Success		       = $false
							Error			   = "PS-Remoting registry query failed"
							Method			   = 'REG(PSRemoting)'
						}
					}
					
					if ($regResult -and $regResult.Success)
					{
						if (-not $info.SystemManufacturer) { $info.SystemManufacturer = $regResult.SystemManufacturer }
						if (-not $info.SystemProductName) { $info.SystemProductName = $regResult.SystemProductName }
						if (-not $info.CPU) { $info.CPU = $regResult.CPU }
						if (-not $info.OSInfo) { $info.OSInfo = $regResult.OSInfo }
						if (-not $info.BIOS -and $regResult.BIOS) { $info.BIOS = $regResult.BIOS }
						$info.Method = $regResult.Method
						$info.Success = $true
						$info.Error = $null
					}
					elseif (-not $info.Error)
					{
						if ($regResult) { $info.Error = $regResult.Error; $info.Method = 'REG(PSRemoting)' }
					}
				}
				
				# --- 4) REG.EXE (starts RemoteRegistry if needed, then restores) ---
				if (-not $info.Success)
				{
					$wasRunning = $false
					$wasDisabled = $false
					try
					{
						$state = sc.exe "\\$NodeName" query RemoteRegistry 2>$null | Select-String 'STATE'
						$start = sc.exe "\\$NodeName" qc    RemoteRegistry 2>$null | Select-String 'START_TYPE'
						if ($state -and $state.Line -match 'RUNNING') { $wasRunning = $true }
						if ($start -and $start.Line -match 'DISABLED') { $wasDisabled = $true }
						if ($wasDisabled) { sc.exe "\\$NodeName" config RemoteRegistry start= demand 2>$null | Out-Null }
						if (-not $wasRunning) { sc.exe "\\$NodeName" start  RemoteRegistry 2>$null | Out-Null }
						
						# tiny inline getter: returns only the DATA (no "REG_SZ")
						$getVal = {
							param ($hive,
								$path,
								$name)
							$patternName = [regex]::Escape($name)
							$raw = reg.exe QUERY "\\$NodeName\$hive\$path" /v $name 2>$null
							if (-not $raw) { return $null }
							$line = $raw | Select-String -Pattern "^\s*$patternName\s+REG_\w+\s+.+$" | Select-Object -First 1
							if (-not $line) { return $null }
							$data = $line.Line -replace "^\s*$patternName\s+REG_\w+\s+", ""
							$data.Trim()
						}
						
						$manuf = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS'               'SystemManufacturer'
						$prod = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS'               'SystemProductName'
						$cpu = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\CentralProcessor\0' 'ProcessorNameString'
						$osN = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'ProductName'
						$osB = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'DisplayVersion'
						$osV = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'CurrentBuild'
						
						# BIOS from REG.EXE
						$biosVendor = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSVendor'
						$biosVerStr = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSVersion'
						if (-not $biosVerStr)
						{
							$biosVerStr = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'SystemBiosVersion'
						}
						$biosDate = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSReleaseDate'
						if ($biosVerStr)
						{
							$biosVerStr = $biosVerStr -replace '\x00', ' '
							$biosVerStr = ($biosVerStr -split '\s{2,}' | Where-Object { $_ }) -join ' '
							$biosVerStr = $biosVerStr.Trim()
						}
						$biosOut = $null
						if ($biosVendor -and $biosVerStr) { $biosOut = "$biosVendor $biosVerStr" }
						elseif ($biosVerStr) { $biosOut = $biosVerStr }
						elseif ($biosVendor) { $biosOut = $biosVendor }
						if ($biosDate)
						{
							if ($biosOut) { $biosOut = "$biosOut ($biosDate)" }
							else { $biosOut = $biosDate }
						}
						
						# If CurrentBuild came as "0x... (7601)" keep the number in parentheses
						if ($osV -and $osV -match '\((\d+)\)') { $osV = $matches[1] }
						
						# build a clean OS string
						$osInfo = $null
						if ($osN) { $osInfo = $osN }
						if ($osB)
						{
							if ($osInfo) { $osInfo = "$osInfo ($osB)" }
							else { $osInfo = $osB }
						}
						elseif ($osV)
						{
							if ($osInfo) { $osInfo = "$osInfo (Build $osV)" }
							else { $osInfo = "Build $osV" }
						}
						
						if ($manuf -or $prod -or $cpu -or $osInfo -or $biosOut)
						{
							if (-not $info.SystemManufacturer) { $info.SystemManufacturer = $manuf }
							if (-not $info.SystemProductName) { $info.SystemProductName = $prod }
							if (-not $info.CPU) { $info.CPU = $cpu }
							if (-not $info.OSInfo) { $info.OSInfo = $osInfo }
							if (-not $info.BIOS -and $biosOut) { $info.BIOS = $biosOut }
							$info.Method = "REG.EXE"
							$info.Success = $true
							$info.Error = $null
						}
						else
						{
							$info.Method = "REG.EXE"
							$info.Success = $false
							$info.Error = "REG queries returned no data"
						}
					}
					catch
					{
						$info.Method = "REG.EXE"
						$info.Success = $false
						$info.Error = "REG queries failed"
					}
					finally
					{
						try
						{
							if (-not $wasRunning) { sc.exe "\\$NodeName" stop RemoteRegistry 2>$null | Out-Null }
							if ($wasDisabled) { sc.exe "\\$NodeName" config RemoteRegistry start= disabled 2>$null | Out-Null }
						}
						catch { }
					}
				}
				
				# --- 5) INI augmentation (optional; fills gaps if any fields still null) ---
				$returnInfo = [PSCustomObject]@{
					Machine    = $NodeName
					NodeNumber = $NodeNumber
					Info	   = $info
					IniFound   = $false
					IniPath    = $null
				}
				try
				{
					if ($iniPath)
					{
						$iniFolder = Split-Path $iniPath -Parent
						$iniLeaf = Split-Path $iniPath -Leaf
						$iniFile = Get-ChildItem -Path $iniFolder -Filter $iniLeaf -ErrorAction SilentlyContinue |
						Sort-Object LastWriteTime -Descending |
						Select-Object -First 1
						if ($iniFile -and (Test-Path $iniFile.FullName))
						{
							$returnInfo.IniFound = $true
							$returnInfo.IniPath = $iniFile.FullName
							$iniLines = Get-Content $iniFile.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
							
							# tiny parser
							$sections = @{ }
							$curSec = ""
							foreach ($line in $iniLines)
							{
								if ($line -match '^\[(.+)\]$')
								{
									$curSec = $matches[1]; $sections[$curSec] = @{ }
								}
								elseif ($line -match '^\s*([^=]+?)\s*=\s*(.*)$' -and $curSec)
								{
									$sections[$curSec][$matches[1].Trim()] = $matches[2].Trim()
								}
							}
							if (-not $info.CPU -and $sections.ContainsKey('PROCESSOR') -and $sections['PROCESSOR'].ContainsKey('Cores'))
							{
								$cores = $sections['PROCESSOR']['Cores']
								$arch = $sections['PROCESSOR']['Architecture']
								if ($arch) { $info.CPU = "$cores cores ($arch)" }
								else { $info.CPU = "$cores cores" }
							}
							if ($info.RAM -eq $null -and $sections.ContainsKey('Memory') -and $sections['Memory'].ContainsKey('PhysicalMemory'))
							{
								$ramMb = $sections['Memory']['PhysicalMemory']
								if ($ramMb -match '^\d+$') { $info.RAM = [math]::Round([double]$ramMb/1024, 1) }
								else { $info.RAM = $ramMb }
							}
							if (-not $info.OSInfo -and $sections.ContainsKey('OperatingSystem') -and $sections['OperatingSystem'].ContainsKey('ProductName'))
							{
								$info.OSInfo = $sections['OperatingSystem']['ProductName']
							}
						}
					}
				}
				catch { }
				
				return $returnInfo
			}
			
			$jobs += $job
			$pending[$job.Id] = $NodeName
			
			# Throttle
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
						
						# Build display line
						$line = "Machine Name: $remoteName |"
						if ($info.Success)
						{
							$line += " Manufacturer: $($info.SystemManufacturer) | Model: $($info.SystemProductName) | CPU: $($info.CPU)"
							if ($info.RAM -ne $null) { $line += " | RAM: $($info.RAM) GB" }
							if ($info.OSInfo) { $line += " | OS: $($info.OSInfo)" }
							$biosText = $info.BIOS; if (-not $biosText) { $biosText = "Unknown" }
							$line += " | BIOS: $biosText"
							$line += " | Method: $($info.Method)"
							Write_Log "Processed $remoteName ($($section.Name)): Success [$($info.Method)]" "green"
						}
						elseif ($info.Error)
						{
							$line += " [Hardware info unavailable] Error: $($info.Error)"
							$biosText = $info.BIOS; if (-not $biosText) { $biosText = "Unknown" }
							$line += " | BIOS: $biosText"
							$line += " | Method: $($info.Method)"
							Write_Log "Processed $remoteName ($($section.Name)): Error [$($info.Method)] - $($info.Error)" "red"
						}
						else
						{
							$line += " [No hardware info found]"
							$biosText = $info.BIOS; if (-not $biosText) { $biosText = "Unknown" }
							$line += " | BIOS: $biosText"
							$line += " | Method: $($info.Method)"
							Write_Log "Processed $remoteName ($($section.Name)): No info [$($info.Method)]" "yellow"
						}
						
						# Append to info-lines
						$infolines = Get-Variable -Name $($section.InfoLinesVar) -ValueOnly
						$infolines += $line
						Set-Variable -Name $($section.InfoLinesVar) -Value $infolines
						
						# Store rich info for programmatic use
						$results = Get-Variable -Name $($section.ResultsVar) -ValueOnly
						$results[$remoteName] = $info
						Set-Variable -Name $($section.ResultsVar) -Value $results
						
						# Cleanup
						Stop-Job $j -ErrorAction SilentlyContinue
						Remove-Job $j -Force -ErrorAction SilentlyContinue
						$jobs = $jobs | Where-Object { $_.Id -ne $j.Id }
						$pending.Remove($j.Id)
					}
				}
				else { break }
			}
		} # end foreach selected
		
		# Drain remaining jobs
		if ($jobs.Count -gt 0)
		{
			Wait-Job -Job $jobs -Timeout 60 | Out-Null
			foreach ($j in $jobs)
			{
				$remoteName = $pending[$j.Id]
				$result = Receive-Job $j 2>$null
				$info = $result.Info
				
				$line = "Machine Name: $remoteName |"
				if ($info.Success)
				{
					$line += " Manufacturer: $($info.SystemManufacturer) | Model: $($info.SystemProductName) | CPU: $($info.CPU)"
					if ($info.RAM -ne $null) { $line += " | RAM: $($info.RAM) GB" }
					if ($info.OSInfo) { $line += " | OS: $($info.OSInfo)" }
					$biosText = $info.BIOS; if (-not $biosText) { $biosText = "Unknown" }
					$line += " | BIOS: $biosText"
					$line += " | Method: $($info.Method)"
					Write_Log "Processed $remoteName ($($section.Name)): Success [$($info.Method)]" "green"
				}
				elseif ($info.Error)
				{
					$line += " [Hardware info unavailable] Error: $($info.Error)"
					$biosText = $info.BIOS; if (-not $biosText) { $biosText = "Unknown" }
					$line += " | BIOS: $biosText"
					$line += " | Method: $($info.Method)"
					Write_Log "Processed $remoteName ($($section.Name)): Error [$($info.Method)] - $($info.Error)" "red"
				}
				else
				{
					$line += " [No hardware info found]"
					$biosText = $info.BIOS; if (-not $biosText) { $biosText = "Unknown" }
					$line += " | BIOS: $biosText"
					$line += " | Method: $($info.Method)"
					Write_Log "Processed $remoteName ($($section.Name)): No info [$($info.Method)]" "yellow"
				}
				
				$infolines = Get-Variable -Name $($section.InfoLinesVar) -ValueOnly
				$infolines += $line
				Set-Variable -Name $($section.InfoLinesVar) -Value $infolines
				
				$results = Get-Variable -Name $($section.ResultsVar) -ValueOnly
				$results[$remoteName] = $info
				Set-Variable -Name $($section.ResultsVar) -Value $results
				
				Stop-Job $j -ErrorAction SilentlyContinue
				Remove-Job $j -Force -ErrorAction SilentlyContinue
			}
		}
		
		# Write file (stable sorting: IPs zero-padded; otherwise lexical)
		$infolines = Get-Variable -Name $($section.InfoLinesVar) -ValueOnly
		if ($infolines.Count)
		{
			$sortedLines = $infolines | Sort-Object {
				if ($_ -match "^Machine Name: ([^|]+)")
				{
					$name = $matches[1].Trim()
					if ($name -match '^\d{1,3}(\.\d{1,3}){3}$')
					{
						$oct = $name -split '\.'
						'{0:D3}.{1:D3}.{2:D3}.{3:D3}' -f [int]$oct[0], [int]$oct[1], [int]$oct[2], [int]$oct[3]
					}
					else { $name }
				}
				else { $_ }
			}
			$filePath = Join-Path $section.Dir $section.FileName
			$sortedLines -join "`r`n" | Set-Content -Path $filePath -Encoding Default
			Write_Log "Exported $($section.Name) info to $filePath" "green"
		}
		
		# Publish programmatic map
		$mapOut = Get-Variable -Name $($section.ResultsVar) -ValueOnly
		if ($section.Name -eq 'Lanes') { $script:LaneHardwareInfo = $mapOut }
		elseif ($section.Name -eq 'Scales') { $script:ScaleHardwareInfo = $mapOut }
		elseif ($section.Name -eq 'BackOffices') { $script:BackofficeHardwareInfo = $mapOut }
		
		Write_Log "Completed processing $($section.Name).`r`n" "green"
	} # end sections
	
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
		[Parameter(Mandatory, Position = 0, HelpMessage = "The directory path where files and folders will be deleted.")]
		[ValidateNotNullOrEmpty()]
		[string]$Path,
		[Parameter(HelpMessage = "Specific file/folder patterns to delete. Wildcards supported.")]
		[string[]]$SpecifiedFiles,
		[Parameter(HelpMessage = "Patterns to exclude from deletion. Wildcards supported.")]
		[string[]]$Exclusions,
		[Parameter(HelpMessage = "Run the deletion as a background job.")]
		[switch]$AsJob
	)
	
	if ($AsJob)
	{
		$scriptBlock = {
			param ($Path,
				$SpecifiedFiles,
				$Exclusions)
			
			$deletedCount = 0
			$rp = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
			if (-not $rp) { return 0 }
			$targetPath = $rp.Path
			$gciBase = Join-Path $targetPath '*'
			
			try
			{
				# Build Get-ChildItem params safely
				$gciParams = @{ Path = $gciBase; Recurse = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
				if ($SpecifiedFiles -and $SpecifiedFiles.Count -gt 0) { $gciParams['Include'] = $SpecifiedFiles }
				if ($Exclusions -and $Exclusions.Count -gt 0) { $gciParams['Exclude'] = $Exclusions }
				
				$items = Get-ChildItem @gciParams
				if ($items)
				{
					foreach ($it in $items)
					{
						try
						{
							Remove-Item -LiteralPath $it.FullName -Force -Recurse -ErrorAction Stop
							$deletedCount++
						}
						catch { }
					}
				}
				return $deletedCount
			}
			catch
			{
				return $deletedCount
			}
		}
		
		return Start-Job -ScriptBlock $scriptBlock -ArgumentList $Path, $SpecifiedFiles, $Exclusions
	}
	
	# -------- Synchronous path --------
	$deletedCount = 0
	$rp = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
	if (-not $rp)
	{
		Write_Log "The specified path '$Path' does not exist." "red"
		return 0
	}
	$targetPath = $rp.Path
	$gciBase = Join-Path $targetPath '*'
	
	try
	{
		$gciParams = @{ Path = $gciBase; Recurse = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
		if ($SpecifiedFiles -and $SpecifiedFiles.Count -gt 0) { $gciParams['Include'] = $SpecifiedFiles }
		if ($Exclusions -and $Exclusions.Count -gt 0) { $gciParams['Exclude'] = $Exclusions }
		
		$items = Get-ChildItem @gciParams
		if ($items)
		{
			foreach ($it in $items)
			{
				try
				{
					Remove-Item -LiteralPath $it.FullName -Force -Recurse -ErrorAction Stop
					$deletedCount++
				}
				catch { }
			}
		}
		else
		{
			Write_Log "No items matched in '$targetPath'." "yellow"
		}
		
		Write_Log "Total items deleted: $deletedCount" "blue"
		return $deletedCount
	}
	catch
	{
		Write_Log "An error occurred during deletion. $_" "red"
		return $deletedCount
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
	
	# Accept all supported forms (lane number, machine name, POSxxx, etc)
	$selectedLanes = $selection.Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($LaneMachineNames.ContainsKey($_)) { $_ }
		elseif ($LaneNumToMachineName.ContainsKey($_)) { $_ }
		else { $_ }
	}
	# Make sure only valid 3-digit lane numbers that are in the mapping
	$selectedLanes = $selectedLanes | Where-Object { $LaneNumToMachineName.ContainsKey($_) }
	
	if (-not $selectedLanes -or $selectedLanes.Count -eq 0)
	{
		Write_Log "No valid lanes selected for processing." "Yellow"
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
#                               FUNCTION: Organize_Desktop_Items
# ---------------------------------------------------------------------------------------------------
# Purpose:
#   Moves non-system/non-excluded items from the current user's Desktop into a single folder,
#   ensuring specified excluded folders exist. Designed to be called independently of other
#   system tuning steps.
#
# Parameters:
#   - UnorganizedFolderName [string]
#       Target folder name on Desktop where unorganized items will be moved. Default: "Unorganized Items"
#   - AdditionalExclusions [string[]]
#       Optional extra file/folder names (case-insensitive) to exclude from moving.
#
# Notes:
#   - Uses Write_Log for consistent logging.
#   - Skips common system shortcuts and script launchers.
#   - Does NOT restart Explorer (kept side-effect free).
# ---------------------------------------------------------------------------------------------------

function Organize_Desktop_Items
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$UnorganizedFolderName = "Unorganized Items",
		[Parameter(Mandatory = $false)]
		[string[]]$AdditionalExclusions
	)
	
	Write_Log "`r`n==================== Starting Organize_Desktop_Items ====================`r`n" "blue"
	
	try
	{
		# -- Resolve Desktop path
		$DesktopPath = [Environment]::GetFolderPath("Desktop")
		
		# -- Compute the destination folder path
		$UnorganizedFolder = Join-Path -Path $DesktopPath -ChildPath $UnorganizedFolderName
		
		# -- Best-effort current script/launcher name (used for exclusions)
		#    $scriptName could be defined elsewhere; fall back to the current command name if available
		$scriptName = if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Name)
		{
			$MyInvocation.MyCommand.Name
		}
		else
		{
			"TBS_Maintenance_Script.exe"
		}
		
		# -- Known system/launcher items to preserve on Desktop (case-insensitive match by name)
		$systemIcons = @(
			"This PC.lnk", "Network.lnk", "Control Panel.lnk", "Recycle Bin.lnk", "User's Files.lnk",
			"Execute(TBS_Maintenance_Script).bat", "Execute(MiniGhost).bat",
			"TBS_Maintenance_Script.exe", "MiniGhost.exe",
			$scriptName
		)
		
		# -- Default excluded folders that must remain on Desktop
		$defaultExcludedFolders = @("Lanes", "Scales", "BackOffices", $UnorganizedFolderName)
		
		# -- Merge additional exclusions if provided
		$excludedSet = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
		foreach ($n in $systemIcons) { [void]$excludedSet.Add($n) }
		foreach ($n in $defaultExcludedFolders) { [void]$excludedSet.Add($n) }
		if ($AdditionalExclusions)
		{
			foreach ($n in $AdditionalExclusions) { if ($n) { [void]$excludedSet.Add($n) } }
		}
		
		# -- Ensure destination and commonly-used excluded folders exist
		foreach ($ensure in ($defaultExcludedFolders | Select-Object -Unique))
		{
			$ensurePath = Join-Path -Path $DesktopPath -ChildPath $ensure
			if (-not (Test-Path -LiteralPath $ensurePath))
			{
				New-Item -ItemType Directory -Path $ensurePath | Out-Null
				Write_Log "Created folder: $ensurePath" "green"
			}
		}
		
		# -- Ensure the target "Unorganized" folder exists
		if (-not (Test-Path -LiteralPath $UnorganizedFolder))
		{
			New-Item -ItemType Directory -Path $UnorganizedFolder | Out-Null
			Write_Log "Created folder: $UnorganizedFolder" "green"
		}
		else
		{
			Write_Log "Folder already exists: $UnorganizedFolder" "cyan"
		}
		
		# -- Enumerate Desktop items; keep both files and folders
		#    We exclude *.lnk shortcuts unless they are explicitly allowed (systemIcons list).
		$desktopItems = Get-ChildItem -LiteralPath $DesktopPath -Force |
		Where-Object {
			# Skip items explicitly excluded by name
			-not $excludedSet.Contains($_.Name) -and
			# Skip .lnk shortcuts unless the name is explicitly allowed
			(-not ($_.Extension -ieq ".lnk"))
		}
		
		# -- Move everything else into the Unorganized folder
		foreach ($item in $desktopItems)
		{
			try
			{
				Move-Item -LiteralPath $item.FullName -Destination $UnorganizedFolder -Force
				Write_Log "Moved item: $($item.Name)" "green"
			}
			catch
			{
				Write_Log "Failed to move item: $($item.Name). Error: $_" "red"
			}
		}
		
		Write_Log "Desktop organization complete." "green"
		Write_Log "==================== Organize_Desktop_Items Completed ====================" "blue"
	}
	catch
	{
		Write_Log "Unexpected error in Organize_Desktop_Items: $_" "red"
	}
}


# ===================================================================================================
#                               FUNCTION: Configure_System_Settings
# ---------------------------------------------------------------------------------------------------
# Description (UPDATED - Desktop organizing removed):
#   Configures power plan, services, and visual settings. No Desktop file moves here.
#
#   1) Power Settings:
#       - High Performance plan
#       - Sleep disabled (AC/DC)
#       - Minimum processor performance to 100% (AC/DC)
#       - Monitor off after 15 minutes (AC)
#
#   2) Services:
#       - Sets fdPHost, FDResPub, SSDPSRV, upnphost to Automatic and starts them
#
#   3) Visual Settings:
#       - Show thumbnails instead of icons
#       - Font smoothing + ClearType
#       - Restarts Explorer to ensure settings apply
#
# Parameters:
#   (none)
#
# Notes:
#   - The prior "Organize Desktop" section has been removed from this function.
#   - Call Organize_Desktop_Items separately when desired.
# ---------------------------------------------------------------------------------------------------

function Configure_System_Settings
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Configure_System_Settings ====================`r`n" "blue"
	
	try
	{
		# ==========================================================================================
		# [CHANGE] Removed: Desktop organizing block
		#          Moved into new function: Organize_Desktop_Items
		# ==========================================================================================
		
		# ===========================================
		# 1. Configure Power Settings
		# ===========================================
		Write_Log "`r`nConfiguring power plan and performance settings..." "blue"
		
		# High Performance plan GUID
		$highPerfGUID = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
		
		try
		{
			powercfg /s $highPerfGUID
			Write_Log "Power plan set to High Performance." "green"
		}
		catch
		{
			Write_Log "Failed to set power plan to High Performance. Error: $_" "red"
		}
		
		# Disable sleep (AC/DC)
		try
		{
			powercfg /change standby-timeout-ac 0
			powercfg /change standby-timeout-dc 0
			Write_Log "System sleep disabled for AC/DC." "green"
		}
		catch
		{
			Write_Log "Failed to disable system sleep. Error: $_" "red"
		}
		
		# Minimum processor performance to 100% (AC/DC)
		try
		{
			# SUB_PROCESSOR         = 54533251-82be-4824-96c1-47b60b740d00
			# PROCTHROTTLEMIN       = 893dee8e-2bef-41e0-89c6-b55d0929964c
			powercfg /setacvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setdcvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setactive $highPerfGUID
			Write_Log "Minimum processor performance set to 100% (AC/DC)." "green"
		}
		catch
		{
			Write_Log "Failed to set processor performance. Error: $_" "red"
		}
		
		# Turn off screen after 15 minutes (AC)
		try
		{
			powercfg /change monitor-timeout-ac 15
			Write_Log "Monitor timeout (AC) set to 15 minutes." "green"
		}
		catch
		{
			Write_Log "Failed to set monitor timeout. Error: $_" "red"
		}
		
		Write_Log "Power and performance settings complete." "green"
		
		# ===========================================
		# 2. Configure Services
		# ===========================================
		Write_Log "`r`nConfiguring services to start automatically..." "blue"
		
		$servicesToConfigure = @("fdPHost", "FDResPub", "SSDPSRV", "upnphost")
		foreach ($service in $servicesToConfigure)
		{
			try
			{
				Set-Service -Name $service -StartupType Automatic -ErrorAction Stop
				Write_Log "Set service '$service' to Automatic." "green"
				
				$svc = Get-Service -Name $service -ErrorAction Stop
				if ($svc.Status -ne 'Running')
				{
					Start-Service -Name $service -ErrorAction Stop
					Write_Log "Started service '$service'." "green"
				}
				else
				{
					Write_Log "Service '$service' is already running." "cyan"
				}
			}
			catch
			{
				Write_Log "Failed to configure service '$service'. Error: $_" "red"
			}
		}
		
		Write_Log "Service configuration complete." "green"
		
		# ===========================================
		# 3. Configure Visual Settings
		# ===========================================
		Write_Log "`r`nConfiguring visual settings..." "blue"
		
		try
		{
			# Show thumbnails (IconsOnly = 0)
			Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
							 -Name "IconsOnly" -Value 0 -Type DWord -ErrorAction Stop
			Write_Log "Enabled 'Show thumbnails instead of icons'." "green"
		}
		catch
		{
			Write_Log "Failed to enable thumbnails. Error: $_" "red"
		}
		
		try
		{
			# Font smoothing + ClearType
			Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothing" -Value "2" -Type String -ErrorAction Stop
			Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothingType" -Value 2 -Type DWord -ErrorAction Stop
			Write_Log "Enabled 'Smooth edges of screen fonts' (ClearType)." "green"
		}
		catch
		{
			Write_Log "Failed to enable font smoothing. Error: $_" "red"
		}
		
		# Restart Explorer to apply visual tweaks (kept here since it's visual-settings-related)
		Write_Log "Restarting Explorer to apply changes..." "yellow"
		try
		{
			Stop-Process -Name explorer -Force
			Write_Log "Explorer restarted." "green"
		}
		catch
		{
			Write_Log "Explorer restart may already be complete or not required. Info: $_" "cyan"
		}
		
		Write_Log "All system configurations have been applied." "green"
		Write_Log "`r`n==================== Configure_System_Settings Completed ====================`r`n" "blue"
	}
	catch
	{
		Write_Log "An unexpected error occurred in Configure_System_Settings: $_" "red"
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
	# ==== Prompt to Restart All Programs on successful lanes ====
	if ($totalRefreshed -gt 0)
	{
		# Build $refreshedLanes (only add a lane if any file was actually refreshed)
		$refreshedLanes = @()
		foreach ($laneObj in $Lanes)
		{
			$laneNum = if ($laneObj -is [pscustomobject] -and $laneObj.LaneNumber) { $laneObj.LaneNumber }
			else { $laneObj }
			$laneNum = $laneNum.PadLeft(3, '0')
			if ($LaneNumToMachineName.ContainsKey($laneNum))
			{
				$machineName = $LaneNumToMachineName[$laneNum]
				$targetPath = "\\$machineName\Storeman\XchDev\EMVConfig\"
				foreach ($file in $fileExtensions)
				{
					$filePath = Join-Path -Path $targetPath -ChildPath $file
					if (Test-Path -Path $filePath)
					{
						$lastWrite = (Get-Item -Path $filePath).LastWriteTime
						if ($lastWrite -ge (Get-Date).AddMinutes(-5))
						{
							$refreshedLanes += $laneNum
							break
						}
					}
				}
			}
		}
		$refreshedLanes = $refreshedLanes | Select-Object -Unique
		if ($refreshedLanes.Count -gt 0)
		{
			Add-Type -AssemblyName System.Windows.Forms
			$laneListStr = ($refreshedLanes | Sort-Object) -join ", "
			$msg = "Do you want to send 'Restart All Programs' to the following lanes?`nStore $StoreNumber Lanes: $laneListStr"
			$result = [System.Windows.Forms.MessageBox]::Show($msg, "Restart All Programs?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
			if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User chose to send 'Restart All Programs' to lanes: $laneListStr" "cyan"
				Send_Restart_All_Programs -StoreNumber $StoreNumber -LaneNumbers $refreshedLanes
			}
			else
			{
				Write_Log "User cancelled 'Restart All Programs' action." "yellow"
			}
		}
	}
	Write_Log "`r`n==================== Refresh_PIN_Pad_Files Function Completed ====================" "blue"
}

# ===================================================================================================
#                                   FUNCTION: Install_FUNCTIONS_Into_SMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Always installs BOTH single-function and multi-function deploy artifacts for SMS:
#     - DEPLOY_SYS.sql            (includes "Multiple Functions" menu entry)
#     - DEPLOY_ONE_FCT.sqm        (single function deployment)
#     - DEPLOY_MULTI_FCT.sqm      (CSV / ranges like 123,234,300-305)
#
# Encoding/format:
#   - Writes files as ANSI (Windows-1252), CRLF line endings, NO BOM.
#
# Parameters:
#   - StoreNumber  [string] : Optional (kept for signature compatibility; not used in content).
#   - OfficePath   [string] : Optional; if omitted tries $script:BasePath, then $BasePath.
#
# Changes in this revision:
#   - CHANGE: Removed nested helper function; write logic is inlined per file (no nested function).
#   - CHANGE: Always includes "Multiple Functions" option in DEPLOY_SYS.sql.
#   - CHANGE: Ensures ® macro marker is injected via $($reg) everywhere it must evaluate.
#   - CHANGE: Normalizes line endings to CRLF before each write.
# ===================================================================================================

function Install_FUNCTIONS_Into_SMS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StoreNumber,
		# kept for compatibility / future use
		[Parameter(Mandatory = $false)]
		[string]$OfficePath # if not provided, we try common script-scoped fallbacks
	)
	
	Write_Log "`r`n==================== Starting Install_FUNCTIONS_Into_SMS ====================`r`n" "blue"
	
	# Registered macro marker (®) as a safe char literal; we always inject via $($reg) to avoid encoding surprises.
	$reg = [char]0x00AE
	
	# --------------------------------------------------------------------------------------------------------------
	# Resolve/validate OfficePath (prefer explicit param; fallback to common script-scoped variables)
	# --------------------------------------------------------------------------------------------------------------
	if (-not $OfficePath -or [string]::IsNullOrWhiteSpace($OfficePath))
	{
		if ($script:BasePath) { $OfficePath = $script:BasePath } # Prefer script-scoped base path
		elseif ($BasePath) { $OfficePath = $BasePath } # Fallback to legacy/global base path
	}
	if (-not $OfficePath -or -not (Test-Path -LiteralPath $OfficePath))
	{
		Write_Log "Office path not found or not provided: '$OfficePath'." "red"
		return
	}
	
	# --- Destination paths (always three files) ---
	$DeploySysPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_SYS.sql"
	$DeployMultiFctPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_MULTI_FCT.sqm"
	
	# --- Encoding: ANSI (Windows-1252), no BOM ---
	$ansi = [System.Text.Encoding]::GetEncoding(1252)
	
	# ===============================================================================================================
	#                                             DEPLOY_SYS.sql (ALWAYS includes "Multiple Functions")
	# ===============================================================================================================
	$DeploySysContent = @"
@FMT(CMP,@dbHot(FINDFIRST,UD_DEPLOY_SYS.SQL)=,$($reg)WIZRPL(UD_RUN=0));
@FMT(CMP,@WIZGET(UD_RUN)=,'$($reg)EXEC(SQL=UD_DEPLOY_SYS)$($reg)FMT(CHR,27)');

@FMT(CMP,@TOOLS(MESSAGEDLG,"!TO KEEP THE LANE'S REFERENCE SAMPLE UP TO DATE YOU SHOULD USE THE "REFERENCE SAMPLE MECHANISM". DO YOU WANT TO CONTINUE?",,NO,YES)=1,'$($reg)FMT(CHR,27)');

@EXEC(INI=HOST_OFFICE[DEPLOY_SYS]);

@WIZRPL(STYLE=SIL);
@WIZRPL(TARGET_FILTER=@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE));

@EXEC(sqi=USERB_DEPLOY_SYS);

@WIZINIT;
@WIZMENU(ONESQM=What do you want to send,
    Functions (NEW)=DEPLOY_MULTI_FCT,
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

@FMT(CMP,@dbSelect(select distinct 1 from lnk_tab where F1000='@Wizget(Target)' and f1056='999')=,,"$($reg)EXEC(msg=!*****_can_not_deploy_system_tables_to_a_host_****);$($reg)FMT(CHR,27);")

@WIZINIT;
@WIZMENU(ACTION=Action on the target database,Add or replace=ADDRPL,Add only=ADD,Replace only=UPDATE,Clean and load=LOAD);
@WIZDISPLAY;

/* SEND ONLY ONE / MULTI / OR OTHERS */

@FMT(CMP,@wizget(ONESQM)=tlz_load,$($reg)EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(ONESQM)=fcz_load,$($reg)EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(ONESQM)=fct_load,$($reg)EXEC(SQM=fct_load));
@FMT(CMP,@wizget(ONESQM)=dril_file_load,$($reg)EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(ONESQM)=dril_page_load,$($reg)EXEC(SQM=DRIL_PAGE_LOAD));
@FMT(CMP,@wizget(ONESQM)=DEPLOY_MULTI_FCT,$($reg)EXEC(SQM=DEPLOY_MULTI_FCT));

@FMT(CMP,@WIZGET(ONESQM)=ALL,,'$($reg)EXEC(SQM=exe_activate_accept_sys)$($reg)fmt(chr,27)');

@FMT(CMP,@wizget(tlz_load)=0,,$($reg)EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(fcz_load)=0,,$($reg)EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(fct_load)=0,,$($reg)EXEC(SQM=fct_load));
@FMT(CMP,@wizget(DRIL_FILE_LOAD)=0,,$($reg)EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(DRIL_PAGE_LOAD)=0,,$($reg)EXEC(SQM=DRIL_PAGE_LOAD));
@FMT(CMP,@wizget(DEPLOY_MULTI_FCT)=0,,$($reg)EXEC(SQM=DEPLOY_MULTI_FCT));

@FMT(CMP,@wizget(exe_activate_accept_all)=0,,$($reg)EXEC(SQM=exe_activate_accept_sys));
@FMT(CMP,@wizget(exe_refresh_menu)=1,$($reg)EXEC(SQM=exe_refresh_menu));

@EXEC(sqi=USERE_DEPLOY_SYS);
"@
	
	# ===============================================================================================================
	#                                      DEPLOY_FCT.sqm (template)
	#   Supports CSV and ranges (e.g., 123,234,300-305).
	# ===============================================================================================================
	$DeployMultiFctContent = @"
INSERT INTO HEADER_DCT VALUES
('HC','00000001','001901','001001',,,1997001,0000,1997001,0001,,'LOAD','CREATE DCT',,,,,,'1/1.0','V1.0',,);

CREATE TABLE FCT_DCT(@MAP_FROM_QUERY);

INSERT INTO HEADER_DCT VALUES
('HM','00000001','001901','001001',,,1997001,0000,1997001,0001,,'@WIZGET(ACTION)','@WIZGET(ACTION) SELECTED FUNCTIONS',,,,,,'1/1.0','V1.0','F1063',);

CREATE VIEW FCT_CHG AS SELECT @FIELDS_FROM_QUERY FROM FCT_DCT;

INSERT INTO FCT_CHG VALUES

/* EXTRACT SECTION */

@DBHOT(HOT_WIZ,PARAMTOLINE,PARAMSAV_DEPLOY_MULTI_FCT);
@FMT(CMP,'@WIZGET(TARGET)<>','$($reg)WIZRPL(TARGET_FILTER=@WIZGET(TARGET))');

@WIZINIT;
@WIZEDIT(FCT_LIST=Type ''ALL'' to deploy all functions,'Function IDs/ranges (122,129;122-129)');
@WIZDISPLAY;

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

@dbEXEC(SET ANSI_NULLS ON)
@dbEXEC(SET QUOTED_IDENTIFIER ON)
@dbEXEC(SET ANSI_PADDING ON)
@dbEXEC(SET ANSI_WARNINGS ON)
@dbEXEC(SET CONCAT_NULL_YIELDS_NULL ON)
@dbEXEC(SET NUMERIC_ROUNDABORT OFF)
@dbEXEC(SET ARITHABORT ON)

@MAP_DEPLOY
SELECT FCT.F1056,FCT.F1056+FCT.F1057 AS F1000,@dbFld(FCT_TAB,FCT.,F1000) FROM
    (SELECT LNI.F1056,LNI.F1057,FCT.*,ROW_NUMBER() OVER (PARTITION BY FCT.F1063,LNI.F1056,LNI.F1057 ORDER BY CASE WHEN FCT.F1000='PAL' THEN 1 ELSE 2 END DESC) AS F1301
    FROM FCT_TAB FCT
    JOIN LNK_TAB LNI ON FCT.F1000=LNI.F1000
    JOIN LNK_TAB LNO ON LNI.F1056=LNO.F1056 AND LNI.F1057=LNO.F1057
    WHERE LNO.F1000='@WIZGET(TARGET_FILTER)' AND (
          UPPER(LTRIM(RTRIM('@WIZGET(FCT_LIST)'))) = 'ALL'
          OR EXISTS
          (
            SELECT 1
            FROM
            (
              SELECT
                CASE WHEN CHARINDEX('-',token)=0 AND token IS NOT NULL AND PATINDEX('%[^0-9]%',token)=0 THEN TRY_CAST(token AS INT) ELSE NULL END AS SingleNum,
                CASE WHEN CHARINDEX('-',token)>0 AND PATINDEX('%[^0-9]%',LEFT(token,CHARINDEX('-',token)-1))=0 AND PATINDEX('%[^0-9]%',SUBSTRING(token,CHARINDEX('-',token)+1,8000))=0 THEN TRY_CAST(LEFT(token,CHARINDEX('-',token)-1) AS INT) ELSE NULL END AS StartNum,
                CASE WHEN CHARINDEX('-',token)>0 AND PATINDEX('%[^0-9]%',LEFT(token,CHARINDEX('-',token)-1))=0 AND PATINDEX('%[^0-9]%',SUBSTRING(token,CHARINDEX('-',token)+1,8000))=0 THEN TRY_CAST(SUBSTRING(token,CHARINDEX('-',token)+1,8000) AS INT) ELSE NULL END AS EndNum
              FROM
              (
                SELECT T.N.value('.','nvarchar(100)') AS token
                FROM
                (
                  SELECT CAST('<i>'+REPLACE(REPLACE(REPLACE(REPLACE('@WIZGET(FCT_LIST)',CHAR(13),''),CHAR(10),''),' ',''),',','</i><i>')+'</i>' AS XML) AS xdoc
                ) X
                CROSS APPLY xdoc.nodes('/i') AS T(N)
              ) S
              WHERE ISNULL(token,'')<>''
            ) FF
            WHERE
                   (FF.SingleNum IS NOT NULL AND FCT.F1063 IS NOT NULL AND ISNUMERIC(CONVERT(VARCHAR(50),FCT.F1063))=1 AND CAST(CONVERT(VARCHAR(50),FCT.F1063) AS INT)=FF.SingleNum)
                OR (FF.StartNum IS NOT NULL AND FF.EndNum IS NOT NULL AND FF.EndNum>=FF.StartNum AND FCT.F1063 IS NOT NULL AND ISNUMERIC(CONVERT(VARCHAR(50),FCT.F1063))=1 AND CAST(CONVERT(VARCHAR(50),FCT.F1063) AS INT) BETWEEN FF.StartNum AND FF.EndNum)
          )
        )
    ) FCT
WHERE FCT.F1301=1
ORDER BY F1000,F1063;

/* RESTORE INITITAL PARAMETER POOL */
@WIZRESET;
@DBHOT(HOT_WIZ,LINETOPARAM,PARAMSAV_DEPLOY_MULTI_FCT);
@DBHOT(HOT_WIZ,CLR,PARAMSAV_DEPLOY_MULTI_FCT);
"@
	
	# ===============================================================================================================
	#                                                  WRITE FILES (Always three)
	#   CHANGE: No nested helper; inline: normalize to CRLF → write ANSI → clear attributes → log.
	# ===============================================================================================================
	
	# -- DEPLOY_SYS.sql --
	try
	{
		$norm = [regex]::Replace($DeploySysContent, "(`r)?`n", "`r`n") # normalize to CRLF
		[System.IO.File]::WriteAllText($DeploySysPath, $norm, $ansi) # write as ANSI, no BOM
		Set-ItemProperty -LiteralPath $DeploySysPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Updated 'DEPLOY_SYS.sql' in '$OfficePath'." "green"
	}
	catch
	{
		Write_Log "Failed to write 'DEPLOY_SYS.sql'. Error: $_" "red"
	}
	
	# -- DEPLOY_MULTI_FCT.sqm --
	try
	{
		$norm = [regex]::Replace($DeployMultiFctContent, "(`r)?`n", "`r`n") # normalize to CRLF
		[System.IO.File]::WriteAllText($DeployMultiFctPath, $norm, $ansi) # write as ANSI, no BOM
		Set-ItemProperty -LiteralPath $DeployMultiFctPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write_Log "Wrote 'DEPLOY_MULTI_FCT.sqm' (All + CSV + ranges) to '$OfficePath'." "green"
	}
	catch
	{
		Write_Log "Failed to write 'DEPLOY_FCT.sqm'. Error: $_" "red"
	}
	
	Write_Log "`r`n==================== Install_FUNCTIONS_Into_SMS Completed ====================`r`n" "blue"
}

# ===================================================================================================
#  FUNCTION: Install_And_Check_LOC_SMS_Options_On_Lanes  (PowerShell 5.1)
# ---------------------------------------------------------------------------------------------------
#  Key behaviors:
#    • Reinstall: uses FirstLoad to rewrite files, but NEVER deletes folders (no Remove-Item, no robocopy /MIR or /PURGE).
#    • Install: first-time uses FirstLoad; if already present, just overwrite/add missing files (no deletes).
#    • Root Inbox (Options\<Option>\Inbox) goes to Office\XF<Store><Lane>.
#    • FirstLoad\Inbox -> XF, FirstLoad\Lbz -> Office\Lbz, Xch* -> Storeman\Xch* (FirstLoad Xch* if first-install/reinstall).
#    • Cgi -> Office\CGI; Htm/Html -> Office\HTM (english only; ignore Cgi_* / Htm_*).
#    • Generic top-level folders (e.g. Images, Layouts…): copy to lane's existing Office\<Folder> or Storeman\<Folder>;
#      if neither exists, create Office\<Folder> and copy there.
#    • Options\<Option> content is copied to Storeman\Options\<Option> (no duplication, no deletes).
#    • Action picker enables "Reinstall" ONLY if at least one selected option already exists on at least one selected lane.
#    • No ternary operator anywhere. Robust UNC copies via robocopy /E (no purge).
# ===================================================================================================

function Install_And_Check_LOC_SMS_Options_On_Lanes
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[string]$BasePath,
		[string]$OptionsRoot,
		[ValidateSet('All', 'Bank', 'Device', 'Link', 'Plugin', 'Promo', 'Xchange', 'Option', 'Others')]
		[string]$Category = 'All',
		[string[]]$OptionName,
		[int]$MaxConcurrency = 6
	)
	
	Write_Log "Install_And_Check_LOC_SMS_Options_On_Lanes: starting..." 'Cyan'
	
	# ---------------- base paths ----------------
	if ([string]::IsNullOrWhiteSpace($BasePath))
	{
		if ($script:BasePath) { $BasePath = $script:BasePath }
		else { $BasePath = 'C:\storeman' } # comment: default base
	}
	if ([string]::IsNullOrWhiteSpace($OptionsRoot))
	{
		$OptionsRoot = Join-Path $BasePath 'Install\Options' # comment: default repo
	}
	Write_Log ("Local BasePath : {0}" -f $BasePath)  'Cyan'
	Write_Log ("Options Root   : {0}" -f $OptionsRoot) 'Cyan'
	if (-not (Test-Path $OptionsRoot)) { Write_Log ("OptionsRoot not found: {0}" -f $OptionsRoot) 'Red'; return }
	
	# ---------------- zip support ----------------
	$zipLoaded = $false
	try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop; $zipLoaded = $true }
	catch { Write_Log "ZIP library unavailable - .zip options will be skipped." 'Yellow' }
	
	# ---------------- scan repo ----------------
	$knownCats = @('Bank', 'Device', 'Link', 'Plugin', 'Promo', 'Xchange', 'Option', 'Others')
	$optionFS = @()
	try
	{
		$optionFS = Get-ChildItem -Path $OptionsRoot -Force -ErrorAction SilentlyContinue |
		Where-Object { $_.PSIsContainer -or $_.Extension -match '\.zip$' }
	}
	catch { }
	if (-not $optionFS -or $optionFS.Count -eq 0) { Write_Log "No options found in repository." 'Yellow'; return }
	
	$optionEntries = @()
	foreach ($it in $optionFS)
	{
		$bn = $it.BaseName
		$cat = 'Others'
		if ($bn -match '^([A-Za-z]+)_')
		{
			$pref = $matches[1]
			if ($pref -match '^(?i)application$') { $pref = 'Option' } # normalize
			foreach ($kc in $knownCats) { if ($kc -ieq $pref) { $cat = $kc; break } }
		}
		else
		{
			$leaf = Split-Path -Path $it.DirectoryName -Leaf
			if ($leaf) { foreach ($kc in $knownCats) { if ($kc -ieq $leaf) { $cat = $kc; break } } }
		}
		if ($Category -ne 'All' -and $cat -ne $Category) { continue }
		if ($OptionName -and $OptionName.Count -gt 0)
		{
			$ok = $false; foreach ($p in $OptionName) { if ($bn -like $p) { $ok = $true; break } }
			if (-not $ok) { continue }
		}
		$e = [PSCustomObject]@{ Name = $bn; Category = $cat; SourcePath = $it.FullName; DisplayName = ("$cat\$bn") }
		$e | Add-Member ScriptMethod ToString { $this.DisplayName } -Force
		$optionEntries += $e
	}
	if (-not $optionEntries -or $optionEntries.Count -eq 0) { Write_Log "No options matched current filters." 'Yellow'; return }
	
	# ---------------- OPTIONS PICKER ----------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$formOpt = New-Object System.Windows.Forms.Form
	$formOpt.Text = "Select LOC Options"
	$formOpt.Size = New-Object System.Drawing.Size(780, 560)
	$formOpt.StartPosition = "CenterScreen"
	$formOpt.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$formOpt.MaximizeBox = $false; $formOpt.MinimizeBox = $false; $formOpt.ShowInTaskbar = $true
	
	$panelCats = New-Object System.Windows.Forms.Panel
	$panelCats.Location = New-Object System.Drawing.Point(12, 12)
	$panelCats.Size = New-Object System.Drawing.Size(744, 36)
	$formOpt.Controls.Add($panelCats)
	
	$catNames = @('All', 'Bank', 'Device', 'Link', 'Plugin', 'Promo', 'Xchange', 'Option', 'Others')
	$presentCats = @('All'); $presentSet = @{ }
	foreach ($e in $optionEntries) { if (-not $presentSet.ContainsKey($e.Category)) { $presentSet[$e.Category] = $true; $presentCats += $e.Category } }
	
	$catButtons = @{ }; $btnX = 0
	foreach ($cn in $catNames)
	{
		$b = New-Object System.Windows.Forms.Button
		$b.Text = $cn; $b.Tag = $cn
		$b.Location = New-Object System.Drawing.Point($btnX, 4)
		$b.Size = New-Object System.Drawing.Size(82, 28)
		if (($cn -ne 'All') -and (-not ($presentCats -contains $cn))) { $b.Enabled = $false }
		else { $b.Enabled = $true }
		[void]$panelCats.Controls.Add($b); $catButtons[$cn] = $b
		$btnX = $btnX + 84
	}
	
	$lblSearch = New-Object System.Windows.Forms.Label
	$lblSearch.Text = "Search:"; $lblSearch.AutoSize = $true
	$lblSearch.Location = New-Object System.Drawing.Point(12, 56)
	$formOpt.Controls.Add($lblSearch)
	
	$txtSearch = New-Object System.Windows.Forms.TextBox
	$txtSearch.Location = New-Object System.Drawing.Point(72, 52)
	$txtSearch.Size = New-Object System.Drawing.Size(684, 24)
	$formOpt.Controls.Add($txtSearch)
	
	$clbOpts = New-Object System.Windows.Forms.CheckedListBox
	$clbOpts.Location = New-Object System.Drawing.Point(12, 84)
	$clbOpts.Size = New-Object System.Drawing.Size(744, 380)
	$clbOpts.CheckOnClick = $true
	$formOpt.Controls.Add($clbOpts)
	
	$btnSelAll = New-Object System.Windows.Forms.Button
	$btnSelAll.Text = "Select All (Filtered)"
	$btnSelAll.Location = New-Object System.Drawing.Point(12, 472)
	$btnSelAll.Size = New-Object System.Drawing.Size(160, 30)
	$formOpt.Controls.Add($btnSelAll)
	
	$btnDesAll = New-Object System.Windows.Forms.Button
	$btnDesAll.Text = "Deselect All (Filtered)"
	$btnDesAll.Location = New-Object System.Drawing.Point(178, 472)
	$btnDesAll.Size = New-Object System.Drawing.Size(170, 30)
	$formOpt.Controls.Add($btnDesAll)
	
	$lblCount = New-Object System.Windows.Forms.Label
	$lblCount.Text = "Selected: 0"; $lblCount.AutoSize = $true
	$lblCount.Location = New-Object System.Drawing.Point(358, 478)
	$formOpt.Controls.Add($lblCount)
	
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"; $btnOK.Location = New-Object System.Drawing.Point(538, 472)
	$btnOK.Size = New-Object System.Drawing.Size(90, 30)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$formOpt.Controls.Add($btnOK); $formOpt.AcceptButton = $btnOK
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"; $btnCancel.Location = New-Object System.Drawing.Point(646, 472)
	$btnCancel.Size = New-Object System.Drawing.Size(90, 30)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$formOpt.Controls.Add($btnCancel); $formOpt.CancelButton = $btnCancel
	
	$currentCategory = 'All'; $checkedState = @{ }
	
	if ($OptionName -and $OptionName.Count -gt 0)
	{
		foreach ($e in $optionEntries) { foreach ($p in $OptionName) { if ($e.Name -like $p) { $checkedState[$e.Name] = $true; break } } }
	}
	
	$updateCountLabel = {
		$n = 0; foreach ($kv in $checkedState.GetEnumerator()) { if ($kv.Value) { $n = $n + 1 } }
		$lblCount.Text = ("Selected: {0}" -f $n)
	}
	$refreshList = {
		$q = ""; if ($txtSearch.Text) { $q = "$($txtSearch.Text)".Trim() }
		$filtered = @()
		foreach ($e in $optionEntries)
		{
			if ($currentCategory -ne 'All' -and $e.Category -ne $currentCategory) { continue }
			if ($q -ne "")
			{
				$hay = ($e.Name + " " + $e.Category + " " + $e.DisplayName)
				if ($hay.ToLower().IndexOf($q.ToLower()) -lt 0) { continue }
			}
			$filtered += $e
		}
		$clbOpts.BeginUpdate(); $clbOpts.Items.Clear()
		foreach ($e in ($filtered | Sort-Object Category, Name))
		{
			$idx = $clbOpts.Items.Add($e)
			if ($checkedState.ContainsKey($e.Name)) { if ($checkedState[$e.Name]) { $clbOpts.SetItemChecked($idx, $true) } }
		}
		$clbOpts.EndUpdate()
		& $updateCountLabel
	}
	
	foreach ($cn in $catNames)
	{
		$b = $catButtons[$cn]; [void]$b.Add_Click({
				param ($s,
					$e) $currentCategory = [string]$s.Tag; & $refreshList
			})
	}
	[void]$txtSearch.Add_TextChanged({ & $refreshList })
	[void]$clbOpts.Add_ItemCheck({
			$i = $_.Index
			if ($i -ge 0 -and $i -lt $clbOpts.Items.Count)
			{
				$it = $clbOpts.Items[$i]
				if ($it -and $it.PSObject.Properties['Name'])
				{
					if ($_.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { $checkedState[$it.Name] = $true }
					else { $checkedState[$it.Name] = $false }
					& $updateCountLabel
				}
			}
		})
	[void]$btnSelAll.Add_Click({
			$i = 0; while ($i -lt $clbOpts.Items.Count)
			{
				$checkedState[$clbOpts.Items[$i].Name] = $true
				$clbOpts.SetItemChecked($i, $true)
				$i = $i + 1
			}
			& $updateCountLabel
		})
	[void]$btnDesAll.Add_Click({
			$i = 0; while ($i -lt $clbOpts.Items.Count)
			{
				$item = $clbOpts.Items[$i]
				if ($item -and $item.PSObject.Properties['Name']) { $checkedState[$item.Name] = $false }
				$clbOpts.SetItemChecked($i, $false)
				$i = $i + 1
			}
			& $updateCountLabel
		})
	
	& $refreshList
	if ($formOpt.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write_Log "No options selected. Aborting." 'Yellow'; return }
	$selectedNames = @(); foreach ($kv in $checkedState.GetEnumerator()) { if ($kv.Value) { $selectedNames += $kv.Key } }
	if (-not $selectedNames -or $selectedNames.Count -eq 0) { Write_Log "No options checked. Aborting." 'Yellow'; return }
	Write_Log ("Options selected: {0}" -f (($selectedNames | Sort-Object) -join ", ")) 'DarkCyan'
	
	# ---------------- LANE PICKER ----------------
	$sel = $null
	try { $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Lanes for LOC Options" }
	catch { Write_Log ("Lane picker failed: {0}" -f $_.Exception.Message) 'Red'; return }
	if (-not $sel -or -not $sel.Lanes -or $sel.Lanes.Count -eq 0) { Write_Log "No lanes selected. Aborting." 'Yellow'; return }
	
	$laneNumToMachine = @{ }
	if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('LaneNumToMachineName')) { $laneNumToMachine = $script:FunctionResults['LaneNumToMachineName'] }
	
	$lanePlan = @()
	foreach ($ln in ($sel.Lanes | Sort-Object))
	{
		$m = $null; if ($laneNumToMachine.ContainsKey($ln)) { $m = $laneNumToMachine[$ln] }
		if ([string]::IsNullOrWhiteSpace($m)) { Write_Log ("Lane {0} has no machine mapping - skipping." -f $ln) 'Yellow' }
		else { $lanePlan += (New-Object PSObject -Property @{ Lane = $ln; Machine = $m }) }
	}
	if ($lanePlan.Count -eq 0) { Write_Log "No lanes resolved to machine names - aborting." 'Red'; return }
	$pairs = @(); foreach ($lp in $lanePlan) { $pairs += ("{0}={1}" -f $lp.Lane, $lp.Machine) }
	Write_Log ("Lanes selected : {0}" -f ($pairs -join ", ")) 'DarkCyan'
	
	# ---------------- resolve remote paths ----------------
	$laneTargets = @()
	foreach ($lp in $lanePlan)
	{
		$mach = $lp.Machine
		$candidates = @("\\$mach\storeman", "\\$mach\c$\Storeman", "\\$mach\d$\Storeman")
		$storemanRoot = $null
		foreach ($p in $candidates)
		{
			try { if (Test-Path $p) { $storemanRoot = $p; break } }
			catch { }
		}
		if (-not $storemanRoot) { Write_Log ("Lane {0} ({1}): Storeman root not reachable." -f $lp.Lane, $mach) 'Yellow'; continue }
		
		$officeRoot = Join-Path $storemanRoot 'Office'
		if (-not (Test-Path $officeRoot)) { Write_Log ("Lane {0} ({1}): missing Office. Skipping." -f $lp.Lane, $mach) 'Yellow'; continue }
		
		$optionsRootLane = Join-Path $storemanRoot 'Options'
		if (-not (Test-Path $optionsRootLane)) { Write_Log ("Lane {0} ({1}): missing Options. Skipping." -f $lp.Lane, $mach) 'Yellow'; continue }
		
		$laneTargets += (New-Object PSObject -Property @{
				Lane			   = $lp.Lane
				Machine		       = $mach
				RemoteStoremanRoot = $storemanRoot
				RemoteOfficePath   = $officeRoot
				RemoteOptionsRoot  = $optionsRootLane
			})
	}
	if ($laneTargets.Count -eq 0) { Write_Log "No reachable lanes with Office/Options present." 'Red'; return }
	
	Write_Log "Remote Office/Storeman paths per lane:" 'Blue'
	Write_Log ((
			$laneTargets | Select-Object Lane, Machine, RemoteStoremanRoot, RemoteOfficePath, RemoteOptionsRoot |
			Sort-Object Lane | Format-Table -AutoSize | Out-String
		)) 'Gray'
	
	# ---------------- stage: extract + parse ----------------
	$selectedFS = @(); foreach ($it in $optionFS) { if ($selectedNames -contains $it.BaseName) { $selectedFS += $it } }
	
	$stagingRecords = @(); $tempDirs = New-Object System.Collections.Generic.List[string]
	foreach ($item in $selectedFS)
	{
		$bn = $item.BaseName
		
		# extract zip if needed
		$extractedRoot = $item.FullName
		try
		{
			if (-not $item.PSIsContainer)
			{
				if ($zipLoaded)
				{
					$dest = Join-Path $env:TEMP ("LOC_Option_" + $bn + "_" + (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
					New-Item -ItemType Directory -Path $dest -Force | Out-Null
					[System.IO.Compression.ZipFile]::ExtractToDirectory($item.FullName, $dest)
					$extractedRoot = $dest
					[void]$tempDirs.Add($dest)
				}
				else { Write_Log ("Skipping ZIP option '{0}' (no ZIP lib)." -f $item.Name) 'Yellow'; continue }
			}
		}
		catch { Write_Log ("Extraction failed for '{0}': {1}" -f $bn, $_.Exception.Message) 'Yellow'; continue }
		
		# inner vendor root: Options\<Option>\...
		$scanBase = $null
		try { $lvl1 = Join-Path $extractedRoot 'Options'; $lvl2 = Join-Path $lvl1 $bn; if (Test-Path $lvl2 -PathType Container) { $scanBase = $lvl2 } }
		catch { }
		if (-not $scanBase)
		{
			try
			{
				$cand = Get-ChildItem -Path $extractedRoot -Directory -Recurse -Force -ErrorAction SilentlyContinue |
				Where-Object { (Split-Path $_.Parent.FullName -Leaf) -ieq 'Options' -and $_.Name -ieq $bn } |
				Select-Object -First 1
				if ($cand) { $scanBase = $cand.FullName }
			}
			catch { }
		}
		if (-not $scanBase) { $scanBase = $extractedRoot }
		
		# collect top-level (english only Cgi/Htm)
		$TopCgiDir = $null; $TopHtmDir = $null; $RootInboxDir = $null; $TopOfficeDir = $null
		$XchDirs = @(); $FirstLoadXchDirs = @(); $FirstLoadInboxFiles = @(); $LbzFiles = @(); $OtherTopDirs = @()
		try
		{
			$top = Get-ChildItem -Path $scanBase -Force -ErrorAction SilentlyContinue
			foreach ($e in $top)
			{
				if (-not $e.PSIsContainer) { continue }
				$nmLower = $e.Name.ToLower()
				
				if ($nmLower -eq 'cgi') { $TopCgiDir = $e.FullName; continue }
				if ($nmLower -eq 'htm' -or $nmLower -eq 'html') { $TopHtmDir = $e.FullName; continue }
				if ($nmLower -eq 'inbox') { $RootInboxDir = $e.FullName; continue }
				if ($nmLower -eq 'office') { $TopOfficeDir = $e.FullName; continue }
				
				if ($nmLower -eq 'firstload')
				{
					$fl = Get-ChildItem -Path $e.FullName -Force -ErrorAction SilentlyContinue
					foreach ($fd in $fl)
					{
						$sn = $fd.Name.ToLower()
						if ($sn -eq 'inbox')
						{
							$FirstLoadInboxFiles += Get-ChildItem -Path $fd.FullName -Recurse -File -Force -ErrorAction SilentlyContinue |
							Where-Object { @('.sqi', '.sqm', '.sql', '.txt') -contains $_.Extension.ToLower() } |
							Select-Object -ExpandProperty FullName
						}
						elseif ($sn -eq 'lbz')
						{
							$LbzFiles += Get-ChildItem -Path $fd.FullName -Recurse -File -Force -ErrorAction SilentlyContinue |
							Where-Object { @('.lbz', '.lbt') -contains $_.Extension.ToLower() } |
							Select-Object -ExpandProperty FullName
						}
						elseif ($sn.Length -ge 3 -and ($sn.Substring(0, 3)) -eq 'xch')
						{
							$FirstLoadXchDirs += $fd.FullName
						}
					}
					continue
				}
				
				if ($nmLower.Length -ge 3 -and ($nmLower.Substring(0, 3)) -eq 'xch') { $XchDirs += $e.FullName; continue }
				
				# generic top-level folder (Images, Layouts, etc.)
				$OtherTopDirs += $e.FullName
			}
		}
		catch { }
		
		# XF sets (Office + Root Inbox + loose root)
		$XF_Office = @(); $XF_RootInbox = @(); $XF_LooseRoot = @()
		try
		{
			if ($TopOfficeDir)
			{
				$XF_Office += Get-ChildItem -Path $TopOfficeDir -Recurse -File -Force -ErrorAction SilentlyContinue |
				Where-Object { @('.sqi', '.sqm', '.sql', '.txt') -contains $_.Extension.ToLower() } |
				Select-Object -ExpandProperty FullName
			}
		}
		catch { }
		try
		{
			if ($RootInboxDir)
			{
				$XF_RootInbox += Get-ChildItem -Path $RootInboxDir -Recurse -File -Force -ErrorAction SilentlyContinue |
				Where-Object { @('.sqi', '.sqm', '.sql', '.txt') -contains $_.Extension.ToLower() } |
				Select-Object -ExpandProperty FullName
			}
		}
		catch { }
		try
		{
			$rootFiles = Get-ChildItem -Path $scanBase -File -Force -ErrorAction SilentlyContinue
			foreach ($rf in $rootFiles)
			{
				$lx = ($rf.Extension).ToLower()
				$ok = $false; foreach ($x in @('.sqi', '.sqm', '.sql', '.txt')) { if ($lx -eq $x) { $ok = $true; break } }
				if ($ok) { $XF_LooseRoot += $rf.FullName }
			}
		}
		catch { }
		
		$stagingRecords += (New-Object PSObject -Property @{
				Name			    = $bn
				ScanBase		    = $scanBase
				TopCgiDir		    = $TopCgiDir
				TopHtmDir		    = $TopHtmDir
				RootInboxDir	    = $RootInboxDir
				TopOfficeDir	    = $TopOfficeDir
				XchDirs			    = $XchDirs
				FirstLoadXchDirs    = $FirstLoadXchDirs
				FirstLoadInboxFiles = $FirstLoadInboxFiles
				LbzFiles		    = $LbzFiles
				XF_Office		    = $XF_Office
				XF_RootInbox	    = $XF_RootInbox
				XF_LooseRoot	    = $XF_LooseRoot
				OtherTopDirs	    = $OtherTopDirs
			})
	}
	if (-not $stagingRecords -or $stagingRecords.Count -eq 0) { Write_Log "Selected options contained no usable content." 'Yellow'; return }
	
	# ---------------- check if Reinstall allowed ----------------
	$reinstallAllowed = $false
	foreach ($lt in $laneTargets)
	{
		foreach ($st in $stagingRecords)
		{
			$candidate = Join-Path $lt.RemoteOptionsRoot $st.Name
			if (Test-Path $candidate) { $reinstallAllowed = $true; break }
		}
		if ($reinstallAllowed) { break }
	}
	if ($reinstallAllowed) { $reinstallStatus = 'Yes' }
	else { $reinstallStatus = 'No' }
	Write_Log ("Reinstall available for selection: {0}" -f $reinstallStatus) 'DarkGray'
	
	# ---------------- ACTION PICKER ----------------
	$formMode = New-Object System.Windows.Forms.Form
	$formMode.Text = "Action"
	$formMode.Size = New-Object System.Drawing.Size(440, 220)
	$formMode.StartPosition = "CenterScreen"
	$formMode.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$formMode.MaximizeBox = $false; $formMode.MinimizeBox = $false; $formMode.ShowInTaskbar = $true
	
	$grp = New-Object System.Windows.Forms.GroupBox
	$grp.Text = "What do you want to do?"
	$grp.Location = New-Object System.Drawing.Point(12, 12)
	$grp.Size = New-Object System.Drawing.Size(410, 120)
	$formMode.Controls.Add($grp)
	
	$rbAudit = New-Object System.Windows.Forms.RadioButton
	$rbAudit.Text = "Audit only"
	$rbAudit.AutoSize = $true
	$rbAudit.Location = New-Object System.Drawing.Point(16, 25)
	$grp.Controls.Add($rbAudit)
	
	$rbInstall = New-Object System.Windows.Forms.RadioButton
	$rbInstall.Text = "Install / Repair (outside only if already installed)"
	$rbInstall.AutoSize = $true
	$rbInstall.Location = New-Object System.Drawing.Point(16, 50)
	$rbInstall.Checked = $true
	$grp.Controls.Add($rbInstall)
	
	$rbReinstall = New-Object System.Windows.Forms.RadioButton
	$rbReinstall.Text = "First Load / Reinstall (FirstLoad + outside; FirstLoad wins)"
	$rbReinstall.AutoSize = $true
	$rbReinstall.Location = New-Object System.Drawing.Point(16, 75)
	$rbReinstall.Enabled = $reinstallAllowed
	$grp.Controls.Add($rbReinstall)
	
	$btnModeOK = New-Object System.Windows.Forms.Button
	$btnModeOK.Text = "OK"; $btnModeOK.Location = New-Object System.Drawing.Point(240, 150)
	$btnModeOK.Size = New-Object System.Drawing.Size(80, 28)
	$btnModeOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$formMode.Controls.Add($btnModeOK); $formMode.AcceptButton = $btnModeOK
	
	$btnModeCancel = New-Object System.Windows.Forms.Button
	$btnModeCancel.Text = "Cancel"; $btnModeCancel.Location = New-Object System.Drawing.Point(332, 150)
	$btnModeCancel.Size = New-Object System.Drawing.Size(80, 28)
	$btnModeCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$formMode.Controls.Add($btnModeCancel); $formMode.CancelButton = $btnModeCancel
	
	$modeDlg = $formMode.ShowDialog()
	if ($modeDlg -ne [System.Windows.Forms.DialogResult]::OK) { Write_Log "User cancelled action selection." 'Yellow'; return }
	$ActionMode = 1
	if ($rbAudit.Checked) { $ActionMode = 0 }
	if ($rbInstall.Checked) { $ActionMode = 1 }
	if ($rbReinstall.Checked) { $ActionMode = 2 }
	
	# ---------------- per-lane processing ----------------
	$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
	$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, [Math]::Max(1, $MaxConcurrency), $iss, $Host)
	$pool.Open()
	$asyncHandles = New-Object System.Collections.Generic.List[System.IAsyncResult]
	$psList = New-Object System.Collections.Generic.List[System.Management.Automation.PowerShell]
	
	foreach ($lt in $laneTargets)
	{
		$ps = [PowerShell]::Create(); $ps.RunspacePool = $pool
		[void]$ps.AddScript({
				param ($laneRec,
					$stagedOptions,
					[int]$actionMode,
					[string]$storeNumberArg)
				
				$lane = $laneRec.Lane; $machine = $laneRec.Machine
				$storemanRoot = $laneRec.RemoteStoremanRoot
				$officeRoot = $laneRec.RemoteOfficePath
				$optionsRoot = $laneRec.RemoteOptionsRoot
				
				$laneMsgs = New-Object System.Collections.Generic.List[string]
				$laneRows = New-Object System.Collections.Generic.List[psobject]
				
				# XF name (must already exist; we don't create)
				$laneInt = $null; $lanePadded = $null
				try { $laneInt = [int]$lane }
				catch { }
				if ($laneInt -ne $null) { $lanePadded = ('{0:D3}' -f $laneInt) }
				else { $lanePadded = "$lane" }
				$xfFolderName = ("XF{0}{1}" -f $storeNumberArg, $lanePadded)
				$xfFolderPath = Join-Path $officeRoot $xfFolderName
				
				foreach ($opt in $stagedOptions)
				{
					$name = $opt.Name
					$scanBase = $opt.ScanBase
					$cgiDir = $opt.TopCgiDir
					$htmDir = $opt.TopHtmDir
					$rootInboxDir = $opt.RootInboxDir
					$officeDir = $opt.TopOfficeDir
					$xchDirsOutside = $opt.XchDirs
					$xchDirsFirst = $opt.FirstLoadXchDirs
					$lbzFilesFirst = $opt.LbzFiles
					$xfOffice = $opt.XF_Office
					$xfRootInbox = $opt.XF_RootInbox
					$xfLooseRoot = $opt.XF_LooseRoot
					$otherTopDirs = $opt.OtherTopDirs
					
					$optFolderPath = Join-Path $optionsRoot $name
					$installedBefore = Test-Path $optFolderPath
					
					$doAudit = $false; $doInstall = $false; $doReinstall = $false
					if ($actionMode -eq 0) { $doAudit = $true }
					if ($actionMode -eq 1) { $doInstall = $true }
					if ($actionMode -eq 2) { $doReinstall = $true }
					
					# First load if reinstall OR not installed yet
					$useFirstLoad = $false
					if ($doReinstall) { $useFirstLoad = $true }
					if ($doInstall -and -not $installedBefore) { $useFirstLoad = $true }
					
					$failed = 0
					$copiedOpt = 0; $copiedCgi = 0; $copiedHtm = 0; $copiedXch = 0; $copiedLBZ = 0; $copiedXF = 0; $copiedOther = 0
					
					if (-not $doAudit)
					{
						# 1) Options\<Option> - safe to create option subfolder
						try
						{
							if (-not (Test-Path $optFolderPath)) { New-Item -ItemType Directory -Path $optFolderPath -Force | Out-Null }
							$rc = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" `
												-ArgumentList @("`"$scanBase`"", "`"$optFolderPath`"", "/E", "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:1", "/W:1") `
												-NoNewWindow -PassThru -Wait
							$copiedOpt = 1
						}
						catch { $failed = $failed + 1 }
						
						# 2) Cgi -> Office\Cgi  (english only; do not create dest)
						if ($cgiDir)
						{
							$destCgi = Join-Path $officeRoot 'Cgi'
							if (Test-Path $destCgi)
							{
								try
								{
									$rc = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" `
														-ArgumentList @("`"$cgiDir`"", "`"$destCgi`"", "/E", "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:1", "/W:1") `
														-NoNewWindow -PassThru -Wait
									$copiedCgi = 1
								}
								catch { $failed = $failed + 1 }
							}
							else
							{
								$laneMsgs.Add(("[WARN] {0} {1} | {2}: Office\Cgi not found. Skipped Cgi copy." -f $lane, $machine, $name))
							}
						}
						
						# 3) Htm/Html -> Office\Htm  (english only; do not create dest)
						if ($htmDir)
						{
							$destHtm = Join-Path $officeRoot 'Htm'
							if (Test-Path $destHtm)
							{
								try
								{
									$rc = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" `
														-ArgumentList @("`"$htmDir`"", "`"$destHtm`"", "/E", "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:1", "/W:1") `
														-NoNewWindow -PassThru -Wait
									$copiedHtm = 1
								}
								catch { $failed = $failed + 1 }
							}
							else
							{
								$laneMsgs.Add(("[WARN] {0} {1} | {2}: Office\Htm not found. Skipped Htm copy." -f $lane, $machine, $name))
							}
						}
						
						# 4) Xch*  -- merge logic
						#    First load: copy FirstLoad Xch* first, then copy outside Xch* with /XC so FL wins.
						#    Repair: copy only outside Xch*.
						$srcXchSetOrder = @()
						if ($useFirstLoad)
						{
							foreach ($p in $xchDirsFirst) { $srcXchSetOrder += ('FL|' + $p) }
							foreach ($p in $xchDirsOutside) { $srcXchSetOrder += ('OUT|' + $p) }
						}
						else
						{
							foreach ($p in $xchDirsOutside) { $srcXchSetOrder += ('OUT|' + $p) }
						}
						foreach ($tagged in $srcXchSetOrder)
						{
							$parts = $tagged.Split('|', 2)
							$kind = $parts[0]; $srcDir = $parts[1]
							$xname = Split-Path -Path $srcDir -Leaf
							$destX = Join-Path $storemanRoot $xname
							try
							{
								if (-not (Test-Path $destX)) { New-Item -ItemType Directory -Path $destX -Force | Out-Null } # allowed to create Xch*
								$args = @("`"$srcDir`"", "`"$destX`"", "/E", "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:1", "/W:1")
								if ($kind -eq 'OUT' -and $useFirstLoad)
								{
									# After FL, do not overwrite changed files from outside
									$args += "/XC"
								}
								$rc = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList $args -NoNewWindow -PassThru -Wait
								$copiedXch = $copiedXch + 1
							}
							catch { $failed = $failed + 1 }
						}
						
						# 5) LBZ (FirstLoad only, first load) -> Office\Lbz (do not create dest)
						if ($useFirstLoad -and $lbzFilesFirst -and $lbzFilesFirst.Count -gt 0)
						{
							$destLBZ = Join-Path $officeRoot 'Lbz'
							if (Test-Path $destLBZ)
							{
								foreach ($f in $lbzFilesFirst)
								{
									try { Copy-Item -Path $f -Destination (Join-Path $destLBZ (Split-Path $f -Leaf)) -Force; $copiedLBZ = $copiedLBZ + 1 }
									catch { $failed = $failed + 1 }
								}
							}
							else
							{
								$laneMsgs.Add(("[WARN] {0} {1} | {2}: Office\Lbz not found. Skipped LBZ." -f $lane, $machine, $name))
							}
						}
						
						# 6) XF drops - file-by-file control to honor "outside if not duplicate; FL wins"
						if (Test-Path $xfFolderPath)
						{
							try
							{
								# Build FirstLoad name set (leaf names only; XF has no subfolders)
								$flLeaf = New-Object System.Collections.Generic.HashSet[string]
								if ($useFirstLoad -and $opt.FirstLoadInboxFiles -and $opt.FirstLoadInboxFiles.Count -gt 0)
								{
									foreach ($f in $opt.FirstLoadInboxFiles)
									{
										try { [void]$flLeaf.Add((Split-Path $f -Leaf).ToLower()) }
										catch { }
									}
								}
								# Outside set (Office + Root Inbox + Loose Root): skip if leaf collides with FL set (when in FL mode)
								$outsideFiles = @()
								if ($opt.XF_Office) { $outsideFiles += $opt.XF_Office }
								if ($opt.XF_RootInbox) { $outsideFiles += $opt.XF_RootInbox }
								if ($opt.XF_LooseRoot) { $outsideFiles += $opt.XF_LooseRoot }
								
								foreach ($src in $outsideFiles)
								{
									$leaf = (Split-Path $src -Leaf)
									$skip = $false
									if ($useFirstLoad)
									{
										if ($flLeaf.Contains($leaf.ToLower())) { $skip = $true }
									}
									if (-not $skip)
									{
										try { Copy-Item -Path $src -Destination (Join-Path $xfFolderPath $leaf) -Force; $copiedXF = $copiedXF + 1 }
										catch { $failed = $failed + 1 }
									}
								}
								# FirstLoad inbox files last (they win)
								if ($useFirstLoad -and $opt.FirstLoadInboxFiles)
								{
									foreach ($src in $opt.FirstLoadInboxFiles)
									{
										$leaf = (Split-Path $src -Leaf)
										try { Copy-Item -Path $src -Destination (Join-Path $xfFolderPath $leaf) -Force; $copiedXF = $copiedXF + 1 }
										catch { $failed = $failed + 1 }
									}
								}
							}
							catch { $failed = $failed + 1 }
						}
						else
						{
							$laneMsgs.Add(("[WARN] {0} {1} | {2}: XF folder {3} not found. Skipped XF drops." -f $lane, $machine, $name, $xfFolderName))
						}
						
						# 7) Other generic top-level folders - copy only if dest already exists; do NOT create
						foreach ($srcOther in $otherTopDirs)
						{
							$oname = Split-Path -Path $srcOther -Leaf
							$destOffice = Join-Path $officeRoot $oname
							$destStore = Join-Path  $storemanRoot $oname
							$target = $null
							if (Test-Path $destOffice) { $target = $destOffice }
							elseif (Test-Path $destStore) { $target = $destStore }
							
							if ($target -ne $null)
							{
								try
								{
									$rc = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" `
														-ArgumentList @("`"$srcOther`"", "`"$target`"", "/E", "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:1", "/W:1") `
														-NoNewWindow -PassThru -Wait
									$copiedOther = $copiedOther + 1
								}
								catch { $failed = $failed + 1 }
							}
							else
							{
								$laneMsgs.Add(("[WARN] {0} {1} | {2}: Neither Office\{3} nor Storeman\{3} exists. Skipped." -f $lane, $machine, $name, $oname))
							}
						}
					}
					
					# after
					$optFolderAfter = Test-Path $optFolderPath
					$installedAfter = $optFolderAfter
					
					# mode text
					$modeTxt = 'Audit'
					if ($actionMode -eq 1) { $modeTxt = 'Install' }
					if ($actionMode -eq 2) { $modeTxt = 'Reinstall' }
					
					# concise summary
					$summary = ("{0} {1} | {2}: Opt={3} Cgi={4} Htm={5} Xch={6} LBZ={7} XF={8} Other={9} {10}" -f `
						$lane, $machine, $name, $copiedOpt, $copiedCgi, $copiedHtm, $copiedXch, $copiedLBZ, $copiedXF, $copiedOther, $modeTxt)
					$laneMsgs.Add($summary)
					
					$laneRows.Add((New-Object PSObject -Property @{
								Lane		    = $lane
								Machine		    = $machine
								Option		    = $name
								InstalledBefore = $installedBefore
								InstalledAfter  = $installedAfter
								Copy_Options    = $copiedOpt
								Copy_Cgi	    = $copiedCgi
								Copy_Htm	    = $copiedHtm
								Copy_Xch	    = $copiedXch
								Copy_LBZ	    = $copiedLBZ
								Copy_XF		    = $copiedXF
								Copy_OtherDirs  = $copiedOther
								Failures	    = $failed
								Mode		    = $modeTxt
							}))
				}
				
				New-Object PSObject -Property @{ Messages = $laneMsgs; Items = $laneRows }
			}).AddArgument($lt).AddArgument($stagingRecords).AddArgument([int]$ActionMode).AddArgument([string]$StoreNumber)
		
		$async = $ps.BeginInvoke()
		[void]$asyncHandles.Add($async)
		[void]$psList.Add($ps)
	}
	
	# ---------------- collect ----------------
	$allMessages = New-Object System.Collections.Generic.List[string]
	$allItems = New-Object System.Collections.Generic.List[psobject]
	$i = 0
	while ($i -lt $asyncHandles.Count)
	{
		$ar = $asyncHandles[$i]
		try
		{
			[void]$ar.AsyncWaitHandle.WaitOne()
			$output = $psList[$i].EndInvoke($ar)
			foreach ($block in $output)
			{
				if ($block -and $block.Messages) { foreach ($m in $block.Messages) { [void]$allMessages.Add($m) } }
				if ($block -and $block.Items) { foreach ($it in $block.Items) { [void]$allItems.Add($it) } }
			}
		}
		catch { Write_Log ("Runspace error: {0}" -f $_.Exception.Message) 'Yellow' }
		finally
		{
			try { $psList[$i].Dispose() }
			catch { }
		}
		$i = $i + 1
	}
	try { $pool.Close(); $pool.Dispose() }
	catch { }
	
	foreach ($m in ($allMessages | Sort-Object)) { Write_Log $m 'Gray' }
	
	Write_Log "=================== Summary ===================" 'Blue'
	Write_Log ((
			$allItems | Sort-Object Lane, Option |
			Select-Object Lane, Machine, Option, InstalledBefore, InstalledAfter, Copy_Options, Copy_Cgi, Copy_Htm, Copy_Xch, Copy_LBZ, Copy_XF, Copy_OtherDirs, Failures, Mode |
			Format-Table -AutoSize | Out-String
		)) 'Gray'
	
	foreach ($d in $tempDirs)
	{
		try { if (Test-Path $d) { Remove-Item -Path $d -Recurse -Force -ErrorAction SilentlyContinue } }
		catch { }
	}
	
	Write_Log "Install_And_Check_LOC_SMS_Options_On_Lanes: done." 'Cyan'
	return $allItems
}

# ===================================================================================================
#                                           FUNCTION: INI_Editor
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generic INI editor for files under \storeman\<RelativeIniPath>. Pick Server/Lane as source,
#   load the INI, select/TYPE any section, edit ALL keys (add/edit/delete/reorder), save, then
#   optionally deploy to other lanes as either:
#      • Merge only your changes (update just the selected section's keys), or
#      • Replace the whole INI file.
#
# UX changes in this build:
#   - INI row (label + path + Load INI) is forced onto one line using a FlowLayoutPanel that resizes
#     the combo so the button never wraps.
#   - More compact top area (tighter margins/padding), so the grid is larger without enlarging the form.
#   - Copy-to-lanes is disabled until you Save; any edit/add/delete/reorder marks the form dirty and
#     re-disables Copy.
#   - Reordering is smoother; Saved file preserves the on-screen key order; no stray blank line added.
# ---------------------------------------------------------------------------------------------------

function INI_Editor
{
	param (
		[string]$RelativeIniPath = 'office\Setup.ini',
		[string[]]$PredefinedIniPaths = @('office\Setup.ini', 'office\Setting.ini', 'office\System.ini', 'XchDev\ApiVerifoneMX.ini')
	)
	
	Write_Log "`r`n==================== Starting INI_Editor ====================`r`n" "blue"
	
	# ==============================================================================================
	#                                    HELPERS: paths, parse, write
	# ==============================================================================================
	
	$getKnownLanes = {
		$known = @()
		if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'])
		{ $known = $script:FunctionResults['LaneMachines'].Values | Where-Object { $_ } | Select-Object -Unique }
		elseif ($script:FunctionResults.ContainsKey('LaneNumToMachineName') -and $script:FunctionResults['LaneNumToMachineName'])
		{ $known = $script:FunctionResults['LaneNumToMachineName'].Values | Where-Object { $_ } | Select-Object -Unique }
		return, ($known | Sort-Object)
	}
	
	$getLaneRoot = {
		param ([string]$ComputerName)
		$s1 = "\\$ComputerName\Storeman"; if (Test-Path -LiteralPath $s1) { return $s1 }
		$s2 = "\\$ComputerName\c$\storeman"; if (Test-Path -LiteralPath $s2) { return $s2 }
		return $null
	}
	
	$getServerRoot = {
		if (Test-Path -LiteralPath 'C:\storeman') { return 'C:\storeman' }
		if (Test-Path -LiteralPath 'D:\storeman') { return 'D:\storeman' }
		return "$($env:SystemDrive)\storeman"
	}
	
	$buildFullPath = {
		param ([string]$Root,
			[string]$Rel)
		$r = ($Rel -as [string]).Trim().TrimStart('\', '/')
		return (Join-Path $Root $r)
	}
	
	# ---------- INI reader (keeps order of sections and keys in memory) ----------
	$readIni = {
		param ([string]$Path)
		if (-not (Test-Path -LiteralPath $Path)) { return $null }
		
		$raw = Get-Content -LiteralPath $Path -Encoding ASCII
		$sections = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]'
		$current = $null
		
		foreach ($line in $raw)
		{
			$t = $line.Trim()
			if ($t -match '^\s*[;#]') { continue }
			if ($t -match '^\[(?<sec>[^\]]+)\]\s*$')
			{
				$sec = $matches['sec']
				if (-not $sections.ContainsKey($sec))
				{
					$od = New-Object System.Collections.Specialized.OrderedDictionary
					$sections.Add($sec, $od)
				}
				$current = $sec
				continue
			}
			if ($current -and $t -match '^(?<k>[^=]+?)\s*=\s*(?<v>.*)$')
			{
				$k = $matches['k'].Trim()
				$v = $matches['v']
				if (-not $sections[$current].Contains($k)) { $sections[$current].Add($k, $v) }
				else { $sections[$current][$k] = $v }
			}
		}
		return @{ RawLines = $raw; Sections = $sections }
	}
	
	# ---------- INI section writer (no extra blank line; preserves key order from OrderedDictionary) ----------
	$writeIniSection = {
		param (
			[string[]]$RawLines,
			[string]$SectionName,
			[System.Collections.Specialized.OrderedDictionary]$NewKeyValues
		)
		$lines = New-Object System.Collections.Generic.List[string]; $lines.AddRange($RawLines)
		$start = -1; $end = $lines.Count
		
		for ($i = 0; $i -lt $lines.Count; $i++)
		{
			if ($lines[$i] -match "^\s*\[$([regex]::Escape($SectionName))\]\s*$") { $start = $i; break }
		}
		
		if ($start -lt 0)
		{
			if ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Trim() -ne '') { $lines.Add('') }
			$lines.Add("[$SectionName]")
			foreach ($k in $NewKeyValues.Keys) { $lines.Add("$k=$($NewKeyValues[$k])") }
			return, $lines.ToArray()
		}
		
		for ($i = $start + 1; $i -lt $lines.Count; $i++)
		{
			if ($lines[$i] -match '^\s*\[.+\]\s*$') { $end = $i; break }
		}
		
		# find last nonblank/noncomment inside section
		$lastContent = $start
		for ($i = [Math]::Min($end, $lines.Count) - 1; $i -gt $start; $i--)
		{
			$t = $lines[$i].Trim()
			if ($t -ne '' -and $t -notmatch '^[;#]') { $lastContent = $i; break }
		}
		$insertAt = $lastContent + 1
		
		# update existing keys
		$updated = @{ }
		for ($i = $start + 1; $i -lt $end; $i++)
		{
			$L = $lines[$i]
			if ($L.Trim() -match '^[;#]') { continue }
			if ($L -match '^(?<k>[^=]+?)\s*=\s*(?<v>.*)$')
			{
				$currKey = $matches['k'].Trim()
				$targetKey = $null
				foreach ($nk in $NewKeyValues.Keys) { if ($nk -ieq $currKey) { $targetKey = $nk; break } }
				if ($targetKey) { $lines[$i] = "$currKey=$($NewKeyValues[$targetKey])"; $updated[$targetKey] = $true }
			}
		}
		
		# add new keys (preserve order from $NewKeyValues)
		$pending = @()
		foreach ($k in $NewKeyValues.Keys) { if (-not $updated.ContainsKey($k)) { $pending += "$k=$($NewKeyValues[$k])" } }
		if ($pending.Count -gt 0)
		{
			$offset = 0
			foreach ($nl in $pending) { $lines.Insert($insertAt + $offset, $nl); $offset++ }
		}
		return, $lines.ToArray()
	}
	
	$lanesToMachines = {
		param ($LaneNums)
		$out = @()
		$map = $script:FunctionResults['LaneNumToMachineName']
		$known = & $getKnownLanes
		foreach ($ln in $LaneNums)
		{
			$pad = $null; try { $pad = "{0:D3}" -f ([int]$ln) }
			catch { $pad = "$ln" }
			$m = $null
			if ($map)
			{
				if ($map.ContainsKey($pad)) { $m = $map[$pad] }
				elseif ($map.ContainsKey([int]$ln)) { $m = $map[[int]$ln] }
				elseif ($map.ContainsKey("$([int]$ln)")) { $m = $map["$([int]$ln)"] }
			}
			if (-not $m)
			{
				$guess = $known | Where-Object { "$_" -like "*$pad" }
				if ($guess) { $m = $guess[0] }
			}
			if ($m) { $out += $m }
		}
		return, (@($out | Where-Object { $_ } | Select-Object -Unique))
	}
	
	# ==============================================================================================
	#                                           UI LAYOUT
	# ==============================================================================================
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	[void][System.Windows.Forms.Application]::EnableVisualStyles()
	
	$frm = New-Object System.Windows.Forms.Form
	$frm.Text = "INI_Editor  -  dynamic section/key editor"
	$frm.Size = New-Object System.Drawing.Size(820, 600) # keep size modest; we'll give the grid more room
	$frm.StartPosition = 'CenterScreen'
	$frm.FormBorderStyle = 'FixedDialog'
	$frm.MaximizeBox = $false
	$frm.MinimizeBox = $false
	$frm.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)
	
	# Root: 1 column; make grid take the most space (Percent=100 row at the end)
	$root = New-Object System.Windows.Forms.TableLayoutPanel
	$root.Dock = 'Fill'
	$root.Padding = New-Object System.Windows.Forms.Padding(8)
	$root.ColumnCount = 1
	$root.RowCount = 6
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # header
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # source group
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # INI row (flow)
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # Section row
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # GRID gets the rest
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # bottom buttons
	$frm.Controls.Add($root)
	
	# ---------- Header (small + tight margins so the grid gains space) ----------
	$hdr = New-Object System.Windows.Forms.Label
	$hdr.Text = "Pick Server/Lane, load an INI, choose/TYPE a section, edit keys, Save, then optionally deploy."
	$hdr.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
	$hdr.AutoSize = $true
	$hdr.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 4)
	$root.Controls.Add($hdr, 0, 0)
	
	# ---------- Source group (compacted) ----------
	$grpSrc = New-Object System.Windows.Forms.GroupBox
	$grpSrc.Text = "Source"
	$grpSrc.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
	$grpSrc.Dock = 'Top'
	$root.Controls.Add($grpSrc, 0, 1)
	
	$srcGrid = New-Object System.Windows.Forms.TableLayoutPanel
	$srcGrid.Dock = 'Top'
	$srcGrid.ColumnCount = 2
	$srcGrid.RowCount = 2
	[void]$srcGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$srcGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$srcGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$srcGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$grpSrc.Controls.Add($srcGrid)
	
	$rbSrv = New-Object System.Windows.Forms.RadioButton
	$rbSrv.Text = "Server (this computer)"
	$rbSrv.Checked = $true
	$rbSrv.AutoSize = $true
	$rbSrv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 2)
	$srcGrid.Controls.Add($rbSrv, 0, 0)
	$srcGrid.SetColumnSpan($rbSrv, 2)
	
	$rbLane = New-Object System.Windows.Forms.RadioButton
	$rbLane.Text = "Lane (source)"
	$rbLane.AutoSize = $true
	$rbLane.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 2)
	$srcGrid.Controls.Add($rbLane, 0, 1)
	
	$laneInline = New-Object System.Windows.Forms.FlowLayoutPanel
	$laneInline.FlowDirection = 'LeftToRight'
	$laneInline.WrapContents = $false
	$laneInline.AutoSize = $true
	$laneInline.AutoSizeMode = 'GrowAndShrink'
	$laneInline.Dock = 'Fill'
	$laneInline.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
	$srcGrid.Controls.Add($laneInline, 1, 1)
	
	$lblLane = New-Object System.Windows.Forms.Label
	$lblLane.Text = "Select lane:"
	$lblLane.AutoSize = $true
	$lblLane.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 2)
	$laneInline.Controls.Add($lblLane)
	
	$cboLane = New-Object System.Windows.Forms.ComboBox
	$cboLane.DropDownStyle = 'DropDownList'
	$cboLane.Enabled = $false
	$cboLane.Width = 260
	$cboLane.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
	$laneInline.Controls.Add($cboLane)
	
	(& $getKnownLanes) | ForEach-Object { [void]$cboLane.Items.Add($_) }
	$rbLane.Add_CheckedChanged({
			$cboLane.Enabled = $rbLane.Checked
			if ($rbLane.Checked -and $cboLane.Items.Count -gt 0 -and -not $cboLane.SelectedItem) { $cboLane.SelectedIndex = 0 }
		})
	
	# ---------- INI path row (ONE ROW via FlowLayoutPanel; resizes combo so the button never wraps) ----------
	$iniFlow = New-Object System.Windows.Forms.FlowLayoutPanel
	$iniFlow.FlowDirection = 'LeftToRight'
	$iniFlow.WrapContents = $false
	$iniFlow.Dock = 'Top'
	$iniFlow.AutoSize = $false
	$iniFlow.Height = 30
	$iniFlow.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
	$root.Controls.Add($iniFlow, 0, 2)
	
	$lblRel = New-Object System.Windows.Forms.Label
	$lblRel.Text = "INI path under \storeman\"
	$lblRel.AutoSize = $true
	$lblRel.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
	$iniFlow.Controls.Add($lblRel)
	
	$cboRel = New-Object System.Windows.Forms.ComboBox
	$cboRel.DropDownStyle = 'DropDown' # editable
	$cboRel.Margin = New-Object System.Windows.Forms.Padding(6, 3, 6, 3)
	$cboRel.Width = 520
	foreach ($p in ($PredefinedIniPaths | Select-Object -Unique)) { [void]$cboRel.Items.Add($p) }
	if ($RelativeIniPath -and -not ($cboRel.Items -contains $RelativeIniPath)) { [void]$cboRel.Items.Add($RelativeIniPath) }
	$cboRel.Text = $RelativeIniPath
	$iniFlow.Controls.Add($cboRel)
	
	$btnLoad = New-Object System.Windows.Forms.Button
	$btnLoad.Text = "Load INI"
	$btnLoad.Width = 100
	$btnLoad.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
	$iniFlow.Controls.Add($btnLoad)
	
	# Resize handler to keep all three controls on ONE line
	$iniFlow.Add_Resize({
			$total = $iniFlow.ClientSize.Width
			$wLabel = $lblRel.PreferredSize.Width
			$wBtn = $btnLoad.Width
			$pad = 24 # rough margins between controls
			$newW = [Math]::Max(220, $total - $wLabel - $wBtn - $pad)
			$cboRel.Width = $newW
		})
	
	# ---------- Section row (tight) ----------
	$rowSec = New-Object System.Windows.Forms.TableLayoutPanel
	$rowSec.Dock = 'Top'
	$rowSec.ColumnCount = 4
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 320)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$rowSec.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 2)
	$root.Controls.Add($rowSec, 0, 3)
	
	$lblSec = New-Object System.Windows.Forms.Label
	$lblSec.Text = "Section:"
	$lblSec.AutoSize = $true
	$lblSec.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 0)
	$rowSec.Controls.Add($lblSec, 0, 0)
	
	$cboSection = New-Object System.Windows.Forms.ComboBox
	$cboSection.DropDownStyle = 'DropDown'
	$cboSection.Width = 320
	$cboSection.Margin = New-Object System.Windows.Forms.Padding(0, 3, 6, 3)
	$rowSec.Controls.Add($cboSection, 1, 0)
	
	$lblPath = New-Object System.Windows.Forms.Label
	$lblPath.Text = ""
	$lblPath.AutoSize = $true
	$lblPath.ForeColor = [System.Drawing.Color]::DimGray
	$lblPath.Margin = New-Object System.Windows.Forms.Padding(12, 6, 6, 0)
	$rowSec.Controls.Add($lblPath, 2, 0)
	
	$btnReload = New-Object System.Windows.Forms.Button
	$btnReload.Text = "Reload"
	$btnReload.Width = 90
	$btnReload.Margin = New-Object System.Windows.Forms.Padding(6, 3, 6, 3)
	$rowSec.Controls.Add($btnReload, 3, 0)
	
	# ---------- Grid (make it take the space) ----------
	$grid = New-Object System.Windows.Forms.DataGridView
	$grid.Dock = 'Fill'
	$grid.AllowUserToAddRows = $false
	$grid.AllowUserToDeleteRows = $true
	$grid.AllowDrop = $true
	$grid.AutoSizeColumnsMode = 'Fill'
	$grid.RowHeadersVisible = $false
	$grid.SelectionMode = 'CellSelect'
	$grid.MultiSelect = $false
	$grid.BackgroundColor = [System.Drawing.Color]::White
	$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 243, 245)
	$grid.EnableHeadersVisualStyles = $false
	$root.Controls.Add($grid, 0, 4)
	
	$dt = New-Object System.Data.DataTable
	[void]$dt.Columns.Add('Key', [string])
	[void]$dt.Columns.Add('Value', [string])
	$grid.DataSource = $dt
	$grid.EditMode = 'EditOnEnter'
	$grid.Columns['Key'].SortMode = 'NotSortable'
	$grid.Columns['Value'].SortMode = 'NotSortable'
	
	# ---------- Bottom buttons ----------
	$btnRow = New-Object System.Windows.Forms.FlowLayoutPanel
	$btnRow.FlowDirection = 'RightToLeft'
	$btnRow.Dock = 'Bottom'
	$btnRow.AutoSize = $true
	$btnRow.Padding = New-Object System.Windows.Forms.Padding(0, 6, 6, 6)
	$frm.Controls.Add($btnRow)
	
	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = "Close"
	$btnClose.Width = 90
	$btnRow.Controls.Add($btnClose)
	
	$btnCopy = New-Object System.Windows.Forms.Button
	$btnCopy.Text = "Copy to other lanes..."
	$btnCopy.Width = 170
	$btnCopy.Enabled = $false # stays off until a successful Save
	$btnRow.Controls.Add($btnCopy)
	
	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Text = "Save to source"
	$btnSave.Width = 130
	$btnRow.Controls.Add($btnSave)
	
	$frm.AcceptButton = $btnSave
	$frm.CancelButton = $btnClose
	
	# ==============================================================================================
	#                                   CONTEXT MENU: Add/Delete
	# ==============================================================================================
	
	$ctx = New-Object System.Windows.Forms.ContextMenuStrip
	$miAdd = $ctx.Items.Add("Add key")
	$miDel = $ctx.Items.Add("Delete key")
	$grid.ContextMenuStrip = $ctx
	
	$script:_ctxRow = -1
	$grid.Add_CellMouseDown({
			param ($sender,
				$e)
			if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right)
			{
				if ($e.RowIndex -ge 0)
				{
					$grid.ClearSelection()
					$script:_ctxRow = $e.RowIndex
					$col = if ($e.ColumnIndex -ge 0) { $e.ColumnIndex }
					else { 0 }
					$grid.CurrentCell = $grid.Rows[$e.RowIndex].Cells[$col]
					$grid.Rows[$e.RowIndex].Selected = $true
				}
				else { $script:_ctxRow = -1 }
			}
		})
	
	# Add key at END; start editing; mark dirty (Copy disabled)
	$miAdd.Add_Click({
			$insertAt = $dt.Rows.Count
			$newRow = $dt.NewRow()
			$newRow['Key'] = ''
			$newRow['Value'] = ''
			[void]$dt.Rows.Add($newRow)
			
			$grid.Focus()
			try { $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $insertAt) }
			catch { }
			$grid.CurrentCell = $grid.Rows[$insertAt].Cells[0]
			$grid.BeginEdit($true)
			
			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})
	
	$miDel.Add_Click({
			$rowIdx = -1
			if ($grid.CurrentCell) { $rowIdx = $grid.CurrentCell.RowIndex }
			if ($rowIdx -lt 0 -and $script:_ctxRow -ge 0) { $rowIdx = $script:_ctxRow }
			if ($rowIdx -lt 0 -or $rowIdx -ge $grid.Rows.Count) { return }
			
			$ans = [System.Windows.Forms.MessageBox]::Show(
				"Delete key '" + [string]$dt.Rows[$rowIdx]['Key'] + "'?",
				"Confirm delete", [System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question)
			if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
			
			$dt.Rows.RemoveAt($rowIdx)
			
			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})
	
	# Any edit => dirty => Copy disabled
	$grid.Add_CellValueChanged({ $state.IsDirty = $true; $btnCopy.Enabled = $false })
	$grid.Add_UserDeletedRow({ $state.IsDirty = $true; $btnCopy.Enabled = $false })
	
	# ==============================================================================================
	#                               DRAG & DROP ROW REORDERING
	# ==============================================================================================
	
	$script:_dragRow = -1
	$script:_dragging = $false
	$script:_dragStart = [System.Drawing.Point]::Empty
	
	$grid.Add_MouseDown({
			param ($s,
				$e)
			$script:_dragStart = [System.Drawing.Point]::new($e.X, $e.Y)
			$hit = $grid.HitTest($e.X, $e.Y)
			if ($hit.RowIndex -ge 0 -and $hit.RowIndex -lt $dt.Rows.Count) { $script:_dragRow = $hit.RowIndex }
			else { $script:_dragRow = -1 }
		})
	
	$grid.Add_MouseMove({
			param ($s,
				$e)
			if ($e.Button -band [System.Windows.Forms.MouseButtons]::Left -and $script:_dragRow -ge 0 -and -not $script:_dragging)
			{
				$dx = [Math]::Abs($e.X - $script:_dragStart.X)
				$dy = [Math]::Abs($e.Y - $script:_dragStart.Y)
				if ($dx -gt 4 -or $dy -gt 4)
				{
					$script:_dragging = $true
					$null = $grid.DoDragDrop("row-move", [System.Windows.Forms.DragDropEffects]::Move)
					$script:_dragging = $false
				}
			}
		})
	
	$grid.Add_DragOver({
			param ($s,
				$e)
			if ($script:_dragRow -lt 0) { $e.Effect = [System.Windows.Forms.DragDropEffects]::None; return }
			$pt = $grid.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
			$hit = $grid.HitTest($pt.X, $pt.Y)
			$target = if ($hit.RowIndex -lt 0) { $dt.Rows.Count - 1 }
			else { $hit.RowIndex }
			$e.Effect = if ($target -ne $script:_dragRow) { [System.Windows.Forms.DragDropEffects]::Move }
			else { [System.Windows.Forms.DragDropEffects]::None }
		})
	
	$grid.Add_DragDrop({
			param ($s,
				$e)
			if ($script:_dragRow -lt 0) { return }
			$source = $script:_dragRow
			$script:_dragRow = -1
			
			$pt = $grid.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
			$hit = $grid.HitTest($pt.X, $pt.Y)
			$target = $hit.RowIndex
			if ($target -lt 0) { $target = $dt.Rows.Count - 1 }
			if ($target -gt $dt.Rows.Count - 1) { $target = $dt.Rows.Count - 1 }
			if ($target -eq $source) { return }
			
			$keyVal = [string]$dt.Rows[$source]['Key']
			$valueVal = [string]$dt.Rows[$source]['Value']
			
			$dt.Rows.RemoveAt($source)
			if ($target -gt $source) { $target-- }
			
			$newRow = $dt.NewRow()
			$newRow['Key'] = $keyVal
			$newRow['Value'] = $valueVal
			if ($target -ge 0 -and $target -lt $dt.Rows.Count) { $dt.Rows.InsertAt($newRow, $target) }
			else { [void]$dt.Rows.Add($newRow); $target = $dt.Rows.Count - 1 }
			
			try { $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $target) }
			catch { }
			$grid.CurrentCell = $grid.Rows[$target].Cells[0]
			$grid.Rows[$target].Selected = $true
			
			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})
	
	# ==============================================================================================
	#                                STATE + LOAD / SAVE / COPY
	# ==============================================================================================
	
	$state = @{
		SourceIsServer = $true
		SourceLane	   = $null
		FullPath	   = $null
		IniObj		   = $null
		DidSave	       = $false
		DidDeploy	   = $false
	}
	
	$loadSectionFromCombo = {
		$dt.Rows.Clear()
		if (-not $state.IniObj) { return }
		$secRequested = ($cboSection.Text -as [string]).Trim()
		if (-not $secRequested) { return }
		$secActual = $null
		foreach ($k in $state.IniObj.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
		if (-not $secActual) { return }
		$od = $state.IniObj.Sections[$secActual]
		foreach ($k in $od.Keys) { [void]$dt.Rows.Add($k, [string]$od[$k]) }
		$state.IsDirty = $false
		$btnCopy.Enabled = $false
	}
	
	$doLoad = {
		$rel = ($cboRel.Text -as [string]).Trim()
		if (-not $rel)
		{
			[System.Windows.Forms.MessageBox]::Show("Enter or select a relative INI path (e.g. office\Setup.ini).", "Missing path",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
			return
		}
		
		if ($rbSrv.Checked)
		{
			$state.SourceIsServer = $true
			$state.SourceLane = $null
			$rootSrv = & $getServerRoot
			$full = & $buildFullPath $rootSrv $rel
		}
		else
		{
			if (-not $cboLane.SelectedItem)
			{
				[System.Windows.Forms.MessageBox]::Show("Select a lane.", "Missing lane",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				return
			}
			$state.SourceIsServer = $false
			$state.SourceLane = [string]$cboLane.SelectedItem
			$rootLane = & $getLaneRoot $state.SourceLane
			if (-not $rootLane)
			{
				[System.Windows.Forms.MessageBox]::Show("Lane storeman root not accessible (\\$($state.SourceLane)\Storeman or \\$($state.SourceLane)\C$\storeman).", "Path error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				return
			}
			$full = & $buildFullPath $rootLane $rel
		}
		
		$state.FullPath = $full
		$lblPath.Text = $state.FullPath
		
		$ini = & $readIni $state.FullPath
		if (-not $ini)
		{
			Write_Log "[Info] INI not found, will be created on save: $($state.FullPath)" "yellow"
			$ini = @{
				RawLines = @("; Created by INI_Editor", "")
				Sections = (New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]')
			}
		}
		$state.IniObj = $ini
		
		$cboSection.Items.Clear()
		foreach ($secName in $state.IniObj.Sections.Keys) { [void]$cboSection.Items.Add($secName) }
		if ($cboSection.Items.Count -gt 0) { $cboSection.SelectedIndex = 0 }
		else { $cboSection.Text = "" }
		
		$loadSectionFromCombo.Invoke()
	}
	
	$btnLoad.Add_Click({ & $doLoad })
	$btnReload.Add_Click({ & $doLoad })
	$cboSection.Add_SelectedIndexChanged({ $loadSectionFromCombo.Invoke() })
	$cboSection.Add_TextChanged({ $loadSectionFromCombo.Invoke() })
	
	# ---------- Save ----------
	$btnSave.Add_Click({
			if (-not $state.FullPath)
			{
				[System.Windows.Forms.MessageBox]::Show("Load an INI first.", "No file",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			if (-not $state.IniObj) { return }
			
			$secRequested = ($cboSection.Text -as [string]).Trim()
			if (-not $secRequested)
			{
				[System.Windows.Forms.MessageBox]::Show("Enter or select a section name first.", "No section",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$kv = New-Object System.Collections.Specialized.OrderedDictionary
			foreach ($row in $dt.Rows)
			{
				$k = [string]$row['Key']; $v = [string]$row['Value']
				if ([string]::IsNullOrWhiteSpace($k)) { continue }
				if ($kv.Contains($k)) { $kv.Remove($k) } # last one wins
				$kv.Add($k, $v)
			}
			
			$secActual = $null
			foreach ($k in $state.IniObj.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
			if (-not $secActual)
			{
				$secActual = $secRequested
				$od = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($k in $kv.Keys) { $od.Add($k, $kv[$k]) }
				$state.IniObj.Sections.Add($secActual, $od)
			}
			else
			{
				$od = $state.IniObj.Sections[$secActual]
				foreach ($k in @($od.Keys)) { $od.Remove($k) }
				foreach ($k in $kv.Keys) { $od.Add($k, $kv[$k]) }
			}
			
			$state.IniObj.RawLines = & $writeIniSection $state.IniObj.RawLines $secActual $kv
			
			try
			{
				$dir = Split-Path -Parent $state.FullPath
				if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
				Set-Content -LiteralPath $state.FullPath -Value $state.IniObj.RawLines -Encoding ASCII -Force
				
				Write_Log "Saved [$secActual] to $($state.FullPath)" "green"
				[System.Windows.Forms.MessageBox]::Show("Saved section [$secActual] to:`r`n$($state.FullPath)", "Saved",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				
				$state.DidSave = $true
				$state.IsDirty = $false
				$btnCopy.Enabled = $true
			}
			catch
			{
				Write_Log "Save failed for $($state.FullPath): $($_.Exception.Message)" "red"
				[System.Windows.Forms.MessageBox]::Show("Save failed:`r`n$($_.Exception.Message)", "Error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
			}
		})
	
	# ---------- Copy (blocked if dirty or not saved) ----------
	$btnCopy.Add_Click({
			if ($state.IsDirty)
			{
				[System.Windows.Forms.MessageBox]::Show("Please Save first. Copy is disabled until the INI is saved.", "Save required",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			if (-not $state.FullPath -or -not (Test-Path -LiteralPath $state.FullPath))
			{
				[System.Windows.Forms.MessageBox]::Show("Nothing to copy. Save the INI first.", "No file",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$secRequested = ($cboSection.Text -as [string]).Trim()
			
			# deploy dialog
			$dlg = New-Object System.Windows.Forms.Form
			$dlg.Text = "Deploy options"
			$dlg.StartPosition = 'CenterParent'
			$dlg.FormBorderStyle = 'FixedDialog'
			$dlg.MaximizeBox = $false
			$dlg.MinimizeBox = $false
			$dlg.ClientSize = New-Object System.Drawing.Size(460, 160)
			
			$lbl = New-Object System.Windows.Forms.Label
			$lbl.AutoSize = $true
			$lbl.MaximumSize = New-Object System.Drawing.Size(440, 0)
			$lbl.Location = New-Object System.Drawing.Point(10, 10)
			$lbl.Text = "How would you like to deploy to other lanes?`r`n`r`n" +
			"* Merge - update only section [$secRequested] keys on targets.`r`n" +
			"* Copy  - replace the entire INI file on targets."
			$dlg.Controls.Add($lbl)
			
			$btnMerge = New-Object System.Windows.Forms.Button
			$btnMerge.Text = "Merge"; $btnMerge.Size = New-Object System.Drawing.Size(90, 28)
			$btnMerge.Location = New-Object System.Drawing.Point(140, 110)
			$btnMerge.DialogResult = [System.Windows.Forms.DialogResult]::Yes
			
			$btnCopyAll = New-Object System.Windows.Forms.Button
			$btnCopyAll.Text = "Copy"; $btnCopyAll.Size = New-Object System.Drawing.Size(90, 28)
			$btnCopyAll.Location = New-Object System.Drawing.Point(240, 110)
			$btnCopyAll.DialogResult = [System.Windows.Forms.DialogResult]::No
			
			$btnCancel = New-Object System.Windows.Forms.Button
			$btnCancel.Text = "Cancel"; $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
			$btnCancel.Location = New-Object System.Drawing.Point(340, 110)
			$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			
			$dlg.Controls.AddRange(@($btnMerge, $btnCopyAll, $btnCancel))
			$dlg.AcceptButton = $btnMerge
			$dlg.CancelButton = $btnCancel
			
			$choice = $dlg.ShowDialog()
			if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
			$mergeOnly = ($choice -eq [System.Windows.Forms.DialogResult]::Yes)
			
			$excluded = @()
			if (-not $state.SourceIsServer -and $state.SourceLane) { $excluded = @($state.SourceLane) }
			
			$StoreNumber = $script:FunctionResults['StoreNumber']
			$sel = $null
			try { $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select destination lanes" -ExcludedNodes $excluded }
			catch { $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select destination lanes" }
			if (-not $sel) { Write_Log "Copy aborted: no lanes selected." "yellow"; return }
			
			$laneNums = @()
			if ($sel -is [System.Collections.IDictionary])
			{
				if ($sel.Contains('Lanes') -and $sel['Lanes']) { $laneNums = @($sel['Lanes']) }
				elseif ($sel.Contains('LaneNumbers') -and $sel['LaneNumbers']) { $laneNums = @($sel['LaneNumbers']) }
			}
			elseif ($sel -is [System.Collections.IEnumerable] -and -not ($sel -is [string])) { $laneNums = @($sel) }
			elseif ($sel -is [string]) { $laneNums = @($sel -split '[,\s]+' | Where-Object { $_ }) }
			$laneNums = @($laneNums | ForEach-Object {
					try { "{0:D3}" -f ([int]$_) }
					catch { "$_" }
				}) | Where-Object { $_ -match '^\d{3}$' }
			if (-not $laneNums -or $laneNums.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("No destination lanes selected.", "No selection",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$destMachines = & $lanesToMachines $laneNums
			if (-not $state.SourceIsServer -and $state.SourceLane) { $destMachines = $destMachines | Where-Object { $_ -ne $state.SourceLane } }
			$destMachines = $destMachines | Select-Object -Unique
			if ($destMachines.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("No valid destination machines resolved.", "No destinations",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$relPath = ($cboRel.Text -as [string]).Trim()
			$srcDir = Split-Path -Parent $state.FullPath
			$srcFile = Split-Path -Leaf $state.FullPath
			$ok = 0; $fail = 0
			
			if ($mergeOnly)
			{
				$kv = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($row in $dt.Rows)
				{
					$k = [string]$row['Key']; $v = [string]$row['Value']
					if ([string]::IsNullOrWhiteSpace($k)) { continue }
					if ($kv.Contains($k)) { $kv.Remove($k) }
					$kv.Add($k, $v)
				}
				
				foreach ($m in $destMachines)
				{
					$rootLane = & $getLaneRoot $m
					if (-not $rootLane) { Write_Log "[Dest Unreachable] $m has no \\Storeman nor \\C$\storeman." "red"; $fail++; continue }
					$dstFull = & $buildFullPath $rootLane $relPath
					$dstDir = Split-Path -Parent $dstFull
					try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
					catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $fail++; continue }
					
					$destIni = & $readIni $dstFull
					if (-not $destIni)
					{
						$destIni = @{ RawLines = @("; Created by INI_Editor (merge)", ""); Sections = (New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]') }
					}
					$secActual = $null
					foreach ($k in $destIni.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
					if (-not $secActual) { $secActual = $secRequested }
					
					$destIni.RawLines = & $writeIniSection $destIni.RawLines $secActual $kv
					try
					{
						Set-Content -LiteralPath $dstFull -Value $destIni.RawLines -Encoding ASCII -Force
						Write_Log "Merged [$secActual] into $dstFull on $m" "green"; $ok++
					}
					catch
					{
						Write_Log "Merge failed on $m ($dstFull): $($_.Exception.Message)" "red"; $fail++
					}
				}
			}
			else
			{
				foreach ($m in $destMachines)
				{
					$rootLane = & $getLaneRoot $m
					if (-not $rootLane) { Write_Log "[Dest Unreachable] $m has no \\Storeman nor \\C$\storeman." "red"; $fail++; continue }
					$dstFull = & $buildFullPath $rootLane $relPath
					$dstDir = Split-Path -Parent $dstFull
					try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
					catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $fail++; continue }
					
					Write_Log "Replacing INI on $m ($dstDir\$srcFile)" "gray"
					$args = @($srcDir, $dstDir, $srcFile, '/COPY:DAT', '/R:2', '/W:2', '/NFL', '/NDL', '/NP', '/MT:8')
					$null = & robocopy @args
					$code = $LASTEXITCODE
					if ($code -ge 8) { Write_Log "Copy to $m FAILED (robocopy $code)" "red"; $fail++ }
					else { Write_Log "Copy to $m OK (robocopy $code)" "green"; $ok++ }
				}
			}
			if ($ok -gt 0) { $state.DidDeploy = $true } # <- at least one target updated OK
			[System.Windows.Forms.MessageBox]::Show("Deploy complete.`r`nOK: $ok   Fail: $fail", "Done",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
		})
	
	# run no-op/activity log when the form actually closes (covers X, Alt+F4, Close button)	$btnClose.Add_Click({
	$frm.Add_FormClosed({
			if (-not $state.DidSave -and -not $state.DidDeploy)
			{
				Write_Log "[No operations] Editor closed without saving or deploying." "gray"
			}
			else
			{
				# Optional: a tiny summary
				$acts = @()
				if ($state.DidSave) { $acts += "save" }
				if ($state.DidDeploy) { $acts += "deploy" }
				Write_Log "[Activity] Completed: $($acts -join ', ')." "gray"
			}
		})
	
	$btnClose.Add_Click({ $frm.Close() })
	
	# ==============================================================================================
	#                                         STARTUP
	# ==============================================================================================
	& $doLoad
	$iniFlow.PerformLayout() # ensure combo gets sized on first draw
	[void]$frm.ShowDialog()
	
	Write_Log "`r`n==================== INI_Editor Completed ====================`r`n" "blue"
}

# ===================================================================================================
#                                FUNCTION: Copy_Files_Between_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Copy one or more \storeman\ folders/files from a chosen source (Server or Lane) to selected lanes.
#   • Quick menu is driven by $QuickItems below; add more entries and they'll appear next run.
#   • Adds the REAL full path (e.g., \\POS001\Storeman\office\System.ini) to the list.
#   • Validates existence on the selected source before adding (quick or manual).
#   • Lane root resolution: \\<Lane>\Storeman → fallback \\<Lane>\C$\storeman.
#   • Files and folders supported. /MIR ignored for single-file operations (safety).
#   • PS 5.1-friendly. No helpers outside this function.
# ---------------------------------------------------------------------------------------------------

function Copy_Files_Between_Nodes
{
	Write_Log "`r`n==================== Starting Copy_Files_Between_Nodes ====================`r`n" "blue"
	
	# =========================================================================================
	# CONFIG: Quick items + optional extra server roots
	#   Label   -> text shown in the Quick... context menu (tip: start with '\storeman\...')
	#   Rel     -> relative path under the storeman root (NO leading slash)
	#   Type    -> "Folder" or "File" (informational)
	#   LaneOnly-> $true to enable only when Lane is selected as source
	# =========================================================================================
	$QuickItems = @(
		@{ Label = "\storeman\office\Htm"; Rel = "office\Htm"; Type = "Folder"; LaneOnly = $false },
		@{ Label = "\storeman\BitMaps"; Rel = "BitMaps"; Type = "Folder"; LaneOnly = $false },
		@{ Label = "\storeman\office\Setting.ini"; Rel = "office\Setting.ini"; Type = "File"; LaneOnly = $true },
		@{ Label = "\storeman\office\System.ini"; Rel = "office\System.ini"; Type = "File"; LaneOnly = $true }
		# Add more here...
		# @{ Label = "\storeman\XchDev\ApiVerifoneMX.ini"; Rel="XchDev\ApiVerifoneMX.ini"; Type="File"; LaneOnly=$false }
	)
	
	# Optional extra server roots to try AFTER $BasePath (if present in any scope)
	$ExtraServerRoots = @(
		# "E:\storeman"
	)
	
	# ---------------- Known lanes for combo (reads your cached maps if present) ----------------
	$knownLanes = @()
	if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'])
	{
		$knownLanes = $script:FunctionResults['LaneMachines'].Values | Where-Object { $_ } | Select-Object -Unique
	}
	elseif ($script:FunctionResults.ContainsKey('LaneNumToMachineName') -and $script:FunctionResults['LaneNumToMachineName'])
	{
		$knownLanes = $script:FunctionResults['LaneNumToMachineName'].Values | Where-Object { $_ } | Select-Object -Unique
	}
	
	# ========================= UI - DPI-safe layout =========================
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	[void][System.Windows.Forms.Application]::EnableVisualStyles()
	
	# --- Form shell ---
	$frm = New-Object System.Windows.Forms.Form
	$frm.Text = "Copy Files Between Nodes - Pick Source and Folders"
	$frm.StartPosition = 'CenterScreen'
	$frm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
	$frm.Font = New-Object System.Drawing.Font('Segoe UI', 9)
	$frm.MinimizeBox = $false
	$frm.MaximizeBox = $false
	$frm.FormBorderStyle = 'FixedDialog'
	$frm.Size = New-Object System.Drawing.Size(760, 560)
	$frm.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
	
	# --- Main layout (header, source, items, bottom) ---
	$layoutMain = New-Object System.Windows.Forms.TableLayoutPanel
	$layoutMain.Dock = 'Fill'
	$layoutMain.Padding = New-Object System.Windows.Forms.Padding(12)
	$layoutMain.ColumnCount = 1
	$layoutMain.RowCount = 4
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$frm.Controls.Add($layoutMain)
	
	# Header
	$lbl = New-Object System.Windows.Forms.Label
	$lbl.Text = "Choose the source and which folder(s)/file(s) to copy to the selected lanes:"
	$lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
	$lbl.AutoSize = $true
	$lbl.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
	$layoutMain.Controls.Add($lbl, 0, 0)
	
	# ---------------- Source group ----------------
	$grpSrc = New-Object System.Windows.Forms.GroupBox
	$grpSrc.Text = "Source"
	$grpSrc.Dock = 'Top'
	$grpSrc.AutoSize = $true
	$grpSrc.AutoSizeMode = 'GrowAndShrink'
	$grpSrc.Padding = New-Object System.Windows.Forms.Padding(12)
	$grpSrc.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
	$layoutMain.Controls.Add($grpSrc, 0, 1)
	
	# Source row: Server (alone), then Lane + combo same row
	$layoutSrc = New-Object System.Windows.Forms.TableLayoutPanel
	$layoutSrc.Dock = 'Top'
	$layoutSrc.AutoSize = $true
	$layoutSrc.AutoSizeMode = 'GrowAndShrink'
	$layoutSrc.ColumnCount = 3
	$layoutSrc.RowCount = 2
	[void]$layoutSrc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutSrc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutSrc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutSrc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutSrc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$grpSrc.Controls.Add($layoutSrc)
	
	$rbServer = New-Object System.Windows.Forms.RadioButton
	$rbServer.Text = "Server (this computer)"
	$rbServer.Checked = $true
	$rbServer.AutoSize = $true
	$layoutSrc.Controls.Add($rbServer, 0, 0)
	$layoutSrc.SetColumnSpan($rbServer, 3)
	
	$rbLane = New-Object System.Windows.Forms.RadioButton
	$rbLane.Text = "Lane (source)"
	$rbLane.AutoSize = $true
	$layoutSrc.Controls.Add($rbLane, 0, 1)
	
	$lblLane = New-Object System.Windows.Forms.Label
	$lblLane.Text = "Select lane:"
	$lblLane.AutoSize = $true
	$lblLane.Margin = New-Object System.Windows.Forms.Padding(12, 3, 6, 0)
	$layoutSrc.Controls.Add($lblLane, 1, 1)
	
	$cboLane = New-Object System.Windows.Forms.ComboBox
	$cboLane.DropDownStyle = 'DropDownList'
	$cboLane.Enabled = $false
	$cboLane.Dock = 'Top'
	$layoutSrc.Controls.Add($cboLane, 2, 1)
	
	$knownLanes | Sort-Object | ForEach-Object { [void]$cboLane.Items.Add($_) }
	$rbLane.Add_CheckedChanged({
			$cboLane.Enabled = $rbLane.Checked
			if ($rbLane.Checked -and $cboLane.Items.Count -gt 0 -and -not $cboLane.SelectedItem) { $cboLane.SelectedIndex = 0 }
		})
	
	# ---------------- Item picker group ----------------
	$grpItems = New-Object System.Windows.Forms.GroupBox
	$grpItems.Text = "Folder(s)/File(s) to Copy"
	$grpItems.Dock = 'Fill'
	$grpItems.Padding = New-Object System.Windows.Forms.Padding(12)
	$grpItems.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
	$layoutMain.Controls.Add($grpItems, 0, 2)
	
	$layoutItems = New-Object System.Windows.Forms.TableLayoutPanel
	$layoutItems.Dock = 'Fill'
	$layoutItems.ColumnCount = 2
	$layoutItems.RowCount = 2
	[void]$layoutItems.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutItems.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutItems.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutItems.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$grpItems.Controls.Add($layoutItems)
	
	$clb = New-Object System.Windows.Forms.CheckedListBox
	$clb.CheckOnClick = $true
	$clb.Dock = 'Fill'
	$clb.HorizontalScrollbar = $true # show long UNC paths completely
	$layoutItems.Controls.Add($clb, 0, 0)
	
	$pnlBtns = New-Object System.Windows.Forms.FlowLayoutPanel
	$pnlBtns.FlowDirection = 'TopDown'
	$pnlBtns.WrapContents = $false
	$pnlBtns.AutoSize = $true
	$pnlBtns.Dock = 'Top'
	$pnlBtns.Padding = New-Object System.Windows.Forms.Padding(0)
	$pnlBtns.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
	$layoutItems.Controls.Add($pnlBtns, 1, 0)
	
	# ---- Add button now says "Add Folder/Files..." and shows a small context menu ----
	$btnAdd = New-Object System.Windows.Forms.Button
	$btnAdd.Text = "Add Folder/Files..."
	$btnAdd.Width = 160; $btnAdd.Height = 30
	$pnlBtns.Controls.Add($btnAdd)
	
	$btnQuick = New-Object System.Windows.Forms.Button
	$btnQuick.Text = "Quick..."
	$btnQuick.Width = 160; $btnQuick.Height = 30
	$btnQuick.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
	$pnlBtns.Controls.Add($btnQuick)
	
	$btnRemove = New-Object System.Windows.Forms.Button
	$btnRemove.Text = "Remove"
	$btnRemove.Width = 160; $btnRemove.Height = 30
	$btnRemove.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
	$pnlBtns.Controls.Add($btnRemove)
	
	$chkMirror = New-Object System.Windows.Forms.CheckBox
	$chkMirror.Text = "Mirror (delete extras on targets)"
	$chkMirror.AutoSize = $true
	$chkMirror.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
	$layoutItems.Controls.Add($chkMirror, 0, 1)
	
	$toolTip = New-Object System.Windows.Forms.ToolTip
	$toolTip.SetToolTip($chkMirror, "Robocopy /MIR will delete files on the destination that are not present in the source.")
	
	# ---------------- Bottom buttons ----------------
	$pnlBottom = New-Object System.Windows.Forms.TableLayoutPanel
	$pnlBottom.AutoSize = $true
	$pnlBottom.ColumnCount = 2
	$pnlBottom.RowCount = 1
	[void]$pnlBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$pnlBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$pnlBottom.Dock = 'Right'
	$layoutMain.Controls.Add($pnlBottom, 0, 3)
	
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Width = 110; $btnOK.Height = 32
	$btnOK.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
	$pnlBottom.Controls.Add($btnOK, 0, 0)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Width = 110; $btnCancel.Height = 32
	$btnCancel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
	$pnlBottom.Controls.Add($btnCancel, 1, 0)
	
	$frm.AcceptButton = $btnOK
	$frm.CancelButton = $btnCancel
	
	# ===================== QUICK CONTEXT MENU (built once; no auto-reopen) =====================
	$cmsQuick = New-Object System.Windows.Forms.ContextMenuStrip
	foreach ($qi in $QuickItems)
	{
		if (-not $qi) { continue }
		
		$mi = New-Object System.Windows.Forms.ToolStripMenuItem
		$mi.Text = $qi.Label
		$mi.Tag = $qi # keep metadata with the menu item
		
		# Click handler MUST use param($s,$e); $s = sender (the menu item)
		$mi.Add_Click({
				param ($s,
					$e)
				
				# Resolve storeman root based on radio selection
				$root = $null
				if ($rbServer.Checked)
				{
					# Prefer any scoped $BasePath, then extras
					$candidates = @($script:BasePath, $global:BasePath, $BasePath) + $ExtraServerRoots
					foreach ($cand in ($candidates | Where-Object { $_ } | Select-Object -Unique))
					{
						if (Test-Path -LiteralPath $cand) { $root = $cand; break }
					}
					if (-not $root)
					{
						[System.Windows.Forms.MessageBox]::Show(
							"Server storeman root not found using `$BasePath (or extras).",
							"Source Unreachable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error
						) | Out-Null
						Write_Log "[Source Unreachable] Server storeman root not found via `$BasePath or extras." "red"
						return
					}
				}
				else
				{
					if (-not $cboLane.SelectedItem)
					{
						[System.Windows.Forms.MessageBox]::Show("Pick a lane as source first.", "Missing Lane",
							[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
						return
					}
					$ln = [string]$cboLane.SelectedItem
					$try1 = "\\$ln\Storeman"; if (Test-Path -LiteralPath $try1) { $root = $try1 }
					if (-not $root) { $try2 = "\\$ln\c$\storeman"; if (Test-Path -LiteralPath $try2) { $root = $try2 } }
					if (-not $root)
					{
						[System.Windows.Forms.MessageBox]::Show("\\$ln\Storeman and \\$ln\C$\storeman are not accessible.",
							"Source Unreachable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
						Write_Log "[Source Unreachable] $ln has no \\Storeman or \\C$\storeman." "red"
						return
					}
				}
				
				# Build full path safely (trim any leading slash in Rel so Join-Path doesn't drop root)
				$meta = $s.Tag
				$rel = ([string]$meta.Rel).Trim().TrimStart('\', '/')
				$full = Join-Path $root $rel
				
				# Verify exists BEFORE adding
				if (-not (Test-Path -LiteralPath $full))
				{
					[System.Windows.Forms.MessageBox]::Show("That item does not exist on the selected source:`r`n$full",
						"Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
					Write_Log "[Source Missing] $full not found; not added to list." "yellow"
					return
				}
				
				# Add real full path (avoid duplicates)
				if (-not ($clb.Items -contains $full))
				{
					[void]$clb.Items.Add($full, $true)
					Write_Log "Added item: $full" "gray"
				}
			})
		
		[void]$cmsQuick.Items.Add($mi)
	}
	
	# On open: toggle LaneOnly items enabled/disabled, then show menu (no auto-reopen)
	$btnQuick.Add_Click({
			foreach ($mi in $cmsQuick.Items)
			{
				if ($mi -is [System.Windows.Forms.ToolStripMenuItem])
				{
					$meta = $mi.Tag
					if ($meta -and $meta.ContainsKey('LaneOnly') -and $meta.LaneOnly) { $mi.Enabled = $rbLane.Checked }
					else { $mi.Enabled = $true }
				}
			}
			$cmsQuick.Show($btnQuick, 0, $btnQuick.Height)
		})
	
	# ===================== "Add Folder/Files..." context menu =====================
	$cmsAdd = New-Object System.Windows.Forms.ContextMenuStrip
	$miAddFolder = New-Object System.Windows.Forms.ToolStripMenuItem
	$miAddFolder.Text = "Add Folder..."
	[void]$cmsAdd.Items.Add($miAddFolder)
	
	$miAddFiles = New-Object System.Windows.Forms.ToolStripMenuItem
	$miAddFiles.Text = "Add File(s)..."
	[void]$cmsAdd.Items.Add($miAddFiles)
	
	# Helper (inline): resolve initial root based on selected source
	$resolveInitialRoot = {
		$initialPath = $null
		if ($rbServer.Checked)
		{
			$candidates = @($script:BasePath, $global:BasePath, $BasePath) + $ExtraServerRoots
			foreach ($cand in ($candidates | Where-Object { $_ } | Select-Object -Unique))
			{
				if (Test-Path -LiteralPath $cand) { $initialPath = $cand; break }
			}
			if (-not $initialPath)
			{
				[System.Windows.Forms.MessageBox]::Show(
					"Server storeman root not found using `$BasePath (or extras).",
					"Source Unreachable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
				Write_Log "[Source Unreachable] Server storeman root not found via `$BasePath or extras." "red"
				return $null
			}
		}
		else
		{
			if (-not $cboLane.SelectedItem)
			{
				[System.Windows.Forms.MessageBox]::Show("Pick a lane as source first.", "Missing Lane",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				return $null
			}
			$ln = [string]$cboLane.SelectedItem
			$try1 = "\\$ln\Storeman"; if (Test-Path -LiteralPath $try1) { $initialPath = $try1 }
			if (-not $initialPath) { $try2 = "\\$ln\c$\storeman"; if (Test-Path -LiteralPath $try2) { $initialPath = $try2 } }
			if (-not $initialPath)
			{
				[System.Windows.Forms.MessageBox]::Show("\\$ln\Storeman and \\$ln\C$\storeman are not accessible.",
					"Source Unreachable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				Write_Log "[Source Unreachable] $ln has no \\Storeman or \\C$\storeman." "red"
				return $null
			}
		}
		return $initialPath
	}
	
	# --- Add Folder... (keeps your old logic, just under the new menu) ---
	$miAddFolder.Add_Click({
			$initialPath = & $resolveInitialRoot
			if (-not $initialPath) { return }
			
			$fd = New-Object System.Windows.Forms.FolderBrowserDialog
			$fd.Description = "Pick a folder under \storeman\ (based on your selected source)."
			$fd.ShowNewFolderButton = $false
			$fd.RootFolder = [System.Environment+SpecialFolder]::Desktop
			$fd.SelectedPath = $initialPath
			
			$res = $fd.ShowDialog()
			if ($res -ne [System.Windows.Forms.DialogResult]::OK -or -not $fd.SelectedPath) { return }
			
			if ($fd.SelectedPath -notmatch '(?i)[\\/](storeman)[\\/]')
			{
				[System.Windows.Forms.MessageBox]::Show("Please choose a folder under \storeman\.", "Not Allowed",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				return
			}
			if (-not (Test-Path -LiteralPath $fd.SelectedPath))
			{
				[System.Windows.Forms.MessageBox]::Show("That folder does not exist on the selected source:`r`n$($fd.SelectedPath)",
					"Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
				Write_Log "[Source Missing] $($fd.SelectedPath) not found; not added to list." "yellow"
				return
			}
			if (-not ($clb.Items -contains $fd.SelectedPath))
			{
				[void]$clb.Items.Add($fd.SelectedPath, $true)
				Write_Log "Added item: $($fd.SelectedPath)" "gray"
			}
		})
	
	# --- Add File(s)... (new) ---
	$miAddFiles.Add_Click({
			$initialPath = & $resolveInitialRoot
			if (-not $initialPath) { return }
			
			$ofd = New-Object System.Windows.Forms.OpenFileDialog
			$ofd.Title = "Pick file(s) under \storeman\ (based on your selected source)"
			$ofd.InitialDirectory = $initialPath
			$ofd.Multiselect = $true
			$ofd.Filter = "All files (*.*)|*.*"
			
			$res = $ofd.ShowDialog()
			if ($res -ne [System.Windows.Forms.DialogResult]::OK -or -not $ofd.FileNames -or $ofd.FileNames.Count -eq 0) { return }
			
			foreach ($f in $ofd.FileNames)
			{
				if ($f -notmatch '(?i)[\\/](storeman)[\\/]')
				{
					[System.Windows.Forms.MessageBox]::Show("Only items under \storeman\ are allowed:`r`n$f",
						"Not Allowed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
					continue
				}
				if (-not (Test-Path -LiteralPath $f))
				{
					[System.Windows.Forms.MessageBox]::Show("That file no longer exists on the selected source:`r`n$f",
						"Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
					Write_Log "[Source Missing] $f not found; not added to list." "yellow"
					continue
				}
				if (-not ($clb.Items -contains $f))
				{
					[void]$clb.Items.Add($f, $true)
					Write_Log "Added item: $f" "gray"
				}
			}
		})
	
	# Show the mini menu when clicking "Add Folder/Files..." (no re-open behavior)
	$btnAdd.Add_Click({ $cmsAdd.Show($btnAdd, 0, $btnAdd.Height) })
	
	# Remove entries
	$btnRemove.Add_Click({
			$toRemove = @()
			foreach ($idx in $clb.CheckedIndices) { $toRemove += $idx }
			if (-not $toRemove -and $clb.SelectedIndex -ge 0) { $toRemove = @($clb.SelectedIndex) }
			$toRemove | Sort-Object -Descending | ForEach-Object { $clb.Items.RemoveAt($_) }
		})
	
	# ---------------- OK / Cancel logic ----------------
	$btnOK.Add_Click({
			$formValid = $true
			if ($rbLane.Checked -and (-not $cboLane.SelectedItem))
			{
				[System.Windows.Forms.MessageBox]::Show("Pick a lane as the source.", "Missing Source",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				$formValid = $false
			}
			if ($clb.CheckedItems.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("Select at least one folder/file to copy.", "Missing Selection",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				$formValid = $false
			}
			if ($formValid -and $chkMirror.Checked)
			{
				$warn = [System.Windows.Forms.MessageBox]::Show(
					"MIRROR mode will delete files in the destination that are not in the source. Continue?",
					"Confirm Mirror",
					[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning
				)
				if ($warn -ne [System.Windows.Forms.DialogResult]::Yes) { $formValid = $false }
			}
			if ($formValid) { $script:__CopyMaps_DialogResult = 'OK'; $frm.Close() }
		})
	$btnCancel.Add_Click({ $script:__CopyMaps_DialogResult = 'Cancel'; $frm.Close() })
	
	[void]$frm.ShowDialog()
	$dialogResult = $script:__CopyMaps_DialogResult
	Remove-Variable -Name __CopyMaps_DialogResult -Scope Script -ErrorAction SilentlyContinue
	if ($dialogResult -ne 'OK')
	{
		Write_Log "User cancelled source/folder selection." "yellow"
		Write_Log "`r`n==================== Copy_Files_Between_Nodes Function Completed ====================" "blue"
		return
	}
	
	# ===================== Gather choices =====================
	$useServerAsSource = $rbServer.Checked
	$sourceMachine = if ($useServerAsSource) { $env:COMPUTERNAME }
	else { [string]$cboLane.SelectedItem }
	$itemsToCopy = @(); foreach ($item in $clb.CheckedItems) { $itemsToCopy += [string]$item }
	$doMirror = [bool]$chkMirror.Checked
	Write_Log "Source machine: $sourceMachine  |  Mirror: $doMirror" "cyan"
	Write_Log "Selected item(s): $($itemsToCopy -join ' | ')" "cyan"
	
	# ===================== Destination selection =====================
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$selection = $null; $triedExclude = $false
	try
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Destination Lanes for Copy" -ExcludedNodes @($sourceMachine)
		$triedExclude = $true
	}
	catch
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Destination Lanes for Copy"
	}
	if ($selection -eq $null) { Write_Log "Destination selection canceled by user." "yellow"; return }
	
	$Lanes = @()
	if ($selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		if ($selection.Lanes[0] -is [PSCustomObject] -and $selection.Lanes[0].PSObject.Properties.Name -contains 'LaneNumber')
		{
			$Lanes = $selection.Lanes | ForEach-Object { $_.LaneNumber }
		}
		else { $Lanes = $selection.Lanes }
	}
	elseif ($selection -is [System.Collections.IEnumerable] -and -not ($selection -is [string])) { $Lanes = @($selection) }
	elseif ($selection -is [string]) { $Lanes = @($selection -split '[,\s]+' | Where-Object { $_ }) }
	else { Write_Log "No destination lanes selected." "yellow"; return }
	
	# Map lane numbers -> machine names
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "[Mapping Missing] LaneNumToMachineName not available; cannot map lanes to machines." "red"; return
	}
	
	$destMachines = @()
	foreach ($ln in $Lanes)
	{
		$s = ("$ln").Trim()
		if ($s -match '[A-Za-z]')
		{
			if ($s -match '^(.*?)[\s]*\(\d{3}\)\s*$') { $s = $matches[1].Trim() } # drop "(001)"
			$destMachines += $s; continue
		}
		if ($s -match '^\d+$')
		{
			$pad3 = ('{0:D3}' -f ([int]$s))
			if ($LaneNumToMachineName.ContainsKey($pad3)) { $destMachines += $LaneNumToMachineName[$pad3] }
			elseif ($LaneNumToMachineName.ContainsKey($s)) { $destMachines += $LaneNumToMachineName[$s] }
			else
			{
				$guess = $knownLanes | Where-Object { "$_" -like "*$pad3" } | Select-Object -First 1
				if ($guess) { $destMachines += $guess }
				else { Write_Log "[Unmapped Lane] No machine for lane $pad3." "yellow" }
			}
			continue
		}
		if ($s -match '\((\d{3})\)$' -or $s -match '(\d{3})$')
		{
			$code = $matches[1]
			if ($LaneNumToMachineName.ContainsKey($code)) { $destMachines += $LaneNumToMachineName[$code] }
			else
			{
				$guess = $knownLanes | Where-Object { "$_" -like "*$code" } | Select-Object -First 1
				if ($guess) { $destMachines += $guess }
				else { Write_Log "[Unmapped Lane] No machine for lane $code." "yellow" }
			}
			continue
		}
		Write_Log "[Parse Warning] Could not interpret destination entry: '$s'." "yellow"
	}
	$destMachines = $destMachines | Where-Object { $_ -and ($_ -ne $sourceMachine) } | Select-Object -Unique
	if (-not $destMachines -or $destMachines.Count -eq 0)
	{
		$msg = if ($triedExclude) { "No non-source lanes selected. Pick at least one destination lane." }
		else { "No non-source lanes selected (picker may not support -ExcludedNodes). Pick at least one lane." }
		[System.Windows.Forms.MessageBox]::Show($msg, "No Destinations",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
		Write_Log "No destination lanes selected." "yellow"; return
	}
	Write_Log "Destinations ($($destMachines.Count)): $($destMachines -join ', ')" "gray"
	
	# ===================== Validate destination roots =====================
	$destRootMap = @{ }
	$unreachable = @()
	foreach ($d in $destMachines)
	{
		$root = $null
		$try1 = "\\$d\Storeman"; if (Test-Path -LiteralPath $try1) { $root = $try1 }
		if (-not $root) { $try2 = "\\$d\c$\storeman"; if (Test-Path -LiteralPath $try2) { $root = $try2 } }
		if ($root) { $destRootMap[$d] = $root }
		else { $unreachable += $d }
	}
	if ($unreachable.Count -gt 0)
	{
		Write_Log "Unreachable lane roots (no \\Storeman or \\C$\storeman): $($unreachable -join ', ')" "red"
		$destMachines = $destMachines | Where-Object { $destRootMap.ContainsKey($_) }
		if (-not $destMachines -or $destMachines.Count -eq 0) { Write_Log "No reachable destinations remain. Aborting." "red"; return }
	}
	
	# ===================== Copy loop (folder OR single file) =====================
	$failList = New-Object System.Collections.Generic.List[string]
	$okCount = 0
	
	foreach ($picked in $itemsToCopy)
	{
		# relative path under \storeman\
		$rel = $null
		$m = [regex]::Match($picked, '(?i)[\\/](storeman)[\\/](?<sub>.*)$')
		if ($m.Success) { $rel = $m.Groups['sub'].Value }
		else { Write_Log "[Invalid Path] '$picked' is not under \storeman\." "red"; $failList.Add("$sourceMachine :: $picked (not under \storeman\)"); continue }
		
		# resolve source full path again (validate)
		if ($useServerAsSource)
		{
			if (-not (Test-Path -LiteralPath $picked)) { Write_Log "[Source Missing] $picked not found on server." "red"; $failList.Add("$sourceMachine :: $picked (missing)"); continue }
			$srcFull = $picked
		}
		else
		{
			$srcRoot = $null
			$tryS1 = "\\$sourceMachine\Storeman"; if (Test-Path -LiteralPath $tryS1) { $srcRoot = $tryS1 }
			if (-not $srcRoot) { $tryS2 = "\\$sourceMachine\c$\storeman"; if (Test-Path -LiteralPath $tryS2) { $srcRoot = $tryS2 } }
			if (-not $srcRoot) { Write_Log "[Source Root Missing] No accessible storeman root on $sourceMachine." "red"; $failList.Add("$sourceMachine :: (no root)"); continue }
			$srcFull = Join-Path $srcRoot $rel
			if (-not (Test-Path -LiteralPath $srcFull)) { Write_Log "[Source Missing] $srcFull not found on lane." "red"; $failList.Add("$sourceMachine :: $srcFull (missing)"); continue }
		}
		
		$isFile = $false
		try { $isFile = Test-Path -LiteralPath $srcFull -PathType Leaf }
		catch { $isFile = $false }
		
		foreach ($dest in $destMachines)
		{
			$dstRoot = $destRootMap[$dest]
			if (-not $dstRoot) { $failList.Add("$dest :: (no root)"); continue }
			
			if ($isFile)
			{
				# Single file copy (no /MIR)
				$dstFull = Join-Path $dstRoot $rel
				$dstDir = Split-Path -Path $dstFull -Parent
				try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
				catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $failList.Add("$dest :: $dstDir (create failed)"); continue }
				
				$srcDir = Split-Path -Path $srcFull -Parent
				$fileName = Split-Path -Path $srcFull -Leaf
				$null = & robocopy @($srcDir, $dstDir, $fileName, '/R:2', '/W:2', '/NFL', '/NDL', '/NP')
				$exit = $LASTEXITCODE
				$failed = ($exit -ge 8)
				$status = if ($failed) { "FAIL ($exit)" }
				elseif ($exit -ge 4) { "OK* ($exit)" }
				else { "OK ($exit)" }
				if ($failed) { Write_Log "Copy file to $dest FAILED: $status" "red"; $failList.Add("$dest :: $rel (robocopy $exit)") }
				else { Write_Log "Copied file '$rel' to $dest $status" "green"; $okCount++ }
				if ($doMirror) { Write_Log "Note: Mirror ignored for single-file item '$rel'." "gray" }
			}
			else
			{
				# Directory copy
				$dstFull = Join-Path $dstRoot $rel
				try { if (-not (Test-Path -LiteralPath $dstFull)) { New-Item -ItemType Directory -Path $dstFull -Force | Out-Null } }
				catch { Write_Log "[Create Failed] $dstFull : $($_.Exception.Message)" "red"; $failList.Add("$dest :: $dstFull (create failed)"); continue }
				
				$args = @($srcFull, $dstFull, '/E', '/COPY:DAT', '/R:2', '/W:2', '/NFL', '/NDL', '/NP', '/MT:8')
				if ($doMirror) { $args += '/MIR' }
				$null = & robocopy @args
				$exit = $LASTEXITCODE
				$failed = ($exit -ge 8)
				$status = if ($failed) { "FAIL ($exit)" }
				elseif ($exit -ge 4) { "OK* ($exit)" }
				else { "OK ($exit)" }
				if ($failed) { Write_Log "Copy to $dest FAILED: $status" "red"; $failList.Add("$dest :: $rel (robocopy $exit)") }
				else { Write_Log "Copy to $dest $status" "green"; $okCount++ }
			}
		}
	}
	
	# ===================== Summary =====================
	$totalOps = ($itemsToCopy.Count) * ($destMachines.Count)
	$failCount = $failList.Count
	Write_Log "`r`n-------------------- Copy Files Summary --------------------" "blue"
	Write_Log "Operations attempted : $totalOps" "gray"
	Write_Log "Successful           : $okCount" "green"
	if ($failCount -gt 0)
	{
		Write_Log "Failed               : $failCount" "red"
		foreach ($f in $failList) { Write_Log "  - $f" "red" }
	}
	else
	{
		Write_Log "Failed               : 0" "green"
	}
	Write_Log "------------------------------------------------------------" "blue"
	Write_Log "`r`n==================== Copy_Files_Between_Nodes Function Completed ====================" "blue"
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
#                                 FUNCTION: Edit_TBS_SCL_ver520_Table
# ---------------------------------------------------------------------------------------------------
# Interactive editor for [TBS_SCL_ver520] WITHOUT DRAGGING and WITHOUT SAVE CHECKBOX.
#   • Auto Sort: BIZERBA first (by numeric IP), then ISHIDA; renumbers BIZERBA from 10,
#                ISHIDA from (last BIZERBA)+10 (or 20 if none). Applies naming/BufferTime
#                template and SAVES immediately.
#   • Save: validates and writes EXACTLY the values visible in the grid (no auto renumber).
#   • Refresh: reloads from DB.
#   • Close/Cancel buttons provided.
#   • IPNetwork shown before IPDevice. All visible columns editable.
#   • Hidden Orig* columns anchor UPDATEs so edits to visible keys are safe.
#   • Auto Sort uses deep snapshots to avoid detached DataRow issues.
#   • Active must be Y/N (uppercased on save). IPNetwork must match local /24 (when detectable).
#   • FIX: All SQL-escape replacements use -replace "'", "''" to avoid parser errors.
#
# CHANGE: No nested helper functions or helper scriptblock variables. All logic is inline.
#         Unavoidable scriptblocks remain only for WinForms event handlers and Sort-Object expressions.
# ===================================================================================================

function Edit_TBS_SCL_ver520_Table
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Edit_TBS_SCL_ver520_Table (No Drag / No Checkbox) ====================`r`n" "blue"
	
	# ----------------------------------------------------------------------------------------------
	# 1) Resolve DB info and build connection string (SQL Auth preferred, else Integrated)
	# ----------------------------------------------------------------------------------------------
	$server = $script:FunctionResults['DBSERVER']
	$database = $script:FunctionResults['DBNAME']
	$connectionString = $script:FunctionResults['ConnectionString']
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	
	if (-not $server -or -not $database)
	{
		Write_Log "DB server or DB name missing - cannot proceed." "red"
		return
	}
	
	if (-not $connectionString)
	{
		$sqlUser = $script:FunctionResults['SQL_UserID']
		$sqlPwd = $script:FunctionResults['SQL_UserPwd']
		if ($sqlUser -and $sqlPwd)
		{
			# CHANGE: Build SQL Auth connection string inline (no helper).
			$connectionString = "Server=$server;Database=$database;User ID=$sqlUser;Password=$sqlPwd;TrustServerCertificate=True;"
			Write_Log "Built SQL Authentication connection string from cached credentials." "cyan"
		}
		else
		{
			# CHANGE: Build Integrated Security connection string inline (no helper).
			$connectionString = "Server=$server;Database=$database;Integrated Security=SSPI;TrustServerCertificate=True;"
			Write_Log "Built Integrated Security connection string (no SQL creds found)." "cyan"
		}
		$script:FunctionResults['ConnectionString'] = $connectionString
	}
	
	# ----------------------------------------------------------------------------------------------
	# 2) Pick SQL runner (Invoke-Sqlcmd if available, else .NET SqlClient)
	# ----------------------------------------------------------------------------------------------
	$InvokeSqlCmd = $null
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch
		{
			Write_Log "SQL module '$SqlModuleName' not loaded. Using .NET SqlClient fallback." "yellow"
		}
	}
	else
	{
		Write_Log "No SQL module name provided. Using .NET SqlClient fallback." "yellow"
	}
	$supportsConnectionString = $InvokeSqlCmd -and ($InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString')
	
	# ----------------------------------------------------------------------------------------------
	# 3) Discover the /24 prefix in use; prefer NIC that can reach the SQL server; else NIC w/ gateway; else first RFC1918
	#     Result: $allowedPrefix like '192.168.5' (used to validate/normalize IPNetwork).
	# ----------------------------------------------------------------------------------------------
	$allowedPrefix = $null
	try
	{
		# Resolve SQL server IPv4s (best effort; instance names may not resolve)
		$serverIPv4s = @()
		try
		{
			$serverIPs = [System.Net.Dns]::GetHostAddresses($server)
			foreach ($ip in $serverIPs)
			{
				if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork)
				{
					$serverIPv4s += $ip
				}
			}
		}
		catch { }
		
		$allNics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object {
			$_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up
		}
		
		# CHANGE: Inline subnet comparison (no helper function). Compare (addr & mask) for equality.
		$pickedIP = $null
		
		# Pass 1: NIC whose IPv4 shares subnet with SQL server
		if ($serverIPv4s.Count -gt 0)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					if ($ua.IPv4Mask -eq $null) { continue }
					
					$aBytes = $ua.Address.GetAddressBytes()
					$mBytes = $ua.IPv4Mask.GetAddressBytes()
					
					$matchFound = $false
					foreach ($srv in $serverIPv4s)
					{
						$bBytes = $srv.GetAddressBytes()
						$same = $true
						for ($i = 0; $i -lt 4; $i++)
						{
							if (($aBytes[$i] -band $mBytes[$i]) -ne ($bBytes[$i] -band $mBytes[$i]))
							{
								$same = $false
								break
							}
						}
						if ($same) { $matchFound = $true; break }
					}
					if ($matchFound) { $pickedIP = $ua.Address.ToString(); break }
				}
				if ($pickedIP) { break }
			}
		}
		
		# Pass 2: NIC with default gateway
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				$hasGw = $false
				foreach ($gw in $ipProps.GatewayAddresses)
				{
					if ($gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
				}
				if (-not $hasGw) { continue }
				
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$pickedIP = $ua.Address.ToString()
					break
				}
				if ($pickedIP) { break }
			}
		}
		
		# Pass 3: first RFC1918 IPv4
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$ip = $ua.Address.ToString()
					if ($ip.StartsWith('10.') -or $ip.StartsWith('192.168.') -or (
							$ip.StartsWith('172.') -and (([int]($ip.Split('.')[1])) -ge 16) -and (([int]($ip.Split('.')[1])) -le 31)
						))
					{
						$pickedIP = $ip
						break
					}
				}
				if ($pickedIP) { break }
			}
		}
		
		if ($pickedIP)
		{
			$parts = $pickedIP.Split('.')
			if ($parts.Length -ge 3)
			{
				$allowedPrefix = "$($parts[0]).$($parts[1]).$($parts[2])"
			}
		}
	}
	catch { }
	
	if ($allowedPrefix)
	{
		Write_Log "IPNetwork allowed prefix detected (active NIC): $allowedPrefix" "cyan"
	}
	else
	{
		Write_Log "WARNING: Could not determine active NIC / RFC1918 IPv4. IPNetwork validation will be skipped." "yellow"
	}
	
	# ----------------------------------------------------------------------------------------------
	# 4) Load table -> DataTable (include Orig* columns to anchor updates)
	# ----------------------------------------------------------------------------------------------
	$selectQuery = @"
SELECT 
    ScaleCode, ScaleName, ScaleBrand, ScaleModel, IPNetwork, IPDevice, BufferTime, Active
FROM [TBS_SCL_ver520]
ORDER BY ScaleCode ASC;
"@
	
	Write_Log "Loading table for edit..." "blue"
	
	# CHANGE: Inline query execution (no $runQuery helper). Return rows.
	$rows = $null
	try
	{
		if ($InvokeSqlCmd)
		{
			if ($supportsConnectionString)
			{
				$rows = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
			}
			else
			{
				$rows = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
			}
		}
		else
		{
			$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
			$csb.ConnectionString = $connectionString
			$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
			$conn.Open()
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $selectQuery; $cmd.CommandTimeout = 0
			$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
			$td = New-Object System.Data.DataTable
			[void]$da.Fill($td)
			$conn.Close()
			$rows = $td
		}
	}
	catch
	{
		Write_Log "SQL error: $($_.Exception.Message)" "red"
		return
	}
	if (-not $rows) { return }
	
	$dt = New-Object System.Data.DataTable
	foreach ($c in 'ScaleCode', 'ScaleName', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice', 'BufferTime', 'Active',
		'OrigScaleBrand', 'OrigScaleModel', 'OrigIPNetwork', 'OrigIPDevice')
	{
		[void]$dt.Columns.Add($c)
	}
	
	foreach ($r in $rows)
	{
		$nr = $dt.NewRow()
		$nr['ScaleCode'] = [string]$r.ScaleCode
		$nr['ScaleName'] = [string]$r.ScaleName
		$nr['ScaleBrand'] = [string]$r.ScaleBrand
		$nr['ScaleModel'] = [string]$r.ScaleModel
		$nr['IPNetwork'] = [string]$r.IPNetwork
		$nr['IPDevice'] = [string]$r.IPDevice
		$nr['BufferTime'] = [string]$r.BufferTime
		$nr['Active'] = [string]$r.Active
		# ORIGINAL KEYS SNAPSHOT
		$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
		$nr['OrigScaleModel'] = [string]$r.ScaleModel
		$nr['OrigIPNetwork'] = [string]$r.IPNetwork
		$nr['OrigIPDevice'] = [string]$r.IPDevice
		[void]$dt.Rows.Add($nr)
	}
	
	# ----------------------------------------------------------------------------------------------
	# 5) WinForms UI - simple layout (no drag UI, no checkbox)
	# ----------------------------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Edit Placings - TBS_SCL_ver520"
	$form.StartPosition = 'CenterScreen'
	$form.Size = New-Object System.Drawing.Size(1100, 720)
	$form.MinimumSize = New-Object System.Drawing.Size(960, 600)
	$form.AutoScaleMode = 'Dpi'
	$form.TopMost = $true
	
	$grid = New-Object System.Windows.Forms.DataGridView
	$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
	$grid.AutoGenerateColumns = $false
	$grid.AllowUserToAddRows = $false
	$grid.AllowUserToDeleteRows = $false
	$grid.ReadOnly = $false
	$grid.SelectionMode = 'FullRowSelect'
	$grid.MultiSelect = $false
	$grid.AutoSizeColumnsMode = 'Fill'
	$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
	$grid.AllowUserToOrderColumns = $true
	$grid.RowHeadersWidth = 30
	
	# CHANGE: Inline column creation (removed $AddCol helper).
	$col_ScaleCode = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleCode.HeaderText = 'ScaleCode'
	$col_ScaleCode.DataPropertyName = 'ScaleCode'
	$col_ScaleCode.ReadOnly = $false
	$col_ScaleCode.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleCode) | Out-Null
	
	$col_ScaleName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleName.HeaderText = 'ScaleName'
	$col_ScaleName.DataPropertyName = 'ScaleName'
	$col_ScaleName.ReadOnly = $false
	$col_ScaleName.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleName) | Out-Null
	
	$col_ScaleBrand = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleBrand.HeaderText = 'ScaleBrand'
	$col_ScaleBrand.DataPropertyName = 'ScaleBrand'
	$col_ScaleBrand.ReadOnly = $false
	$col_ScaleBrand.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleBrand) | Out-Null
	
	$col_ScaleModel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleModel.HeaderText = 'ScaleModel'
	$col_ScaleModel.DataPropertyName = 'ScaleModel'
	$col_ScaleModel.ReadOnly = $false
	$col_ScaleModel.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleModel) | Out-Null
	
	$col_IPNetwork = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_IPNetwork.HeaderText = 'IPNetwork'
	$col_IPNetwork.DataPropertyName = 'IPNetwork'
	$col_IPNetwork.ReadOnly = $false
	$col_IPNetwork.SortMode = 'NotSortable'
	$grid.Columns.Add($col_IPNetwork) | Out-Null
	
	$col_IPDevice = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_IPDevice.HeaderText = 'IPDevice'
	$col_IPDevice.DataPropertyName = 'IPDevice'
	$col_IPDevice.ReadOnly = $false
	$col_IPDevice.SortMode = 'NotSortable'
	$grid.Columns.Add($col_IPDevice) | Out-Null
	
	$col_BufferTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_BufferTime.HeaderText = 'BufferTime'
	$col_BufferTime.DataPropertyName = 'BufferTime'
	$col_BufferTime.ReadOnly = $false
	$col_BufferTime.SortMode = 'NotSortable'
	$grid.Columns.Add($col_BufferTime) | Out-Null
	
	$col_Active = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_Active.HeaderText = 'Active'
	$col_Active.DataPropertyName = 'Active'
	$col_Active.ReadOnly = $false
	$col_Active.SortMode = 'NotSortable'
	$grid.Columns.Add($col_Active) | Out-Null
	
	# Hidden originals
	$col_OrigScaleBrand = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigScaleBrand.HeaderText = 'OrigScaleBrand'
	$col_OrigScaleBrand.DataPropertyName = 'OrigScaleBrand'
	$col_OrigScaleBrand.ReadOnly = $true
	$col_OrigScaleBrand.Visible = $false
	$col_OrigScaleBrand.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigScaleBrand) | Out-Null
	
	$col_OrigScaleModel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigScaleModel.HeaderText = 'OrigScaleModel'
	$col_OrigScaleModel.DataPropertyName = 'OrigScaleModel'
	$col_OrigScaleModel.ReadOnly = $true
	$col_OrigScaleModel.Visible = $false
	$col_OrigScaleModel.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigScaleModel) | Out-Null
	
	$col_OrigIPNetwork = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigIPNetwork.HeaderText = 'OrigIPNetwork'
	$col_OrigIPNetwork.DataPropertyName = 'OrigIPNetwork'
	$col_OrigIPNetwork.ReadOnly = $true
	$col_OrigIPNetwork.Visible = $false
	$col_OrigIPNetwork.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigIPNetwork) | Out-Null
	
	$col_OrigIPDevice = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigIPDevice.HeaderText = 'OrigIPDevice'
	$col_OrigIPDevice.DataPropertyName = 'OrigIPDevice'
	$col_OrigIPDevice.ReadOnly = $true
	$col_OrigIPDevice.Visible = $false
	$col_OrigIPDevice.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigIPDevice) | Out-Null
	
	$grid.DataSource = $dt
	$form.Controls.Add($grid)
	
	# Bottom bar
	$bottom = New-Object System.Windows.Forms.Panel
	$bottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
	$bottom.Height = 86
	$form.Controls.Add($bottom)
	
	$lbl = New-Object System.Windows.Forms.Label
	$lbl.AutoSize = $true
	$lbl.Location = New-Object System.Drawing.Point(10, 10)
	if ($allowedPrefix)
	{
		$lbl.Text = "Edit cells directly. Auto Sort applies standard order + numbering and saves immediately.  Active must be Y/N.  IPNetwork must match $allowedPrefix.* (accepted: '$allowedPrefix', '$allowedPrefix.0', '$allowedPrefix.*', '$allowedPrefix.x', '$allowedPrefix.')"
	}
	else
	{
		$lbl.Text = "Edit cells directly. Auto Sort applies standard order + numbering and saves immediately.  Active must be Y/N."
	}
	$bottom.Controls.Add($lbl)
	
	$btnAuto = New-Object System.Windows.Forms.Button
	$btnAuto.Text = "Auto Sort (BIZERBA first)"
	$btnAuto.Size = New-Object System.Drawing.Size(220, 28)
	$btnAuto.Location = New-Object System.Drawing.Point(10, 44)
	$bottom.Controls.Add($btnAuto)
	
	$btnRefresh = New-Object System.Windows.Forms.Button
	$btnRefresh.Text = "Refresh"
	$btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
	$btnRefresh.Location = New-Object System.Drawing.Point(240, 44)
	$bottom.Controls.Add($btnRefresh)
	
	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = "Close"
	$btnClose.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnClose)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnCancel)
	
	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Text = "Save"
	$btnSave.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnSave)
	
	# CHANGE: Inline initial placement (removed $placeButtons helper)
	$bw = [int]$bottom.ClientSize.Width
	$btnSave.Location = New-Object System.Drawing.Point(($bw - 110), 44)
	$btnCancel.Location = New-Object System.Drawing.Point(($bw - 220), 44)
	$btnClose.Location = New-Object System.Drawing.Point(($bw - 330), 44)
	
	# CHANGE: Inline resize handler computes placements directly
	$form.add_Resize({
			$bw2 = [int]$bottom.ClientSize.Width
			$btnSave.Location = New-Object System.Drawing.Point(($bw2 - 110), 44)
			$btnCancel.Location = New-Object System.Drawing.Point(($bw2 - 220), 44)
			$btnClose.Location = New-Object System.Drawing.Point(($bw2 - 330), 44)
		})
	
	# ----------------------------------------------------------------------------------------------
	# Helper-less inline SAVE logic (duplicated in Save and Auto Sort to keep "all inline" constraint)
	# ----------------------------------------------------------------------------------------------
	$sqlSaveTemplate = @'
BEGIN TRY
  BEGIN TRAN;

  DECLARE @Placings TABLE(
      OrigScaleBrand nvarchar(100) NOT NULL,
      OrigScaleModel nvarchar(100) NOT NULL,
      OrigIPNetwork  nvarchar(100) NOT NULL,
      OrigIPDevice   nvarchar(100) NOT NULL,
      NewScaleBrand  nvarchar(100) NOT NULL,
      NewScaleModel  nvarchar(100) NOT NULL,
      NewIPNetwork   nvarchar(100) NOT NULL,
      NewIPDevice    nvarchar(100) NOT NULL,
      NewScaleCode   int           NOT NULL,
      NewBufferTime  nvarchar(20)  NOT NULL,
      NewScaleName   nvarchar(200) NOT NULL,
      NewActive      nvarchar(20)  NOT NULL
  );

  INSERT INTO @Placings (
      OrigScaleBrand, OrigScaleModel, OrigIPNetwork, OrigIPDevice,
      NewScaleBrand, NewScaleModel, NewIPNetwork, NewIPDevice,
      NewScaleCode, NewBufferTime, NewScaleName, NewActive
  )
  VALUES
  $(VALUES_BLOCK);

  UPDATE tgt
     SET tgt.ScaleBrand = src.NewScaleBrand,
         tgt.ScaleModel = src.NewScaleModel,
         tgt.IPNetwork  = src.NewIPNetwork,
         tgt.IPDevice   = src.NewIPDevice,
         tgt.ScaleCode  = src.NewScaleCode,
         tgt.BufferTime = src.NewBufferTime,
         tgt.ScaleName  = src.NewScaleName,
         tgt.Active     = src.NewActive
  FROM [TBS_SCL_ver520] AS tgt
  JOIN @Placings AS src
    ON  tgt.ScaleBrand = src.OrigScaleBrand
    AND tgt.ScaleModel = src.OrigScaleModel
    AND tgt.IPNetwork  = src.OrigIPNetwork
    AND tgt.IPDevice   = src.OrigIPDevice;

  COMMIT TRAN;
END TRY
BEGIN CATCH
  IF @@TRANCOUNT > 0 ROLLBACK TRAN;
  DECLARE @msg nvarchar(4000) = ERROR_MESSAGE();
  RAISERROR(@msg, 16, 1);
END CATCH;
'@
	
	# ----------------------------------------------------------------------------------------------
	# 6) Button: Save (validate + write exactly what's in grid; no renumber)
	# ----------------------------------------------------------------------------------------------
	$btnSave.add_Click({
			# Commit any in-progress grid edits (inline, no helper)
			if ($grid -and $grid.IsCurrentCellInEditMode) { [void]$grid.EndEdit() }
			else { if ($grid) { [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } }
			
			# Validate ScaleCode unique & numeric; Active Y/N; IPNetwork matches allowed /24 if available
			$seenCodes = @{ }
			$rowIdx = 0
			foreach ($row in $dt.Rows)
			{
				$raw = $row['ScaleCode']
				$val = $null
				if (-not [int]::TryParse([string]$raw, [ref]$val))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid ScaleCode at row #$($rowIdx + 1): '$raw' is not a number.", "Invalid Code", 'OK', 'Error')
					return
				}
				if ($seenCodes.ContainsKey($val))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Duplicate ScaleCode '$val' detected (row #$($rowIdx + 1)).", "Duplicate Code", 'OK', 'Error')
					return
				}
				$seenCodes[$val] = $true
				
				$actRaw = [string]$row['Active']; if ($null -eq $actRaw) { $actRaw = '' }
				$actUp = $actRaw.Trim().ToUpper()
				if ($actUp -ne 'Y' -and $actUp -ne 'N')
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid Active value at row #$($rowIdx + 1): '$actRaw'. Use Y or N.", "Invalid Active", 'OK', 'Error')
					return
				}
				$row['Active'] = $actUp
				
				if ($allowedPrefix)
				{
					$ipnRaw = [string]$row['IPNetwork']; if ($null -eq $ipnRaw) { $ipnRaw = '' }
					$ipnRaw = $ipnRaw.Trim()
					$m = [regex]::Match($ipnRaw, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if (-not $m.Success)
					{
						[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #$($rowIdx + 1): '$ipnRaw'. Expected '$allowedPrefix', '$allowedPrefix.0' or '$allowedPrefix.*'.", "Invalid IPNetwork", 'OK', 'Error')
						return
					}
					$prefix = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)"
					if ($prefix -ne $allowedPrefix)
					{
						[void][System.Windows.Forms.MessageBox]::Show("IPNetwork at row #$($rowIdx + 1) must be within '$allowedPrefix.*'. Got '$ipnRaw'.", "IPNetwork Out of Range", 'OK', 'Error')
						return
					}
					# Normalize to bare prefix (no trailing .x)
					$row['IPNetwork'] = $allowedPrefix
				}
				
				$rowIdx = $rowIdx + 1
			}
			
			# Build VALUES block with proper SQL escaping
			$values = New-Object System.Collections.Generic.List[string]
			foreach ($row in $dt.Rows)
			{
				$nBrand = ([string]$row['ScaleBrand']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nModel = ([string]$row['ScaleModel']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nIPNet = ([string]$row['IPNetwork']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nIPDev = ([string]$row['IPDevice']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nCode = [int]$row['ScaleCode']
				$nBuf = ([string]$row['BufferTime']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nName = ([string]$row['ScaleName']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nAct = ([string]$row['Active']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				
				$oBrand = (([string]$row['OrigScaleBrand'])) -replace "'", "''"
				$oModel = (([string]$row['OrigScaleModel'])) -replace "'", "''"
				$oIPNet = (([string]$row['OrigIPNetwork'])) -replace "'", "''"
				$oIPDev = (([string]$row['OrigIPDevice'])) -replace "'", "''"
				
				$values.Add("('$oBrand','$oModel','$oIPNet','$oIPDev','$nBrand','$nModel','$nIPNet','$nIPDev',$nCode,'$nBuf','$nName','$nAct')")
			}
			if ($values.Count -eq 0)
			{
				Write_Log "Nothing to save." "yellow"
				return
			}
			
			$sqlSave = $sqlSaveTemplate.Replace('$(VALUES_BLOCK)', ([string]::Join(",`r`n  ", $values)))
			
			Write_Log "Saving to database..." "blue"
			
			# Execute non-query inline
			$ok = $true
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						& $InvokeSqlCmd -ConnectionString $connectionString -Query $sqlSave -ErrorAction Stop | Out-Null
					}
					else
					{
						& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $sqlSave -ErrorAction Stop | Out-Null
					}
				}
				else
				{
					$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb.ConnectionString = $connectionString
					$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
					$conn.Open()
					$cmd = $conn.CreateCommand(); $cmd.CommandText = $sqlSave; $cmd.CommandTimeout = 0
					[void]$cmd.ExecuteNonQuery()
					$conn.Close()
				}
			}
			catch
			{
				$ok = $false
				Write_Log "Save failed: $($_.Exception.Message)" "red"
				[void][System.Windows.Forms.MessageBox]::Show("Save failed. Check logs.", "Error", 'OK', 'Error')
			}
			if (-not $ok) { return }
			
			# Refresh view from DB
			Write_Log "Refreshing view..." "blue"
			$fresh = $null
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						$fresh = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
					}
					else
					{
						$fresh = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					}
				}
				else
				{
					$csb2 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb2.ConnectionString = $connectionString
					$conn2 = New-Object System.Data.SqlClient.SqlConnection $csb2.ConnectionString
					$conn2.Open()
					$cmd2 = $conn2.CreateCommand(); $cmd2.CommandText = $selectQuery; $cmd2.CommandTimeout = 0
					$da2 = New-Object System.Data.SqlClient.SqlDataAdapter $cmd2
					$td2 = New-Object System.Data.DataTable
					[void]$da2.Fill($td2)
					$conn2.Close()
					$fresh = $td2
				}
			}
			catch
			{
				Write_Log "SQL error (refresh): $($_.Exception.Message)" "red"
				return
			}
			
			if ($fresh)
			{
				$dt.Rows.Clear() | Out-Null
				foreach ($r in $fresh)
				{
					$nr = $dt.NewRow()
					$nr['ScaleCode'] = [string]$r.ScaleCode
					$nr['ScaleName'] = [string]$r.ScaleName
					$nr['ScaleBrand'] = [string]$r.ScaleBrand
					$nr['ScaleModel'] = [string]$r.ScaleModel
					$nr['IPNetwork'] = [string]$r.IPNetwork
					$nr['IPDevice'] = [string]$r.IPDevice
					$nr['BufferTime'] = [string]$r.BufferTime
					$nr['Active'] = [string]$r.Active
					$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
					$nr['OrigScaleModel'] = [string]$r.ScaleModel
					$nr['OrigIPNetwork'] = [string]$r.IPNetwork
					$nr['OrigIPDevice'] = [string]$r.IPDevice
					[void]$dt.Rows.Add($nr)
				}
				$grid.Refresh()
			}
			
			[void][System.Windows.Forms.MessageBox]::Show("Saved and refreshed.", "Success", 'OK', 'Information')
			Write_Log "Saved and refreshed." "green"
		})
	
	# ----------------------------------------------------------------------------------------------
	# 7) Auto Sort: BIZERBA by numeric IP then ISHIDA, renumber, template names, immediate SAVE
	# ----------------------------------------------------------------------------------------------
	$btnAuto.add_Click({
			# Deep snapshots split by brand to avoid DataRow detach issues
			$snapB = @()
			$snapI = @()
			foreach ($row in $dt.Rows)
			{
				$o = [ordered]@{
					ScaleCode	   = $row['ScaleCode']; ScaleName = $row['ScaleName']; ScaleBrand = $row['ScaleBrand']; ScaleModel = $row['ScaleModel']
					IPNetwork	   = $row['IPNetwork']; IPDevice = $row['IPDevice']; BufferTime = $row['BufferTime']; Active = $row['Active']
					OrigScaleBrand = $row['OrigScaleBrand']; OrigScaleModel = $row['OrigScaleModel']; OrigIPNetwork = $row['OrigIPNetwork']; OrigIPDevice = $row['OrigIPDevice']
				}
				if ($o.ScaleBrand -eq 'BIZERBA') { $snapB += $o }
				elseif ($o.ScaleBrand -eq 'ISHIDA') { $snapI += $o }
			}
			
			# Sort by numeric IPDevice (non-numeric last), then by raw IPDevice
			$snapB = $snapB | Sort-Object @{ Expression = { $n = $null; if ([int]::TryParse([string]$_.IPDevice, [ref]$n)) { $n }
					else { [int]::MaxValue } } },
										  @{ Expression = { $_.IPDevice } }
			$snapI = $snapI | Sort-Object @{ Expression = { $n = $null; if ([int]::TryParse([string]$_.IPDevice, [ref]$n)) { $n }
					else { [int]::MaxValue } } },
										  @{ Expression = { $_.IPDevice } }
			
			$dt.BeginLoadData()
			try
			{
				$dt.Rows.Clear()
				foreach ($s in $snapB) { $nr = $dt.NewRow(); foreach ($k in $s.Keys) { $nr[$k] = [string]$s[$k] }; [void]$dt.Rows.Add($nr) }
				foreach ($s in $snapI) { $nr = $dt.NewRow(); foreach ($k in $s.Keys) { $nr[$k] = [string]$s[$k] }; [void]$dt.Rows.Add($nr) }
			}
			finally { $dt.EndLoadData() }
			
			# Template naming + BufferTime for BIZERBA; Ishida WMAI name rule
			$firstB = $true
			foreach ($r in $dt.Rows)
			{
				if ($r['ScaleBrand'] -eq 'BIZERBA')
				{
					$r['ScaleName'] = "Scale $($r['IPDevice'])"
					if ($firstB) { $r['BufferTime'] = '1'; $firstB = $false }
					else { $r['BufferTime'] = '5' }
				}
			}
			$ishWmai = @()
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'ISHIDA' -and $r['ScaleModel'] -eq 'WMAI') { $ishWmai += $r } }
			if ($ishWmai.Count -gt 1) { foreach ($r in $ishWmai) { $r['ScaleName'] = "Ishida Wrapper $($r['IPDevice'])" } }
			elseif ($ishWmai.Count -eq 1) { $ishWmai[0]['ScaleName'] = 'Ishida Wrapper' }
			
			# Renumber: BIZERBA start at 10; ISHIDA start at (last BIZERBA)+10 or 20 if none
			$codeB = 10
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'BIZERBA') { $r['ScaleCode'] = [string]$codeB; $codeB = $codeB + 1 } }
			$ishidaStart = 20
			if ($codeB -gt 10) { $ishidaStart = ($codeB - 1) + 10 }
			$codeI = $ishidaStart
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'ISHIDA') { $r['ScaleCode'] = [string]$codeI; $codeI = $codeI + 1 } }
			
			$grid.Refresh()
			
			# Immediately SAVE (same validation rules as Save button)
			if ($grid -and $grid.IsCurrentCellInEditMode) { [void]$grid.EndEdit() }
			else { if ($grid) { [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } }
			
			$seenCodes2 = @{ }
			$rowIdx2 = 0
			foreach ($row2 in $dt.Rows)
			{
				$raw2 = $row2['ScaleCode']
				$val2 = $null
				if (-not [int]::TryParse([string]$raw2, [ref]$val2))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid ScaleCode at row #$($rowIdx2 + 1): '$raw2' is not a number.", "Invalid Code", 'OK', 'Error')
					return
				}
				if ($seenCodes2.ContainsKey($val2))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Duplicate ScaleCode '$val2' detected (row #$($rowIdx2 + 1)).", "Duplicate Code", 'OK', 'Error')
					return
				}
				$seenCodes2[$val2] = $true
				
				$actRaw2 = [string]$row2['Active']; if ($null -eq $actRaw2) { $actRaw2 = '' }
				$actUp2 = $actRaw2.Trim().ToUpper()
				if ($actUp2 -ne 'Y' -and $actUp2 -ne 'N')
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid Active value at row #$($rowIdx2 + 1): '$actRaw2'. Use Y or N.", "Invalid Active", 'OK', 'Error')
					return
				}
				$row2['Active'] = $actUp2
				
				if ($allowedPrefix)
				{
					$ipnRaw2 = [string]$row2['IPNetwork']; if ($null -eq $ipnRaw2) { $ipnRaw2 = '' }
					$ipnRaw2 = $ipnRaw2.Trim()
					$m2 = [regex]::Match($ipnRaw2, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if (-not $m2.Success)
					{
						[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #$($rowIdx2 + 1): '$ipnRaw2'. Expected '$allowedPrefix', '$allowedPrefix.0' or '$allowedPrefix.*'.", "Invalid IPNetwork", 'OK', 'Error')
						return
					}
					$prefix2 = "$($m2.Groups[1].Value).$($m2.Groups[2].Value).$($m2.Groups[3].Value)"
					if ($prefix2 -ne $allowedPrefix)
					{
						[void][System.Windows.Forms.MessageBox]::Show("IPNetwork at row #$($rowIdx2 + 1) must be within '$allowedPrefix.*'. Got '$ipnRaw2'.", "IPNetwork Out of Range", 'OK', 'Error')
						return
					}
					$row2['IPNetwork'] = $allowedPrefix
				}
				
				$rowIdx2 = $rowIdx2 + 1
			}
			
			$values2 = New-Object System.Collections.Generic.List[string]
			foreach ($row3 in $dt.Rows)
			{
				$nBrand2 = ([string]$row3['ScaleBrand']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nModel2 = ([string]$row3['ScaleModel']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nIPNet2 = ([string]$row3['IPNetwork']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nIPDev2 = ([string]$row3['IPDevice']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nCode2 = [int]$row3['ScaleCode']
				$nBuf2 = ([string]$row3['BufferTime']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nName2 = ([string]$row3['ScaleName']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nAct2 = ([string]$row3['Active']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				
				$oBrand2 = (([string]$row3['OrigScaleBrand'])) -replace "'", "''"
				$oModel2 = (([string]$row3['OrigScaleModel'])) -replace "'", "''"
				$oIPNet2 = (([string]$row3['OrigIPNetwork'])) -replace "'", "''"
				$oIPDev2 = (([string]$row3['OrigIPDevice'])) -replace "'", "''"
				
				$values2.Add("('$oBrand2','$oModel2','$oIPNet2','$oIPDev2','$nBrand2','$nModel2','$nIPNet2','$nIPDev2',$nCode2,'$nBuf2','$nName2','$nAct2')")
			}
			if ($values2.Count -eq 0)
			{
				Write_Log "Nothing to save after Auto Sort." "yellow"
				return
			}
			
			$sqlSave2 = $sqlSaveTemplate.Replace('$(VALUES_BLOCK)', ([string]::Join(",`r`n  ", $values2)))
			
			Write_Log "Saving to database after Auto Sort..." "blue"
			
			$ok2 = $true
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						& $InvokeSqlCmd -ConnectionString $connectionString -Query $sqlSave2 -ErrorAction Stop | Out-Null
					}
					else
					{
						& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $sqlSave2 -ErrorAction Stop | Out-Null
					}
				}
				else
				{
					$csb3 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb3.ConnectionString = $connectionString
					$conn3 = New-Object System.Data.SqlClient.SqlConnection $csb3.ConnectionString
					$conn3.Open()
					$cmd3 = $conn3.CreateCommand(); $cmd3.CommandText = $sqlSave2; $cmd3.CommandTimeout = 0
					[void]$cmd3.ExecuteNonQuery()
					$conn3.Close()
				}
			}
			catch
			{
				$ok2 = $false
				Write_Log "Save (Auto Sort) failed: $($_.Exception.Message)" "red"
				[void][System.Windows.Forms.MessageBox]::Show("Save after Auto Sort failed. Check logs.", "Error", 'OK', 'Error')
			}
			if (-not $ok2) { return }
			
			# Refresh view from DB
			Write_Log "Refreshing view after Auto Sort..." "blue"
			$fresh2 = $null
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						$fresh2 = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
					}
					else
					{
						$fresh2 = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					}
				}
				else
				{
					$csb4 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb4.ConnectionString = $connectionString
					$conn4 = New-Object System.Data.SqlClient.SqlConnection $csb4.ConnectionString
					$conn4.Open()
					$cmd4 = $conn4.CreateCommand(); $cmd4.CommandText = $selectQuery; $cmd4.CommandTimeout = 0
					$da4 = New-Object System.Data.SqlClient.SqlDataAdapter $cmd4
					$td4 = New-Object System.Data.DataTable
					[void]$da4.Fill($td4)
					$conn4.Close()
					$fresh2 = $td4
				}
			}
			catch
			{
				Write_Log "SQL error (refresh after Auto Sort): $($_.Exception.Message)" "red"
				return
			}
			
			if ($fresh2)
			{
				$dt.Rows.Clear() | Out-Null
				foreach ($r in $fresh2)
				{
					$nr = $dt.NewRow()
					$nr['ScaleCode'] = [string]$r.ScaleCode
					$nr['ScaleName'] = [string]$r.ScaleName
					$nr['ScaleBrand'] = [string]$r.ScaleBrand
					$nr['ScaleModel'] = [string]$r.ScaleModel
					$nr['IPNetwork'] = [string]$r.IPNetwork
					$nr['IPDevice'] = [string]$r.IPDevice
					$nr['BufferTime'] = [string]$r.BufferTime
					$nr['Active'] = [string]$r.Active
					$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
					$nr['OrigScaleModel'] = [string]$r.ScaleModel
					$nr['OrigIPNetwork'] = [string]$r.IPNetwork
					$nr['OrigIPDevice'] = [string]$r.IPDevice
					[void]$dt.Rows.Add($nr)
				}
				$grid.Refresh()
			}
			
			[void][System.Windows.Forms.MessageBox]::Show("Auto Sort applied, saved, and refreshed.", "Success", 'OK', 'Information')
			Write_Log "Auto Sort applied, saved, and refreshed." "green"
		})
	
	# ----------------------------------------------------------------------------------------------
	# 8) Refresh / Close / Cancel wiring (inline)
	# ----------------------------------------------------------------------------------------------
	$btnRefresh.add_Click({
			Write_Log "Refreshing view from database..." "blue"
			$fresh = $null
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						$fresh = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
					}
					else
					{
						$fresh = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					}
				}
				else
				{
					$csb5 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb5.ConnectionString = $connectionString
					$conn5 = New-Object System.Data.SqlClient.SqlConnection $csb5.ConnectionString
					$conn5.Open()
					$cmd5 = $conn5.CreateCommand(); $cmd5.CommandText = $selectQuery; $cmd5.CommandTimeout = 0
					$da5 = New-Object System.Data.SqlClient.SqlDataAdapter $cmd5
					$td5 = New-Object System.Data.DataTable
					[void]$da5.Fill($td5)
					$conn5.Close()
					$fresh = $td5
				}
			}
			catch
			{
				Write_Log "SQL error (refresh): $($_.Exception.Message)" "red"
				return
			}
			
			if ($fresh)
			{
				$dt.Rows.Clear() | Out-Null
				foreach ($r in $fresh)
				{
					$nr = $dt.NewRow()
					$nr['ScaleCode'] = [string]$r.ScaleCode
					$nr['ScaleName'] = [string]$r.ScaleName
					$nr['ScaleBrand'] = [string]$r.ScaleBrand
					$nr['ScaleModel'] = [string]$r.ScaleModel
					$nr['IPNetwork'] = [string]$r.IPNetwork
					$nr['IPDevice'] = [string]$r.IPDevice
					$nr['BufferTime'] = [string]$r.BufferTime
					$nr['Active'] = [string]$r.Active
					$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
					$nr['OrigScaleModel'] = [string]$r.ScaleModel
					$nr['OrigIPNetwork'] = [string]$r.IPNetwork
					$nr['OrigIPDevice'] = [string]$r.IPDevice
					[void]$dt.Rows.Add($nr)
				}
				$grid.Refresh()
			}
		})
	
	$btnClose.add_Click({ $form.Close() })
	$btnCancel.add_Click({
			Write_Log "Cancel: reloading from database..." "yellow"
			$btnRefresh.PerformClick()
		})
	
	# ----------------------------------------------------------------------------------------------
	# 9) Show dialog
	# ----------------------------------------------------------------------------------------------
	[void]$form.ShowDialog()
	Write_Log "==================== Edit_TBS_SCL_ver520_Table Closed ====================" "blue"
}

# ===================================================================================================
#                               FUNCTION: Troubleshoot_ScaleCommApp
# ---------------------------------------------------------------------------------------------------
# Description:
#   Audits & (optionally) repairs ScaleCommApp *.exe.config files:
#     - StoreName                  => script's store name
#     - FistScaleIP                => FIRST active scale full IP from IPNetwork + IPDevice
#                                     (robust parser + table/active-NIC fallbacks)
#     - SQL_InstanceName           => script's SQL instance
#     - SQL_databaseName           => script's DB name
#     - Recs_BatchSendFull         => "10000"
#
# Logging policy (consolidated):
#   - Start/End banners.
#   - DB/NIC discovery logged ONCE.
#   - Show ONE "Desired Settings" block and ONE "Existing (first file)" snapshot.
#   - No per-file OK/mismatch spam; only a final summary and list of saved files.
#
# Implementation:
#   - Uses Invoke-Sqlcmd (if available) or .NET SqlClient.
#   - Saves UTF-8 (no BOM), timestamped .bak, respects -Force and -WhatIf.
#   - Honors -AllMatches (else only first file), -NoAutoFix to audit-only.
#   - **NO TERNARY / NO '??'** - PS 5.1 compatible.
# ===================================================================================================

function Troubleshoot_ScaleCommApp
{
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[string]$ConfigPath,
		[string[]]$ScaleCommRoots = @('C:\ScaleCommApp', 'D:\ScaleCommApp'),
		[switch]$AllMatches,
		[switch]$NoAutoFix,
		[switch]$Force,
		[string]$ExpectedStoreName,
		[string]$ExpectedSqlInstance,
		[string]$ExpectedDatabase
	)
	
	Write_Log "`r`n==================== Starting Troubleshoot_ScaleCommApp ====================`r`n" "blue"
	
	# ------------------------------------------------------------------------------------------------
	# Resolve expectations & connection (same behavior; no ternary)
	# ------------------------------------------------------------------------------------------------
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']
	$connStr = $script:FunctionResults['ConnectionString']
	$sqlModule = $script:FunctionResults['SqlModuleName']
	
	if (-not $ExpectedDatabase -and $dbName) { $ExpectedDatabase = $dbName }
	if (-not $ExpectedSqlInstance -and $server) { $ExpectedSqlInstance = $server }
	if (-not $ExpectedStoreName)
	{
		$ExpectedStoreName = $script:FunctionResults['StoreName']
		if (-not $ExpectedStoreName) { $ExpectedStoreName = $script:FunctionResults['STORENAME'] }
	}
	
	if (-not $connStr -and $server -and $dbName)
	{
		$sqlUser = $script:FunctionResults['SQL_UserID']
		$sqlPwd = $script:FunctionResults['SQL_UserPwd']
		if ($sqlUser -and $sqlPwd)
		{
			$connStr = "Server=$server;Database=$dbName;User ID=$sqlUser;Password=$sqlPwd;TrustServerCertificate=True;"
		}
		else
		{
			$connStr = "Server=$server;Database=$dbName;Integrated Security=SSPI;TrustServerCertificate=True;"
		}
		$script:FunctionResults['ConnectionString'] = $connStr
	}
	
	# ------------------------------------------------------------------------------------------------
	# Active NIC prefix (fallback) - logged ONCE
	# ------------------------------------------------------------------------------------------------
	$ActiveNicPrefix = $null
	try
	{
		$serverIPv4s = @()
		try
		{
			[System.Net.Dns]::GetHostAddresses($server) | ForEach-Object {
				if ($_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $serverIPv4s += $_.ToString() }
			}
		}
		catch { }
		
		$allNics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
		Where-Object { $_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up }
		
		function _SameSubnet($aStr, $bStr, $maskStr)
		{
			try
			{
				$a = [System.Net.IPAddress]::Parse($aStr).GetAddressBytes()
				$b = [System.Net.IPAddress]::Parse($bStr).GetAddressBytes()
				$m = [System.Net.IPAddress]::Parse($maskStr).GetAddressBytes()
				for ($i = 0; $i -lt 4; $i++) { if (($a[$i] -band $m[$i]) -ne ($b[$i] -band $m[$i])) { return $false } }
				return $true
			}
			catch { return $false }
		}
		
		$pickedIP = $null
		foreach ($ni in $allNics)
		{
			$ipProps = $ni.GetIPProperties()
			foreach ($ua in $ipProps.UnicastAddresses)
			{
				if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
				if (-not $ua.IPv4Mask) { continue }
				foreach ($srv in $serverIPv4s)
				{
					if (_SameSubnet $ua.Address.ToString() $srv $ua.IPv4Mask.ToString()) { $pickedIP = $ua.Address.ToString(); break }
				}
				if ($pickedIP) { break }
			}
			if ($pickedIP) { break }
		}
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				$hasGw = $false
				foreach ($gw in $ipProps.GatewayAddresses)
				{
					if ($gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
				}
				if (-not $hasGw) { continue }
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$pickedIP = $ua.Address.ToString(); break
				}
				if ($pickedIP) { break }
			}
		}
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$ip = $ua.Address.ToString()
					if ($ip.StartsWith('10.') -or $ip.StartsWith('192.168.')) { $pickedIP = $ip; break }
					if ($ip.StartsWith('172.'))
					{
						$parts = $ip.Split('.')
						if ($parts.Length -ge 2)
						{
							$sec = 0
							if ([int]::TryParse($parts[1], [ref]$sec)) { if ($sec -ge 16 -and $sec -le 31) { $pickedIP = $ip; break } }
						}
					}
				}
				if ($pickedIP) { break }
			}
		}
		if ($pickedIP)
		{
			$p = $pickedIP.Split('.')
			if ($p.Length -ge 3) { $ActiveNicPrefix = "$($p[0]).$($p[1]).$($p[2])" }
		}
	}
	catch { }
	
	if ($ActiveNicPrefix) { Write_Log "Active NIC /24 prefix: $ActiveNicPrefix" "cyan" }
	else { Write_Log "Active NIC /24 prefix not detected (fallback may be unavailable)." "yellow" }
	
	# ------------------------------------------------------------------------------------------------
	# FIRST active scale full IP (robust) - logged ONCE
	# ------------------------------------------------------------------------------------------------
	$firstScaleIP = $null
	if ($server -and $dbName -and $connStr)
	{
		$InvokeSqlCmd = $null
		if ($sqlModule -and $sqlModule -ne 'None') { try { Import-Module $sqlModule -ErrorAction Stop; $InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $sqlModule -ErrorAction Stop }
			catch { $InvokeSqlCmd = $null } }
		$supportsConnStr = $false
		if ($InvokeSqlCmd) { $supportsConnStr = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString' }
		
		function _RunQuery([string]$query)
		{
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnStr) { return Invoke-Sqlcmd -ConnectionString $connStr -Query $query -ErrorAction Stop }
					else { return Invoke-Sqlcmd -ServerInstance $server -Database $dbName -Query $query -ErrorAction Stop }
				}
				else
				{
					$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb.ConnectionString = $connStr
					$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
					$conn.Open()
					$cmd = $conn.CreateCommand(); $cmd.CommandText = $query; $cmd.CommandTimeout = 30
					$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
					$dt = New-Object System.Data.DataTable
					[void]$da.Fill($dt)
					$conn.Close()
					return $dt
				}
			}
			catch { return $null }
		}
		
		$q = @"
SELECT IPNetwork, IPDevice, ScaleCode, Active
FROM dbo.TBS_SCL_ver520
ORDER BY
    CASE WHEN TRY_CAST(ScaleCode AS INT) IS NULL THEN 1 ELSE 0 END,
    TRY_CAST(ScaleCode AS INT),
    ScaleCode;
"@
		$rows = _RunQuery $q
		
		if ($rows -and (($rows -is [System.Data.DataTable] -and $rows.Rows.Count -gt 0) -or ($rows -isnot [System.Data.DataTable] -and $rows.Count -gt 0)))
		{
			Write_Log "Loaded rows from TBS_SCL_ver520 for first-scale detection." "cyan"
			
			$IsActive = {
				param ($v)
				if ($null -eq $v) { return $true }
				$s = ([string]$v).Trim().ToUpper()
				if ($s -eq '') { return $true }
				if ($s -eq 'Y' -or $s -eq '1' -or $s -eq 'TRUE') { return $true }
				return $false
			}
			$ParsePrefix = {
				param ($ipn)
				if ([string]::IsNullOrWhiteSpace($ipn)) { return $null }
				$m = [regex]::Match($ipn.Trim(), '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
				if ($m.Success) { return "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)" }
				return $null
			}
			$ParseOctet = {
				param ($d)
				if ($null -eq $d) { return $null }
				$tmp = 0
				if ([int]::TryParse(([string]$d).Trim(), [ref]$tmp))
				{
					if ($tmp -ge 0 -and $tmp -le 255) { return $tmp }
				}
				return $null
			}
			
			$enum = @()
			if ($rows -is [System.Data.DataTable]) { $enum = $rows.Rows }
			else { $enum = $rows }
			
			$pfxCounts = @{ }
			foreach ($r in $enum)
			{
				$pref = & $ParsePrefix ($r.IPNetwork)
				if ($pref)
				{
					if (-not $pfxCounts.ContainsKey($pref)) { $pfxCounts[$pref] = 0 }
					$pfxCounts[$pref]++
				}
			}
			$MostCommonPrefix = $null
			if ($pfxCounts.Count -gt 0)
			{
				$MostCommonPrefix = ($pfxCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
				Write_Log "Most common IPNetwork prefix in table: $MostCommonPrefix" "cyan"
			}
			
			$candidates = @()
			foreach ($r in $enum) { if (& $IsActive ($r.Active)) { $candidates += $r } }
			
			foreach ($r in $candidates)
			{
				$p = & $ParsePrefix ($r.IPNetwork)
				if (-not $p -and $r.IPNetwork)
				{
					$m4 = [regex]::Match(([string]$r.IPNetwork).Trim(), '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\s*$')
					if ($m4.Success)
					{
						$p = "$($m4.Groups[1].Value).$($m4.Groups[2].Value).$($m4.Groups[3].Value)"
						if (-not $r.IPDevice) { $r.IPDevice = $m4.Groups[4].Value }
					}
				}
				$d = & $ParseOctet ($r.IPDevice)
				if ($p -and ($null -ne $d)) { $firstScaleIP = "$p.$d"; break }
			}
			
			if (-not $firstScaleIP -and $MostCommonPrefix)
			{
				foreach ($r in $candidates)
				{
					$d = & $ParseOctet ($r.IPDevice)
					if ($null -ne $d) { $firstScaleIP = "$MostCommonPrefix.$d"; break }
				}
			}
			
			if (-not $firstScaleIP -and $ActiveNicPrefix)
			{
				foreach ($r in $candidates)
				{
					$d = & $ParseOctet ($r.IPDevice)
					if ($null -ne $d) { $firstScaleIP = "$ActiveNicPrefix.$d"; break }
				}
			}
			
			if ($firstScaleIP) { Write_Log "First scale IP resolved: $firstScaleIP" "cyan" }
			else { Write_Log "First scale IP NOT resolved (need prefix + device)." "yellow" }
		}
		else
		{
			Write_Log "Could not read rows from TBS_SCL_ver520 (DB issue or empty table)." "yellow"
		}
	}
	else
	{
		Write_Log "DB connection info incomplete; first-scale detection skipped." "yellow"
	}
	
	# ------------------------------------------------------------------------------------------------
	# Locate target configs
	# ------------------------------------------------------------------------------------------------
	$targets = @()
	if ($ConfigPath)
	{
		if (Test-Path -LiteralPath $ConfigPath) { $targets += (Get-Item -LiteralPath $ConfigPath) }
		else { Write_Log "Config path not found: $ConfigPath" "red"; return }
	}
	else
	{
		foreach ($root in $ScaleCommRoots)
		{
			if (Test-Path -LiteralPath $root)
			{
				$found = Get-ChildItem -LiteralPath $root -File -Filter '*.exe.config' -ErrorAction SilentlyContinue
				if ($found) { $targets += $found }
			}
		}
	}
	$targets = $targets | Sort-Object FullName -Unique
	if (-not $targets -or $targets.Count -eq 0) { Write_Log "No *.exe.config found under ScaleCommApp roots." "red"; return }
	if (-not $AllMatches -and $targets.Count -gt 1)
	{
		Write_Log ("Multiple configs found; using first. Use -AllMatches to update all." + "`r`n - " + (($targets | ForEach-Object FullName) -join "`r`n - ")) "yellow"
		$targets = @($targets[0])
	}
	
	# ------------------------------------------------------------------------------------------------
	# ONE-TIME "Desired Settings" + "Existing (first file)" snapshot (NO TERNARY)
	# ------------------------------------------------------------------------------------------------
	$previewPath = $targets[0].FullName
	$previewXml = $null
	try { [xml]$previewXml = Get-Content -LiteralPath $previewPath -Raw -ErrorAction Stop }
	catch { }
	function _GetVal([xml]$xmlDoc, $key)
	{
		if (-not $xmlDoc) { return $null }
		$cfg = $xmlDoc.configuration; if (-not $cfg) { $cfg = $xmlDoc.DocumentElement }
		if (-not $cfg) { return $null }
		$as = $cfg.appSettings; if (-not $as) { return $null }
		$n = $as.SelectSingleNode("add[@key='$key']"); if ($n) { return [string]$n.GetAttribute('value') }
		return $null
	}
	$curStorePreview = _GetVal $previewXml 'StoreName'
	$curFirstIPPreview = _GetVal $previewXml 'FistScaleIP'
	$curSqlInstPreview = _GetVal $previewXml 'SQL_InstanceName'
	$curDbNamePreview = _GetVal $previewXml 'SQL_databaseName'
	$curBatchPreview = _GetVal $previewXml 'Recs_BatchSendFull'
	
	# -- Build log strings without '??' (PowerShell 5.1 safe)
	$dStore = '<unchanged/skip>'; if ($ExpectedStoreName) { $dStore = $ExpectedStoreName }
	$dIP = '<skip>'; if ($firstScaleIP) { $dIP = $firstScaleIP }
	$dInst = '<unchanged/skip>'; if ($ExpectedSqlInstance) { $dInst = $ExpectedSqlInstance }
	$dDb = '<unchanged/skip>'; if ($ExpectedDatabase) { $dDb = $ExpectedDatabase }
	
	Write_Log "Desired Settings:" "blue"
	Write_Log ("  StoreName          : {0}" -f $dStore) "green"
	Write_Log ("  FistScaleIP        : {0}" -f $dIP) "green"
	Write_Log ("  SQL_InstanceName   : {0}" -f $dInst) "green"
	Write_Log ("  SQL_databaseName   : {0}" -f $dDb) "green"
	Write_Log ("  Recs_BatchSendFull : 10000") "green"
	
	$leafPreview = Split-Path -Leaf $previewPath
	Write_Log ("Existing (first file: {0}):" -f $leafPreview) "blue"
	$eStore = $curStorePreview; if (-not $eStore) { $eStore = '<missing>' }
	$eIP = $curFirstIPPreview; if (-not $eIP) { $eIP = '<missing>' }
	$eInst = $curSqlInstPreview; if (-not $eInst) { $eInst = '<missing>' }
	$eDb = $curDbNamePreview; if (-not $eDb) { $eDb = '<missing>' }
	$eBatch = $curBatchPreview; if (-not $eBatch) { $eBatch = '<missing>' }
	Write_Log ("  StoreName          : {0}" -f $eStore) "yellow"
	Write_Log ("  FistScaleIP        : {0}" -f $eIP) "yellow"
	Write_Log ("  SQL_InstanceName   : {0}" -f $eInst) "yellow"
	Write_Log ("  SQL_databaseName   : {0}" -f $eDb) "yellow"
	Write_Log ("  Recs_BatchSendFull : {0}" -f $eBatch) "yellow"
	
	# ------------------------------------------------------------------------------------------------
	# Apply to all targets (quiet per-file; track counts + saved list)
	# ------------------------------------------------------------------------------------------------
	$autoFix = (-not $NoAutoFix.IsPresent)
	[int]$audited = 0
	[int]$updated = 0
	[int]$skipped = 0
	$savedFiles = New-Object System.Collections.Generic.List[string]
	
	foreach ($t in $targets)
	{
		$audited++
		$path = $t.FullName
		
		try { [xml]$xml = Get-Content -LiteralPath $path -Raw -ErrorAction Stop }
		catch { $skipped++; continue }
		$cfg = $xml.configuration; if (-not $cfg) { $cfg = $xml.DocumentElement }
		if (-not $cfg) { $skipped++; continue }
		$as = $cfg.appSettings
		if (-not $as) { $skipped++; continue }
		
		$getVal = { param ($key) $n = $as.SelectSingleNode("add[@key='$key']"); if ($n) { return [string]$n.GetAttribute('value') }
			else { return $null } }
		$setVal = {
			param ($key,
				$val)
			$n = $as.SelectSingleNode("add[@key='$key']"); if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', $key); $null = $as.AppendChild($n) }
			$null = $n.SetAttribute('value', [string]$val)
		}
		
		$changedHere = $false
		$curStore = & $getVal 'StoreName'
		$curFirstIP = & $getVal 'FistScaleIP'
		$curInst = & $getVal 'SQL_InstanceName'
		$curDb = & $getVal 'SQL_databaseName'
		$curBatch = & $getVal 'Recs_BatchSendFull'
		
		if ($ExpectedStoreName -and $curStore -ne $ExpectedStoreName) { if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set StoreName')) { & $setVal 'StoreName' $ExpectedStoreName; $changedHere = $true } }
		if ($firstScaleIP -and $curFirstIP -ne $firstScaleIP) { if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set FistScaleIP')) { & $setVal 'FistScaleIP' $firstScaleIP; $changedHere = $true } }
		if ($ExpectedSqlInstance -and $curInst -ne $ExpectedSqlInstance) { if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set SQL_InstanceName')) { & $setVal 'SQL_InstanceName' $ExpectedSqlInstance; $changedHere = $true } }
		if ($ExpectedDatabase -and $curDb -ne $ExpectedDatabase) { if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set SQL_databaseName')) { & $setVal 'SQL_databaseName' $ExpectedDatabase; $changedHere = $true } }
		if ($curBatch -ne '10000') { if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set Recs_BatchSendFull=10000')) { & $setVal 'Recs_BatchSendFull' '10000'; $changedHere = $true } }
		
		if ($changedHere -and $autoFix)
		{
			try
			{
				$item = Get-Item -LiteralPath $path -ErrorAction Stop
				if ($item.Attributes -band [IO.FileAttributes]::ReadOnly)
				{
					if ($Force) { $item.Attributes = ($item.Attributes -bxor [IO.FileAttributes]::ReadOnly) }
					else { $skipped++; continue }
				}
			}
			catch { $skipped++; continue }
			
			try
			{
				$dir = Split-Path -LiteralPath $path -Parent
				$leaf = Split-Path -LiteralPath $path -Leaf
				$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
				$bak = Join-Path $dir "$leaf.$stamp.bak"
				Copy-Item -LiteralPath $path -Destination $bak -Force
			}
			catch { }
			
			try
			{
				$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
				$sw = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
				$xml.Save($sw); $sw.Close()
				$updated++; $savedFiles.Add((Split-Path -Leaf $path)) | Out-Null
			}
			catch { $skipped++ }
		}
	}
	
	# ------------------------------------------------------------------------------------------------
	# Final summary (NO TERNARY)
	# ------------------------------------------------------------------------------------------------
	$sumIP = 'n/a'
	if ($firstScaleIP) { $sumIP = $firstScaleIP }
	
	Write_Log ("`r`nTroubleshoot_ScaleCommApp: Audited={0}, Updated={1}, Skipped={2}, FirstScaleIP={3}" -f $audited, $updated, $skipped, $sumIP) "blue"
	if ($updated -gt 0)
	{
		Write_Log ("Saved files: {0}" -f ($savedFiles -join ', ')) "green"
	}
	
	Write_Log "`r`n==================== Troubleshoot_ScaleCommApp Completed ====================" "blue"
}

# ===================================================================================================
#                                 FUNCTION: Repair_BMS (All-in-one, PS5.1-safe)
# ---------------------------------------------------------------------------------------------------
# Description:
#   Forces "BMS" to stop with escalating actions, then repairs it:
#     - Disable service (prevent recovery auto-restart while working)
#     - Stop-Service -Force  ->  sc stop  ->  net stop /y  ->  WMI StopService()
#     - Kill by Service PID (WMI/sc queryex), by service association (taskkill /FI "SERVICES eq BMS"),
#       and by known process names (ScaleManagementApp.exe, BMSSrv.exe, BMS.exe)
#     - Clean toBizerba, delete service, re-register (BMSSrv.exe -reg), start service
#
# Notes:
#   - Uses Write_Log for colored output.
#   - Includes waits and verification loops to avoid timing issues.
#   - PS5.1-compatible: no nested/local helper functions; everything is inline.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - BMSSrvPath (Optional): Full path to BMSSrv.exe (default: C:\Bizerba\RetailConnect\BMS\BMSSrv.exe)
#   - UpdateSpecialsPath (Optional): Full path to ScaleManagementAppUpdateSpecials.exe
#                                    (default: C:\ScaleCommApp\ScaleManagementAppUpdateSpecials.exe)
# ---------------------------------------------------------------------------------------------------

function Repair_BMS
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $false)]
		[string]$BMSSrvPath = "$env:SystemDrive\Bizerba\RetailConnect\BMS\BMSSrv.exe",
		[Parameter(Mandatory = $false)]
		[string]$UpdateSpecialsPath = "$env:SystemDrive\ScaleCommApp\ScaleManagementAppUpdateSpecials.exe"
	)
	
	Write_Log "`r`n==================== Starting Repair_BMS Function ====================`r`n" "blue"
	
	# ---------------------------
	# 0) Elevation check
	# ---------------------------
	$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
	IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	if (-not $isAdmin)
	{
		Write_Log "Insufficient permissions. Please run this script as an Administrator." "red"
		return
	}
	
	# ---------------------------
	# 1) Validate BMSSrv.exe path
	# ---------------------------
	if (-not (Test-Path -LiteralPath $BMSSrvPath))
	{
		Write_Log "BMSSrv.exe not found at path: $BMSSrvPath" "red"
		return
	}
	
	# ---------------------------
	# Constants and state flags
	# ---------------------------
	$serviceName = "BMS"
	$toBizerbaPath = Join-Path -Path "$env:SystemDrive\Bizerba\RetailConnect\BMS" -ChildPath "toBizerba"
	$processNames = @("ScaleManagementApp", "BMSSrv", "BMS") # UI/client first, then service binaries
	$needHardKill = $false
	
	# Helper note:
	#   Since the user requested "all one function", we inline the logic for:
	#   - Safe service retrieval
	#   - PID discovery
	#   - Wait loops for state changes
	
	# -------------------------------------------
	# Inline: Safe get-service (returns $null if missing)
	# -------------------------------------------
	$svcObj = $null
	try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcObj = $null }
	
	# ---------------------------------------------------------
	# 2) Disable service (block auto-restart/recovery while we work)
	# ---------------------------------------------------------
	if ($svcObj)
	{
		try
		{
			if ($svcObj.StartType -ne 'Disabled')
			{
				Write_Log "Disabling '$serviceName' temporarily to prevent auto-restart during repair..." "blue"
				Set-Service -Name $serviceName -StartupType Disabled
			}
		}
		catch
		{
			Write_Log "Failed to set '$serviceName' startup type to Disabled: $_" "yellow"
		}
	}
	else
	{
		Write_Log "'$serviceName' service not found before stop attempts." "yellow"
	}
	
	# ---------------------------------------------------------
	# 3) Stop service via escalating methods (each with inline waits)
	# ---------------------------------------------------------
	
	# 3.1 Stop-Service -Force
	$svcObj = $null; try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcObj = $null }
	if ($svcObj -and $svcObj.Status -ne 'Stopped')
	{
		Write_Log "Stopping '$serviceName' via Stop-Service -Force..." "blue"
		try { Stop-Service -Name $serviceName -Force -ErrorAction Stop }
		catch { Write_Log "Stop-Service failed: $_" "yellow" }
		
		# Inline wait up to 10s for Stopped
		$elapsed = 0; $desired = 'Stopped'; $timeoutSec = 10; $stopped = $false
		while ($elapsed -lt $timeoutSec)
		{
			$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
			catch { $svcCheck = $null }
			if (-not $svcCheck) { $stopped = $true; break } # Missing => acceptable as stopped
			if ($svcCheck.Status -eq $desired) { $stopped = $true; break }
			Start-Sleep -Seconds 1; $elapsed++
		}
		if (-not $stopped) { Write_Log "Service did not stop after Stop-Service. Escalating..." "yellow" }
	}
	
	# 3.2 sc.exe stop
	$svcObj = $null; try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcObj = $null }
	if ($svcObj -and $svcObj.Status -ne 'Stopped')
	{
		Write_Log "Stopping '$serviceName' via sc.exe stop..." "blue"
		try { sc.exe stop $serviceName | Out-Null }
		catch { Write_Log "sc stop failed: $_" "yellow" }
		
		# Inline wait up to 10s
		$elapsed = 0; $desired = 'Stopped'; $timeoutSec = 10; $stopped = $false
		while ($elapsed -lt $timeoutSec)
		{
			$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
			catch { $svcCheck = $null }
			if (-not $svcCheck) { $stopped = $true; break }
			if ($svcCheck.Status -eq $desired) { $stopped = $true; break }
			Start-Sleep -Seconds 1; $elapsed++
		}
		if (-not $stopped) { Write_Log "Service did not stop after sc stop. Escalating..." "yellow" }
	}
	
	# 3.3 net stop /y
	$svcObj = $null; try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcObj = $null }
	if ($svcObj -and $svcObj.Status -ne 'Stopped')
	{
		Write_Log "Stopping '$serviceName' via 'net stop /y' (with dependencies)..." "blue"
		try { cmd /c "net stop $serviceName /y" | Out-Null }
		catch { Write_Log "net stop failed: $_" "yellow" }
		
		# Inline wait up to 10s
		$elapsed = 0; $desired = 'Stopped'; $timeoutSec = 10; $stopped = $false
		while ($elapsed -lt $timeoutSec)
		{
			$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
			catch { $svcCheck = $null }
			if (-not $svcCheck) { $stopped = $true; break }
			if ($svcCheck.Status -eq $desired) { $stopped = $true; break }
			Start-Sleep -Seconds 1; $elapsed++
		}
		if (-not $stopped) { Write_Log "Service did not stop after 'net stop'. Escalating..." "yellow" }
	}
	
	# 3.4 WMI/CIM StopService()
	$svcObj = $null; try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcObj = $null }
	if ($svcObj -and $svcObj.Status -ne 'Stopped')
	{
		Write_Log "Stopping '$serviceName' via WMI StopService()..." "blue"
		try
		{
			$svcWmi = $null
			try { $svcWmi = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $serviceName) -ErrorAction Stop }
			catch { $svcWmi = $null }
			if ($svcWmi)
			{
				$ret = $null
				try { $ret = Invoke-CimMethod -InputObject $svcWmi -MethodName StopService -ErrorAction SilentlyContinue }
				catch { $ret = $null }
				if ($ret) { Write_Log ("WMI StopService() returned: {0}" -f ($ret.ReturnValue)) "yellow" }
			}
		}
		catch
		{
			Write_Log "WMI StopService failed: $_" "yellow"
		}
		
		# Inline wait up to 8s; if still not stopped, set hard-kill flag
		$elapsed = 0; $desired = 'Stopped'; $timeoutSec = 8; $stopped = $false
		while ($elapsed -lt $timeoutSec)
		{
			$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
			catch { $svcCheck = $null }
			if (-not $svcCheck) { $stopped = $true; break }
			if ($svcCheck.Status -eq $desired) { $stopped = $true; break }
			Start-Sleep -Seconds 1; $elapsed++
		}
		if (-not $stopped) { Write_Log "Service still not stopped after WMI StopService. Preparing hard kill..." "yellow"; $needHardKill = $true }
	}
	
	# ---------------------------------------------------------
	# 4) Hard kill: by PID, by service association, and by names
	# ---------------------------------------------------------
	$svcObj = $null; try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcObj = $null }
	if ($needHardKill -or ($svcObj -and $svcObj.Status -ne 'Stopped'))
	{
		# 4.1 Kill by PID (WMI first, then sc queryex)
		$pid = $null
		
		# Try WMI/CIM for PID
		try
		{
			$svcWmi = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $serviceName) -ErrorAction Stop
			if ($svcWmi -and $svcWmi.ProcessId -gt 0) { $pid = [int]$svcWmi.ProcessId }
		}
		catch { }
		
		# Fallback: sc queryex
		if (-not $pid)
		{
			try
			{
				$scOut = sc.exe queryex $serviceName 2>&1
				if ($scOut)
				{
					foreach ($line in $scOut)
					{
						if ($line -match 'PID\s*:\s*(\d+)') { $pid = [int]$Matches[1]; break }
					}
				}
			}
			catch { }
		}
		
		# Kill the PID if still running
		if ($pid -and (Get-Process -Id $pid -ErrorAction SilentlyContinue))
		{
			try
			{
				Write_Log ("Killing service PID {0} (from queryex/WMI)..." -f $pid) "blue"
				Stop-Process -Id $pid -Force -ErrorAction Stop
			}
			catch
			{
				Write_Log "Killing PID $pid failed: $_" "yellow"
			}
		}
		
		# 4.2 Kill any process hosting the service (SERVICES filter)
		try
		{
			Write_Log "Killing any process hosting '$serviceName' (taskkill /F /FI 'SERVICES eq BMS')..." "blue"
			cmd /c 'taskkill /F /FI "SERVICES eq BMS"' | Out-Null
		}
		catch
		{
			Write_Log "taskkill by SERVICES filter failed (may be fine): $_" "yellow"
		}
		
		# 4.3 Kill well-known binaries by name
		Write_Log "Force-terminating related processes by name: $($processNames -join ', ') ..." "blue"
		foreach ($p in $processNames)
		{
			try
			{
				$procs = Get-Process -Name $p -ErrorAction SilentlyContinue
				if ($procs)
				{
					$procs | ForEach-Object { Write_Log ("Killing process {0} (PID {1})" -f $_.ProcessName, $_.Id) "yellow" }
					$procs | Stop-Process -Force -ErrorAction Stop
				}
			}
			catch
			{
				Write_Log "Could not kill process '$p': $_" "yellow"
			}
		}
		
		Start-Sleep -Seconds 2
	}
	
	# ---------------------------------------------------------
	# Final verification: ensure service is stopped (or gone)
	# ---------------------------------------------------------
	$elapsed = 0; $desired = 'Stopped'; $timeoutSec = 5; $stopped = $false
	while ($elapsed -lt $timeoutSec)
	{
		$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
		catch { $svcCheck = $null }
		if (-not $svcCheck) { $stopped = $true; break }
		if ($svcCheck.Status -eq $desired) { $stopped = $true; break }
		Start-Sleep -Seconds 1; $elapsed++
	}
	
	if (-not $stopped)
	{
		$svcObj = $null; try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
		catch { $svcObj = $null }
		if ($svcObj -and $svcObj.Status -ne 'Stopped')
		{
			Write_Log "WARNING: '$serviceName' still not stopped after all methods; proceeding with delete anyway." "yellow"
		}
		else
		{
			Write_Log "'$serviceName' appears stopped." "green"
		}
	}
	else
	{
		Write_Log "'$serviceName' confirmed stopped." "green"
	}
	
	# ---------------------------------------------------------
	# 5) Clean the toBizerba staging folder (if present)
	# ---------------------------------------------------------
	if (Test-Path -LiteralPath $toBizerbaPath)
	{
		try
		{
			Write_Log "Cleaning folder: $toBizerbaPath" "blue"
			$items = Get-ChildItem -LiteralPath $toBizerbaPath -Force -ErrorAction SilentlyContinue
			foreach ($it in $items)
			{
				try { Remove-Item -LiteralPath $it.FullName -Recurse -Force -ErrorAction Stop }
				catch { Write_Log "Failed to remove '$($it.FullName)': $_" "yellow" }
			}
			Write_Log "toBizerba folder cleaned." "green"
		}
		catch
		{
			Write_Log "Failed to clean toBizerba folder: $_" "yellow"
		}
	}
	else
	{
		Write_Log "toBizerba folder not found at: $toBizerbaPath (skipping cleanup)" "yellow"
	}
	
	# ---------------------------------------------------------
	# 6) Delete the BMS service (best-effort)
	# ---------------------------------------------------------
	$svcObj = $null; try { $svcObj = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcObj = $null }
	if ($svcObj)
	{
		Write_Log "Deleting '$serviceName' service via sc.exe delete..." "blue"
		try
		{
			sc.exe delete $serviceName | Out-Null
			
			# Wait up to 15s for SCM to reflect deletion
			$wait = 0
			while ($wait -lt 15)
			{
				Start-Sleep -Seconds 1; $wait++
				$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
				catch { $svcCheck = $null }
				if (-not $svcCheck) { break }
			}
			
			$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
			catch { $svcCheck = $null }
			if (-not $svcCheck)
			{
				Write_Log "'$serviceName' service deleted." "green"
			}
			else
			{
				Write_Log "Warning: '$serviceName' still appears in SCM. Continuing..." "yellow"
			}
		}
		catch
		{
			Write_Log "Failed to delete '$serviceName' service: $_" "yellow"
		}
	}
	else
	{
		Write_Log "'$serviceName' service not present; skipping delete." "yellow"
	}
	
	# ---------------------------------------------------------
	# 7) Register the service again (BMSSrv.exe -reg)
	# ---------------------------------------------------------
	Write_Log "Registering service with: `"$BMSSrvPath -reg`"" "blue"
	try
	{
		$proc = Start-Process -FilePath $BMSSrvPath -ArgumentList "-reg" -NoNewWindow -Wait -PassThru
		if ($proc.ExitCode -ne 0)
		{
			Write_Log ("BMSSrv.exe registration failed with exit code {0}." -f $proc.ExitCode) "red"
			return
		}
		Write_Log "BMSSrv.exe registered successfully." "green"
	}
	catch
	{
		Write_Log "An error occurred while registering BMSSrv.exe: $_" "red"
		return
	}
	
	# Verify service appears (wait up to 15s)
	$appeared = $false
	for ($i = 0; $i -lt 15; $i++)
	{
		Start-Sleep -Seconds 1
		$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
		catch { $svcCheck = $null }
		if ($svcCheck) { $appeared = $true; break }
	}
	if (-not $appeared)
	{
		Write_Log "BMS service did not appear after registration. Aborting." "red"
		return
	}
	
	# Restore startup to Automatic (we disabled earlier)
	try { Set-Service -Name $serviceName -StartupType Automatic }
	catch { Write_Log "Failed to set '$serviceName' startup type to Automatic: $_" "yellow" }
	
	# ---------------------------------------------------------
	# 8) Start the BMS service (robust)
	# ---------------------------------------------------------
	Write_Log "Starting '$serviceName' service..." "blue"
	$started = $false
	try
	{
		Start-Service -Name $serviceName -ErrorAction Stop
		$started = $true
	}
	catch
	{
		Write_Log "Start-Service failed: $_" "yellow"
		try
		{
			sc.exe start $serviceName | Out-Null
			Start-Sleep -Seconds 2
			$svcCheck = $null; try { $svcCheck = Get-Service -Name $serviceName -ErrorAction Stop }
			catch { $svcCheck = $null }
			if ($svcCheck -and $svcCheck.Status -eq 'Running') { $started = $true }
		}
		catch
		{
			Write_Log "sc.exe start also failed: $_" "red"
		}
	}
	
	if ($started)
	{
		Write_Log "'$serviceName' service started successfully." "green"
	}
	else
	{
		Write_Log "Failed to start '$serviceName' service." "red"
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
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected or selection cancelled." "yellow"
		Write_Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	# Normalize to 3-digit lane numbers (always as strings)
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$lanesToProcess = $selection.Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($LaneNumToMachineName.ContainsKey($_)) { $_ }
		else { $_ }
	}
	$lanesToProcess = $lanesToProcess | Where-Object { $LaneNumToMachineName.ContainsKey($_) }
	if (-not $lanesToProcess -or $lanesToProcess.Count -eq 0)
	{
		Write_Log "No valid lanes selected for processing." "yellow"
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
	# STEP 1: Use Show_Node_Selection_Form to select registers/lanes
	# --------------------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No registers selected or selection cancelled." "yellow"
		Write_Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
		return
	}
	
	# Normalize to 3-digit lane/register numbers (always as strings)
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$registersToProcess = $selection.Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($LaneNumToMachineName.ContainsKey($_)) { $_ }
		else { $_ }
	}
	$registersToProcess = $registersToProcess | Where-Object { $LaneNumToMachineName.ContainsKey($_) }
	if (-not $registersToProcess -or $registersToProcess.Count -eq 0)
	{
		Write_Log "No valid registers selected for processing." "yellow"
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
		[Parameter(Mandatory = $false)]
		$Selection,
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string[]]$NodeTypes = @("Lane", "Scale", "Backoffice")
	)
	Write_Log "`r`n==================== Starting Reboot_Nodes Function ====================`r`n" "blue"
	
	# Load global node data
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
	
	# --- Use the shared selection dialog ---
	if ($Selection)
	{
		$selection = $Selection
	}
	else
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes $NodeTypes
	}
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
#                      FUNCTION: Enable_SQL_Protocols_On_Selected_Lanes  (Runspaces, No Nested Funcs)
# ---------------------------------------------------------------------------------------------------
# Description:
#   Enables TCP/IP, Named Pipes, and Shared Memory protocols for all SQL Server instances
#   on the selected remote lane machines. Sets static TCP port (default: 1433), disables dynamic ports,
#   restarts the SQL service when needed, and optionally creates & live-tests a Linked Server
#   to determine working transport. Works in parallel using a runspace pool (NOT background jobs).
#
# Parameters:
#   - [string]$StoreNumber            3-digit store, e.g. "123"
#   - [string]$tcpPort                static TCP port (default "1433")
#   - [switch]$CreateLinkedServers    on by default; pass -CreateLinkedServers:$false to skip
#
# Requirements:
#   - RemoteRegistry reachable on lanes (function will start it demand-mode)
#   - $script:FunctionResults['LaneNumToMachineName'] populated (lane -> machine)
#   - $script:FunctionResults['DBSERVER'] local SQL instance for linked-server management
#   - Write_Log, Show_Node_Selection_Form, Send_Restart_All_Programs provided by your script
#
# Implementation notes (important):
#   - NO nested functions anywhere. The runspace payload is a single script block with param().
#   - Concurrency uses RunspacePool (MTA). Degree of parallelism auto-throttled.
#   - Logs are buffered per-lane inside the runspace and emitted when that lane completes.
#     (Simple + robust; avoids cross-thread UI/update issues.)
# ===================================================================================================

function Enable_SQL_Protocols_On_Selected_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$tcpPort = "1433",
		[Parameter(Mandatory = $false)]
		[switch]$CreateLinkedServers
	)
	
	# ---- Default CreateLinkedServers to $true if omitted -------------------------------------------
	if (-not $PSBoundParameters.ContainsKey('CreateLinkedServers')) { $CreateLinkedServers = $true }
	
	Write_Log "`r`n==================== Starting Enable_SQL_Protocols_On_Selected_Lanes ====================`r`n" "blue"
	
	# ---- Resolve lane -> machine -------------------------------------------------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "No lanes available. Please retrieve nodes first." "red"
		Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Completed ====================" "blue"
		return
	}
	
	# ---- Local SQL instance for linked-server operations -------------------------------------------
	$localInstance = $script:FunctionResults['DBSERVER']
	if ([string]::IsNullOrWhiteSpace($localInstance) -or $localInstance -eq 'N/A') { $localInstance = 'localhost' }
	
	# ---- Lane selection UI -------------------------------------------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
		Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Completed ====================" "blue"
		return
	}
	
	# ---- Normalize lane list (ensure 3-digit strings) ----------------------------------------------
	$lanes =
	$selection.Lanes |
	ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		else { $_ }
	} |
	ForEach-Object { $_.ToString().PadLeft(3, '0') } |
	Sort-Object -Unique
	
	Write_Log "Selected lanes: $($lanes -join ', ')" "green"
	
	# ---- Prepare caches used elsewhere in your script ----------------------------------------------
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
	if (-not $script:ProtocolResults) { $script:ProtocolResults = @() }
	
	# ---- Create a runspace pool (MTA). Throttle to a reasonable degree of parallelism --------------
	$maxThreads = [Math]::Min($lanes.Count, [Math]::Max(2, [Environment]::ProcessorCount))
	$pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
	$pool.ApartmentState = 'MTA'
	$pool.Open()
	
	# ---- Script that each runspace will execute (NO nested functions; just inline code) ------------
	$laneWork = @'
param($machine,$lane,$StoreNumber,$tcpPort,$createLinks,$localInstance)

# Per-lane output buffer; we send it back as a single PSCustomObject
$out = @()
$laneNeedsRestart = $false
$protocol = "File"           # default; changed after probe/DMV
$creationTransport = $null   # remember which transport we CREATE with (TCP/NP)
$createdAndTested = $false   # becomes true if linked server SELECT works

try {
    # --- Opening banner ---------------------------------------------------------------------------
    $out += [pscustomobject]@{ Text = "`r`n--- Processing Machine: $machine (Store $StoreNumber, Lane $lane) ---"; Color = "blue" }
    $out += [pscustomobject]@{ Text = "Ensuring RemoteRegistry is running on $machine..."; Color = "gray" }

    # Start RemoteRegistry on-demand; ignore console noise
    sc.exe "\\$machine" config RemoteRegistry start= demand | Out-Null
    sc.exe "\\$machine" start RemoteRegistry | Out-Null

    # --- Open remote HKLM -------------------------------------------------------------------------
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $machine)

    # --- Enumerate SQL instances from both 64/32-bit locations -----------------------------------
    $instances = @{}
    foreach ($root in @(
        "SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL",
        "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL"
    )) {
        $k = $reg.OpenSubKey($root)
        if ($k) {
            foreach ($n in $k.GetValueNames()) {
                $id = $k.GetValue($n)
                if ($id -and -not $instances.ContainsKey($n)) { $instances[$n] = $id }
            }
            $k.Close()
        }
    }

    if ($instances.Count -eq 0) {
        $out += [pscustomobject]@{ Text = "No SQL instances found on $machine."; Color = "yellow" }
        $reg.Close()
    }
    else {
        foreach ($instanceName in $instances.Keys) {
            $instanceID = $instances[$instanceName]
            $out += [pscustomobject]@{ Text = "Processing SQL Instance: $instanceName (ID: $instanceID)"; Color = "blue" }
            $needsRestart = $false

            # ---- Enable Mixed Mode (LoginMode=2) -------------------------------------------------
            foreach ($authPath in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer"
            )) {
                $k = $reg.OpenSubKey($authPath, $true)
                if ($k) {
                    if ($k.GetValue("LoginMode", 1) -ne 2) {
                        $k.SetValue("LoginMode", 2, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "Mixed Mode Authentication enabled at $authPath."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "Mixed Mode Authentication already enabled at $authPath."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Enable TCP ----------------------------------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    if ($k.GetValue('Enabled', 0) -ne 1) {
                        $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "TCP/IP protocol enabled at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "TCP/IP already enabled at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Set static port and clear dynamic -----------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    $curPort = $k.GetValue('TcpPort', "")
                    $curDyn  = $k.GetValue('TcpDynamicPorts', "")
                    if ($curPort -ne $tcpPort -or $curDyn -ne "") {
                        $k.SetValue('TcpPort', $tcpPort, [Microsoft.Win32.RegistryValueKind]::String)
                        $k.SetValue('TcpDynamicPorts', '', [Microsoft.Win32.RegistryValueKind]::String)
                        $out += [pscustomobject]@{ Text = "Registry port set to $tcpPort at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "TCP port already set to $tcpPort at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Enable Named Pipes --------------------------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    if ($k.GetValue('Enabled', 0) -ne 1) {
                        $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "Named Pipes protocol enabled at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "Named Pipes already enabled at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Enable Shared Memory ------------------------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    if ($k.GetValue('Enabled', 0) -ne 1) {
                        $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "Shared Memory protocol enabled at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "Shared Memory already enabled at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Restart SQL service if any change ----------------------------------------------
            if ($needsRestart) {
                $svcName = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instanceName" }
                $out += [pscustomobject]@{ Text = "Restarting SQL Service $svcName on $machine..."; Color = "gray" }
                sc.exe "\\$machine" stop $svcName | Out-Null
                Start-Sleep -Seconds 10
                sc.exe "\\$machine" start $svcName | Out-Null
                Start-Sleep -Seconds 5
                $out += [pscustomobject]@{ Text = "SQL Service $svcName restarted successfully on $machine."; Color = "green" }
                $laneNeedsRestart = $true
            } else {
                $out += [pscustomobject]@{ Text = "No protocol/auth changes for $instanceName on $machine. No restart needed."; Color = "green" }
            }
        } # foreach instance

        $reg.Close()
    }

    # --- Quick transport probe: TCP then NP --------------------------------------------------------
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $task = $tcpClient.ConnectAsync($machine, [int]$tcpPort)
        if ($task.Wait(1500) -and $tcpClient.Connected) {
            $tcpClient.Close()
            $protocol = "TCP"
        }
        else {
            $npCs = "Server=np:\\\\$machine\\pipe\\sql\\query;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=5"
            $cn = New-Object System.Data.SqlClient.SqlConnection $npCs
            try {
                $cn.Open()
                $cmd = $cn.CreateCommand(); $cmd.CommandText = "SELECT 1"; $cmd.CommandTimeout = 2
                [void]$cmd.ExecuteScalar()
                $protocol = "Named Pipes"
            } catch { } finally { $cn.Close() }
        }
    } catch { }

    if ($protocol -eq 'File') {
        $out += [pscustomobject]@{ Text = "Pre-probe did not confirm TCP or NP (likely timeouts). Will use Linked Server live test to confirm."; Color = "yellow" }
    }

    # --- Linked server creation and live test (optional) -------------------------------------------
    if ($createLinks) {
        # Provider preference; try to filter by installed list (best effort).
        $pref = @('SQLOLEDB','MSOLEDBSQL','SQLNCLI11','SQLNCLI','MSDASQL')
        try {
            $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=8"
            $cn = New-Object System.Data.SqlClient.SqlConnection $cs
            $cn.Open()
            try {
                $cmd = $cn.CreateCommand(); $cmd.CommandTimeout = 8; $cmd.CommandText = "EXEC master.sys.sp_enum_oledb_providers;"
                $rdr = $cmd.ExecuteReader()
                $dt = New-Object System.Data.DataTable
                $dt.Load($rdr)
                $installed = @()
                foreach ($row in $dt.Rows) {
                    $col = if ($dt.Columns.Contains('name')) { 'name' } elseif ($dt.Columns.Contains('PROGID')) { 'PROGID' } else { $dt.Columns[0].ColumnName }
                    $installed += [string]$row[$col]
                }
                if ($installed.Count -gt 0) {
                    $pref2 = @()
                    foreach ($p in $pref) { if ($installed -contains $p) { $pref2 += $p } }
                    if ($pref2.Count -gt 0) { $pref = $pref2 }
                }
            } finally { $cn.Close() }
        } catch { }

        $tcpDataSource = "tcp:$machine,$tcpPort"
        $linkName = $machine
        $firstCreatedProv = $null
        $firstCreatedTsql = $null

        foreach ($prov in $pref) {
            try {
                $optTsql = "BEGIN TRY EXEC master.dbo.sp_MSset_oledb_prop N'$prov', N'AllowInProcess', 1; END TRY BEGIN CATCH END CATCH;
BEGIN TRY EXEC master.dbo.sp_MSset_oledb_prop N'$prov', N'DynamicParameters', 1; END TRY BEGIN CATCH END CATCH;"

                if ($prov -ne 'MSDASQL') {
                    if ($protocol -eq 'TCP') {
                        $creationTransport = 'TCP'
                        $createTsql = @"
IF EXISTS(SELECT 1 FROM sys.servers WHERE name=N'$linkName')
    EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';
$optTsql
EXEC master.dbo.sp_addlinkedserver @server=N'$linkName', @srvproduct=N'', @provider=N'$prov', @datasrc=N'$tcpDataSource';
EXEC master.dbo.sp_serveroption N'$linkName', N'data access', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'use remote collation', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'rpc out', N'true';
IF NOT EXISTS (SELECT 1 FROM sys.linked_logins ll JOIN sys.servers s ON s.server_id=ll.server_id WHERE s.name=N'$linkName')
    EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'$linkName', @useself='TRUE';
"@
                    }
                    else {
                        $creationTransport = 'Named Pipes'
                        $createTsql = @"
IF EXISTS(SELECT 1 FROM sys.servers WHERE name=N'$linkName')
    EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';
$optTsql
EXEC master.dbo.sp_addlinkedserver
    @server=N'$linkName', @srvproduct=N'', @provider=N'$prov',
    @datasrc=N'\\$machine\pipe\sql\query', @provstr=N'Network Library=dbnmpntw';
EXEC master.dbo.sp_serveroption N'$linkName', N'data access', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'use remote collation', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'rpc out', N'true';
IF NOT EXISTS (SELECT 1 FROM sys.linked_logins ll JOIN sys.servers s ON s.server_id=ll.server_id WHERE s.name=N'$linkName')
    EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'$linkName', @useself='TRUE';
"@
                    }
                }
                else {
                    if ($protocol -eq 'TCP') {
                        $creationTransport = 'TCP'
                        $provstr18 = "Driver={ODBC Driver 18 for SQL Server};Server=$machine,$tcpPort;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                        $provstr17 = "Driver={ODBC Driver 17 for SQL Server};Server=$machine,$tcpPort;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                    }
                    else {
                        $creationTransport = 'Named Pipes'
                        $npServer = "np:\\\\$machine\\pipe\\sql\\query"
                        $provstr18 = "Driver={ODBC Driver 18 for SQL Server};Server=$npServer;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                        $provstr17 = "Driver={ODBC Driver 17 for SQL Server};Server=$npServer;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                    }
                    $createTsql = @"
IF EXISTS(SELECT 1 FROM sys.servers WHERE name=N'$linkName')
    EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';
$optTsql
DECLARE @ok bit = 0;
BEGIN TRY
    EXEC master.dbo.sp_addlinkedserver @server=N'$linkName', @srvproduct=N'', @provider=N'MSDASQL', @provstr=N'$provstr18';
    SET @ok = 1;
END TRY BEGIN CATCH END CATCH;
IF (@ok = 0)
BEGIN TRY
    EXEC master.dbo.sp_addlinkedserver @server=N'$linkName', @srvproduct=N'', @provider=N'MSDASQL', @provstr=N'$provstr17';
    SET @ok = 1;
END TRY BEGIN CATCH END CATCH;
IF (@ok = 1)
BEGIN
    EXEC master.dbo.sp_serveroption N'$linkName', N'data access', N'true';
    EXEC master.dbo.sp_serveroption N'$linkName', N'use remote collation', N'true';
    EXEC master.dbo.sp_serveroption N'$linkName', N'rpc out', N'true';
    IF NOT EXISTS (SELECT 1 FROM sys.linked_logins ll JOIN sys.servers s ON s.server_id=ll.server_id WHERE s.name=N'$linkName')
        EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'$linkName', @useself='TRUE';
END
ELSE
BEGIN
    RAISERROR('MSDASQL addlinkedserver failed for both ODBC 18 and 17', 11, 1);
END
"@
                }

                # -- Execute the CREATE script
                $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=30"
                $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                $cn.Open()
                try {
                    $cmd = $cn.CreateCommand()
                    $cmd.CommandTimeout = 30
                    $cmd.CommandText = $createTsql
                    [void]$cmd.ExecuteNonQuery()
                } finally { $cn.Close() }

                if (-not $firstCreatedProv) { $firstCreatedProv = $prov; $firstCreatedTsql = $createTsql }

                # -- Live test (no MSDTC): query remote master via linked server
                $testOk = $false
                $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=8"
                $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                $cn.Open()
                try {
                    $cmd = $cn.CreateCommand()
                    $cmd.CommandTimeout = 8
                    $cmd.CommandText = "SELECT TOP 1 name FROM [$linkName].master.sys.databases"
                    $rdr = $cmd.ExecuteReader()
                    $dt = New-Object System.Data.DataTable
                    $dt.Load($rdr)
                    $testOk = ($dt -ne $null)
                } catch {
                    $testOk = $false
                } finally { $cn.Close() }

                if ($testOk) {
                    $createdAndTested = $true
                    # Try to detect the actual transport on the remote side via DMV
                    $normalized = $false
                    try {
                        $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=8"
                        $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                        $cn.Open()
                        try {
                            $cmd = $cn.CreateCommand()
                            $cmd.CommandTimeout = 8
                            $cmd.CommandText = "EXEC ('SELECT TOP 1 net_transport FROM sys.dm_exec_connections WHERE session_id = @@SPID') AT [$linkName]"
                            $rdr = $cmd.ExecuteReader()
                            $dt = New-Object System.Data.DataTable
                            $dt.Load($rdr)
                            if ($dt -and $dt.Rows.Count -gt 0) {
                                $rt = [string]$dt.Rows[0][0]
                                if ($rt -match 'TCP')    { $protocol = 'TCP';          $normalized = $true }
                                elseif ($rt -match 'Named') { $protocol = 'Named Pipes'; $normalized = $true }
                                elseif ($rt -match 'Shared') { $protocol = 'Shared Memory'; $normalized = $true }
                                $out += [pscustomobject]@{ Text = "Remote net_transport via [$linkName]: $rt" + ($(if ($normalized) { " (persisting as '$protocol')" } else { "" })); Color = "gray" }
                            }
                        } finally { $cn.Close() }
                    } catch { }

                    if (-not $normalized -or [string]::IsNullOrWhiteSpace($protocol) -or $protocol -eq 'File') { $protocol = $creationTransport }
                    $out += [pscustomobject]@{ Text = "Linked Server [$linkName] created via provider '$prov' over $protocol - **LIVE TEST PASSED**."; Color = "green" }
                    break
                }
                else {
                    # Drop and try next provider
                    try {
                        $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15"
                        $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                        $cn.Open()
                        try {
                            $cmd = $cn.CreateCommand()
                            $cmd.CommandTimeout = 15
                            $cmd.CommandText = "EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';"
                            [void]$cmd.ExecuteNonQuery()
                        } finally { $cn.Close() }
                    } catch { }
                    $out += [pscustomobject]@{ Text = "Provider '$prov' created [$linkName] over $creationTransport but live test FAILED - trying next provider..."; Color = "yellow" }
                }
            }
            catch {
                $out += [pscustomobject]@{ Text = "Provider '$prov' failed to create Linked Server [$linkName] - trying next provider..."; Color = "gray" }
            }
        } # foreach provider

        if (-not $createdAndTested -and $firstCreatedProv) {
            # Leave a creatable provider so admin can map a SQL login later
            try {
                $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=30"
                $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                $cn.Open()
                try {
                    $cmd = $cn.CreateCommand()
                    $cmd.CommandTimeout = 30
                    $cmd.CommandText = $firstCreatedTsql
                    [void]$cmd.ExecuteNonQuery()
                } finally { $cn.Close() }
            } catch { }
            if ($protocol -eq 'File' -or [string]::IsNullOrWhiteSpace($protocol)) { $protocol = $creationTransport }
            $out += [pscustomobject]@{ Text = "No provider passed live test. Re-created [$linkName] using '$firstCreatedProv' over $protocol so you can map a SQL login if needed."; Color = "yellow" }
        }
        elseif (-not $createdAndTested -and -not $firstCreatedProv) {
            $out += [pscustomobject]@{ Text = "All provider attempts failed creating Linked Server [$linkName]."; Color = "red" }
        }
    }
    else {
        $out += [pscustomobject]@{ Text = "CreateLinkedServers disabled; skipped Linked Server creation."; Color = "gray" }
    }
}
catch {
    $out += [pscustomobject]@{ Text = "Failed to process [$machine]: $_"; Color = "red" }
    $protocol = "File"
}

# Final guard to avoid returning 'File' on success
if ($createdAndTested -and ($protocol -eq 'File' -or [string]::IsNullOrWhiteSpace($protocol))) {
    $protocol = $creationTransport
}

# Return structured result (parent thread will log and persist)
[pscustomobject]@{
    Output    = $out
    Protocol  = $protocol
    Lane      = $lane
    Machine   = $machine
    Restarted = $laneNeedsRestart
}
'@
	
	# ---- Launch all lanes in parallel via runspaces -----------------------------------------------
	$tasks = @()
	foreach ($lane in $lanes)
	{
		$machine = $LaneNumToMachineName[$lane]
		if (-not $machine)
		{
			Write_Log "Machine name not found for lane $lane. Skipping." "yellow"
			continue
		}
		
		# Create PowerShell instance bound to the runspace pool
		$ps = [PowerShell]::Create()
		$ps.RunspacePool = $pool
		# IMPORTANT: pass a pure boolean for $createLinks
		$ps.AddScript($laneWork).AddArgument($machine).AddArgument($lane).AddArgument($StoreNumber).AddArgument($tcpPort).AddArgument([bool]$CreateLinkedServers.IsPresent).AddArgument($localInstance) | Out-Null
		
		$handle = $ps.BeginInvoke()
		$tasks += [pscustomobject]@{ Lane = $lane; Machine = $machine; PS = $ps; Handle = $handle; Done = $false }
	}
	
	# ---- Collect results as lanes complete (quasi-live) -------------------------------------------
	$restarted = @()
	while ($true)
	{
		$allDone = $true
		foreach ($t in $tasks)
		{
			if (-not $t.Done)
			{
				if ($t.Handle.IsCompleted)
				{
					# Finish and dispose
					$res = $t.PS.EndInvoke($t.Handle)
					$t.PS.Dispose()
					$t.Done = $true
					
					# Emit buffered logs for this lane
					foreach ($line in $res.Output) { Write_Log $line.Text $line.Color }
					
					# Persist in-memory protocol maps
					$script:LaneProtocols[$res.Lane] = $res.Protocol
					$script:ProtocolResults = $script:ProtocolResults | Where-Object { $_.Lane -ne $res.Lane }
					$script:ProtocolResults += [PSCustomObject]@{ Lane = $res.Lane; Protocol = $res.Protocol }
					
					# Persist to file (dedup by lane)
					$protocolResultsFile = 'C:\Tecnica_Systems\Alex_C.T\Setup_Files\Protocol_Results.txt'
					$laneStr = $res.Lane.ToString().PadLeft(3, '0')
					$proto = $res.Protocol
					$dir = [System.IO.Path]::GetDirectoryName($protocolResultsFile)
					if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
					if (-not (Test-Path $protocolResultsFile)) { New-Item -Path $protocolResultsFile -ItemType File -Force | Out-Null }
					$all = @()
					if (Test-Path $protocolResultsFile)
					{
						$all = Get-Content -LiteralPath $protocolResultsFile -ErrorAction SilentlyContinue | Where-Object { $_ -match '\S' }
					}
					$all = $all | Where-Object { -not ($_ -match "^\s*0*${laneStr}\s*,") }
					$all += "$laneStr,$proto"
					$sorted = $all | Sort-Object { ($_ -split ',')[0] -as [int] }
					[System.IO.File]::WriteAllLines($protocolResultsFile, $sorted, [System.Text.Encoding]::UTF8)
					
					if ($res.Restarted) { $restarted += $res.Lane }
					
					# Final per-lane summary line
					Write_Log "Protocol detected for $($res.Machine) (Lane $($res.Lane)): $proto" "magenta"
				}
				else
				{
					$allDone = $false
				}
			}
		}
		if ($allDone) { break }
		Start-Sleep -Milliseconds 150
	}
	
	# ---- Cleanup pool -----------------------------------------------------------------------------
	$pool.Close()
	$pool.Dispose()
	
	# ---- Post-actions for lanes that needed restart -----------------------------------------------
	if ($restarted.Count -gt 0)
	{
		Send_Restart_All_Programs -StoreNumber $StoreNumber -LaneNumbers $restarted -Silent
		Write_Log "Restart All Programs sent to restarted lanes: $($restarted -join ', ')" "green"
	}
	
	Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Completed ====================" "blue"
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

# =========================================================================================
# FUNCTION: Add_Scale_Credentials  (fast, parallel, silent; PS 5.1 compatible)
# -----------------------------------------------------------------------------------------
# - Gathers unique scale IPs from ScaleCodeToIPInfo (uses .FullIP or IPNetwork+IPDevice).
# - For each IP, in parallel:
#     * Quick preflight: TCP 445 reachable? If not, skip.
#     * If \\IP\c$ already accessible, skip (no creds needed).
#     * Try creds in order: bizuser/bizerba, then bizuser/biyerba using "net use".
#       On success: remove that mapping and persist creds with CMDKEY.
# - No output, no logging, returns immediately (background runspaces).
# - Keeps handles in $script:ScaleCredTasks so they don't get GC'd; you can clean them later.
# =========================================================================================

function Add_Scale_Credentials
{
	param (
		[Parameter(Mandatory = $true)]
		[hashtable]$ScaleCodeToIPInfo,
		[int]$MaxParallel = 16
	)
	
	# Collect unique IPs
	$scaleIPs = $ScaleCodeToIPInfo.Values | ForEach-Object {
		if ($_.FullIP) { $_.FullIP }
		elseif ($_.IPNetwork -and $_.IPDevice) { "$($_.IPNetwork)$($_.IPDevice)" }
	} | Where-Object { $_ } | Select-Object -Unique
	
	if (-not $scaleIPs -or $scaleIPs.Count -eq 0) { return }
	
	# Runspace pool
	$min = 1
	if ($MaxParallel -lt 1) { $MaxParallel = 1 }
	$pool = [runspacefactory]::CreateRunspacePool($min, $MaxParallel)
	$pool.ApartmentState = 'MTA'
	$pool.Open()
	
	if (-not $script:ScaleCredTasks) { $script:ScaleCredTasks = @{ } }
	
	foreach ($ip in $scaleIPs)
	{
		$ps = [powershell]::Create()
		$null = $ps.AddScript({
				param ($ip)
				
				# ---- quick SMB reachability (port 445) ----
				$smbOk = $false
				try
				{
					$client = New-Object System.Net.Sockets.TcpClient
					$ar = $client.BeginConnect($ip, 445, $null, $null)
					if ($ar.AsyncWaitHandle.WaitOne(400))
					{
						try { $client.EndConnect($ar) }
						catch { }
						if ($client.Connected) { $smbOk = $true }
					}
					$client.Close()
				}
				catch { }
				
				if (-not $smbOk) { return }
				
				# ---- already accessible without creds? ----
				$already = $false
				try
				{
					if (Test-Path "\\$ip\c$") { $already = $true }
				}
				catch { }
				
				if ($already) { return }
				
				# ---- try credentials quickly via NET USE (no persistence) ----
				$tryCreds = @(
					@{ User = 'bizuser'; Pass = 'bizerba' },
					@{ User = 'bizuser'; Pass = 'biyerba' }
				)
				
				foreach ($c in $tryCreds)
				{
					# map attempt (non-persistent)
					$cmd = "net use \\$ip\c$ /user:$($c.User) $($c.Pass) /persistent:no"
					$null = & cmd.exe /c $cmd 2>$null
					$rc = $LASTEXITCODE
					
					if ($rc -eq 0)
					{
						# we proved credentials work; tear down mapping then persist with cmdkey
						$null = & cmd.exe /c "net use \\$ip\c$ /delete /y" 2>$null
						$null = & cmdkey.exe /add:$ip /user:$($c.User) /pass:$($c.Pass) 2>$null
						return
					}
					else
					{
						# ensure no lingering mapping before next attempt
						$null = & cmd.exe /c "net use \\$ip\c$ /delete /y" 2>$null
					}
				}
			}).AddArgument($ip)
		
		$ps.RunspacePool = $pool
		$handle = $ps.BeginInvoke()
		
		# Stash to prevent GC and allow later cleanup if you want
		$script:ScaleCredTasks[$ip] = @{ PS = $ps; Handle = $handle }
	}
	
	# Note: We do not wait. This function returns immediately while tasks run in background.
	# Optional: You can later clean completed entries like this:
	# foreach ($k in @($script:ScaleCredTasks.Keys)) {
	#   $st = $script:ScaleCredTasks[$k]; if ($st.Handle.IsCompleted) { $st.PS.EndInvoke($st.Handle); $st.PS.Dispose(); $script:ScaleCredTasks.Remove($k) }
	# }
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

function Remove-DuplicateFilesByContentAndLines {
    param([string]`$Path)

    # --------- Remove exact duplicates (keep oldest) ----------
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

    # --------- Remove files whose lines all exist in another file (any order) ----------
    # Reload the file list after removing exact duplicates
    `$files = Get-ChildItem -Path `$Path -File -ErrorAction SilentlyContinue

    for (`$i = 0; `$i -lt `$files.Count; `$i++) {
        `$fileA = `$files[`$i]
        `$linesA = Get-Content -LiteralPath `$fileA.FullName -ErrorAction SilentlyContinue | Where-Object { `$_.Trim() -ne "" }
        if (-not `$linesA -or `$linesA.Count -eq 0) { continue }
        for (`$j = 0; `$j -lt `$files.Count; `$j++) {
            if (`$i -eq `$j) { continue }
            `$fileB = `$files[`$j]
            `$linesB = Get-Content -LiteralPath `$fileB.FullName -ErrorAction SilentlyContinue | Where-Object { `$_.Trim() -ne "" }
            if (`$linesB.Count -lt `$linesA.Count) { continue } # Only look for supersets
            # Check if every line in A is in B (case-insensitive)
            `$allFound = `$linesA | ForEach-Object { `$lineA = `$_.Trim(); `$linesB -contains `$lineA }
            if (`$allFound -notcontains `$false) {
                # Try to delete fileA (the smaller one)
                `$canDelete = `$true
                try {
                    `$stream = [System.IO.File]::Open(`$fileA.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                    `$stream.Close()
                } catch { `$canDelete = `$false }
                if (`$canDelete) {
                    try {
                        Remove-Item `$fileA.FullName -Force
                        Add-Content -Path `$LogPath -Value "`$(Get-Date): Removed `$($fileA.FullName) (all lines found in `$($fileB.FullName))"
                    } catch {
                        Add-Content -Path `$LogPath -Value "`$(Get-Date): Failed to remove `$($fileA.FullName): `$_"
                    }
                } else {
                    Add-Content -Path `$LogPath -Value "`$(Get-Date): `$($fileA.FullName) is in use, skipped deletion"
                }
                # After deleting, break both loops to avoid index issues
                `$files = Get-ChildItem -Path `$Path -File -ErrorAction SilentlyContinue
                `$i = -1; break
            }
        }
    }
}

Add-Content -Path `$LogPath -Value "`$(Get-Date): Monitor started with interval `$IntervalSeconds seconds"
while (`$true) {
    Remove-DuplicateFilesByContentAndLines -Path `$Path
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
	$dedupedLanes = @{ }
	foreach ($kv in $LaneNumToMachineName.GetEnumerator())
	{
		# Only write one file per machine name
		$machineName = $kv.Value
		if ($machineName -and -not $dedupedLanes.ContainsKey($machineName))
		{
			$dedupedLanes[$machineName] = $true
			if ($machineName.ToUpper() -match '^(POS|SCO)\d+$')
			{
				$fileName = "$($machineName.ToUpper()).vnc"
			}
			else
			{
				$fileName = "Lane_$($kv.Key).vnc"
			}
			$filePath = Join-Path $lanesDir $fileName
			$parent = Split-Path $filePath -Parent
			if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory | Out-Null }
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
	}
	Write_Log "$laneCount lane VNC files written to $lanesDir`r`n" "blue"
	
	# ---- Scales ---- #
	Write_Log "-------------------- Exporting Scale VNC Files --------------------" "blue"
	$scaleCount = 0
	$dedupedScales = @{ }
	foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
	{
		$scaleObj = $kv.Value
		# Make sure to dedupe by IP address
		$ip = if ($scaleObj.FullIP) { $scaleObj.FullIP }
		elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)" }
		else { $null }
		if ($ip -and -not $dedupedScales.ContainsKey($ip))
		{
			$dedupedScales[$ip] = $true
			$octets = $ip -split '\.'
			$lastOctet = $octets[-1]
			$brandRaw = ($scaleObj.ScaleBrand -as [string]).Trim()
			$model = ($scaleObj.ScaleModel -as [string]).Trim()
			$brand = if ($brandRaw)
			{
				($brandRaw -split ' ' | ForEach-Object {
						if ($_.Length -gt 0) { $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower() }
						else { $_ }
					}) -join ' '
			}
			else { "" }
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
	}
	Write_Log "$scaleCount scale VNC files written to $scalesDir`r`n" "blue"
	
	# ---- Backoffices ---- #
	Write_Log "-------------------- Exporting Backoffice VNC Files --------------------" "blue"
	$boCount = 0
	$dedupedBOs = @{ }
	foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
	{
		$boName = $kv.Value
		$terminal = $kv.Key
		if ($boName -and -not $dedupedBOs.ContainsKey($boName))
		{
			$dedupedBOs[$boName] = $true
			$fileName = "Backoffice_${terminal}.vnc"
			$filePath = Join-Path $backofficesDir $fileName
			$parent = Split-Path $filePath -Parent
			if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory | Out-Null }
			$content = $vncTemplate.Replace('%%HOST%%', $boName).Replace('%%PASSWORD%%', $DefaultVNCPassword)
			[System.IO.File]::WriteAllText($filePath, $content, $script:ansiPcEncoding)
			Write_Log "Created: $filePath" "green"
			$boCount++
		}
	}
	Write_Log "$boCount backoffice VNC files written to $backofficesDir`r`n" "blue"
	
	Write_Log "VNC file export complete!" "green"
	Write_Log "`r`n==================== Export_VNCFiles_ForAllNodes Completed ====================" "blue"
}

# ===================================================================================================
#                           FUNCTION: Schedule_LocalDB_Backup
# ---------------------------------------------------------------------------------------------------
# Description:
#   Interactive GUI tool to configure, schedule, and maintain automated SQL Server database backups
#   for the local store server. Prompts user for preferred time, frequency, and retention policy.
#   Generates and schedules a PowerShell script that:
#     - Backs up the local database to a dated .bak file
#     - Deletes oldest backups, keeping only the specified number
#     - Uses Write_Log for all logging and status messages
#   Schedules as a SYSTEM task for maximum reliability.
#
# Parameters:
#   None (uses environment and $script:FunctionResults for database info)
#
# Usage:
#   Schedule_LocalDB_Backup
#
# Author: Alex_C.T
# ===================================================================================================

function Schedule_LocalDB_Backup
{
	Write_Log "`r`n==================== Starting Schedule_LocalDB_Backup ====================`r`n" "blue"
	
	try
	{
		# --- Validate FunctionResults for DB and Paths
		$dbName = $script:FunctionResults['DBNAME']
		$dbServer = $script:FunctionResults['DBSERVER']
		$backupDir = $script:BackupRoot
		$scriptsDir = $script:ScriptsFolder
		$serverFolder = Join-Path $backupDir $LocalHost
		
		if (-not $dbName -or -not $dbServer)
		{
			Write_Log "DBNAME or DBSERVER not found in FunctionResults. Aborting." "red"
			return
		}
		if (-not $backupDir) { $backupDir = "C:\Tecnica_Systems\Backups\" }
		if (-not $scriptsDir) { $scriptsDir = "C:\Tecnica_Systems\Alex_C.T\Scripts" }
		
		# --- Prompt User: Backup Time, Freq, Retention
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
		
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Configure Local DB Backup Scheduler"
		$form.Size = New-Object System.Drawing.Size(375, 245)
		$form.StartPosition = "CenterScreen"
		$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$form.MaximizeBox = $false
		$form.MinimizeBox = $false
		
		$lblTime = New-Object System.Windows.Forms.Label
		$lblTime.Text = "Time to run backup (24h, HH:mm):"
		$lblTime.Location = New-Object System.Drawing.Point(10, 18)
		$lblTime.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblTime)
		
		$txtTime = New-Object System.Windows.Forms.MaskedTextBox
		$txtTime.Mask = "00:00"
		$txtTime.Text = "01:00"
		$txtTime.Location = New-Object System.Drawing.Point(220, 16)
		$txtTime.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($txtTime)
		
		$lblFreq = New-Object System.Windows.Forms.Label
		$lblFreq.Text = "Frequency (every X days):"
		$lblFreq.Location = New-Object System.Drawing.Point(10, 58)
		$lblFreq.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblFreq)
		
		$numFreq = New-Object System.Windows.Forms.NumericUpDown
		$numFreq.Minimum = 1
		$numFreq.Maximum = 31
		$numFreq.Value = 1
		$numFreq.Location = New-Object System.Drawing.Point(220, 56)
		$numFreq.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($numFreq)
		
		$lblKeep = New-Object System.Windows.Forms.Label
		$lblKeep.Text = "How many backups to keep:"
		$lblKeep.Location = New-Object System.Drawing.Point(10, 98)
		$lblKeep.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblKeep)
		
		$numKeep = New-Object System.Windows.Forms.NumericUpDown
		$numKeep.Minimum = 1
		$numKeep.Maximum = 99
		$numKeep.Value = 3
		$numKeep.Location = New-Object System.Drawing.Point(220, 96)
		$numKeep.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($numKeep)
		
		$btnOK = New-Object System.Windows.Forms.Button
		$btnOK.Text = "OK"
		$btnOK.Location = New-Object System.Drawing.Point(70, 150)
		$btnOK.Size = New-Object System.Drawing.Size(90, 30)
		$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$form.AcceptButton = $btnOK
		$form.Controls.Add($btnOK)
		
		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Location = New-Object System.Drawing.Point(185, 150)
		$btnCancel.Size = New-Object System.Drawing.Size(90, 30)
		$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$form.CancelButton = $btnCancel
		$form.Controls.Add($btnCancel)
		
		$result = $form.ShowDialog()
		if ($result -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "Backup scheduling cancelled by user." "yellow"
			Write_Log "`r`n==================== Schedule_LocalDB_Backup Completed ====================" "blue"
			return
		}
		
		$timeInput = $txtTime.Text
		$freqInput = [int]$numFreq.Value
		$keepInput = [int]$numKeep.Value
		
		if ($timeInput -notmatch '^\d{2}:\d{2}$')
		{
			[System.Windows.Forms.MessageBox]::Show("Time must be in HH:mm (24h) format.", "Input Error", 0, 16)
			Write_Log "Invalid time format entered." "red"
			return
		}
		$hours, $minutes = $timeInput.Split(":")
		if ([int]$hours -gt 23 -or [int]$minutes -gt 59)
		{
			[System.Windows.Forms.MessageBox]::Show("Invalid time entered.", "Input Error", 0, 16)
			Write_Log "Invalid time (out of range)." "red"
			return
		}
		
		# --- Compose the backup script (no nested here-string)
		$backupScript = @"
# Auto-generated by Schedule_${LocalHost}_DB_Backup (Alex_C.T)
`$ErrorActionPreference = 'Stop'
`$BackupRoot = `"$serverFolder`"
`$Database = `"$dbName`"
`$SqlInstance = `"$dbServer`"
`$MaxBackups = $keepInput

# Compose backup file path with timestamp
`$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
`$backupFile = Join-Path `$BackupRoot ("{0}_{1}.bak" -f `$Database, `$timestamp)

# Ensure folder exists
if (-not (Test-Path `$BackupRoot)) {
    New-Item -Path `$BackupRoot -ItemType Directory -Force | Out-Null
}

# Delete oldest backups if needed
`$bakFiles = Get-ChildItem -Path `$BackupRoot -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
while (`$bakFiles.Count -ge `$MaxBackups) {
    Remove-Item -Path `$bakFiles[0].FullName -Force
    `$bakFiles = Get-ChildItem -Path `$BackupRoot -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
}

# Run the backup
`$tsql = "BACKUP DATABASE [`$Database] TO DISK = N'`$backupFile' WITH NOFORMAT, NOINIT, NAME = N'`$Database-FullBackup-`$timestamp', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
& sqlcmd -S `"$dbServer`" -Q `$tsql -b
`$exitCode = `$LASTEXITCODE

if (`$exitCode -eq 0) {
    # Success log
    "[`$(Get-Date)] Backup complete: `$backupFile" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
}
else {
    # Failure log 
    "[`$(Get-Date)] Backup FAILED for [`$Database] on [`$SqlInstance] (exit code: `$exitCode)" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
}
"@
		
		# --- Write the backup script to the scripts folder
		$scriptName = "Run_${LocalHost}_DB_Backup.ps1"
		$backupScriptPath = Join-Path $scriptsDir $scriptName
		if (-not (Test-Path $scriptsDir))
		{
			New-Item -Path $scriptsDir -ItemType Directory -Force | Out-Null
		}
		Set-Content -Path $backupScriptPath -Value $backupScript -Encoding UTF8
		
		# --- Build Task Scheduler arguments ---
		$hour = "{0:D2}" -f [int]$hours
		$min = "{0:D2}" -f [int]$minutes
		$startTime = "$hour`:$min"
		
		$freqArg = "/SC DAILY"
		if ($freqInput -gt 1)
		{
			$freqArg = "/SC DAILY /MO $freqInput"
		}
		
		$taskName = "${LocalHost}_DB_Backup_Schedule"
		$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"`"$backupScriptPath`"`""
		
		# --- Remove existing task if present ---
		schtasks.exe /Delete /TN "$taskName" /F 2>$null | Out-Null
		
		# --- Create scheduled task ---
		$cmd = "schtasks.exe /Create /RU SYSTEM /RL HIGHEST /TN `"$taskName`" /TR `"$action`" $freqArg /ST $startTime /F"
		Write_Log "Creating scheduled task with command:" "yellow"
		Write_Log $cmd "gray"
		Invoke-Expression $cmd
		
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Backup task scheduled successfully (`"$taskName`") at $startTime every $freqInput day(s)." "green"
			[System.Windows.Forms.MessageBox]::Show("Backup task scheduled successfully.`nScript: $backupScriptPath", "Scheduled!", 0, 64)
		}
		else
		{
			Write_Log "Failed to schedule backup task. Check permissions or path." "red"
			[System.Windows.Forms.MessageBox]::Show("Failed to schedule backup task.`nCheck permissions or path.", "Error", 0, 16)
		}
	}
	catch
	{
		Write_Log "Fatal error in Schedule_LocalDB_Backup: $($_.Exception.Message)" "red"
		[System.Windows.Forms.MessageBox]::Show("Fatal error: $($_.Exception.Message)", "Error", 0, 16)
	}
	Write_Log "`r`n==================== Schedule_LocalDB_Backup Completed ====================" "blue"
}

# ===================================================================================================
#                           FUNCTION: Schedule_Storeman_Zip_Backup
# ---------------------------------------------------------------------------------------------------
# Description:
#   Interactive GUI tool to configure, schedule, and maintain automated ZIP backups of the Storeman
#   folder. Prompts user for preferred time, frequency, and retention policy.
#   Generates and schedules a PowerShell script that:
#     - Zips the Storeman folder to a dated .zip file
#     - Deletes oldest backups, keeping only the specified number
#     - Uses Write_Log for all logging and status messages
#   Schedules as a SYSTEM task for maximum reliability.
#
# Parameters:
#   None (uses $BasePath as detected Storeman folder)
#
# Usage:
#   Schedule_Storeman_Zip_Backup
#
# Author: Alex_C.T
# ===================================================================================================

function Schedule_Storeman_Zip_Backup
{
	Write_Log "`r`n==================== Starting Schedule_Storeman_Zip_Backup ====================`r`n" "blue"
	
	try
	{
		# --- Use detected Storeman folder ---
		$storemanPath = $BasePath
		if (-not (Test-Path $storemanPath))
		{
			Write_Log "Storeman path ($storemanPath) not found. Aborting." "red"
			return
		}
		$LocalHost = $env:COMPUTERNAME
		$backupRoot = "${script:BackupRoot}${LocalHost}"
		$scriptsDir = $script:ScriptsFolder
		$backupScriptName = "Run_${LocalHost}_Storeman_Zip_Backup.ps1"
		$taskName = "${LocalHost}_Storeman_Zip_Backup"
		
		# --- Prompt User: Backup Time, Frequency, Retention (defaults set for weekly) ---
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
		
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Configure Storeman ZIP Backup Scheduler"
		$form.Size = New-Object System.Drawing.Size(385, 245)
		$form.StartPosition = "CenterScreen"
		$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$form.MaximizeBox = $false
		$form.MinimizeBox = $false
		
		$lblTime = New-Object System.Windows.Forms.Label
		$lblTime.Text = "Time to run backup (24h, HH:mm):"
		$lblTime.Location = New-Object System.Drawing.Point(10, 18)
		$lblTime.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblTime)
		
		$txtTime = New-Object System.Windows.Forms.MaskedTextBox
		$txtTime.Mask = "00:00"
		$txtTime.Text = "02:00"
		$txtTime.Location = New-Object System.Drawing.Point(220, 16)
		$txtTime.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($txtTime)
		
		$lblFreq = New-Object System.Windows.Forms.Label
		$lblFreq.Text = "Frequency (every X days):"
		$lblFreq.Location = New-Object System.Drawing.Point(10, 58)
		$lblFreq.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblFreq)
		
		$numFreq = New-Object System.Windows.Forms.NumericUpDown
		$numFreq.Minimum = 1
		$numFreq.Maximum = 31
		$numFreq.Value = 7 # Default = once a week
		$numFreq.Location = New-Object System.Drawing.Point(220, 56)
		$numFreq.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($numFreq)
		
		$lblKeep = New-Object System.Windows.Forms.Label
		$lblKeep.Text = "How many backups to keep:"
		$lblKeep.Location = New-Object System.Drawing.Point(10, 98)
		$lblKeep.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblKeep)
		
		$numKeep = New-Object System.Windows.Forms.NumericUpDown
		$numKeep.Minimum = 1
		$numKeep.Maximum = 99
		$numKeep.Value = 1
		$numKeep.Location = New-Object System.Drawing.Point(220, 96)
		$numKeep.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($numKeep)
		
		$btnOK = New-Object System.Windows.Forms.Button
		$btnOK.Text = "OK"
		$btnOK.Location = New-Object System.Drawing.Point(80, 150)
		$btnOK.Size = New-Object System.Drawing.Size(90, 30)
		$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$form.AcceptButton = $btnOK
		$form.Controls.Add($btnOK)
		
		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Location = New-Object System.Drawing.Point(195, 150)
		$btnCancel.Size = New-Object System.Drawing.Size(90, 30)
		$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$form.CancelButton = $btnCancel
		$form.Controls.Add($btnCancel)
		
		$result = $form.ShowDialog()
		if ($result -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "Storeman ZIP backup scheduling cancelled by user." "yellow"
			Write_Log "`r`n==================== Schedule_Storeman_Zip_Backup Completed ====================" "blue"
			return
		}
		
		$timeInput = $txtTime.Text
		$freqInput = [int]$numFreq.Value
		$keepInput = [int]$numKeep.Value
		
		if ($timeInput -notmatch '^\d{2}:\d{2}$')
		{
			[System.Windows.Forms.MessageBox]::Show("Time must be in HH:mm (24h) format.", "Input Error", 0, 16)
			Write_Log "Invalid time format entered." "red"
			return
		}
		$hours, $minutes = $timeInput.Split(":")
		if ([int]$hours -gt 23 -or [int]$minutes -gt 59)
		{
			[System.Windows.Forms.MessageBox]::Show("Invalid time entered.", "Input Error", 0, 16)
			Write_Log "Invalid time (out of range)." "red"
			return
		}
		
		# --- Compose the backup script ---
		$backupScript = @"
# Auto-generated by Schedule_${LocalHost}_Storeman_Zip_Backup (Alex_C.T)
`$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

`$storemanPath = `"$storemanPath`"
`$backupRoot = `"$backupRoot`"
if (-not (Test-Path `$storemanPath)) {
    Write-Host "Storeman path (`$storemanPath) not found. Aborting." -ForegroundColor Red
    exit 1
}
if (-not (Test-Path `$backupRoot)) {
    New-Item -Path `$backupRoot -ItemType Directory -Force | Out-Null
}

# Compose backup ZIP filename with timestamp
`$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
`$backupZipPath = Join-Path `$backupRoot ("Storeman_Backup_`$timestamp.zip")

# Delete oldest backups if needed
`$oldZips = Get-ChildItem -Path `$backupRoot -Filter "Storeman_Backup_*.zip" | Sort-Object LastWriteTime -Descending
if (`$oldZips.Count -ge $keepInput) {
    `$oldZips | Select-Object -Skip ($keepInput - 1) | ForEach-Object {
        try {
            Remove-Item `$_.FullName -Force
        } catch {
            Write-Host "Failed to delete old backup: `$_.FullName" -ForegroundColor Yellow
        }
    }
}

# Run the backup
try {
    `$zip = [System.IO.Compression.ZipFile]::Open(`$backupZipPath, 1)  # 1 = Create

    # Gather all root directories starting with "install" (case-insensitive)
    `$rootDirs = Get-ChildItem -Path `$storemanPath -Directory | Select-Object -ExpandProperty FullName
    `$installDirs = `$rootDirs | Where-Object { `$_ -match '(?i)\\install' }

    `$backupFullPathU = [System.IO.Path]::GetFullPath((Join-Path `$storemanPath "BACKUP"))
    `$backupFullPathL = [System.IO.Path]::GetFullPath((Join-Path `$storemanPath "backup"))
    `$logFullPath     = [System.IO.Path]::GetFullPath((Join-Path `$storemanPath "log"))

    `$files = Get-ChildItem -Path `$storemanPath -Recurse -File
    `$countAdded = 0

    foreach (`$file in `$files) {
        `$filePath = [System.IO.Path]::GetFullPath(`$file.FullName)

        # Skip files under any install* folder
        `$skipInstall = `$false
        foreach (`$dir in `$installDirs) {
            if (`$filePath -like "`$dir*") {
                `$skipInstall = `$true
                break
            }
        }
        if (`$skipInstall) {
            continue
        }

        # Skip other excluded folders and log files
        if (
            (`$filePath -like "`$backupFullPathU*") -or
            (`$filePath -like "`$backupFullPathL*") -or
            (`$filePath -like "`$logFullPath*") -or
            (`$file.Extension -eq ".log")
        ) {
            continue
        }

        `$relativePath = `$filePath.Substring(`$storemanPath.Length).TrimStart('\','/')
        try {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(`$zip, `$file.FullName, `$relativePath) | Out-Null
            `$countAdded++
        }
        catch {
            continue
        }
    }
    `$zip.Dispose()
    "`$([DateTime]::Now) Backup complete: `$backupZipPath (`$countAdded files)" | Out-File -FilePath (Join-Path `$backupRoot "backup.log") -Append -Encoding utf8
    Write-Host "Backup complete: `$backupZipPath (`$countAdded files)" -ForegroundColor Green
}
catch {
    if (`$zip) { `$zip.Dispose() }
    "`$([DateTime]::Now) Backup FAILED for `$storemanPath - `$(`$_.Exception.Message)" | Out-File -FilePath (Join-Path `$backupRoot "backup.log") -Append -Encoding utf8
    Write-Host "Backup FAILED: `$(`$_.Exception.Message)" -ForegroundColor Red
}
"@
		
		# --- Write the backup script to the scripts folder
		$backupScriptPath = Join-Path $scriptsDir $backupScriptName
		if (-not (Test-Path $scriptsDir))
		{
			New-Item -Path $scriptsDir -ItemType Directory -Force | Out-Null
		}
		Set-Content -Path $backupScriptPath -Value $backupScript -Encoding UTF8
		
		# --- Build Task Scheduler arguments ---
		$hour = "{0:D2}" -f [int]$hours
		$min = "{0:D2}" -f [int]$minutes
		$startTime = "$hour`:$min"
		$freqArg = "/SC DAILY"
		if ($freqInput -gt 1)
		{
			$freqArg = "/SC DAILY /MO $freqInput"
		}
		$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"`"$backupScriptPath`"`""
		
		# --- Remove existing task if present ---
		schtasks.exe /Delete /TN "$taskName" /F 2>$null | Out-Null
		
		# --- Create scheduled task ---
		$cmd = "schtasks.exe /Create /RU SYSTEM /RL HIGHEST /TN `"$taskName`" /TR `"$action`" $freqArg /ST $startTime /F"
		Write_Log "Creating scheduled task with command:" "yellow"
		Write_Log $cmd "gray"
		Invoke-Expression $cmd
		
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Storeman ZIP backup task scheduled successfully (`"$taskName`") at $startTime every $freqInput day(s)." "green"
			[System.Windows.Forms.MessageBox]::Show("Storeman ZIP backup scheduled successfully.`nScript: $backupScriptPath", "Scheduled!", 0, 64)
		}
		else
		{
			Write_Log "Failed to schedule Storeman ZIP backup task. Check permissions or path." "red"
			[System.Windows.Forms.MessageBox]::Show("Failed to schedule Storeman ZIP backup task.`nCheck permissions or path.", "Error", 0, 16)
		}
	}
	catch
	{
		Write_Log "Fatal error in Schedule_Storeman_Zip_Backup: $($_.Exception.Message)" "red"
		[System.Windows.Forms.MessageBox]::Show("Fatal error: $($_.Exception.Message)", "Error", 0, 16)
	}
	Write_Log "`r`n==================== Schedule_Storeman_Zip_Backup Completed ====================" "blue"
}

# ===================================================================================================
#                      FUNCTION: Schedule_LaneDB_Backup
# ---------------------------------------------------------------------------------------------------
# Description:
#   Lets the user select one or more lanes (via Show_Node_Selection_Form in Lane mode).
#   Only lanes with valid ProtocolResults are allowed.
#   Prompts user for backup schedule details (time, frequency, retention).
#   For each lane, creates a backup script at $ScriptsFolder\$MachineName\Run_${MachineName}DB_Backup.ps1.
#   Schedules a task for each lane using that script.
#   Each backup goes to $BackupRoot\$MachineName\STORESQL_{timestamp}.bak (etc).
#   Always uses the detected protocol (TCP/Named Pipes) for sqlcmd.

# Author: Alex_C.T
# ===================================================================================================

function Schedule_LaneDB_Backup
{
	Write_Log "`r`n==================== Starting Schedule_LaneDB_Backup ====================`r`n" "blue"
	
	# 1. Node selection (lanes only)
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Lanes for DB Backup"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for backup." "yellow"
		return
	}
	
	# 2. Check ProtocolResults for connectivity
	$goodLanes = @()
	foreach ($lane in $selection.Lanes)
	{
		$protocol = $script:LaneProtocols[$lane]
		if ($protocol -and $protocol -ne "File") { $goodLanes += $lane }
		else { Write_Log "Lane $lane not available in ProtocolResults or protocol not valid. Skipping." "yellow" }
	}
	if (-not $goodLanes -or $goodLanes.Count -eq 0)
	{
		Write_Log "No selected lanes are available for DB backup." "red"
		return
	}
	
	# 3. Prompt for backup options, user and password (all in one form, more spacing)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object Windows.Forms.Form
	$form.Text = "Configure Lane DB Backup Scheduler"
	$form.Size = [Drawing.Size]::new(500, 380) # Larger form
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	# -- Backup Settings GroupBox --
	$grpBackup = New-Object Windows.Forms.GroupBox
	$grpBackup.Text = "Backup Settings"
	$grpBackup.Location = [Drawing.Point]::new(20, 15)
	$grpBackup.Size = [Drawing.Size]::new(450, 140)
	$form.Controls.Add($grpBackup)
	
	# Time
	$lblTime = New-Object Windows.Forms.Label
	$lblTime.Text = "Time to run backup (24h, HH:mm):"
	$lblTime.Location = [Drawing.Point]::new(25, 35)
	$lblTime.Size = [Drawing.Size]::new(200, 20)
	$grpBackup.Controls.Add($lblTime)
	$txtTime = New-Object Windows.Forms.MaskedTextBox
	$txtTime.Mask = "00:00"
	$txtTime.Text = "01:00"
	$txtTime.Location = [Drawing.Point]::new(260, 33)
	$txtTime.Size = [Drawing.Size]::new(80, 24)
	$grpBackup.Controls.Add($txtTime)
	
	# Frequency
	$lblFreq = New-Object Windows.Forms.Label
	$lblFreq.Text = "Frequency (every X days):"
	$lblFreq.Location = [Drawing.Point]::new(25, 70)
	$lblFreq.Size = [Drawing.Size]::new(200, 20)
	$grpBackup.Controls.Add($lblFreq)
	$numFreq = New-Object Windows.Forms.NumericUpDown
	$numFreq.Minimum = 1
	$numFreq.Maximum = 31
	$numFreq.Value = 1
	$numFreq.Location = [Drawing.Point]::new(260, 68)
	$numFreq.Size = [Drawing.Size]::new(80, 24)
	$grpBackup.Controls.Add($numFreq)
	
	# Retention
	$lblKeep = New-Object Windows.Forms.Label
	$lblKeep.Text = "How many backups to keep:"
	$lblKeep.Location = [Drawing.Point]::new(25, 105)
	$lblKeep.Size = [Drawing.Size]::new(200, 20)
	$grpBackup.Controls.Add($lblKeep)
	$numKeep = New-Object Windows.Forms.NumericUpDown
	$numKeep.Minimum = 1
	$numKeep.Maximum = 99
	$numKeep.Value = 1
	$numKeep.Location = [Drawing.Point]::new(260, 103)
	$numKeep.Size = [Drawing.Size]::new(80, 24)
	$grpBackup.Controls.Add($numKeep)
	
	# -- Task User/Password GroupBox --
	$grpUser = New-Object Windows.Forms.GroupBox
	$grpUser.Text = "Scheduled Task Credentials"
	$grpUser.Location = [Drawing.Point]::new(20, 165)
	$grpUser.Size = [Drawing.Size]::new(450, 100)
	$form.Controls.Add($grpUser)
	
	# Username (default Administrator)
	$lblUser = New-Object Windows.Forms.Label
	$lblUser.Text = "Task user (default: Administrator):"
	$lblUser.Location = [Drawing.Point]::new(25, 35)
	$lblUser.Size = [Drawing.Size]::new(200, 18)
	$grpUser.Controls.Add($lblUser)
	$txtUser = New-Object Windows.Forms.TextBox
	$txtUser.Text = "Administrator"
	$txtUser.Location = [Drawing.Point]::new(260, 33)
	$txtUser.Size = [Drawing.Size]::new(130, 24)
	$grpUser.Controls.Add($txtUser)
	
	# Password (masked)
	$lblPwd = New-Object Windows.Forms.Label
	$lblPwd.Text = "Task password (blank = only when user logged in):"
	$lblPwd.Location = [Drawing.Point]::new(25, 65)
	$lblPwd.Size = [Drawing.Size]::new(270, 18)
	$grpUser.Controls.Add($lblPwd)
	$txtPwd = New-Object Windows.Forms.TextBox
	$txtPwd.Text = ""
	$txtPwd.Location = [Drawing.Point]::new(305, 63)
	$txtPwd.Size = [Drawing.Size]::new(120, 24)
	$txtPwd.UseSystemPasswordChar = $true
	$grpUser.Controls.Add($txtPwd)
	
	# OK/Cancel
	$btnOK = New-Object Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = [Drawing.Point]::new(120, 285)
	$btnOK.Size = [Drawing.Size]::new(110, 38)
	$btnOK.Font = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = [Drawing.Point]::new(260, 285)
	$btnCancel.Size = [Drawing.Size]::new(110, 38)
	$btnCancel.Font = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	# Show dialog and handle input
	$result = $form.ShowDialog()
	if ($result -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Lane backup scheduling cancelled by user." "yellow"
		return
	}
	$timeInput = $txtTime.Text
	$freqInput = [int]$numFreq.Value
	$keepInput = [int]$numKeep.Value
	$taskUser = $txtUser.Text.Trim()
	if (-not $taskUser) { $taskUser = "Administrator" }
	$taskPwd = $txtPwd.Text
	
	# Validate time
	if ($timeInput -notmatch '^\d{2}:\d{2}$')
	{
		[System.Windows.Forms.MessageBox]::Show("Time must be in HH:mm (24h) format.", "Input Error", 0, 16)
		Write_Log "Invalid time format entered." "red"
		return
	}
	$hours, $minutes = $timeInput.Split(":")
	if ([int]$hours -gt 23 -or [int]$minutes -gt 59)
	{
		[System.Windows.Forms.MessageBox]::Show("Invalid time entered.", "Input Error", 0, 16)
		Write_Log "Invalid time (out of range)." "red"
		return
	}
	
	# 4. Loop over each lane and schedule backup
	foreach ($lane in $goodLanes)
	{
		$laneInfo = $script:FunctionResults['LaneDatabaseInfo'][$lane]
		if (-not $laneInfo)
		{
			$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $lane
			if ($laneInfo) { $script:FunctionResults['LaneDatabaseInfo'][$lane] = $laneInfo }
			else
			{
				Write_Log "Could not get DB info for lane $lane. Skipping." "yellow"
				continue
			}
		}
		$machine = $laneInfo.MachineName
		$dbName = $laneInfo.DBName
		
		# Determine protocol and sqlcmd -S string
		$protocol = $script:LaneProtocols[$lane]
		switch ($protocol)
		{
			"TCP"         { $sqlcmdS = $laneInfo.TcpServer }
			"Named Pipes" { $sqlcmdS = $laneInfo.NamedPipes }
			default       { Write_Log "Unknown/unsupported protocol for $machine, skipping." "yellow"; continue }
		}
		
		# Set backup and script directories (per-lane)
		$backupDir = $script:BackupRoot
		$scriptsDir = $script:ScriptsFolder
		if (-not $backupDir) { $backupDir = $script:BackupRoot }
		if (-not $scriptsDir) { $scriptsDir = $script:ScriptsFolder }
		$machineFolder = Join-Path $backupDir $machine
		
		# ----> KEY: name backup script exactly as Run_${MachineName}_DB_Backup.ps1
		$scriptName = "Run_${machine}_DB_Backup.ps1"
		$backupScriptPath = Join-Path $scriptsDir $scriptName
		
		# Compose the backup script for this lane (matches server style)
		$backupScript = @"
# Auto-generated by Schedule_${machine}_DB_Backup (Alex_C.T)
`$ErrorActionPreference = 'Stop'
`$BackupRoot = `"$machineFolder`"
`$Database = `"$dbName`"
`$SqlcmdS = `"$sqlcmdS`"
`$RemoteBackupFolder = "\\$machine\C$\Tecnica_Systems\Alex_C.T\Backups\$machine"
`$MaxBackups = $keepInput
`$MaxRemoteBackups = 3

# Compose backup file name
`$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
`$bakName = ("{0}_{1}.bak" -f `$Database, `$timestamp)
`$remoteBakFile = Join-Path `$RemoteBackupFolder `$bakName
`$localBakFile  = Join-Path `$BackupRoot `$bakName

# Ensure backup root exists (server)
if (-not (Test-Path `$BackupRoot)) { New-Item -Path `$BackupRoot -ItemType Directory -Force | Out-Null }
# Ensure backup root exists (remote/lane)
if (-not (Test-Path `$RemoteBackupFolder)) { New-Item -Path `$RemoteBackupFolder -ItemType Directory -Force | Out-Null }

# Clean up old remote lane backups (always keep 3)
try {
    `$remoteBakFiles = Get-ChildItem -Path `$RemoteBackupFolder -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
    while (`$remoteBakFiles.Count -ge `$MaxRemoteBackups) {
        Remove-Item -Path `$remoteBakFiles[0].FullName -Force
        `$remoteBakFiles = Get-ChildItem -Path `$RemoteBackupFolder -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
    }
} catch { }

# Clean up old local server backups (per user)
try {
    `$bakFiles = Get-ChildItem -Path `$BackupRoot -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
    while (`$bakFiles.Count -ge `$MaxBackups) {
        Remove-Item -Path `$bakFiles[0].FullName -Force
        `$bakFiles = Get-ChildItem -Path `$BackupRoot -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
    }
} catch { }

# Run the backup on the lane via UNC SQL backup path
`$tsql = "BACKUP DATABASE [`$Database] TO DISK = N'`$remoteBakFile' WITH NOFORMAT, NOINIT, NAME = N'`$Database-FullBackup-`$timestamp', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
& sqlcmd -S `"$SqlcmdS`" -Q `$tsql -b
`$exitCode = `$LASTEXITCODE

if (`$exitCode -eq 0) {
    # Wait for backup to finish (file to appear and not be locked)
    `$waitSec = 0
    while (-not (Test-Path `$remoteBakFile) -and `$waitSec -lt 120) {
        Start-Sleep -Seconds 1; `$waitSec++
    }
    # Optionally, wait until file is not locked (ready for copy)
    `$ready = `$false; `$tries = 0
    while (-not `$ready -and `$tries -lt 60) {
        try {
            `$s = [System.IO.File]::Open(`$remoteBakFile, 'Open', 'Read', 'None')
            `$s.Close(); `$ready = `$true
        } catch { Start-Sleep -Milliseconds 500; `$tries++ }
    }

    if ((Test-Path `$remoteBakFile) -and `$ready) {
        try {
            Copy-Item -Path `$remoteBakFile -Destination `$localBakFile -Force
            "[`$(Get-Date)] Backup complete: `$localBakFile" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
        } catch {
            "[`$(Get-Date)] Backup succeeded but copy failed: $($_.Exception.Message)" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
        }
    } else {
        "[`$(Get-Date)] Backup file not found or not ready for copy after backup." | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
    }
} else {
    # Remove failed/partial file from lane backup folder if exists
    if (Test-Path `$remoteBakFile) { Remove-Item -Path `$remoteBakFile -Force }
    "[`$(Get-Date)] Backup FAILED for [`$Database] on [`$SqlcmdS] (exit code: `$exitCode)" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
}
"@
		
		# Write the script to scripts folder under machine folder
		$parentDir = Split-Path $backupScriptPath -Parent
		if (-not (Test-Path $parentDir)) { New-Item -Path $parentDir -ItemType Directory -Force | Out-Null }
		Set-Content -Path $backupScriptPath -Value $backupScript -Encoding UTF8
		
		# Build Task Scheduler arguments
		$hour = "{0:D2}" -f [int]$hours
		$min = "{0:D2}" -f [int]$minutes
		$startTime = "$hour`:$min"
		$freqArg = "/SC DAILY"
		if ($freqInput -gt 1) { $freqArg = "/SC DAILY /MO $freqInput" }
		$taskName = "${machine}_DB_Backup_Schedule"
		$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"`"$backupScriptPath`"`""
		
		# Task Scheduler command: Default is Administrator, with /IT if no password (interactive, user logged in required)
		if ($taskUser -eq "" -or $taskUser -eq "SYSTEM")
		{
			# Run as SYSTEM
			$cmd = "schtasks.exe /Create /RU SYSTEM /RL HIGHEST /TN `"$taskName`" /TR `"$action`" $freqArg /ST $startTime /F"
		}
		elseif ($taskPwd -eq "")
		{
			# Run as Administrator (or user) INTERACTIVE only (must be logged in), use /IT
			$cmd = "schtasks.exe /Create /RU `"$taskUser`" /IT /RL HIGHEST /TN `"$taskName`" /TR `"$action`" $freqArg /ST $startTime /F"
		}
		else
		{
			# Run as user with password (runs in background even if user not logged in)
			$cmd = "schtasks.exe /Create /RU `"$taskUser`" /RP `"$taskPwd`" /RL HIGHEST /TN `"$taskName`" /TR `"$action`" $freqArg /ST $startTime /F"
		}
		
		Write_Log "Scheduling Lane $lane ($machine): $cmd" "cyan"
		Invoke-Expression $cmd
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Backup task scheduled for $machine ($dbName) at $startTime every $freqInput day(s)." "green"
		}
		else
		{
			Write_Log "Failed to schedule backup for $machine." "red"
		}
	}
	
	Write_Log "`r`n==================== Schedule_LaneDB_Backup Completed ====================" "blue"
}

# ===================================================================================================
#                               FUNCTION: Update_ScaleConfig_And_DB
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user to choose the weighted item marking logic (SCL or POS).
#   Updates all ScaleCommApp XML config files (ProdType_F272) robustly using XML.
#   If the setting doesn't exist, it is inserted; if it exists, it is updated.
#   Runs the correct SQL update/merge using your item key (default F01) for SCL_TAB and POS_TAB.
#   When reverting to Default/POS mode, clears only F272 in SCL_TAB:
#     - If the F272 field allows NULL, sets F272 to NULL.
#     - If the F272 field does not allow NULL, sets F272 to a blank string ('').
#     - Fields like F.1000 (NOT NULL) are **not** touched or updated in any way.
#   Uses Write_Log for all progress, result, and error reporting.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   Folder                  Directory containing ScaleCommApp config files.       (default: C:\ScaleCommApp)
#   MinItem                 Lower bound of item code range.                       (default: 0020000100000)
#   MaxItem                 Upper bound of item code range.                       (default: 0029999900000)
#   POS_Field               Field in POS_TAB for weighted flag.                   (default: F82)
#   POS_Value               Value in POS_TAB.$POS_Field to trigger update.        (default: 0)
#   SCL_Field               Field in SCL_TAB to set (usually F272).               (default: F272)
#   SCL_Value               Value to set in SCL_TAB when in SCL mode.             (default: 3)
#   SCL_Clear_As_Blank      Set to $true to use blank string ('') instead of NULL for F272 when reverting to POS mode.
#   ItemKey                 Column name for unique item code in both tables.      (default: F01)
# ---------------------------------------------------------------------------------------------------

function Update_ScaleConfig_And_DB
{
	Write_Log "`r`n==================== Starting Update_ScaleConfig_And_DB Function ====================`r`n" "blue"
	
	# ---- Set all defaults up front (edit as needed) ----
	$Folder = 'C:\ScaleCommApp'
	$MinItem = '0020000100000'
	$MaxItem = '0029999900000'
	$POS_Field = 'F82'
	$POS_Value = 0
	$SCL_Field = 'F272'
	$SCL_Value = 3
	$ItemKey = 'F01'
	
	# ---------------------------------------------------------------------------
	# Prompt the user for which mode to use (GUI)
	# ---------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Weighted Item Marking Method"
	$form.Size = New-Object System.Drawing.Size(420, 180)
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
	$radioPOS.Text = "Default / Old Way (POS.F82, config to POS, F272 cleared)"
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
	
	# ---------------------------------------------------------------------------
	# XML Config Update (robust, always update or insert key)
	# ---------------------------------------------------------------------------
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
	
	# ---------------------------------------------------------------------------
	# SQL Logic: run the appropriate update/merge (ONLY F272 IS CLEARED)
	# ---------------------------------------------------------------------------
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
    SELECT
        $ItemKey,
        F1000,
        TRY_CAST(FLOOR(CAST(SUBSTRING($ItemKey, 4, LEN($ItemKey) - 8) AS FLOAT)) AS INT) AS F267,
        $POS_Field
    FROM POS_TAB
    WHERE $ItemKey BETWEEN '$MinItem' AND '$MaxItem'
) AS Source
ON Target.$ItemKey = Source.$ItemKey

-- Update F272 only when it's NOT already one of the preserved values
WHEN MATCHED AND (Target.$SCL_Field NOT IN (0,1,3,4,9,10) OR Target.$SCL_Field IS NULL) THEN
    UPDATE SET
        $SCL_Field = CASE WHEN Source.$POS_Field = 1 THEN 0 ELSE $SCL_Value END,
        F1000      = Source.F1000,
        F267       = Source.F267

-- Still refresh F1000/F267, but leave F272 as-is if it's already a preserved value
WHEN MATCHED AND Target.$SCL_Field IN (0,1,3,4,9,10) THEN
    UPDATE SET
        F1000 = Source.F1000,
        F267  = Source.F267

-- New rows: insert with computed F272
WHEN NOT MATCHED BY TARGET THEN
    INSERT ($ItemKey, F1000, $SCL_Field, F267)
    VALUES (
        Source.$ItemKey,
        Source.F1000,
        CASE WHEN Source.$POS_Field = 1 THEN 0 ELSE $SCL_Value END,
        Source.F267
    );

SELECT @@ROWCOUNT AS RowsAffected;
"@
	}
	else
	{
		$SqlQuery = @"
UPDATE SCL_TAB
SET $SCL_Field = NULL
WHERE $ItemKey BETWEEN '$MinItem' AND '$MaxItem'
  AND $SCL_Field IS NOT NULL
  AND $SCL_Field NOT IN (0,1,3,4,9,10);

SELECT @@ROWCOUNT AS RowsAffected;
"@
	}
	try
	{
		if ($supportsConnectionString)
		{
			$result = & $InvokeSqlCmd -ConnectionString $ConnectionString -Query $SqlQuery -ErrorAction Stop -QueryTimeout 0
		}
		else
		{
			$result = & $InvokeSqlCmd -ServerInstance $server -Database $dbName -Query $SqlQuery -ErrorAction Stop -QueryTimeout 0
		}
		$rowsAffected = if ($result -and $result.RowsAffected) { $result.RowsAffected }
		else { 0 }
		Write_Log "Database updated for $Mode mode. ($rowsAffected) items changed." "green"
	}
	catch
	{
		Write_Log "Error executing SQL update for $Mode mode: $_" "red"
	}
	Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
}

# ===================================================================================================
# FUNCTION: Configure_Scale_Sub-Departments
# ---------------------------------------------------------------------------------------------------
# UI/UX:
#   - Professional WinForms dialog, Segoe UI, group box, descriptive subtitles.
#   - Fixed, DPI-aware size to avoid clipping (no PreferredSize math).
#   - ASCII-only labels (FAM > RPC > CAT).
# Behavior:
#   - Auto picks first EMPTY table in order: FAM_TAB > RPC_TAB > CAT_TAB; shows which was used.
#   - Individual choices prompt to Back Out or Overwrite (exact 7 rows).
#   - Updates SdpDefault to OBJ.F16 / OBJ.F18 / OBJ.F17 or restores to "1" (no DB change).
#   - Schema-aware, transactional SQL; PS 5.1 safe (no ternary, no ??, no nested funcs).
# Logging:
#   - Start banner (per your requirement) and the requested completion banner at ALL exits.
# ===================================================================================================

function Configure_Scale_Sub-Departments
{
	[CmdletBinding()]
	param (
		[string]$ServerInstance,
		[string]$Database,
		[string]$ConfigFolder = 'C:\ScaleCommApp'
	)
	
	# ---- REQUIRED START BANNER (keep exactly as requested earlier) ----
	Write_Log "`r`n==================== Starting Update_ScaleConfig_And_DB Function ====================`r`n" "blue"
	
	# -------------------- Resolve Server/DB (params > cache > defaults) --------------------
	try
	{
		if (-not $ServerInstance -or $ServerInstance.Trim() -eq '')
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBSERVER') -and $script:FunctionResults['DBSERVER'])
			{
				$ServerInstance = $script:FunctionResults['DBSERVER']
			}
			else
			{
				$ServerInstance = 'localhost\SQLEXPRESS'
			}
		}
		if (-not $Database -or $Database.Trim() -eq '')
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBNAME') -and $script:FunctionResults['DBNAME'])
			{
				$Database = $script:FunctionResults['DBNAME']
			}
			else
			{
				$Database = 'STORESQL'
			}
		}
		Write_Log ("Target SQL: " + $ServerInstance + " | DB: " + $Database) "cyan"
	}
	catch
	{
		Write_Log ("Failed resolving server/database: " + $_.Exception.Message) "red"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# -------------------- Ensure Invoke-Sqlcmd is available --------------------
	try
	{
		$moduleLoaded = $false
		if (Get-Module -ListAvailable -Name SqlServer | Select-Object -First 1)
		{
			Import-Module SqlServer -ErrorAction Stop
			$moduleLoaded = $true
		}
		elseif (Get-Module -ListAvailable -Name SQLPS | Select-Object -First 1)
		{
			Import-Module SQLPS -DisableNameChecking -ErrorAction Stop
			$moduleLoaded = $true
		}
		else
		{
			Import-Module SqlServer -ErrorAction SilentlyContinue | Out-Null
			Import-Module SQLPS -DisableNameChecking -ErrorAction SilentlyContinue | Out-Null
		}
		if (-not $moduleLoaded)
		{
			Write_Log "Warning: SqlServer/SQLPS module not found; Invoke-Sqlcmd may be unavailable." "yellow"
		}
	}
	catch
	{
		Write_Log ("Module import error: " + $_.Exception.Message) "red"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# ==============================
	# SQL templates (literal strings)
	# ==============================
	$Tsql_Check_For = @"
SET NOCOUNT ON;
DECLARE @Schema sysname = NULL, @Cnt bigint = NULL, @Tbl sysname = N'%TABLE%';
SELECT TOP (1) @Schema = s.name
FROM sys.tables AS t
JOIN sys.schemas AS s ON s.schema_id = t.schema_id
WHERE t.name = @Tbl;
IF @Schema IS NOT NULL
BEGIN
    DECLARE @sql nvarchar(max) =
        N'SELECT @c = COUNT_BIG(1) FROM ' + QUOTENAME(@Schema) + N'.' + QUOTENAME(@Tbl) + N';';
    EXEC sp_executesql @sql, N'@c bigint OUTPUT', @c = @Cnt OUTPUT;
END
SELECT @Schema AS SchemaName, @Cnt AS RowCnt;
"@
	
	$Tsql_Merge_FAM_Keep = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.FAM_TAB AS T
USING src S ON T.F16 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1040=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F16,F49,F1040,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL);
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_FAM_Overwrite = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.FAM_TAB AS T
USING src S ON T.F16 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1040=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F16,F49,F1040,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL)
WHEN NOT MATCHED BY SOURCE THEN
 DELETE;
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_RPC_Keep = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.RPC_TAB AS T
USING src S ON T.F18 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1024=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F18,F49,F1024,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL);
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_RPC_Overwrite = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.RPC_TAB AS T
USING src S ON T.F18 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1024=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F18,F49,F1024,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL)
WHEN NOT MATCHED BY SOURCE THEN
 DELETE;
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_CAT_Keep = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (
    SELECT vID, vName FROM (VALUES
        (1,N'Produce'),
        (2,N'Deli'),
        (3,N'Meat'),
        (4,N'Seafood'),
        (5,N'Cafeteria'),
        (6,N'Hot Foods'),
        (7,N'Bakery')
    ) AS x(vID,vName)
)
MERGE /*+ HOLDLOCK */ %SCHEMA%.CAT_TAB AS T
USING src AS S
    ON T.F17 = S.vID

WHEN MATCHED THEN
    UPDATE SET
        T.F49   = NULL,
        T.F1023 = S.vName,
        T.F1123 = NULL,
        T.F1124 = NULL,
        T.F1168 = NULL,
        T.F1943 = NULL,
        T.F1944 = NULL,
        T.F1945 = NULL,
        T.F1946 = NULL,
        T.F1947 = NULL,
        T.F1965 = NULL,
        T.F1966 = NULL,
        T.F1967 = NULL,
        T.F2653 = NULL,
        T.F2654 = NULL

WHEN NOT MATCHED BY TARGET THEN
    INSERT (
        F17,   F49,  F1023,  F1123, F1124, F1168,
        F1943, F1944, F1945, F1946, F1947,
        F1965, F1966, F1967,
        F2653, F2654
    )
    VALUES (
        S.vID,        -- F17
        NULL,         -- F49
        S.vName,      -- F1023
        NULL,         -- F1123
        NULL,         -- F1124
        NULL,         -- F1168
        NULL,         -- F1943
        NULL,         -- F1944
        NULL,         -- F1945
        NULL,         -- F1946
        NULL,         -- F1947
        NULL,         -- F1965
        NULL,         -- F1966
        NULL,         -- F1967
        NULL,         -- F2653
        NULL          -- F2654
    );

COMMIT;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK;
    THROW;
END CATCH;
"@
	$Tsql_Merge_CAT_Overwrite = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (
    SELECT vID, vName FROM (VALUES
        (1,N'Produce'),
        (2,N'Deli'),
        (3,N'Meat'),
        (4,N'Seafood'),
        (5,N'Cafeteria'),
        (6,N'Hot Foods'),
        (7,N'Bakery')
    ) AS x(vID,vName)
)
MERGE /*+ HOLDLOCK */ %SCHEMA%.CAT_TAB AS T
USING src AS S
    ON T.F17 = S.vID

WHEN MATCHED THEN
    UPDATE SET
        T.F49   = NULL,
        T.F1023 = S.vName,
        T.F1123 = NULL,
        T.F1124 = NULL,
        T.F1168 = NULL,
        T.F1943 = NULL,
        T.F1944 = NULL,
        T.F1945 = NULL,
        T.F1946 = NULL,
        T.F1947 = NULL,
        T.F1965 = NULL,
        T.F1966 = NULL,
        T.F1967 = NULL,
        T.F2653 = NULL,
        T.F2654 = NULL

WHEN NOT MATCHED BY TARGET THEN
    INSERT (
        F17,   F49,  F1023,  F1123, F1124, F1168,
        F1943, F1944, F1945, F1946, F1947,
        F1965, F1966, F1967,
        F2653, F2654
    )
    VALUES (
        S.vID,        -- F17
        NULL,         -- F49
        S.vName,      -- F1023
        NULL,         -- F1123
        NULL,         -- F1124
        NULL,         -- F1168
        NULL,         -- F1943
        NULL,         -- F1944
        NULL,         -- F1945
        NULL,         -- F1946
        NULL,         -- F1947
        NULL,         -- F1965
        NULL,         -- F1966
        NULL,         -- F1967
        NULL,         -- F2653
        NULL          -- F2654
    )

WHEN NOT MATCHED BY SOURCE THEN
    DELETE;

COMMIT;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK;
    THROW;
END CATCH;
"@
	# ===================
	# WinForms UI (no PreferredSize math; fixed DPI-aware size to avoid clipping)
	# ===================
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	[void][System.Windows.Forms.Application]::EnableVisualStyles()
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Scale Sub-Departments Configuration"
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.Topmost = $true
	$form.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	
	# ---- FIX: no PreferredSize/ClientSize math. Set roomy, DPI-aware size and lock minimum. ----
	$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
	$form.AutoSize = $false
	$form.Size = New-Object System.Drawing.Size -ArgumentList 700, 400 # width, height
	$form.MinimumSize = $form.Size
	
	# Root layout
	$root = New-Object System.Windows.Forms.TableLayoutPanel
	$root.Dock = 'Fill'
	$root.AutoSize = $true
	$root.AutoSizeMode = 'GrowAndShrink'
	$root.ColumnCount = 1
	$root.RowCount = 5
	$root.Padding = New-Object System.Windows.Forms.Padding(12, 12, 12, 12)
	$root.BackColor = $form.BackColor
	$form.Controls.Add($root)
	
	# Title
	$title = New-Object System.Windows.Forms.Label
	$title.Text = "Choose how to configure Sub-Deparments for the scales"
	$title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11, [System.Drawing.FontStyle]::Bold)
	$title.AutoSize = $true
	$title.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 8)
	$root.Controls.Add($title)
	
	# Group box
	$grp = New-Object System.Windows.Forms.GroupBox
	$grp.Text = "Mode"
	$grp.Dock = 'Top'
	$grp.AutoSize = $true
	$grp.AutoSizeMode = 'GrowAndShrink'
	$grp.Padding = New-Object System.Windows.Forms.Padding(12, 12, 12, 12)
	$root.Controls.Add($grp)
	
	# Grid inside group
	$grid = New-Object System.Windows.Forms.TableLayoutPanel
	$grid.Dock = 'Top'
	$grid.AutoSize = $true
	$grid.AutoSizeMode = 'GrowAndShrink'
	$grid.ColumnCount = 1
	$grid.RowCount = 8
	$grp.Controls.Add($grid)
	
	# Auto radio + desc
	$rbAuto = New-Object System.Windows.Forms.RadioButton
	$rbAuto.Text = "Auto (pick empty table: FAM > RPC > CAT)"
	$rbAuto.AutoSize = $true
	$rbAuto.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$rbAuto.Checked = $true
	$grid.Controls.Add($rbAuto)
	
	$descAuto = New-Object System.Windows.Forms.Label
	$descAuto.Text = "   Recommended. Populates the first EMPTY table and sets SdpDefault accordingly."
	$descAuto.AutoSize = $true
	$descAuto.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descAuto.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descAuto)
	
	# FAM radio + desc
	$rbFAM = New-Object System.Windows.Forms.RadioButton
	$rbFAM.Text = "Use FAM_TAB (Family) (OBJ.F16)"
	$rbFAM.AutoSize = $true
	$rbFAM.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$grid.Controls.Add($rbFAM)
	
	$descFAM = New-Object System.Windows.Forms.Label
	$descFAM.Text = "   Populates FAM_TAB with 7 rows. If it already has rows, you can back out or overwrite."
	$descFAM.AutoSize = $true
	$descFAM.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descFAM.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descFAM)
	
	# RPC radio + desc
	$rbRPC = New-Object System.Windows.Forms.RadioButton
	$rbRPC.Text = "Use RPC_TAB (Report) (OBJ.F18)"
	$rbRPC.AutoSize = $true
	$rbRPC.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$grid.Controls.Add($rbRPC)
	
	$descRPC = New-Object System.Windows.Forms.Label
	$descRPC.Text = "   Populates RPC_TAB with 7 rows. If it already has rows, you can back out or overwrite."
	$descRPC.AutoSize = $true
	$descRPC.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descRPC.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descRPC)
	
	# CAT radio + desc
	$rbCAT = New-Object System.Windows.Forms.RadioButton
	$rbCAT.Text = "Use CAT_TAB (Category) (OBJ.F17)"
	$rbCAT.AutoSize = $true
	$rbCAT.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$grid.Controls.Add($rbCAT)
	
	$descCAT = New-Object System.Windows.Forms.Label
	$descCAT.Text = "   Populates CAT_TAB with 7 rows. If it already has rows, you can back out or overwrite."
	$descCAT.AutoSize = $true
	$descCAT.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descCAT.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descCAT)
	
	# Restore default
	$chkRestore = New-Object System.Windows.Forms.CheckBox
	$chkRestore.Text = "Restore default SdpDefault=1 (no DB changes)"
	$chkRestore.AutoSize = $true
	$chkRestore.Margin = New-Object System.Windows.Forms.Padding(8, 8, 4, 4)
	$root.Controls.Add($chkRestore)
	
	# Footer: server|db
	$status = New-Object System.Windows.Forms.Label
	$status.Text = "Resolved target: " + $ServerInstance + " | " + $Database
	$status.AutoSize = $true
	$status.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$status.Margin = New-Object System.Windows.Forms.Padding(8, 8, 4, 8)
	$root.Controls.Add($status)
	
	# Buttons row
	$btnPanel = New-Object System.Windows.Forms.TableLayoutPanel
	$btnPanel.AutoSize = $true
	$btnPanel.AutoSizeMode = 'GrowAndShrink'
	$btnPanel.ColumnCount = 3
	$btnPanel.RowCount = 1
	$btnPanel.Dock = 'Top'
	$btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	$btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$root.Controls.Add($btnPanel)
	
	$spacer = New-Object System.Windows.Forms.Label
	$spacer.AutoSize = $true
	$btnPanel.Controls.Add($spacer, 0, 0)
	
	$btnRun = New-Object System.Windows.Forms.Button
	$btnRun.Text = "Run"
	$btnRun.AutoSize = $true
	$btnRun.Margin = New-Object System.Windows.Forms.Padding(4, 0, 8, 0)
	$btnPanel.Controls.Add($btnRun, 1, 0)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.AutoSize = $true
	$btnCancel.Margin = New-Object System.Windows.Forms.Padding(4, 0, 8, 0)
	$btnPanel.Controls.Add($btnCancel, 2, 0)
	
	# Default buttons
	$form.AcceptButton = $btnRun
	$form.CancelButton = $btnCancel
	
	# Tooltips
	$tip = New-Object System.Windows.Forms.ToolTip
	$tip.SetToolTip($rbAuto, "Automatically picks the first EMPTY table: FAM > RPC > CAT.")
	$tip.SetToolTip($rbFAM, "Use FAM_TAB. If it already has rows, you can overwrite or back out.")
	$tip.SetToolTip($rbRPC, "Use RPC_TAB. If it already has rows, you can overwrite or back out.")
	$tip.SetToolTip($rbCAT, "Use CAT_TAB. If it already has rows, you can overwrite or back out.")
	$tip.SetToolTip($chkRestore, "Only updates configs to SdpDefault=1. No database changes.")
	
	# Decide action (no nested funcs)
	$script:_Sdp_Action = $null # "AUTO" | "FAM" | "RPC" | "CAT" | "RESTORE"
	$btnRun.Add_Click({
			if ($chkRestore.Checked)
			{
				$script:_Sdp_Action = "RESTORE"
			}
			else
			{
				if ($rbAuto.Checked) { $script:_Sdp_Action = "AUTO" }
				elseif ($rbFAM.Checked) { $script:_Sdp_Action = "FAM" }
				elseif ($rbRPC.Checked) { $script:_Sdp_Action = "RPC" }
				elseif ($rbCAT.Checked) { $script:_Sdp_Action = "CAT" }
			}
			if (-not $script:_Sdp_Action)
			{
				[System.Windows.Forms.MessageBox]::Show("Pick a mode or check 'Restore default'.", "No Action Selected",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})
	$btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
	
	$dr = $form.ShowDialog()
	if ($dr -ne [System.Windows.Forms.DialogResult]::OK -or -not $script:_Sdp_Action)
	{
		Write_Log "User canceled subdepartment configuration." "yellow"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# =========================
	# Behavior: Restore default (no DB)
	# =========================
	if ($script:_Sdp_Action -eq 'RESTORE')
	{
		Write_Log "Restoring SdpDefault=1 in configs (no DB changes)..." "blue"
		$files = @('ScaleManagementApp.exe.config', 'ScaleManagementAppUpdateSpecials.exe.config', 'ScaleManagementAppSetup.exe.config', 'ScaleManagementApp_FastDEPLOY.exe.config')
		$key = 'SdpDefault'
		$value = '1'
		foreach ($file in $files)
		{
			try
			{
				$full = Join-Path $ConfigFolder $file
				if (-not (Test-Path $full)) { Write_Log ("[" + $file + "] Not found in '" + $ConfigFolder + "' - skipped.") "yellow"; continue }
				[xml]$xml = Get-Content -Path $full -Raw
				if (-not $xml.configuration) { $xml.AppendChild($xml.CreateElement('configuration')) | Out-Null }
				if (-not $xml.configuration.appSettings) { $xml.configuration.AppendChild($xml.CreateElement('appSettings')) | Out-Null }
				$settings = $xml.configuration.appSettings
				$existing = $settings.add | Where-Object { $_.key -eq $key }
				if ($existing) { $existing.value = $value; Write_Log ("[" + $file + "] Updated '" + $key + "' to '1'.") "green" }
				else { $add = $xml.CreateElement('add'); $add.SetAttribute('key', $key); $add.SetAttribute('value', $value); $settings.AppendChild($add) | Out-Null; Write_Log ("[" + $file + "] Inserted '" + $key + "' = '1'.") "green" }
				$xml.Save($full)
			}
			catch { Write_Log ("[" + $file + "] XML update error: " + $_.Exception.Message) "red" }
		}
		Write_Log "SdpDefault restored to '1'." "green"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# =========================
	# Table checks (schema + rowcount)
	# =========================
	$famInfo = $null; $rpcInfo = $null; $catInfo = $null
	try
	{
		$famInfo = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query ($Tsql_Check_For.Replace('%TABLE%', 'FAM_TAB')) -ErrorAction Stop | Select-Object -First 1
		$rpcInfo = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query ($Tsql_Check_For.Replace('%TABLE%', 'RPC_TAB')) -ErrorAction Stop | Select-Object -First 1
		$catInfo = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query ($Tsql_Check_For.Replace('%TABLE%', 'CAT_TAB')) -ErrorAction Stop | Select-Object -First 1
	}
	catch
	{
		Write_Log ("Error reading table info: " + $_.Exception.Message) "red"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	$famCount = 0; if ($famInfo -and $famInfo.RowCnt -ne $null) { $famCount = [int64]$famInfo.RowCnt }
	$rpcCount = 0; if ($rpcInfo -and $rpcInfo.RowCnt -ne $null) { $rpcCount = [int64]$rpcInfo.RowCnt }
	$catCount = 0; if ($catInfo -and $catInfo.RowCnt -ne $null) { $catCount = [int64]$catInfo.RowCnt }
	
	# =========================
	# Execute selected action (parse-safe .Replace and clean Query build)
	# =========================
	$chosenShort = $null
	$chosenSchema = $null
	$sdpValue = $null
	$overwrite = $false
	
	if ($script:_Sdp_Action -eq 'AUTO')
	{
		if ($famInfo -and $famInfo.SchemaName -and $famCount -eq 0)
		{
			$chosenShort = 'FAM_TAB'; $chosenSchema = $famInfo.SchemaName; $sdpValue = 'OBJ.F16'
			$q = $Tsql_Merge_FAM_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') # no space before .Replace
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
		}
		elseif ($rpcInfo -and $rpcInfo.SchemaName -and $rpcCount -eq 0)
		{
			$chosenShort = 'RPC_TAB'; $chosenSchema = $rpcInfo.SchemaName; $sdpValue = 'OBJ.F18'
			$q = $Tsql_Merge_RPC_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']')
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
		}
		elseif ($catInfo -and $catInfo.SchemaName -and $catCount -eq 0)
		{
			$chosenShort = 'CAT_TAB'; $chosenSchema = $catInfo.SchemaName; $sdpValue = 'OBJ.F17'
			$q = $Tsql_Merge_CAT_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']')
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
		}
		else
		{
			$why = "No empty candidate found. FAM rows=" + ($famInfo.RowCnt) + ", RPC rows=" + ($rpcInfo.RowCnt) + ", CAT rows=" + ($catInfo.RowCnt) + "."
			[void][System.Windows.Forms.MessageBox]::Show($why, "Auto: No Empty Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			Write_Log ("Auto aborted: " + $why) "yellow"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		Write_Log ("Auto used: [" + $chosenSchema + "]." + $chosenShort) "green"
		[void][System.Windows.Forms.MessageBox]::Show(("Auto used: [" + $chosenSchema + "]." + $chosenShort), "Auto Selection",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
	}
	elseif ($script:_Sdp_Action -eq 'FAM')
	{
		if (-not ($famInfo -and $famInfo.SchemaName))
		{
			[void][System.Windows.Forms.MessageBox]::Show(("FAM_TAB does not exist in " + $Database + "."), "Missing Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			Write_Log "FAM_TAB missing." "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		$chosenShort = 'FAM_TAB'; $chosenSchema = $famInfo.SchemaName; $sdpValue = 'OBJ.F16'
		if ($famCount -gt 0)
		{
			$msg = "FAM_TAB already has " + $famCount + " row(s)." + [Environment]::NewLine + "Overwrite table to the standard 7 rows?"
			$res = [System.Windows.Forms.MessageBox]::Show($msg, "FAM_TAB Occupied",
				[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if ($res -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User backed out of FAM overwrite." "yellow"
				Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
				return
			}
			$overwrite = $true
		}
		try
		{
			if ($overwrite) { $q = $Tsql_Merge_FAM_Overwrite.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			else { $q = $Tsql_Merge_FAM_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
			Write_Log ("FAM_TAB configured. Overwrite=" + $overwrite) "green"
		}
		catch
		{
			Write_Log ("FAM SQL error: " + $_.Exception.Message) "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
	}
	elseif ($script:_Sdp_Action -eq 'RPC')
	{
		if (-not ($rpcInfo -and $rpcInfo.SchemaName))
		{
			[void][System.Windows.Forms.MessageBox]::Show(("RPC_TAB does not exist in " + $Database + "."), "Missing Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			Write_Log "RPC_TAB missing." "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		$chosenShort = 'RPC_TAB'; $chosenSchema = $rpcInfo.SchemaName; $sdpValue = 'OBJ.F18'
		if ($rpcCount -gt 0)
		{
			$msg = "RPC_TAB already has " + $rpcCount + " row(s)." + [Environment]::NewLine + "Overwrite table to the standard 7 rows?"
			$res = [System.Windows.Forms.MessageBox]::Show($msg, "RPC_TAB Occupied",
				[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if ($res -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User backed out of RPC overwrite." "yellow"
				Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
				return
			}
			$overwrite = $true
		}
		try
		{
			if ($overwrite) { $q = $Tsql_Merge_RPC_Overwrite.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			else { $q = $Tsql_Merge_RPC_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
			Write_Log ("RPC_TAB configured. Overwrite=" + $overwrite) "green"
		}
		catch
		{
			Write_Log ("RPC SQL error: " + $_.Exception.Message) "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
	}
	elseif ($script:_Sdp_Action -eq 'CAT')
	{
		if (-not ($catInfo -and $catInfo.SchemaName))
		{
			[void][System.Windows.Forms.MessageBox]::Show(("CAT_TAB does not exist in " + $Database + "."), "Missing Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			Write_Log "CAT_TAB missing." "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		$chosenShort = 'CAT_TAB'; $chosenSchema = $catInfo.SchemaName; $sdpValue = 'OBJ.F17'
		if ($catCount -gt 0)
		{
			$msg = "CAT_TAB already has " + $catCount + " row(s)." + [Environment]::NewLine + "Overwrite table to the standard 7 rows?"
			$res = [System.Windows.Forms.MessageBox]::Show($msg, "CAT_TAB Occupied",
				[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if ($res -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User backed out of CAT overwrite." "yellow"
				Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
				return
			}
			$overwrite = $true
		}
		try
		{
			if ($overwrite) { $q = $Tsql_Merge_CAT_Overwrite.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			else { $q = $Tsql_Merge_CAT_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
			Write_Log ("CAT_TAB configured. Overwrite=" + $overwrite) "green"
		}
		catch
		{
			Write_Log ("CAT SQL error: " + $_.Exception.Message) "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
	}
	
	# =========================
	# Update SdpDefault in config files
	# =========================
	if ($sdpValue)
	{
		Write_Log ("Updating SdpDefault to '" + $sdpValue + "' in configs...") "blue"
		$files = @('ScaleManagementApp.exe.config', 'ScaleManagementAppUpdateSpecials.exe.config', 'ScaleManagementAppSetup.exe.config', 'ScaleManagementApp_FastDEPLOY.exe.config')
		$key = 'SdpDefault'
		foreach ($file in $files)
		{
			try
			{
				$full = Join-Path $ConfigFolder $file
				if (-not (Test-Path $full)) { Write_Log ("[" + $file + "] Not found in '" + $ConfigFolder + "' - skipped.") "yellow"; continue }
				[xml]$xml = Get-Content -Path $full -Raw
				if (-not $xml.configuration) { $xml.AppendChild($xml.CreateElement('configuration')) | Out-Null }
				if (-not $xml.configuration.appSettings) { $xml.configuration.AppendChild($xml.CreateElement('appSettings')) | Out-Null }
				$settings = $xml.configuration.appSettings
				$existing = $settings.add | Where-Object { $_.key -eq $key }
				if ($existing)
				{
					if ($existing.value -ne $sdpValue) { $existing.value = $sdpValue; Write_Log ("[" + $file + "] Updated '" + $key + "' to '" + $sdpValue + "'.") "green" }
					else { Write_Log ("[" + $file + "] '" + $key + "' already '" + $sdpValue + "' - no change.") "gray" }
				}
				else
				{
					$add = $xml.CreateElement('add'); $add.SetAttribute('key', $key); $add.SetAttribute('value', $sdpValue)
					$settings.AppendChild($add) | Out-Null
					Write_Log ("[" + $file + "] Inserted '" + $key + "' = '" + $sdpValue + "'.") "green"
				}
				$xml.Save($full)
			}
			catch { Write_Log ("[" + $file + "] XML update error: " + $_.Exception.Message) "red" }
		}
		Write_Log ("SdpDefault successfully set to '" + $sdpValue + "'.") "green"
	}
	
	# ---- REQUESTED COMPLETION BANNER (exact string) ----
	Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
}

# ===================================================================================================
#                          FUNCTION: Deploy_Scale_Currency_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user to select one or more scales (via Show_Node_Selection_Form).
#   Prompts for the currency symbol to be used in the currency text files (defaults to "$").
#   Copies the required .txt and .properties files (with correct currency) to each selected scale:
#     \\<ScaleIP>\c$\bizstorecard\bizerba\_fileIO\generic_data\in
#   Uses shared mappings and writes all file contents inline (no external dependencies).
#   Reports detailed status per scale.
#   CHANGE: Writes files as UTF-8 **without BOM** to prevent a leading ï»¿ (or similar) before the currency.
# ===================================================================================================

function Deploy_Scale_Currency_Files
{
	Write_Log "`r`n==================== Starting Deploy_Scale_Currency_Files ====================`r`n" "blue"
	
	# ---- Node selection for scales (show tabbed GUI) ----
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Scale" -OnlyBizerbaScales
	if (-not $selection -or -not $selection.Scales -or $selection.Scales.Count -eq 0)
	{
		Write_Log "No scales selected for deployment." "yellow"
		return
	}
	
	# ---- Load Assemblies ----
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# ---- Create Form ----
	$form = New-Object Windows.Forms.Form
	$form.Text = "Set Currency Symbol"
	$form.Size = [System.Drawing.Size]::new(420, 250)
	$form.FormBorderStyle = 'FixedDialog'
	$form.StartPosition = 'CenterScreen'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
	$form.Icon = [System.Drawing.SystemIcons]::Information
	
	# ---- Centered, bold, wrapped label ----
	$lbl = New-Object Windows.Forms.Label
	$lbl.Text = "Please enter the currency symbol to use`nfor all scale price files:"
	$lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
	$lbl.Width = $form.ClientSize.Width - 40
	$lbl.Height = 50
	$lbl.Location = New-Object System.Drawing.Point(20, 20)
	$lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$lbl.AutoSize = $false
	$lbl.ForeColor = [System.Drawing.Color]::FromArgb(33, 33, 33)
	$form.Controls.Add($lbl)
	
	# ---- ComboBox for common symbols ----
	$cmbCurrency = New-Object Windows.Forms.ComboBox
	$cmbCurrency.Items.AddRange(@('$', [char]0x20AC, [char]0x00A3, [char]0x00A5, [char]0x20B9, [char]0x20BD, [char]0x20A9, [char]0x20BA, [char]0x20AA))
	$cmbCurrency.Text = '$'
	$cmbCurrency.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Regular)
	$cmbCurrency.DropDownStyle = 'DropDown'
	$cmbCurrency.MaxLength = 3
	$cmbCurrency.Width = 100
	$cmbCurrency.Height = 30
	$cmbCurrency.Location = [System.Drawing.Point]::new([Math]::Floor(($form.ClientSize.Width - $cmbCurrency.Width)/2), 80)
	$cmbCurrency.FlatStyle = 'Flat'
	$form.Controls.Add($cmbCurrency)
	
	# ---- Tooltip ----
	$toolTip = New-Object System.Windows.Forms.ToolTip
	$toolTip.SetToolTip($cmbCurrency, "Select a common symbol or type a custom one (up to 3 characters).")
	
	# ---- Preview Label ----
	$previewLabel = New-Object Windows.Forms.Label
	$previewLabel.Text = "Preview: $1.99"
	$previewLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Italic)
	$previewLabel.Width = $form.ClientSize.Width - 40
	$previewLabel.Height = 30
	$previewLabel.Location = New-Object System.Drawing.Point(20, 120)
	$previewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$previewLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
	$form.Controls.Add($previewLabel)
	
	# ---- Update preview on text change ----
	$cmbCurrency.Add_TextChanged({
			$symbol = ($this.Text).Trim()
			if ([string]::IsNullOrWhiteSpace($symbol)) { $symbol = '$' }
			$previewLabel.Text = "Preview: $($symbol)1.99"
		})
	$cmbCurrency.Add_Leave({
			if ([string]::IsNullOrWhiteSpace($cmbCurrency.Text)) { $cmbCurrency.Text = '$' }
		})
	$form.Add_Shown({
			if ([string]::IsNullOrWhiteSpace($cmbCurrency.Text)) { $cmbCurrency.Text = '$' }
			$previewLabel.Text = "Preview: $($cmbCurrency.Text.Trim())1.99"
		})
	
	# ---- OK / Cancel ----
	$btnWidth = 100; $btnHeight = 36; $btnSpacing = 24; $btnY = 170
	$totalBtnWidth = $btnWidth * 2 + $btnSpacing
	$startX = [Math]::Floor(($form.ClientSize.Width - $totalBtnWidth)/2)
	
	$btnOK = New-Object Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Size = [System.Drawing.Size]::new($btnWidth, $btnHeight)
	$btnOK.Location = [System.Drawing.Point]::new($startX, $btnY)
	$btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$btnOK.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
	$btnOK.ForeColor = [System.Drawing.Color]::White
	$btnOK.FlatStyle = 'Flat'
	$btnOK.FlatAppearance.BorderSize = 0
	$btnOK.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 173) })
	$btnOK.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204) })
	$btnOK.Add_Click({
			if ([string]::IsNullOrWhiteSpace($cmbCurrency.Text))
			{
				[System.Windows.Forms.MessageBox]::Show("Please enter a currency symbol.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				return
			}
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close()
		})
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Size = [System.Drawing.Size]::new($btnWidth, $btnHeight)
	$btnCancel.Location = [System.Drawing.Point]::new($startX + $btnWidth + $btnSpacing, $btnY)
	$btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
	$btnCancel.BackColor = [System.Drawing.Color]::FromArgb(232, 72, 85)
	$btnCancel.ForeColor = [System.Drawing.Color]::White
	$btnCancel.FlatStyle = 'Flat'
	$btnCancel.FlatAppearance.BorderSize = 0
	$btnCancel.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(196, 61, 72) })
	$btnCancel.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(232, 72, 85) })
	$btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	$form.Add_Resize({
			$lbl.Width = $form.ClientSize.Width - 40
			$cmbCurrency.Location = [System.Drawing.Point]::new([Math]::Floor(($form.ClientSize.Width - $cmbCurrency.Width)/2), 80)
			$previewLabel.Width = $form.ClientSize.Width - 40
			$totalBtnWidth = $btnWidth * 2 + $btnSpacing
			$startX = [Math]::Floor(($form.ClientSize.Width - $totalBtnWidth)/2)
			$btnOK.Location = [System.Drawing.Point]::new($startX, $btnY)
			$btnCancel.Location = [System.Drawing.Point]::new($startX + $btnWidth + $btnSpacing, $btnY)
		})
	
	# ---- Show the form ----
	$res = $form.ShowDialog()
	if ($res -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Cancelled by user." "yellow"
		return
	}
	$currency = $cmbCurrency.Text.Trim()
	if (-not $currency) { $currency = '$' }
	
	# ---- Inline file content templates ----
	# NOTE: These are what get *inlined* by the scale. If a BOM is present at the start of the file,
	#       the BOM shows up right before the first visible token (your currency). Hence No-BOM writes.
	$totalprice_txt = "=$%$ BT20 =`"$currency *#.##`""
	$unitprice_txt = "=$%$ BT10 =`"$currency *#.##`""
	
	$totalprice_properties = @"
<?xml version="1.0" encoding="UTF-8"?>
<properties>
    <source type="FILE">
        <path>C:\bizstorecard\bizerba\_fileIO\generic_data\in\na_f_totalprice.txt</path>
        <result>VALUE</result>
    </source>
</properties>
"@
	
	$unitprice_properties = @"
<?xml version="1.0" encoding="UTF-8"?>
<properties>
    <source type="FILE">
        <path>C:\bizstorecard\bizerba\_fileIO\generic_data\in\na_f_unitprice.txt</path>
        <result>VALUE</result>
    </source>
</properties>
"@
	
	# ---- Push files to each selected scale ----
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$results = @()
	
	foreach ($scaleCode in $selection.Scales)
	{
		$scaleObj = $ScaleCodeToIPInfo[$scaleCode]
		$scaleIP = $scaleObj.FullIP
		$scaleLabel = $scaleObj.ScaleName
		$targetPath = "\\$scaleIP\c$\bizstorecard\bizerba\_fileIO\generic_data\in"
		
		$result = [PSCustomObject]@{
			Scale = "$scaleLabel [$scaleIP]"
			ScaleCode = $scaleCode
			ScaleIP = $scaleIP
			Result = "Success"
			Details = @()
		}
		
		$isAccessible = $false
		try
		{
			if (Test-Path $targetPath -ErrorAction Stop) { $isAccessible = $true }
		}
		catch
		{
			Write_Log "Scale $scaleLabel [$scaleIP] is offline or share inaccessible. Skipping." "yellow"
			$result.Result = "Failed"
			$result.Details = @("Share not reachable")
			$results += ,$result
			continue
		}
		
		if (-not $isAccessible)
		{
			try
			{
				New-Item -Path $targetPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
				$isAccessible = $true
			}
			catch
			{
				Write_Log "Could not create remote folder on $scaleLabel [$scaleIP]. Skipping." "yellow"
				$result.Result = "Failed"
				$result.Details = @("Failed to create share folder")
				$results += ,$result
				continue
			}
		}
		
		# ==========================================================================================
		# CHANGED: Write files as UTF-8 **without BOM** so the first bytes are the actual content.
		#          Using .NET API here because PS5.1 `-Encoding UTF8` includes a BOM.
		# ==========================================================================================
		try
		{
			$utf8NoBom = New-Object System.Text.UTF8Encoding($false) # $false => NO BOM
			
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_totalprice.txt'), $totalprice_txt, $utf8NoBom) # No BOM
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_unitprice.txt'), $unitprice_txt, $utf8NoBom) # No BOM
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_totalprice.properties'), $totalprice_properties, $utf8NoBom) # No BOM
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_unitprice.properties'), $unitprice_properties, $utf8NoBom) # No BOM
			
			# Optional: ensure CRLF line endings (Windows) on the XML (usually fine as-is).
			# If your environment requires explicit CRLF normalization, uncomment:
			# $crlf = "`r`n"
			# [System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_totalprice.properties'),
			#     ($totalprice_properties -replace "`r?`n",$crlf), $utf8NoBom)
			# [System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_unitprice.properties'),
			#     ($unitprice_properties -replace "`r?`n",$crlf), $utf8NoBom)
			
			$result.Details += "Files copied successfully (UTF-8 without BOM)."
			Write_Log "Deployed price files to $($scaleLabel) [$scaleIP] with Currency ($currency)" "green"
		}
		catch
		{
			$result.Result = "Failed"
			$result.Details = @("File write failed (network/permission).")
			Write_Log "Could not deploy price files to $($scaleLabel) [$scaleIP]. File write failed or network/permission denied." "yellow"
		}
		
		$results += ,$result
	}
	
	# ---- Show summary ----
	Write_Log "Deployment summary:" "Magenta"
	foreach ($r in $results)
	{
		$msg = "$($r.Scale): $($r.Result) - $($r.Details -join '; ')"
		if ($r.Result -eq "Success") { Write_Log $msg "green" }
		else { Write_Log $msg "red" }
	}
	
	# ---- Prompt for reboot of only SUCCESSFUL scales ----
	$successScales = $results | Where-Object { $_.Result -eq "Success" -and $_.PSObject.Properties.Match('ScaleCode') } | Select-Object -ExpandProperty ScaleCode
	$successScalesLabels = $results | Where-Object { $_.Result -eq "Success" } | ForEach-Object { $_.Scale }
	
	if ($successScales.Count -gt 0)
	{
		$scaleListText = ($successScalesLabels -join "`n")
		$rebootMsg = "Do you want to reboot the following successfully deployed scales now to apply changes?`n`n$scaleListText"
		$dialogResult = [System.Windows.Forms.MessageBox]::Show($rebootMsg, "Reboot Scales?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
		
		if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			Write_Log "User chose to reboot all successfully deployed scales." "cyan"
			$rebootSelection = [PSCustomObject]@{ Lanes = @(); Scales = $successScales; Backoffices = @() }
			Reboot_Nodes -StoreNumber $StoreNumber -NodeTypes Scale -Selection $rebootSelection
		}
		else
		{
			Write_Log "User chose not to reboot. The following scales will need to be rebooted to apply changes:" "yellow"
			foreach ($s in $successScalesLabels) { Write_Log "$s" "yellow" }
		}
	}
	
	Write_Log "`r`n==================== Deploy_Scale_Currency_Files Completed ====================" "blue"
}

# ===================================================================================================
# 								FUNCTION: Pick_And_Update_IshidaSDPs
# ---------------------------------------------------------------------------------------------------
# Description:
#   One-shot picker + writer for IshidaSDPs based on SDP_TAB.
#   - Uses F04 as SDP number and F1022 as SDP name.
#   - Shows a GUI to pick which SDPs go into the config.
#   - Pre-checks those already present in IshidaSDPs.
#   - Saves back to <add key="IshidaSDPs" value="..."> as a comma-separated list of numbers,
#     or "Null" when none are selected (per request).
#
# Notes:
#   - Handles missing ConnectionString by building one (SQL Auth preferred, else Integrated).
#   - PS 5.1 compatible WinForms UI.
#   - Logging via Write_Log when available; falls back to Write-Host.
#   - Saves UTF-8 without BOM and creates a timestamped .bak next to the config.
# ===================================================================================================

function Pick_And_Update_IshidaSDPs
{
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[string]$ConfigPath,
		[string[]]$ScaleCommRoots = @('C:\ScaleCommApp', 'D:\ScaleCommApp'),
		[switch]$AllMatches,
		[switch]$IncludeInactive,
		# reserved for future use
		[switch]$Force
	)
	
	# ---------------- Logging shim (Write_Log if present, else host) ----------------
	$logInfo = { param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'cyan' }
		else { Write-Host "[INFO] $m" -ForegroundColor Cyan } }
	$logWarn = { param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'yellow' }
		else { Write-Host "[WARN] $m" -ForegroundColor Yellow } }
	$logErr = { param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'red' }
		else { Write-Host "[ERR ] $m" -ForegroundColor Red } }
	$logOk = { param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'green' }
		else { Write-Host "[ OK ] $m" -ForegroundColor Green } }
	$logBlue = { param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'blue' }
		else { Write-Host $m -ForegroundColor Blue } }
	
	& $logBlue "`r`n==================== Starting Pick_And_Update_IshidaSDPs ====================`r`n"
	
	# ---------------- Resolve DB info from your cache ----------------
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	
	if (-not $dbName -or -not $server -or -not $ConnectionString)
	{
		if (Get-Command Get_Store_And_Database_Info -ErrorAction SilentlyContinue)
		{
			& $logWarn "DB cache incomplete; calling Get_Store_And_Database_Info..."
			try { Get_Store_And_Database_Info | Out-Null }
			catch { & $logErr "Failed to resolve DB info: $($_.Exception.Message)" }
			$dbName = $script:FunctionResults['DBNAME']
			$server = $script:FunctionResults['DBSERVER']
			$ConnectionString = $script:FunctionResults['ConnectionString']
			$SqlModuleName = $script:FunctionResults['SqlModuleName']
		}
	}
	
	# ---------------- Build a ConnectionString when missing ----------------
	if (-not $ConnectionString -and $server -and $dbName)
	{
		$sqlUser = $script:FunctionResults['SQL_UserID']
		$sqlPwd = $script:FunctionResults['SQL_UserPwd']
		if ($sqlUser -and $sqlPwd)
		{
			$ConnectionString = "Server=$server;Database=$dbName;User ID=$sqlUser;Password=$sqlPwd;TrustServerCertificate=True;"
			& $logInfo "Built SQL Authentication connection string from cached credentials."
		}
		else
		{
			$ConnectionString = "Server=$server;Database=$dbName;Integrated Security=SSPI;TrustServerCertificate=True;"
			& $logInfo "Built Integrated Security connection string (no SQL creds found)."
		}
		$script:FunctionResults['ConnectionString'] = $ConnectionString
	}
	
	if (-not $server -or -not $dbName)
	{
		& $logErr "DB server or DB name unresolved. Cannot read SDP_TAB."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	# ---------------- Prepare SQL execution (Invoke-Sqlcmd or .NET) ----------------
	$InvokeSqlCmd = $null
	if ($SqlModuleName -and $SqlModuleName -ne 'None')
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch
		{
			& $logWarn "SQL module '$SqlModuleName' could not be loaded. Will use .NET SqlClient fallback."
		}
	}
	else
	{
		& $logWarn "No SQL module name provided. Will use .NET SqlClient fallback."
	}
	$supportsConnectionString = $false
	if ($InvokeSqlCmd) { $supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString' }
	
	# ---------------- Read SDP_TAB (F04 number, F1022 name) ----------------
	$sql = @"
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='SDP_TAB')
    RAISERROR('SDP_TAB not found',16,1);

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SDP_TAB' AND COLUMN_NAME='F04')
    RAISERROR('Column F04 not found in SDP_TAB',16,1);

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SDP_TAB' AND COLUMN_NAME='F1022')
    RAISERROR('Column F1022 not found in SDP_TAB',16,1);

SELECT
    CAST([F04]   AS int)            AS SdpNumber,
    CAST([F1022] AS nvarchar(200))  AS SdpName
FROM SDP_TAB
ORDER BY CAST([F04] AS int);
"@
	
	try
	{
		if ($InvokeSqlCmd)
		{
			if ($supportsConnectionString -and $ConnectionString)
			{
				$rows = & $InvokeSqlCmd -ConnectionString $ConnectionString -Query $sql -QueryTimeout 60 -ErrorAction Stop
			}
			else
			{
				$rows = & $InvokeSqlCmd -ServerInstance $server -Database $dbName -Query $sql -QueryTimeout 60 -ErrorAction Stop
			}
		}
		else
		{
			$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
			if ($ConnectionString) { $csb.ConnectionString = $ConnectionString }
			else { $csb.DataSource = $server; $csb.InitialCatalog = $dbName; $csb.IntegratedSecurity = $true }
			$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $sql
			$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
			$dt = New-Object System.Data.DataTable
			$conn.Open(); [void]$da.Fill($dt); $conn.Close()
			$rows = $dt
		}
	}
	catch
	{
		& $logErr "Failed to read SDP_TAB (F04/F1022): $($_.Exception.Message)"
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	if (-not $rows -or $rows.Rows.Count -eq 0)
	{
		& $logWarn "SDP_TAB returned 0 rows."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	# Normalize for UI
	$sdpList = @()
	foreach ($r in $rows)
	{
		$sdpList += [pscustomobject]@{
			SdpNumber = [int]$r.SdpNumber
			SdpName   = [string]$r.SdpName
		}
	}
	& $logInfo ("Loaded SDPs: " + ($sdpList.Count))
	
	# ---------------- Find target config files ----------------
	$targets = @()
	if ($ConfigPath)
	{
		if (Test-Path -LiteralPath $ConfigPath) { $targets += (Get-Item -LiteralPath $ConfigPath) }
		else { & $logErr "Config path not found: $ConfigPath" }
	}
	else
	{
		foreach ($root in $ScaleCommRoots)
		{
			if (-not (Test-Path -LiteralPath $root)) { continue }
			$found = Get-ChildItem -LiteralPath $root -File -Filter '*.exe.config' -ErrorAction SilentlyContinue
			if ($found) { $targets += $found }
		}
	}
	$targets = $targets | Sort-Object FullName -Unique
	if (-not $AllMatches -and $targets.Count -gt 1)
	{
		& $logWarn ("Multiple configs found; using first. Use -AllMatches to update all." +
			"`n - " + (($targets | ForEach-Object FullName) -join "`n - "))
		$targets = @($targets[0])
	}
	if (-not $targets -or $targets.Count -eq 0)
	{
		& $logErr "No target *.exe.config found."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	# ---------------- Read IshidaSDPs from first config (for pre-checks) ----------------
	$firstConfig = $targets[0].FullName
	[xml]$xmlFirst = Get-Content -LiteralPath $firstConfig -Raw
	$configuration = $xmlFirst.configuration; if (-not $configuration) { $configuration = $xmlFirst.DocumentElement }
	if (-not $configuration) { & $logErr "Invalid config (no <configuration>): $firstConfig"; return }
	$appSettings = $configuration.appSettings
	if (-not $appSettings)
	{
		$appSettings = $xmlFirst.CreateElement('appSettings')
		[void]$configuration.AppendChild($appSettings)
	}
	$keyName = 'IshidaSDPs'
	$node = $appSettings.SelectSingleNode("add[@key='$keyName']")
	$existingRaw = ''
	$existingNums = @()
	if ($node)
	{
		$existingRaw = [string]$node.GetAttribute('value')
		foreach ($part in ($existingRaw -split ',')) { $p = $part.Trim(); if ($p -match '^\s*(\d+)') { $existingNums += [int]$Matches[1] } }
	}
	& $logInfo ("Current IshidaSDPs in first config: " + ($existingRaw -replace '\s+', ' '))
	
	# ---------- build the "top bar" title text with the already-checked list ----------
	# CHANGE: Put already-in-config SDPs on the form title (top bar). Truncate if too long.
	$existingListText = ($existingNums | Sort-Object -Unique) -join ','
	if ([string]::IsNullOrWhiteSpace($existingListText)) { $existingListText = 'None' }
	$existingShort = $existingListText
	if ($existingShort.Length -gt 90) { $existingShort = $existingShort.Substring(0, 90) + '... (' + $existingNums.Count + ' total)' } # CHANGE
	
	# ---------------- Build WinForms picker ----------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Ishida SDPs  |  Already in config: $existingShort" # CHANGE: show on top bar
	$form.StartPosition = 'CenterScreen'
	$form.Size = New-Object System.Drawing.Size(640, 640)
	$form.MinimizeBox = $false
	$form.MaximizeBox = $false
	$form.TopMost = $true
	
	$search = New-Object System.Windows.Forms.TextBox
	$search.Location = New-Object System.Drawing.Point(10, 10)
	$search.Size = New-Object System.Drawing.Size(505, 24)
	$form.Controls.Add($search)
	
	$btnAll = New-Object System.Windows.Forms.Button
	$btnAll.Text = "Select All"
	$btnAll.Location = New-Object System.Drawing.Point(520, 10)
	$btnAll.Size = New-Object System.Drawing.Size(100, 28)
	$form.Controls.Add($btnAll)
	
	$btnNone = New-Object System.Windows.Forms.Button
	$btnNone.Text = "Clear"
	$btnNone.Location = New-Object System.Drawing.Point(520, 44)
	$btnNone.Size = New-Object System.Drawing.Size(100, 28)
	$form.Controls.Add($btnNone)
	
	$list = New-Object System.Windows.Forms.CheckedListBox
	$list.Location = New-Object System.Drawing.Point(10, 44)
	$list.Size = New-Object System.Drawing.Size(505, 520)
	$list.CheckOnClick = $true
	$form.Controls.Add($list)
	
	$legend = New-Object System.Windows.Forms.Label
	$legend.AutoSize = $true
	$legend.Location = New-Object System.Drawing.Point(10, 570)
	$legend.Text = "Checked = selected. Pre-checked = already in config. Click Save to write."
	$form.Controls.Add($legend)
	
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "Save"
	$btnOK.Location = New-Object System.Drawing.Point(520, 536)
	$btnOK.Size = New-Object System.Drawing.Size(100, 28)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(520, 570)
	$btnCancel.Size = New-Object System.Drawing.Size(100, 28)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	$allEntries = @()
	foreach ($s in $sdpList)
	{
		$display = "{0} - {1}" -f $s.SdpNumber, $s.SdpName
		$null = $list.Items.Add($display)
		$allEntries += [pscustomobject]@{ Display = $display; Number = $s.SdpNumber }
		if ($existingNums -contains [int]$s.SdpNumber) { $list.SetItemChecked(($list.Items.Count - 1), $true) }
	}
	
	$search.Add_TextChanged({
			$needle = $search.Text.Trim()
			$list.Items.Clear()
			if ([string]::IsNullOrWhiteSpace($needle)) { foreach ($e in $allEntries) { [void]$list.Items.Add($e.Display) } }
			else { foreach ($e in $allEntries) { if ($e.Display -like "*$needle*") { [void]$list.Items.Add($e.Display) } } }
			for ($i = 0; $i -lt $list.Items.Count; $i++)
			{
				$text = [string]$list.Items[$i]
				if ($text -match '^\s*(\d+)') { if ($existingNums -contains [int]$Matches[1]) { $list.SetItemChecked($i, $true) } }
			}
		})
	
	$btnAll.Add_Click({ for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) } })
	$btnNone.Add_Click({ for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) } })
	
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		& $logWarn "User cancelled. No changes saved."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	$selectedNums = @()
	foreach ($item in $list.CheckedItems)
	{
		$text = [string]$item
		if ($text -match '^\s*(\d+)') { $selectedNums += [int]$Matches[1] }
	}
	$selectedNums = $selectedNums | Sort-Object -Unique
	
	# ---------- write "Null" when nothing is selected ----------
	# CHANGE: If user leaves everything unchecked, write "Null" to the config value.
	$valueToWrite = if ($selectedNums.Count -gt 0) { ($selectedNums -join ',') }
	else { 'Null' }
	
	& $logInfo ("User selected SDPs -> " + $valueToWrite)
	
	# ---------------- Write back to each target config ----------------
	foreach ($t in $targets)
	{
		$path = $t.FullName
		if (-not $PSCmdlet.ShouldProcess($path, "Set appSettings 'IshidaSDPs'")) { continue }
		
		try
		{
			$item = Get-Item -LiteralPath $path -ErrorAction Stop
			if ($item.Attributes -band [IO.FileAttributes]::ReadOnly)
			{
				if ($Force) { & $logWarn "Removing ReadOnly attribute: $path"; $item.Attributes = ($item.Attributes -bxor [IO.FileAttributes]::ReadOnly) }
				else { & $logErr "File is read-only. Re-run with -Force to clear attribute."; continue }
			}
		}
		catch { & $logErr "Cannot access file attributes for ${path}: $($_.Exception.Message)"; continue }
		
		try
		{
			$dir = Split-Path -LiteralPath $path -Parent
			$base = Split-Path -LiteralPath $path -Leaf
			$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
			$bak = Join-Path $dir "$base.$stamp.bak"
			Copy-Item -LiteralPath $path -Destination $bak -Force
			& $logOk "Backup created: $bak"
		}
		catch { & $logWarn "Backup failed for ${path}: $($_.Exception.Message)" }
		
		try
		{
			[xml]$xml = Get-Content -LiteralPath $path -Raw
			$cfg = $xml.configuration; if (-not $cfg) { $cfg = $xml.DocumentElement }
			if (-not $cfg) { & $logErr "Invalid config: no <configuration> root in $path"; continue }
			
			$as = $cfg.appSettings
			if (-not $as) { $as = $xml.CreateElement('appSettings'); [void]$cfg.AppendChild($as); & $logWarn "Created <appSettings> section in $path" }
			
			$key = 'IshidaSDPs'
			$node = $as.SelectSingleNode("add[@key='$key']")
			if ($node)
			{
				$old = $node.GetAttribute('value')
				if ($old -ne $valueToWrite) { $node.SetAttribute('value', $valueToWrite); & $logOk "Updated ${key}: '$old' -> '$valueToWrite'" }
				else { & $logInfo "No change needed in $path ($key already '$valueToWrite')." }
			}
			else
			{
				$add = $xml.CreateElement('add')
				[void]$add.SetAttribute('key', $key)
				[void]$add.SetAttribute('value', $valueToWrite)
				[void]$as.AppendChild($add)
				& $logOk "Added $key with value '$valueToWrite'"
			}
			
			$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
			$sw = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
			$xml.Save($sw); $sw.Close()
			& $logOk "Saved config: $path"
		}
		catch
		{
			& $logErr "XML update failed for ${path}: $($_.Exception.Message)"
		}
	}
	
	& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
}

# ===================================================================================================
#                                FUNCTION: Sync_Selected_Node_Hosts
# ---------------------------------------------------------------------------------------------------
# Description:
#   Lets the user pick any subset of lanes/backoffices (via node selector), resolves their IP/hostname
#   mapping, ensures the local hosts file is updated (replacing old entries if IPs have changed),
#   and then copies the finished hosts file to all selected nodes.
#   - Always includes the current machine.
#   - Custom node mappings are appended after a blank line (for clarity/compatibility).
#   - Shows a final Write_Log output as a table, sorted by IP, coloring changed rows yellow.
#   - Ensures each IP is mapped to only one hostname (last selected wins if dupe).
#   - Local IP selection is robust: ignores VPN/virtual/tunnel adapters, prefers Ethernet/Wi-Fi with a gateway.
# Usage:
#   Sync_Selected_Node_Hosts -StoreNumber "001"
# Prerequisites:
#   - Retrieve_Nodes and Show_Node_Selection_Form must be available and run.
#   - The script must run as admin for hosts file writes/copies.
# ===================================================================================================

function Sync_Selected_Node_Hosts
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	Write_Log "`r`n==================== Starting Sync_Selected_Node_Hosts ====================`r`n" "blue"
	
	if (-not $script:FunctionResults['LaneNumToMachineName'] -or -not $script:FunctionResults['BackofficeNumToMachineName'])
	{
		Write_Log "Node mappings not found. Please run Retrieve_Nodes first." "red"
		return
	}
	
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes @("Lane", "Backoffice")
	if (-not $selection)
	{
		Write_Log "Node selection cancelled." "yellow"
		return
	}
	$lanes = $selection.Lanes
	$backoffices = $selection.Backoffices
	
	$ipToHost = @{ }
	$changedMappings = @()
	$finalRows = @()
	
	# ---------------- Local machine details (robust local IPv4 selection) ----------------
	$localHostname = $env:COMPUTERNAME
	$localIp = $null
	try
	{
		# Get all UP NICs; exclude obvious VPN/virtual/tunnel/loopback by name/description/type
		$badRegex = '(?i)(vpn|virtual|vmware|hyper[- ]?v|vbox|virtualbox|loopback|hamachi|zerotier|tailscale|anyconnect|forti|juniper|pulse|citrix|openvpn|wireguard|proton|nord|ppp|remote)'
		$nics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
		Where-Object { $_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up }
		
		$candidates = @()
		foreach ($nic in $nics)
		{
			if ($nic.Name -match $badRegex -or $nic.Description -match $badRegex) { continue }
			
			$type = $nic.NetworkInterfaceType.ToString()
			if ($type -notin @('Ethernet', 'Wireless80211')) { continue } # prefer physical/LAN
			
			$props = $nic.GetIPProperties()
			$hasGw = $false
			foreach ($gw in $props.GatewayAddresses)
			{
				if ($gw.Address -and $gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
			}
			
			foreach ($ua in $props.UnicastAddresses)
			{
				if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
				$ip = $ua.Address.ToString()
				if ($ip -eq '0.0.0.0' -or $ip.StartsWith('127.') -or $ip.StartsWith('169.254.')) { continue }
				
				$isPrivate = ($ip -match '^(10\.)|(192\.168\.)|(172\.(1[6-9]|2[0-9]|3[0-1])\.)')
				$score = 0
				if ($type -eq 'Ethernet') { $score += 100 }
				elseif ($type -eq 'Wireless80211') { $score += 80 }
				else { $score += 10 }
				if ($hasGw) { $score += 50 }
				if ($isPrivate) { $score += 20 }
				
				$candidates += [PSCustomObject]@{
					IP		   = $ip
					NicName    = $nic.Name
					Desc	   = $nic.Description
					Type	   = $type
					HasGateway = $hasGw
					IsPrivate  = $isPrivate
					Score	   = $score
				}
			}
		}
		
		if ($candidates.Count -gt 0)
		{
			$chosen = $candidates | Sort-Object Score -Descending | Select-Object -First 1
			$localIp = $chosen.IP
			Write_Log ("Selected local IP: {0} via NIC '{1}' ({2}); HasGateway={3}; Private={4}" -f `
				$chosen.IP, $chosen.NicName, $chosen.Type, $chosen.HasGateway, $chosen.IsPrivate) "gray"
		}
	}
	catch
	{
		# swallow; fallback below
	}
	
	# Fallbacks if the scoring path didn't produce an IP
	if (-not $localIp)
	{
		try { $localIp = ([System.Net.Dns]::GetHostAddresses($localHostname) | Where-Object { $_.AddressFamily -eq 'InterNetwork' -and -not $_.ToString().StartsWith('169.254.') -and -not $_.ToString().StartsWith('127.') } | Select-Object -ExpandProperty IPAddressToString -First 1) }
		catch { }
	}
	if (-not $localIp)
	{
		try { $localIp = (Test-Connection -ComputerName $localHostname -Count 1 -ErrorAction Stop | Select-Object -ExpandProperty IPV4Address | Select-Object -First 1) }
		catch { }
	}
	if (-not $localIp)
	{
		Write_Log "Could not resolve a suitable local IPv4 address for $localHostname. Aborting." "red"
		return
	}
	$ipToHost[$localIp] = $localHostname
	
	# ---------------- Add lanes/backoffices ----------------
	$nodes = @()
	foreach ($lane in $lanes)
	{
		$hostname = $script:FunctionResults['LaneNumToMachineName'][$lane]
		if ($hostname) { $nodes += @{ Hostname = $hostname; LaneNum = $lane; Type = 'Lane' } }
	}
	foreach ($bo in $backoffices)
	{
		$hostname = $script:FunctionResults['BackofficeNumToMachineName'][$bo]
		if ($hostname) { $nodes += @{ Hostname = $hostname; LaneNum = $bo; Type = 'Backoffice' } }
	}
	
	foreach ($n in $nodes)
	{
		if ($n.Hostname -eq $localHostname) { continue } # already added as local
		$ip = $null
		try
		{
			# Prefer DNS A-record style resolution; fall back to ping
			$addr = [System.Net.Dns]::GetHostAddresses($n.Hostname) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1
			if ($addr) { $ip = $addr.IPAddressToString }
		}
		catch { }
		if (-not $ip)
		{
			try { $ip = (Test-Connection -ComputerName $n.Hostname -Count 1 -ErrorAction Stop | Select-Object -ExpandProperty IPV4Address | Select-Object -First 1) }
			catch { }
		}
		if (-not $ip)
		{
			Write_Log "$($n.Type) $($n.LaneNum): Could not resolve IP for $($n.Hostname). Skipping." "red"
			continue
		}
		$ipToHost[$ip] = $n.Hostname
	}
	
	if ($ipToHost.Count -eq 0)
	{
		Write_Log "No valid IP/hostname mappings found for selected nodes." "yellow"
		return
	}
	
	# ---------------- Load previous hosts file mappings ----------------
	$hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
	$oldLines = if (Test-Path $hostsPath) { Get-Content $hostsPath -Raw }
	else { "" }
	$oldMappings = @{ }
	foreach ($line in $oldLines -split "`r?`n")
	{
		if ($line -match '^\s*([0-9\.]+)\s+(\S+)\s*$') { $oldMappings[$matches[2].ToLower()] = $matches[1] }
	}
	
	# Preserve all lines before the first bare "IP hostname" mapping
	$defaultSection = @()
	$customStart = $false
	foreach ($line in $oldLines -split "`r?`n")
	{
		if (-not $customStart -and $line -notmatch '^\s*[0-9\.]+\s+\S+\s*$')
		{
			$defaultSection += $line
		}
		elseif (-not $customStart -and $line -match '^\s*[0-9\.]+\s+\S+\s*$')
		{
			$customStart = $true
		}
	}
	while ($defaultSection.Count -gt 0 -and [string]::IsNullOrWhiteSpace($defaultSection[-1]))
	{
		$defaultSection = $defaultSection[0 .. ($defaultSection.Count - 2)]
	}
	$outputLines = @($defaultSection + '', '') # Always 1 blank line after defaults
	
	# ---------------- Build custom mappings (local first, then others by IP) ----------------
	$orderedMappings = @()
	$orderedMappings += @{ IP = $localIp; Hostname = $localHostname }
	foreach ($ip in ($ipToHost.Keys | Sort-Object))
	{
		if ($ip -eq $localIp) { continue } # already first
		$orderedMappings += @{ IP = $ip; Hostname = $ipToHost[$ip] }
	}
	
	foreach ($entry in $orderedMappings)
	{
		$hn = $entry.Hostname
		$ip = $entry.IP
		$oldIp = $oldMappings[$hn.ToLower()]
		$rowColor = if ($oldIp -and $oldIp -ne $ip) { "yellow" }
		else { "green" }
		$finalRows += @{ IP = $ip; Hostname = $hn; Color = $rowColor }
		$outputLines += "$ip`t$hn"
		if ($rowColor -eq "yellow") { $changedMappings += "$hn ($oldIp => $ip)" }
	}
	
	# ---------------- Write hosts file (local) ----------------
	Set-Content -Path $hostsPath -Value $outputLines -Encoding ascii
	
	# ---------------- Copy hosts file to selected nodes (skip local) ----------------
	foreach ($entry in $orderedMappings)
	{
		$hn = $entry.Hostname
		$ip = $entry.IP
		if ($hn -eq $localHostname) { continue } # Don't network-copy to self
		$targetPath = "\\$hn\C$\Windows\System32\drivers\etc\hosts"
		try
		{
			Copy-Item -Path $hostsPath -Destination $targetPath -Force
			Write_Log "Copied hosts file to $hn [$ip]" "cyan"
		}
		catch
		{
			Write_Log "Failed to copy hosts file to $hn [$ip]: $_" "red"
		}
	}
	
	# ---------------- Table Output ----------------
	Write_Log "`r`nHost file mappings (sorted by IP, local host first):" "blue"
	$maxIpLen = ($finalRows | ForEach-Object { "$($_.IP)".Length } | Measure-Object -Maximum).Maximum
	$maxHostLen = ($finalRows | ForEach-Object { "$($_.Hostname)".Length } | Measure-Object -Maximum).Maximum
	Write_Log ("IP".PadRight($maxIpLen + 2) + "Hostname".PadRight($maxHostLen + 2) + "Changed") "blue"
	foreach ($r in $finalRows)
	{
		$line = "$($r.IP)".PadRight($maxIpLen + 2) + "$($r.Hostname)".PadRight($maxHostLen + 2)
		if ($r.Color -eq "yellow") { $line += "<CHANGED>" }
		Write_Log $line $r.Color
	}
	if ($changedMappings.Count)
	{
		Write_Log "Updated the following mappings:" "yellow"
		foreach ($chg in $changedMappings) { Write_Log "  $chg" "yellow" }
	}
	else
	{
		Write_Log "No mappings changed from previous hosts file." "green"
	}
	Write_Log "`r`n==================== Sync_Selected_Node_Hosts Completed ====================" "blue"
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
#   Extras:
#      -ExcludedNodes         : hide specific nodes (machine names like POS001 and/or 3-digit lane numbers like 001)
#      -SingleLaneOnly        : limit lane selection to exactly one (not required; just a cap)
#      -LaneSelectionLimit <n>: limit lane selection to at most n (not required; just a cap)
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
		[string]$Title = "Select Nodes to Process",
		[string[]]$ExcludedNodes,
		[switch]$SingleLaneOnly,
		# NEW: cap lane selection at 1 (optional)
		[int]$LaneSelectionLimit = 0 # NEW: cap lane selection at N (0 = unlimited)
	)
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Normalize exclusions (names lowercased; lanes as 3-digit)
	$__ex_names = @()
	$__ex_lanes = @()
	if ($ExcludedNodes)
	{
		foreach ($e in $ExcludedNodes)
		{
			if ($null -ne $e -and "$e".Trim())
			{
				$s = "$e".Trim()
				if ($s -match '(\d{1,3})$')
				{
					try { $__ex_lanes += ('{0:D3}' -f ([int]$matches[1])) }
					catch { }
				}
				$__ex_names += $s.ToLower()
			}
		}
		$__ex_names = $__ex_names | Select-Object -Unique
		$__ex_lanes = $__ex_lanes | Select-Object -Unique
	}
	
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
			if ($laneFolders) { $allLanes = $laneFolders | ForEach-Object { $_.Name.Substring($_.Name.Length - 3, 3) } }
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
			
			# Skip excluded by machine name or lane number
			$__skip = $false
			if ($laneMachine -and ($__ex_names -contains $laneMachine.ToLower())) { $__skip = $true }
			elseif ($laneNum -and ($__ex_lanes -contains $laneNum)) { $__skip = $true }
			if ($__skip) { continue }
			
			# Add item
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
		
		# --- Enforce optional selection cap for lanes (single or N) ---
		$__laneCap = if ($SingleLaneOnly) { 1 }
		elseif ($LaneSelectionLimit -gt 0) { [int]$LaneSelectionLimit }
		else { 0 }
		if ($__laneCap -gt 0)
		{
			$clbLanes.Add_ItemCheck({
					# If trying to check and it would exceed the cap, cancel the check.
					if ($_.NewValue -ne [System.Windows.Forms.CheckState]::Checked) { return }
					$idx = $_.Index
					$checkedCount = 0
					for ($i = 0; $i -lt $clbLanes.Items.Count; $i++)
					{
						if ($i -ne $idx -and $clbLanes.GetItemChecked($i)) { $checkedCount++ }
					}
					if (($checkedCount + 1) -gt $__laneCap)
					{
						$_.NewValue = [System.Windows.Forms.CheckState]::Unchecked
						if (-not $script:__laneLimitWarned)
						{
							[System.Windows.Forms.MessageBox]::Show("You can select at most $__laneCap lane(s).", "Selection limit",
								[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
							$script:__laneLimitWarned = $true
						}
					}
				})
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
			if ($OnlyBizerbaScales) { $allScales = $allScales | Where-Object { $_.ScaleBrand -match 'bizerba' } }
		}
		$sortedScales = $allScales | Sort-Object { [int]($_.ScaleCode) }
		
		$uniqueScaleIPs = @{ }
		$dedupedScales = @()
		foreach ($scale in $sortedScales)
		{
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
			# Skip excluded by name/code
			$__skip = $false
			if ($scale.ScaleName -and ($__ex_names -contains ($scale.ScaleName.ToLower()))) { $__skip = $true }
			elseif ($scale.ScaleCode -and ($__ex_lanes -contains ('{0:D3}' -f ([int]$scale.ScaleCode)))) { $__skip = $true }
			if ($__skip) { continue }
			
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
		
		$boDict = if ($script:FunctionResults.ContainsKey('BackofficeNumToMachineName')) { $script:FunctionResults['BackofficeNumToMachineName'] }
		else { @{ } }
		$boLabels = $script:FunctionResults['BackofficeNumToLabel']
		$boPaths = $script:FunctionResults['BackofficeNumToPath']
		
		$seenBonumbers = @{ }
		foreach ($boNumKey in $boDict.Keys | Sort-Object)
		{
			if ($boNumKey -match '(\d{3})')
			{
				$bonum = $matches[1]
				
				# Skip excluded by machine name or BO number
				$__skip = $false
				$machineName = $boDict[$boNumKey]
				if ($machineName -and ($__ex_names -contains $machineName.ToLower())) { $__skip = $true }
				elseif ($__ex_lanes -contains $bonum) { $__skip = $true }
				if ($__skip) { continue }
				
				if (-not $seenBonumbers.ContainsKey($bonum))
				{
					$seenBonumbers[$bonum] = $true
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
			for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { return $false } }
			return $clb.Items.Count -gt 0
		}
		$isAnyChecked = {
			for ($i = 0; $i -lt $clb.Items.Count; $i++) { if ($clb.GetItemChecked($i)) { return $true } }
			return $false
		}
		$btnSelectAll.Add_Click({
				$allChecked = & $isAllChecked
				if (-not $allChecked)
				{
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
					& $setBtnColor $btnSelectAll 2
				}
			})
		$btnDeselectAll.Add_Click({
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
				& $setBtnColor $btnSelectAll 0
				& $setBtnColor $btnDeselectAll 0
			})
		foreach ($event in @("Add_ItemCheck", "Add_MouseUp", "Add_KeyUp"))
		{
			$clb.$event.Invoke({
					Start-Sleep -Milliseconds 30
					$allChecked = & $isAllChecked
					$anyChecked = & $isAnyChecked
					if ($allChecked) { & $setBtnColor $btnSelectAll 2 }
					elseif ($anyChecked) { & $setBtnColor $btnSelectAll 1 }
					else { & $setBtnColor $btnSelectAll 0 }
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
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { $currentTabAllChecked = $false; break } }
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked)
				{
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
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
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
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
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
				else { & $setBtnColor $btnSelectAll 1 }
			})
		$btnDeselectAll.Add_Click({
				$tabName = $tabs.SelectedTab.Text
				$clb = $tabControls[$tabName]
				$tabIndex = $tabs.SelectedIndex
				$noneChecked = $true
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { if ($clb.GetItemChecked($i)) { $noneChecked = $false; break } }
				if ($noneChecked)
				{
					$originalTab = $tabs.SelectedTab
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
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
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
			$clb.Add_MouseUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
			$clb.Add_KeyUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
		}
		$tabs.add_SelectedIndexChanged({
				$tabIndex = $tabs.SelectedIndex
				$clb = $tabControls[$tabs.SelectedTab.Text]
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
				else
				{
					$allChecked = $true
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { $allChecked = $false; break } }
					if ($allChecked -and $clb.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
					else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
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
		Remove-Variable -Name __laneLimitWarned -Scope Script -ErrorAction SilentlyContinue
		return $null
	}
	
	Remove-Variable -Name __laneLimitWarned -Scope Script -ErrorAction SilentlyContinue
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
#                                FUNCTION: Start_Lane_Protocol_Jobs (Runspaces + Polling)
# ---------------------------------------------------------------------------------------------------
# Parallel SQL protocol detector (PS 5.1):
#   - Tries SqlClient to tcp:<lane> then np:<lane>, else "File"
#   - Connect Timeout=1 (fast)
#   - Results cached to: C:\Tecnica_Systems\Alex_C.T\Setup_Files\Protocol_Results.txt
#   - Updates $script:LaneProtocolJobs, $script:LaneProtocols, $script:ProtocolResults
#   - Polling loop keeps WinForms UI responsive (Application.DoEvents)
# ---------------------------------------------------------------------------------------------------.

function Start_Lane_Protocol_Jobs
{
	param (
		[Parameter(Mandatory)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory)]
		[string]$SqlModuleName # kept for signature compatibility, not used inside workers
	)
	
	# -------- Paths / setup ----------
	$script:ProtocolResultsFile = 'C:\Tecnica_Systems\Alex_C.T\Setup_Files\Protocol_Results.txt'
	$resultsDir = [System.IO.Path]::GetDirectoryName($script:ProtocolResultsFile)
	if (-not (Test-Path $resultsDir)) { New-Item -Path $resultsDir -ItemType Directory -Force | Out-Null }
	if (-not (Test-Path $script:ProtocolResultsFile)) { New-Item -Path $script:ProtocolResultsFile -ItemType File -Force | Out-Null }
	
	try { Add-Type -AssemblyName System.Data }
	catch { }
	
	# -------- Globals ----------
	$script:LaneProtocolJobs = @{ }
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
	if (-not $script:ProtocolResults) { $script:ProtocolResults = @() }
	
	# Warm cache from file (if any)
	$existing = (Get-Content -LiteralPath $script:ProtocolResultsFile -ErrorAction SilentlyContinue)
	if ($existing)
	{
		foreach ($line in $existing)
		{
			if ($line -match '^\s*([^,]+),\s*([^,]+)\s*$')
			{
				$lane = $matches[1].Trim()
				$protocol = $matches[2].Trim()
				$script:LaneProtocols[$lane] = $protocol
				$script:ProtocolResults = @($script:ProtocolResults | Where-Object { $_.Lane -ne $lane })
				$script:ProtocolResults += [PSCustomObject]@{ Lane = $lane; Protocol = $protocol }
			}
		}
	}
	
	# -------- RunspacePool ----------
	$minThreads = 1
	$maxThreads = [Math]::Max(8, [Environment]::ProcessorCount * 2)
	$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
	$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool($minThreads, $maxThreads, $iss, $Host)
	try { $pool.ApartmentState = 'MTA' }
	catch { }
	$pool.Open()
	
	# -------- Worker script (single-quoted; no outer $ expansion) ----------
	$worker = @'
param([string]$machine,[string]$lane)

Add-Type -AssemblyName System.Data 2>$null

function Test-SqlConn([string]$dataSource) {
    $cs = 'Data Source=' + $dataSource + ';Initial Catalog=master;Integrated Security=True;Connect Timeout=1'
    $cn = New-Object System.Data.SqlClient.SqlConnection $cs
    try { $cn.Open(); $cn.Close(); return $true }
    catch { return $false }
    finally { $cn.Dispose() }
}

$protocol = 'File'
if (Test-SqlConn ('tcp:' + $machine)) {
    $protocol = 'TCP'
}
elseif (Test-SqlConn ('np:' + $machine)) {
    $protocol = 'Named Pipes'
}

[PSCustomObject]@{ Lane = $lane; Protocol = $protocol }
'@
	
	# -------- Queue workers ----------
	$pending = @{ }
	foreach ($k in $LaneNumToMachineName.Keys)
	{
		$numStr = ($k -replace '[^\d]', '')
		if (-not $numStr) { continue }
		$laneNum = $numStr.PadLeft(3, '0')
		$machine = $LaneNumToMachineName[$k]
		if (-not $machine) { continue }
		
		$ps = [System.Management.Automation.PowerShell]::Create()
		$ps.RunspacePool = $pool
		$null = $ps.AddScript($worker).AddArgument([string]$machine).AddArgument([string]$laneNum)
		$handle = $ps.BeginInvoke()
		$script:LaneProtocolJobs[$laneNum] = @{ PS = $ps; Handle = $handle }
		$pending[$laneNum] = @{ PS = $ps; Handle = $handle }
	}
	
	# -------- Poll until done; update file as results come in ----------
	$lastWriteCount = -1
	while ($pending.Count -gt 0)
	{
		
		$lanesDone = @()
		foreach ($lane in $pending.Keys)
		{
			$task = $pending[$lane]
			$handle = $task.Handle
			if ($handle -and $handle.IsCompleted)
			{
				$ps = $task.PS
				$resultList = $null
				try { $resultList = $ps.EndInvoke($handle) }
				catch { $resultList = $null }
				finally
				{
					try { $ps.Dispose() }
					catch { }
				}
				
				$result = $null
				if ($resultList -and $resultList.Count -ge 1) { $result = $resultList[0] }
				if (-not $result) { $result = [PSCustomObject]@{ Lane = $lane; Protocol = 'File' } }
				
				$rawLane = [string]$result.Lane
				$numericLane = ($rawLane -replace '[^\d]', '').PadLeft(3, '0')
				$protocol = [string]$result.Protocol
				
				# Update caches (multiple keys for convenience)
				$script:LaneProtocols[$numericLane] = $protocol
				$script:LaneProtocols[$rawLane] = $protocol
				if ($script:FunctionResults -and $script:FunctionResults['LaneNumToMachineName'])
				{
					$machineName = $script:FunctionResults['LaneNumToMachineName'][$numericLane]
					if ($machineName)
					{
						$script:LaneProtocols[$machineName] = $protocol
						$script:LaneProtocols[$machineName.ToLower()] = $protocol
					}
				}
				
				$script:ProtocolResults = @($script:ProtocolResults | Where-Object { $_.Lane -ne $rawLane })
				$script:ProtocolResults += [PSCustomObject]@{ Lane = $rawLane; Protocol = $protocol }
				
				$lanesDone += $lane
			}
		}
		
		if ($lanesDone.Count -gt 0)
		{
			foreach ($d in $lanesDone)
			{
				$pending.Remove($d) | Out-Null
				$script:LaneProtocolJobs.Remove($d) | Out-Null
			}
			
			# Write results (sorted by numeric lane)
			$sorted = $script:ProtocolResults | Sort-Object { ($_.Lane -replace '[^\d]', '') -as [int] }
			$lines = foreach ($row in $sorted) { '{0},{1}' -f $row.Lane, $row.Protocol }
			[System.IO.File]::WriteAllLines($script:ProtocolResultsFile, $lines, [System.Text.Encoding]::UTF8)
			$lastWriteCount = $script:ProtocolResults.Count
		}
		
		# Keep UI responsive if WinForms is around
		try
		{
			if ([System.Windows.Forms.Application]::MessageLoop)
			{
				[System.Windows.Forms.Application]::DoEvents()
			}
		}
		catch { }
		
		Start-Sleep -Milliseconds 150
	}
	
	# Done; close pool
	try { $pool.Close(); $pool.Dispose() }
	catch { }
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
	# High DPI awareness (for scaling on modern displays)
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
	
	# Create a timer to refresh the GUI every second
	$refreshTimer = New-Object System.Windows.Forms.Timer
	$refreshTimer.Interval = 1000 # 1 second
	$refreshTimer.add_Tick({
			# Refresh the form to update all controls
			$form.Refresh()
		})
	$refreshTimer.Start()
	
	# Initialize ToolTip with professional delay
	$toolTip = New-Object System.Windows.Forms.ToolTip
	$toolTip.AutoPopDelay = 10000
	$toolTip.InitialDelay = 300
	$toolTip.ReshowDelay = 500
	$toolTip.ShowAlways = $true
	$toolTip.BackColor = [System.Drawing.Color]::LightYellow
	
	# Create the main form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Created by: Alex_C.T   |   Version: $VersionNumber   |   Revised: $VersionDate   |   Powershell Version: $PowerShellVersion"
	$form.Size = New-Object System.Drawing.Size(1006, 570)
	$form.MinimumSize = New-Object System.Drawing.Size(800, 500)
	$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
	$form.BackColor = [System.Drawing.SystemColors]::ControlLight # Light gray background
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9) # Modern font
	
	# -------------------- Idle Close (configurable) + Busy-Safe Watchdog --------------------
	# Config (add these near your other $script: vars)
	if (-not $script:LastActivity) { $script:LastActivity = Get-Date }
	$script:IsBusy = $false
	
	# NEW: reason & ignore list for idle logic
	$script:BusyReason = $null
	# Add any long-lived background jobs here that you want to IGNORE for idle-close:
	# e.g. 'ClearXEFolderJob' or any other job name you see via Get-Job
	$script:IdleIgnoreJobNames = @('ClearXEFolderJob')
	# Optional: hard-close after this many idle minutes even if still "busy" (0 = disabled)
	$script:IdleHardCloseAfterMinutes = 0
	
	# Make sure form sees keystrokes even when a control has focus
	if ($form -and $form -is [System.Windows.Forms.Form])
	{
		try { $form.KeyPreview = $true }
		catch { }
	}
	
	# --- Busy scanner (UI thread) -------------------------------------------------------------
	$BusyTimer = New-Object System.Windows.Forms.Timer
	$BusyTimer.Interval = 800
	$BusyTimer.add_Tick({
			# Reset busy flags each tick
			$script:IsBusy = $false
			$script:BusyReason = $null
			
			try
			{
				# 1) Active lane-protocol runspace tasks?
				if ($script:LaneProtocolJobs)
				{
					foreach ($kv in @($script:LaneProtocolJobs.GetEnumerator()))
					{
						$st = $kv.Value
						if ($st -and $st.Handle -and -not $st.Handle.IsCompleted)
						{
							$script:IsBusy = $true
							$script:BusyReason = "LaneProtocol runspace ($($kv.Key))"
							break
						}
					}
				}
			}
			catch { }
			
			if (-not $script:IsBusy)
			{
				try
				{
					# 2) The lane runspace pool still has in-flight work?
					if ($script:LaneProtocolPool)
					{
						$avail = $script:LaneProtocolPool.GetAvailableRunspaces()
						$max = $script:LaneProtocolPool.MaxRunspaces
						if ($avail -lt $max)
						{
							$script:IsBusy = $true
							$script:BusyReason = "LaneProtocol pool busy ($avail/$max available)"
						}
					}
				}
				catch { }
			}
			
			if (-not $script:IsBusy)
			{
				try
				{
					# 3) Scale credential helper tasks?
					if ($script:ScaleCredTasks)
					{
						foreach ($k in @($script:ScaleCredTasks.Keys))
						{
							$st = $script:ScaleCredTasks[$k]
							if ($st -and $st.Handle -and -not $st.Handle.IsCompleted)
							{
								$script:IsBusy = $true
								$script:BusyReason = "ScaleCred task ($k)"
								break
							}
						}
					}
				}
				catch { }
			}
			
			if (-not $script:IsBusy)
			{
				try
				{
					# 4) The Clear XE job (ignore if whitelisted by name)
					if ($ClearXEJob -and ($ClearXEJob.State -eq 'Running'))
					{
						$name = $null; try { $name = $ClearXEJob.Name }
						catch { }
						if (-not $name) { $name = 'ClearXEFolderJob' } # defensive fallback
						if ($script:IdleIgnoreJobNames -notcontains $name)
						{
							$script:IsBusy = $true
							$script:BusyReason = "Job: $name"
						}
					}
				}
				catch { }
			}
			
			if (-not $script:IsBusy)
			{
				try
				{
					# 5) Any other background jobs (exclude ignored names)
					$jobs = Get-Job -State Running -ErrorAction SilentlyContinue |
					Where-Object { $script:IdleIgnoreJobNames -notcontains $_.Name }
					if ($jobs)
					{
						$script:IsBusy = $true
						$script:BusyReason = "Job: $($jobs[0].Name)"
					}
				}
				catch { }
			}
		})
	$BusyTimer.Start()
	
	# --- Idle watchdog (UI thread) ------------------------------------------------------------
	$script:IdleTimer = New-Object System.Windows.Forms.Timer
	$script:IdleTimer.Interval = 30000 # 30s
	$script:IdleTimer.add_Tick({
			if ($script:IdleMinutesAllowed -le 0) { return }
			try
			{
				$minsIdle = (New-TimeSpan -Start $script:LastActivity -End (Get-Date)).TotalMinutes
				
				# Optionally: hard-close after a longer idle period even if "busy"
				$hardCloseReady = ($script:IdleHardCloseAfterMinutes -gt 0 -and $minsIdle -ge $script:IdleHardCloseAfterMinutes)
				
				if ($minsIdle -ge $script:IdleMinutesAllowed)
				{
					if (-not $script:IsBusy -or $hardCloseReady)
					{
						# if hard-close, explain why we're overriding
						if ($hardCloseReady -and $script:IsBusy -and $script:BusyReason)
						{
							try { Write_Log "Idle $([math]::Round($minsIdle, 1)) min; overriding busy ($script:BusyReason) due to hard-close window." "yellow" }
							catch { }
						}
						else
						{
							try { Write_Log "Idle for $([math]::Round($minsIdle, 1)) minute(s) - closing." "yellow" }
							catch { }
						}
						$script:SuppressClosePrompt = $true
						if ($form -and -not $form.IsDisposed) { $form.Close() }
					}
					else
					{
						# Busy detected at idle threshold; explain who's blocking
						if ($script:BusyReason)
						{
							try { Write_Log "Idle $([math]::Round($minsIdle, 1)) min, but busy: $script:BusyReason (not closing)." "gray" }
							catch { }
						}
						else
						{
							try { Write_Log "Idle $([math]::Round($minsIdle, 1)) min, but busy: <unknown> (not closing)." "gray" }
							catch { }
						}
						# Reset the window so we don't spam checks while still working
						$script:LastActivity = Get-Date
					}
				}
			}
			catch { }
		})
	$script:IdleTimer.Start()
	
	# --- Activity hooks (form + all existing child controls) ---------------------------------
	if ($form -and $form -is [System.Windows.Forms.Form])
	{
		try { $form.add_KeyDown({ $script:LastActivity = Get-Date }) }
		catch { }
		try { $form.add_MouseMove({ $script:LastActivity = Get-Date }) }
		catch { }
		
		try
		{
			# Recursively attach to current controls
			$stack = New-Object System.Collections.Stack
			$stack.Push($form)
			while ($stack.Count -gt 0)
			{
				$parent = [System.Windows.Forms.Control]$stack.Pop()
				foreach ($ctrl in @($parent.Controls))
				{
					try { $ctrl.add_KeyDown({ $script:LastActivity = Get-Date }) }
					catch { }
					try { $ctrl.add_Click({ $script:LastActivity = Get-Date }) }
					catch { }
					try { $ctrl.add_MouseMove({ $script:LastActivity = Get-Date }) }
					catch { }
					if ($ctrl -and $ctrl.HasChildren) { $stack.Push($ctrl) }
				}
			}
		}
		catch { }
	}
	
	# ========================= Banner (header) =========================
	# Static banner background (full width)
	$bannerLabel = New-Object System.Windows.Forms.Label
	$bannerLabel.Text = "PowerShell Script - TBS_Maintenance_Script"
	$bannerLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
	$bannerLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$bannerLabel.Dock = 'Top'
	$bannerLabel.Height = 30 # give the header some breathing room
	$bannerLabel.Cursor = [System.Windows.Forms.Cursors]::Default # parent is NOT clickable
	$form.Controls.Add($bannerLabel)
	
	# Ensure ToolTip exists (shared across the app)
	if (-not $toolTip)
	{
		$toolTip = New-Object System.Windows.Forms.ToolTip
		$toolTip.AutoPopDelay = 8000
		$toolTip.InitialDelay = 300
		$toolTip.ReshowDelay = 100
	}
	
	# Clickable text overlay (only text bounds are clickable)
	$bannerLink = New-Object System.Windows.Forms.LinkLabel
	$bannerLink.Parent = $bannerLabel # render "on top" of the banner area
	$bannerLink.Text = $bannerLabel.Text # mirror the banner's text
	$bannerLink.Font = $bannerLabel.Font # mirror the banner's font
	$bannerLink.AutoSize = $true # <-- hit area == text size
	$bannerLink.BackColor = [System.Drawing.Color]::Transparent
	$bannerLink.Cursor = [System.Windows.Forms.Cursors]::Hand
	
	# Make the entire string a link
	$bannerLink.LinkArea = New-Object System.Windows.Forms.LinkArea(0, $bannerLink.Text.Length)
	$bannerLink.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline # underline on hover
	$bannerLink.LinkColor = $bannerLabel.ForeColor # normal color (usually black)
	$bannerLink.ActiveLinkColor = [System.Drawing.Color]::DodgerBlue # color while clicking
	$bannerLink.VisitedLinkColor = $bannerLabel.ForeColor # keep same color after click
	$bannerLink.DisabledLinkColor = [System.Drawing.Color]::Gray
	
	# Tooltip for the link itself
	$toolTip.SetToolTip($bannerLink, "Open TBS Helpdesk (helpdesk.tecnicasystems.com)")
	
	# Store original link color so we can restore after hover
	$bannerLink.Tag = [pscustomobject]@{ LinkColor = $bannerLink.LinkColor }
	
	# --- Hover color change (this is what you asked for) ---
	$bannerLink.Add_MouseEnter({
			param ($s,
				$e)
			try { $s.LinkColor = [System.Drawing.Color]::DodgerBlue }
			catch { }
		})
	$bannerLink.Add_MouseLeave({
			param ($s,
				$e)
			try
			{
				if ($s.Tag -and $s.Tag.LinkColor) { $s.LinkColor = $s.Tag.LinkColor }
			}
			catch { }
		})
	
	# Center the link text inside the banner area (initial + on resize + on text/font changes)
	$centerBannerLink = {
		# Center horizontally against the FORM (not just the label), so it stays visually centered
		$x = [math]::Max(0, ($form.ClientSize.Width - $bannerLink.Width) / 2)
		# Center vertically inside the banner strip
		$y = [math]::Max(0, ($bannerLabel.Height - $bannerLink.Height) / 2)
		$bannerLink.Location = New-Object System.Drawing.Point([int]$x, [int]$y)
	}
	& $centerBannerLink
	$form.add_Resize({ & $centerBannerLink })
	
	# If someone changes the banner text or font later, keep the link synced and centered
	$bannerLabel.add_TextChanged({
			$bannerLink.Text = $bannerLabel.Text
			$bannerLink.LinkArea = New-Object System.Windows.Forms.LinkArea(0, $bannerLink.Text.Length)
			$bannerLink.AutoSize = $true
			& $centerBannerLink
		})
	$bannerLabel.add_FontChanged({
			$bannerLink.Font = $bannerLabel.Font
			$bannerLink.AutoSize = $true
			& $centerBannerLink
		})
	
	# Click handler (only the text is clickable)
	$bannerLink.Add_LinkClicked({
			$script:LastActivity = Get-Date
			$url = "https://helpdesk.tecnicasystems.com"
			try { Start-Process $url }
			catch
			{
				[System.Windows.Forms.MessageBox]::Show(
					"Couldn't open $url.`r`n$($_.Exception.Message)",
					"Open Link",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
			}
		})
	
	# ========================= Form Closing (X) =========================
	$form.add_FormClosing({
			# Skip confirm if we're closing due to idle timeout
			if ($script:SuppressClosePrompt)
			{
				$script:SuppressClosePrompt = $false
			}
			else
			{
				$confirmResult = [System.Windows.Forms.MessageBox]::Show(
					"Are you sure you want to exit?",
					"Confirm Exit",
					[System.Windows.Forms.MessageBoxButtons]::YesNo,
					[System.Windows.Forms.MessageBoxIcon]::Question
				)
				if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes)
				{
					$_.Cancel = $true
					return
				}
			}
			
			# ---- Timers ----
			try
			{
				if ($script:IdleTimer)
				{
					try { $script:IdleTimer.Stop() }
					catch { }; try { $script:IdleTimer.Dispose() }
					catch { }; $script:IdleTimer = $null
				}
			}
			catch { }
			try
			{
				if ($refreshTimer)
				{
					try { $refreshTimer.Stop() }
					catch { }; try { $refreshTimer.Dispose() }
					catch { }; $refreshTimer = $null
				}
			}
			catch { }
			
			try
			{
				if ($script:protocolTimer)
				{
					try { $script:protocolTimer.Stop() }
					catch { }
					try { $script:protocolTimer.Dispose() }
					catch { }
					$script:protocolTimer = $null
				}
			}
			catch { }
			try
			{
				if ($global:ProtocolFormTimer)
				{
					try { $global:ProtocolFormTimer.Stop() }
					catch { }
					try { $global:ProtocolFormTimer.Dispose() }
					catch { }
					$global:ProtocolFormTimer = $null
				}
			}
			catch { }
			
			# ---- Popup form (optional) ----
			try
			{
				if ($global:ProtocolForm)
				{
					try { $global:ProtocolForm.Hide() }
					catch { }
					try { $global:ProtocolForm.Dispose() }
					catch { }
					$global:ProtocolForm = $null
				}
			}
			catch { }
			
			# ---- Lane protocol runspaces ----
			try
			{
				if ($script:LaneProtocolJobs)
				{
					foreach ($kv in @($script:LaneProtocolJobs.GetEnumerator()))
					{
						$state = $kv.Value
						if ($state)
						{
							try
							{
								if ($state.Handle -and (-not $state.Handle.IsCompleted))
								{
									try
									{
										if ($state.PS) { $null = $state.PS.Stop() }
									}
									catch { }
								}
							}
							catch { }
							try { if ($state.PS) { $state.PS.Dispose() } }
							catch { }
						}
						[void]$script:LaneProtocolJobs.Remove($kv.Key)
					}
				}
			}
			catch { }
			try
			{
				if ($script:LaneProtocolPool)
				{
					try { $script:LaneProtocolPool.Close() }
					catch { }
					try { $script:LaneProtocolPool.Dispose() }
					catch { }
					$script:LaneProtocolPool = $null
				}
			}
			catch { }
			
			# ---- Scale credential runspaces (if you used the background cred helper) ----
			try
			{
				if ($script:ScaleCredReaper)
				{
					try { $script:ScaleCredReaper.Stop() }
					catch { }
					try { $script:ScaleCredReaper.Dispose() }
					catch { }
					$script:ScaleCredReaper = $null
				}
			}
			catch { }
			try
			{
				if ($script:ScaleCredTasks)
				{
					foreach ($k in @($script:ScaleCredTasks.Keys))
					{
						$st = $script:ScaleCredTasks[$k]
						if ($st)
						{
							try { if ($st.Handle -and $st.PS) { $st.PS.EndInvoke($st.Handle) } }
							catch { }
							try { if ($st.PS) { $st.PS.Dispose() } }
							catch { }
						}
						[void]$script:ScaleCredTasks.Remove($k)
					}
				}
			}
			catch { }
			
			# --- Any lingering job from earlier versions (defensive) ---
			try
			{
				Get-Job -Name 'ClearXEFolderJob' -ErrorAction SilentlyContinue | ForEach-Object {
					try { Stop-Job $_ -Force -ErrorAction SilentlyContinue }
					catch { }
					try { Remove-Job $_ -Force -ErrorAction SilentlyContinue }
					catch { }
				}
			}
			catch { }
			
			Write_Log "Form is closing. Performing cleanup." "green"
			Delete_Files -Path "$TempDir" -SpecifiedFiles "*.sqi", "*.sql"
		})
	
	# ========================= Protocol Table Popup =========================
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
		
		# No minimize/maximize
		$global:ProtocolForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$global:ProtocolForm.MaximizeBox = $false
		$global:ProtocolForm.MinimizeBox = $false
		
		# Hide when focus leaves
		$global:ProtocolForm.add_Deactivate({ $global:ProtocolForm.Hide() })
		
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
		
		# Prevent disposal on X; just hide
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
				# Save scroll & selection BEFORE clearing rows
				$prevRowCount = $global:ProtocolGrid.Rows.Count
				$scrollIndex = 0
				if ($prevRowCount -gt 0)
				{
					try { $scrollIndex = $global:ProtocolGrid.FirstDisplayedScrollingRowIndex }
					catch { $scrollIndex = 0 }
				}
				$selIndex = $null
				if ($global:ProtocolGrid.SelectedRows.Count -gt 0)
				{
					try { $selIndex = $global:ProtocolGrid.SelectedRows[0].Index }
					catch { $selIndex = $null }
				}
				
				# Refresh grid
				$global:ProtocolGrid.Rows.Clear()
				
				$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
				if ($script:ProtocolResults)
				{
					# Sort by lane number if numeric, otherwise lexical
					$sorted = $script:ProtocolResults | Sort-Object {
						($_.Lane -replace '[^\d]', '') -as [int]
					}
					foreach ($rowObj in $sorted)
					{
						$r = $global:ProtocolGrid.Rows.Add()
						$machineName = $null
						if ($LaneNumToMachineName) { $machineName = $LaneNumToMachineName[$rowObj.Lane] }
						if (-not $machineName) { $machineName = $rowObj.Lane }
						$global:ProtocolGrid.Rows[$r].Cells[0].Value = $machineName
						$global:ProtocolGrid.Rows[$r].Cells[1].Value = $rowObj.Protocol
					}
				}
				
				# Column widths (Lane fixed, Protocol fills; account for scrollbar)
				$global:ProtocolGrid.Columns[0].Width = 60
				$visibleRowCount = [math]::Floor($global:ProtocolGrid.DisplayRectangle.Height / $global:ProtocolGrid.RowTemplate.Height)
				$scrollBarVisible = $global:ProtocolGrid.Rows.Count -gt $visibleRowCount
				if ($scrollBarVisible)
				{
					$global:ProtocolGrid.Columns[1].Width = $global:ProtocolGrid.Width - 60 - 4 - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
				}
				else
				{
					$global:ProtocolGrid.Columns[1].Width = $global:ProtocolGrid.Width - 60 - 4
				}
				
				# Restore scroll & selection safely
				$rowCount = $global:ProtocolGrid.Rows.Count
				if ($rowCount -gt 0)
				{
					if ($scrollIndex -lt 0) { $scrollIndex = 0 }
					if ($scrollIndex -ge $rowCount) { $scrollIndex = $rowCount - 1 }
					try { $global:ProtocolGrid.FirstDisplayedScrollingRowIndex = $scrollIndex }
					catch { }
				}
				if ($selIndex -ne $null -and $rowCount -gt $selIndex -and $selIndex -ge 0)
				{
					try { $global:ProtocolGrid.Rows[$selIndex].Selected = $true }
					catch { }
				}
			})
		$global:ProtocolFormTimer.Start()
	}
	
	######################################################################################################################
	#                                                                                                                    #
	#                                                    Labels 					                                     #
	#                                                                                                                    #
	######################################################################################################################
	
	# Shared tooltip (create once)  ─────────────────────────────────────────────────────────────────────────────────────
	if (-not $toolTip)
	{
		$toolTip = New-Object System.Windows.Forms.ToolTip
		$toolTip.AutoPopDelay = 8000
		$toolTip.InitialDelay = 300
		$toolTip.ReshowDelay = 100
	}
	
	# Helper as scriptblock (not a function): make a label look like a link (hand + blue + underline on hover)
	$applyLinkStyle = {
		param ([System.Windows.Forms.Label]$Label)
		$Label.Cursor = [System.Windows.Forms.Cursors]::Hand
		if (-not $Label.Tag) { $Label.Tag = [pscustomobject]@{ Color = $Label.ForeColor; Font = $Label.Font } }
		
		$Label.add_MouseEnter({
				param ($s,
					$e)
				try
				{
					$s.ForeColor = 'DodgerBlue'
					$s.Font = New-Object System.Drawing.Font($s.Font, ($s.Font.Style -bor [System.Drawing.FontStyle]::Underline))
				}
				catch { }
			})
		$Label.add_MouseLeave({
				param ($s,
					$e)
				try
				{
					if ($s.Tag -and $s.Tag.Font) { $s.Font = $s.Tag.Font }
					if ($s.Tag -and $s.Tag.Color) { $s.ForeColor = $s.Tag.Color }
				}
				catch { }
			})
	}
	
	# ─────────────────────────────────────────── Create the labels (your same look/positions) ──────────────────────────
	
	# SMS Version (left, row 1) - clickable text only
	$smsVersionLabel = New-Object System.Windows.Forms.Label
	$smsVersionLabel.Text = "SMS Version: N/A"
	$smsVersionLabel.Location = New-Object System.Drawing.Point(50, 30)
	$smsVersionLabel.AutoSize = $true # only letters are clickable (no dead zone)
	$smsVersionLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($smsVersionLabel) | Out-Null
	$toolTip.SetToolTip($smsVersionLabel, "Shows current SMS version. Click to bring SMSStart to front (or launch it).")
	& $applyLinkStyle $smsVersionLabel
	
	# Click: bring SMSStart to front or launch it (inline, no functions)
	$smsVersionLabel.Add_Click({
			$script:LastActivity = Get-Date
			$orig = $smsVersionLabel.ForeColor
			$smsVersionLabel.ForeColor = 'DodgerBlue'
			try
			{
				@"
using System;
using System.Runtime.InteropServices;
public static class NativeWin {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
}
"@ | ForEach-Object {
					if (-not ("NativeWin" -as [type])) { Add-Type -TypeDefinition $_ -ErrorAction SilentlyContinue }
				}
				
				$p = Get-Process -Name 'SMSStart' -ErrorAction SilentlyContinue |
				Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
				if ($p)
				{
					$h = $p.MainWindowHandle
					if ([NativeWin]::IsIconic($h)) { [NativeWin]::ShowWindow($h, 9) | Out-Null }
					[NativeWin]::SetForegroundWindow($h) | Out-Null
					if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "Brought SMSStart to the foreground." "green" }
					return
				}
				$exe = Join-Path $BasePath 'SMSStart.exe'
				$proc = $null
				if (Test-Path -LiteralPath $exe)
				{
					$proc = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe -Parent) -PassThru -ErrorAction Stop
				}
				else
				{
					$lnk = Get-ChildItem -LiteralPath $BasePath -Filter '*SMSStart*.lnk' -ErrorAction SilentlyContinue | Select-Object -First 1
					if ($lnk) { $proc = Start-Process -FilePath $lnk.FullName -PassThru -ErrorAction Stop }
				}
				if (-not $proc)
				{
					if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "SMSStart not found under $BasePath" "yellow" }
					return
				}
				$null = $proc.WaitForInputIdle(5000)
				for ($i = 0; $i -lt 20 -and $proc.MainWindowHandle -eq 0; $i++) { Start-Sleep -Milliseconds 150; $proc.Refresh() }
				if ($proc.MainWindowHandle -ne 0)
				{
					if ([NativeWin]::IsIconic($proc.MainWindowHandle)) { [NativeWin]::ShowWindow($proc.MainWindowHandle, 9) | Out-Null }
					[NativeWin]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
				}
				if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "Launched SMSStart from $BasePath and brought to front." "green" }
			}
			finally
			{
				$smsVersionLabel.ForeColor = $orig
			}
		})
	
	# Store Name (centered, row 1) - clickable text only
	$storeNameLabel = New-Object System.Windows.Forms.Label
	$storeNameLabel.Text = "Store Name: N/A"
	$storeNameLabel.AutoSize = $true
	$storeNameLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$storeNameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$form.Controls.Add($storeNameLabel) | Out-Null
	$toolTip.SetToolTip($storeNameLabel, "Shows store name. Click to ping all nodes (Lanes > Scales > Backoffices).")
	& $applyLinkStyle $storeNameLabel
	$storeNameLabel.Add_Click({
			$script:LastActivity = Get-Date
			$storeNameLabel.ForeColor = 'DodgerBlue'
			try
			{
				Ping_All_Nodes -NodeType "Lane" -StoreNumber $StoreNumber
				Ping_All_Nodes -NodeType "Scale" -StoreNumber $StoreNumber
				Ping_All_Nodes -NodeType "Backoffice" -StoreNumber $StoreNumber
			}
			finally
			{
				$storeNameLabel.ForeColor = 'Black'
			}
		})
	
	# Store Number (right, row 1) - clickable text only
	$script:storeNumberLabel = New-Object System.Windows.Forms.Label
	$storeNumberLabel.Text = "Store Number: N/A"
	$storeNumberLabel.AutoSize = $true
	$storeNumberLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$storeNumberLabel.Anchor = 'Top,Right' # still anchor to right; layout scriptblock will clamp
	$form.Controls.Add($storeNumberLabel) | Out-Null
	$toolTip.SetToolTip($storeNumberLabel, "Click to open the Storeman folder.")
	& $applyLinkStyle $storeNumberLabel
	
	# Hover: refresh tooltip with actual BasePath
	$storeNumberLabel.Add_MouseHover({
			$p = if ($script:BasePath) { $script:BasePath }
			elseif ($BasePath) { $BasePath }
			else { $null }
			$toolTip.SetToolTip($storeNumberLabel, $(if ($p) { "Open Storeman folder: $p" }
					else { "Storeman folder not detected yet." }))
		})
	# Click: open Storeman folder (inline)
	$storeNumberLabel.Add_Click({
			$script:LastActivity = Get-Date
			$path = if ($script:BasePath) { $script:BasePath }
			elseif ($BasePath) { $BasePath }
			else { $null }
			if ($path -and (Test-Path -LiteralPath $path))
			{
				try { Start-Process -FilePath $path }
				catch
				{
					[System.Windows.Forms.MessageBox]::Show("Couldn't open: $path`r`n$($_.Exception.Message)",
						"Open Storeman Folder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				}
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Storeman folder not found.`r`nCurrent value: " + [string]$path,
					"Open Storeman Folder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
			}
		})
	
	# Number of Lanes (centered, row 2) - clickable text only
	$script:NodesStore = New-Object System.Windows.Forms.Label
	$NodesStore.Text = "Number of Lanes: $($Counts.NumberOfLanes)"
	$NodesStore.Location = New-Object System.Drawing.Point(50, 50) # seed; layout will clamp/center
	$NodesStore.AutoSize = $true
	$NodesStore.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($NodesStore) | Out-Null
	$toolTip.SetToolTip($NodesStore, "Shows count of Lanes. Click to ping Lane nodes.")
	& $applyLinkStyle $NodesStore
	$NodesStore.Add_Click({
			$script:LastActivity = Get-Date
			$NodesStore.ForeColor = 'DodgerBlue'
			try { Ping_All_Nodes -NodeType "Lane" -StoreNumber $StoreNumber }
			finally { $NodesStore.ForeColor = 'Black' }
		})
	
	# Number of Scales (right-block, row 2)
	$script:scalesLabel = New-Object System.Windows.Forms.Label
	$scalesLabel.Text = "Number of Scales: $($Counts.NumberOfScales)"
	$scalesLabel.Location = New-Object System.Drawing.Point(420, 50) # seed; layout will position near right block
	$scalesLabel.AutoSize = $true
	$scalesLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($scalesLabel) | Out-Null
	$toolTip.SetToolTip($scalesLabel, "Shows count of Scales. Click to ping Scale nodes.")
	& $applyLinkStyle $scalesLabel
	$scalesLabel.Add_Click({
			$script:LastActivity = Get-Date
			try { Ping_All_Nodes -NodeType "Scale" -StoreNumber $StoreNumber }
			catch { }
		})
	
	# Number of Backoffices (rightmost, row 2)
	$NodesBackoffices = New-Object System.Windows.Forms.Label
	$NodesBackoffices.Text = "Number of Backoffices: N/A"
	$NodesBackoffices.Location = New-Object System.Drawing.Point(785, 50) # seed; layout will right-align
	$NodesBackoffices.AutoSize = $true
	$NodesBackoffices.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($NodesBackoffices) | Out-Null
	$toolTip.SetToolTip($NodesBackoffices, "Shows count of Backoffices. Click to ping Backoffice nodes.")
	& $applyLinkStyle $NodesBackoffices
	$NodesBackoffices.Add_Click({
			$script:LastActivity = Get-Date
			try { Ping_All_Nodes -NodeType "Backoffice" -StoreNumber $StoreNumber }
			catch { }
		})
	
	# ─────────────────────────────────────────── Non-overlap/clamping layout (updated) ───────────────────────────────────
	$TopPadLeft = 50 # left margin for labels
	$Gap = 12 # spacing between neighbors
	
	$layoutTopStrip = {
		if (-not $form -or $form.IsDisposed) { return }
		
		# Y positions for the two rows
		$y1 = 30 # Row 1: SMS (L), Store Name (C), Store Number (R)
		$y2 = 50 # Row 2: Lanes (L), Scales (C), Backoffices (R)
		
		# Right edge to align with = log box right (fallback to client right if logBox not ready)
		$rightEdge = if ($logBox -and -not $logBox.IsDisposed) { $logBox.Left + $logBox.Width }
		else { $form.ClientSize.Width - 20 }
		
		# ---------------- Row 1 ----------------
		if ($smsVersionLabel) { $smsVersionLabel.Top = $y1; $smsVersionLabel.Left = $TopPadLeft }
		
		if ($storeNumberLabel)
		{
			$storeNumberLabel.Top = $y1
			# align the right edge to the logBox right edge
			$storeNumberLabel.Left = [math]::Max(0, $rightEdge - $storeNumberLabel.Width)
		}
		
		if ($storeNameLabel)
		{
			$storeNameLabel.Top = $y1
			
			# clamp centered Store Name between SMS (left) and Store Number (right)
			$leftBound = if ($smsVersionLabel) { $smsVersionLabel.Left + $smsVersionLabel.Width + $Gap }
			else { $TopPadLeft }
			$rightBound = if ($storeNumberLabel) { $storeNumberLabel.Left - $Gap }
			else { $rightEdge }
			
			$centered = [math]::Floor(($form.ClientSize.Width - $storeNameLabel.Width) / 2)
			$safeLeft = [math]::Max($leftBound, [math]::Min($centered, $rightBound - $storeNameLabel.Width))
			$storeNameLabel.Left = [int][math]::Max(0, $safeLeft)
		}
		
		# ---------------- Row 2 ----------------
		if ($NodesStore)
		{
			$NodesStore.Top = $y2
			$NodesStore.Left = $TopPadLeft
		}
		
		if ($NodesBackoffices)
		{
			$NodesBackoffices.Top = $y2
			# align the right edge to the logBox right edge
			$NodesBackoffices.Left = [math]::Max(0, $rightEdge - $NodesBackoffices.Width)
		}
		
		if ($scalesLabel)
		{
			$scalesLabel.Top = $y2
			
			# clamp centered Scales between Lanes (left) and Backoffices (right)
			$leftBound2 = if ($NodesStore) { $NodesStore.Left + $NodesStore.Width + $Gap }
			else { $TopPadLeft }
			$rightBound2 = if ($NodesBackoffices) { $NodesBackoffices.Left - $Gap }
			else { $rightEdge }
			
			$centered2 = [math]::Floor(($form.ClientSize.Width - $scalesLabel.Width) / 2)
			$safeLeft2 = [math]::Max($leftBound2, [math]::Min($centered2, $rightBound2 - $scalesLabel.Width))
			$scalesLabel.Left = [int][math]::Max(0, $safeLeft2)
		}
	}
	
	
	# run once + re-run on any size/text/font change (reacts to AutoSize growth/shrink)
	& $layoutTopStrip
	$form.add_Resize({ & $layoutTopStrip })
	foreach ($lbl in @($smsVersionLabel, $storeNameLabel, $storeNumberLabel, $NodesStore, $scalesLabel, $NodesBackoffices))
	{
		if ($lbl)
		{
			$lbl.add_TextChanged({ & $layoutTopStrip })
			$lbl.add_FontChanged({ & $layoutTopStrip })
			$lbl.add_SizeChanged({ & $layoutTopStrip }) # extra safety for AutoSize
		}
	}
	
	# ───────────────────────────────────────────────────── Log box (unchanged) ─────────────────────────────────────────
	$logBox = New-Object System.Windows.Forms.RichTextBox
	$logBox.Location = New-Object System.Drawing.Point(50, 70)
	$logBox.Size = New-Object System.Drawing.Size(900, 400)
	$logBox.ReadOnly = $true
	$logBox.Font = New-Object System.Drawing.Font("Consolas", 10)
	$form.Controls.Add($logBox) | Out-Null
	$toolTip.SetToolTip($logBox, "Log output. Right-click to clear.")
	$logBox.Add_MouseUp({
			param ($sender,
				$eventArgs)
			if ($eventArgs.Button -eq [System.Windows.Forms.MouseButtons]::Right)
			{
				$logBox.Clear()
				if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "Log Cleared" }
			}
		})
	
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
			$script:LastActivity = Get-Date
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
				) | Out-Null
			}
		})
	[void]$ContextMenuServer.Items.Add($ServerDBMaintenanceItem)
	
	############################################################################
	# 2) Schedule the DB maintenance at the lanes
	############################################################################
	$ServerScheduleRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Server DB Maintenance")
	$ServerScheduleRepairItem.ToolTipText = "Schedule a task to run maintenance at the server database."
	$ServerScheduleRepairItem.Add_Click({
			$script:LastActivity = Get-Date
			Schedule_Server_DB_Maintenance -StoreNumber $StoreNumber
		})
	[void]$ContextMenuServer.Items.Add($ServerScheduleRepairItem)
	
	############################################################################
	# 3) Schedule the local DB backup on the server
	############################################################################
	$ServerScheduleBackupItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Local DB Backup")
	$ServerScheduleBackupItem.ToolTipText = "Schedule a task to run database backup on the server."
	$ServerScheduleBackupItem.Add_Click({
			$script:LastActivity = Get-Date
			Schedule_LocalDB_Backup
		})
	[void]$ContextMenuServer.Items.Add($ServerScheduleBackupItem)
	
	############################################################################
	# 4) Schedule the Storeman ZIP backup on the server
	############################################################################
	$ServerScheduleStoremanZipBackupItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Storeman ZIP Backup")
	$ServerScheduleStoremanZipBackupItem.ToolTipText = "Schedule a task to back up the Storeman folder to a weekly ZIP archive."
	$ServerScheduleStoremanZipBackupItem.Add_Click({
			$script:LastActivity = Get-Date
			Schedule_Storeman_Zip_Backup
		})
	[void]$ContextMenuServer.Items.Add($ServerScheduleStoremanZipBackupItem)
		
	############################################################################
	# 5) Manage SQL 'sa' Account Menu Item
	############################################################################
	$ManageSaAccountItem = New-Object System.Windows.Forms.ToolStripMenuItem("Manage SQL 'sa' Account")
	$ManageSaAccountItem.ToolTipText = "Enable or disable the 'sa' account on the local SQL Server with a predefined password."
	$ManageSaAccountItem.Add_Click({
			$script:LastActivity = Get-Date
			Manage_Sa_Account
		})
	[void]$ContextMenuServer.Items.Add($ManageSaAccountItem)
	
	############################################################################
	# 6) Repair Windows Menu Item
	############################################################################
	$RepairWindowsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair Windows")
	$RepairWindowsItem.ToolTipText = "Perform repairs on the Windows operating system."
	$RepairWindowsItem.Add_Click({
			$script:LastActivity = Get-Date
			Repair_Windows
		})
	[void]$ContextMenuServer.Items.Add($RepairWindowsItem)
	
	############################################################################
	# 7) Configure System Settings Menu Item
	############################################################################
	$ConfigureSystemSettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Configure System Settings")
	$ConfigureSystemSettingsItem.ToolTipText = "Set High Performance power plan, disable sleep, ensure required services are Automatic/running, and apply visual tweaks."
	$ConfigureSystemSettingsItem.Add_Click({
			$script:LastActivity = Get-Date
			$confirmResult = [System.Windows.Forms.MessageBox]::Show(
				"Warning: This will modify power, services, and visual settings. Continue?",
				"Confirm Changes",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				try { Configure_System_Settings }
				catch
				{
					if (Get-Command Write_Log -ErrorAction SilentlyContinue)
					{
						Write_Log "Configure_System_Settings threw an error: $_" "red"
					}
					else
					{
						[System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Configure System Settings",
							[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
					}
				}
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Operation canceled.", "Canceled",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
			}
		})
	[void]$ContextMenuServer.Items.Add($ConfigureSystemSettingsItem)
	
	############################################################################
	# 8) Organize Desktop Items Menu Item
	############################################################################
	$OrganizeDesktopItemsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Organize Desktop Items")
	$OrganizeDesktopItemsItem.ToolTipText = "Move non-system items from the Desktop into the 'Unorganized Items' folder (or a name you choose)."
	$OrganizeDesktopItemsItem.Add_Click({
			$script:LastActivity = Get-Date
			$result = [System.Windows.Forms.MessageBox]::Show(
				"Use a custom folder name for your unorganized Desktop items? (Click 'No' to use the default: 'Unorganized Items')",
				"Organize Desktop Items",
				[System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)
			if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
			{
				[System.Windows.Forms.MessageBox]::Show("Operation canceled.", "Canceled",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			$targetFolderName = "Unorganized Items"
			if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Add-Type -AssemblyName System.Windows.Forms
				Add-Type -AssemblyName System.Drawing
				$inForm = New-Object System.Windows.Forms.Form
				$inForm.Text = "Custom Folder Name"
				$inForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
				$inForm.StartPosition = "CenterParent"
				$inForm.ClientSize = New-Object System.Drawing.Size(420, 120)
				$inForm.MaximizeBox = $false; $inForm.MinimizeBox = $false; $inForm.TopMost = $true
				$lbl = New-Object System.Windows.Forms.Label
				$lbl.Text = "Enter a folder name to create on the Desktop:"; $lbl.AutoSize = $true
				$lbl.Location = New-Object System.Drawing.Point(12, 12); $inForm.Controls.Add($lbl)
				$tb = New-Object System.Windows.Forms.TextBox
				$tb.Size = New-Object System.Drawing.Size(390, 24); $tb.Location = New-Object System.Drawing.Point(12, 40)
				$tb.Text = "Unorganized Items"; $inForm.Controls.Add($tb)
				$okBtn = New-Object System.Windows.Forms.Button
				$okBtn.Text = "OK"; $okBtn.Size = New-Object System.Drawing.Size(90, 28)
				$okBtn.Location = New-Object System.Drawing.Point(230, 80)
				$okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
				$inForm.AcceptButton = $okBtn; $inForm.Controls.Add($okBtn)
				$cancelBtn = New-Object System.Windows.Forms.Button
				$cancelBtn.Text = "Cancel"; $cancelBtn.Size = New-Object System.Drawing.Size(90, 28)
				$cancelBtn.Location = New-Object System.Drawing.Point(324, 80)
				$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
				$inForm.CancelButton = $cancelBtn; $inForm.Controls.Add($cancelBtn)
				$dlgRes = $inForm.ShowDialog()
				if ($dlgRes -ne [System.Windows.Forms.DialogResult]::OK)
				{
					[System.Windows.Forms.MessageBox]::Show("Operation canceled.", "Canceled",
						[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
					return
				}
				$txt = $tb.Text
				if ([string]::IsNullOrWhiteSpace($txt))
				{
					[System.Windows.Forms.MessageBox]::Show("Folder name cannot be empty.", "Invalid Input",
						[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
					return
				}
				$targetFolderName = $txt.Trim()
			}
			$confirm = [System.Windows.Forms.MessageBox]::Show(
				"This will move non-system items from the Desktop into '$targetFolderName'. Continue?",
				"Confirm Desktop Organization",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				[System.Windows.Forms.MessageBox]::Show("Operation canceled.", "Canceled",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			try
			{
				if (Get-Command Organize_Desktop_Items -ErrorAction SilentlyContinue)
				{
					Organize_Desktop_Items -UnorganizedFolderName $targetFolderName
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show(
						"Organize_Desktop_Items is not loaded in this session.",
						"Function Not Found",
						[System.Windows.Forms.MessageBoxButtons]::OK,
						[System.Windows.Forms.MessageBoxIcon]::Error
					) | Out-Null
				}
			}
			catch
			{
				if (Get-Command Write_Log -ErrorAction SilentlyContinue)
				{
					Write_Log "Organize_Desktop_Items threw an error: $_" "red"
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Organize Desktop Items",
						[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				}
			}
		})
	[void]$ContextMenuServer.Items.Add($OrganizeDesktopItemsItem)
	
	############################################################################
	# Show the context menu when the Server Tools button is clicked
	############################################################################
	$ServerToolsButton.Add_Click({ $ContextMenuServer.Show($ServerToolsButton, 0, $ServerToolsButton.Height) })
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
	$LaneToolsButton.Add_Click({ $ContextMenuLane.Show($LaneToolsButton, 0, $LaneToolsButton.Height) })
	$LaneToolsButton.Add_MouseDown({
			param ($sender,
				$e)
			if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right)
			{
				$global:ProtocolForm.Show(); $global:ProtocolForm.BringToFront()
			}
		})
	
	############################################################################
	# 1) Lane DB Maintenance Button
	############################################################################
	$LaneDBMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Lane DB Maintenance")
	$LaneDBMaintenanceItem.ToolTipText = "Runs maintenance at the lane(s) databases for the selected lane(s)."
	$LaneDBMaintenanceItem.Add_Click({
			$script:LastActivity = Get-Date
			Process_Lanes -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($LaneDBMaintenanceItem)
	
	############################################################################
	# 2) Schedule the DB maintenance at the lanes
	############################################################################
	$LaneScheduleMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Lane DB Maintenance")
	$LaneScheduleMaintenanceItem.ToolTipText = "Schedule a task to run maintenance at the lane/s database."
	$LaneScheduleMaintenanceItem.Add_Click({
			$script:LastActivity = Get-Date
			Schedule_Lane_DB_Maintenance -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($LaneScheduleMaintenanceItem)
	
	############################################################################
	# 3) Schedule DB backup of the lanes
	############################################################################
	$LaneScheduleBackupItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Lane DB Backups")
	$LaneScheduleBackupItem.ToolTipText = "Schedule a task to run backups of the selected lanes' databases."
	$LaneScheduleBackupItem.Add_Click({
			$script:LastActivity = Get-Date
			Schedule_LaneDB_Backup
		})
	[void]$ContextMenuLane.Items.Add($LaneScheduleBackupItem)
	
	<############################################################################
	#  X) Install/Check LOC Options (Lanes) - lane picker + options picker
	############################################################################
	$InstallCheckLOCOptionsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Install/Check LOC Options")
	$InstallCheckLOCOptionsItem.ToolTipText = "Pick lanes, then pick LOC Options to audit/install/reinstall (with categories & search)."
	$InstallCheckLOCOptionsItem.Add_Click({
			$script:LastActivity = Get-Date
			Install_And_Check_LOC_SMS_Options_On_Lanes -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($InstallCheckLOCOptionsItem)#>
	
	############################################################################
	#  X) Audit/Repair Lane Databases - lane picker + in-function level picker
	############################################################################
	$RepairLaneDatabasesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Audit/Repair Lane Databases")
	$RepairLaneDatabasesItem.ToolTipText = "Pick lanes, then choose Audit/Quick/Deep in the next dialog. Uses Startup.ini via Get_All_Lanes_Database_Info."
	$RepairLaneDatabasesItem.Add_Click({
			$script:LastActivity = Get-Date
			Repair_LOC_Databases_On_Lanes -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($RepairLaneDatabasesItem)
	
	############################################################################
	# 4) Pump Table to Lane Menu Item
	############################################################################
	$PumpTableToLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Pump Table to Lane")
	$PumpTableToLaneItem.ToolTipText = "Pump the selected tables to the lane/s databases."
	$PumpTableToLaneItem.Add_Click({
			$script:LastActivity = Get-Date
			Pump_Tables -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($PumpTableToLaneItem)
	
	############################################################################
	# 5) Update Lane Configuration Menu Item
	############################################################################
	$UpdateLaneConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Lane Configuration")
	$UpdateLaneConfigItem.ToolTipText = "Update the configuration files for the lanes. Fixes connectivity errors and mistakes made during lane ghosting."
	$UpdateLaneConfigItem.Add_Click({
			$script:LastActivity = Get-Date
			Update_Lane_Config -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($UpdateLaneConfigItem)
	
	############################################################################
	# 6) Close Open Transactions Menu Item
	############################################################################
	$CloseOpenTransItem = New-Object System.Windows.Forms.ToolStripMenuItem("Close Open Transactions")
	$CloseOpenTransItem.ToolTipText = "Close any open transactions at the lane/s."
	$CloseOpenTransItem.Add_Click({
			$script:LastActivity = Get-Date
			Close_Open_Transactions -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($CloseOpenTransItem)
	
	############################################################################
	# 7) Open Lane C$ Share(s)
	############################################################################
	$OpenLaneCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Lane C$ Share(s)")
	$OpenLaneCShareItem.ToolTipText = "Select lanes and open their administrative C$ shares in Explorer."
	$OpenLaneCShareItem.Add_Click({
			$script:LastActivity = Get-Date
			Open_Selected_Node_C_Path -StoreNumber $StoreNumber -NodeTypes Lane
		})
	[void]$ContextMenuLane.Items.Add($OpenLaneCShareItem)
	
	############################################################################
	# 8) Delete DBS Menu Item
	############################################################################
	$DeleteDBSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Delete DBS")
	$DeleteDBSItem.ToolTipText = "Delete the DBS files (*.txt, *.dwr, if selected *.sus as well) at the lane."
	$DeleteDBSItem.Add_Click({
			$script:LastActivity = Get-Date
			Delete_DBS -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($DeleteDBSItem)
	
	############################################################################
	# 9) Refresh PIN Pad Files Menu Item
	############################################################################
	$RefreshPinPadFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh PIN Pad Files")
	$RefreshPinPadFilesItem.ToolTipText = "Refresh the PIN pad files for the lane/s."
	$RefreshPinPadFilesItem.Add_Click({
			$script:LastActivity = Get-Date
			Refresh_PIN_Pad_Files -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($RefreshPinPadFilesItem)
	
	############################################################################
	# 10) Drawer Control Item
	############################################################################
	$DrawerControlItem = New-Object System.Windows.Forms.ToolStripMenuItem("Drawer Control")
	$DrawerControlItem.ToolTipText = "Set the Drawer Control for a lane for testing"
	$DrawerControlItem.Add_Click({
			$script:LastActivity = Get-Date
			Drawer_Control -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($DrawerControlItem)
	
	############################################################################
	# 11) Refresh Database
	############################################################################
	$RefreshDatabaseItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Database")
	$RefreshDatabaseItem.ToolTipText = "Refresh the database at the lane/s"
	$RefreshDatabaseItem.Add_Click({
			$script:LastActivity = Get-Date
			Refresh_Database -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($RefreshDatabaseItem)
	
	############################################################################
	# 12) Send Restart Command Menu Item
	############################################################################
	$SendRestartCommandItem = New-Object System.Windows.Forms.ToolStripMenuItem("Send Restart All Programs")
	$SendRestartCommandItem.ToolTipText = "Send restart all programs to selected lane(s) for the store."
	$SendRestartCommandItem.Add_Click({
			$script:LastActivity = Get-Date
			Send_Restart_All_Programs -StoreNumber "$StoreNumber"
		})
	[void]$ContextMenuLane.Items.Add($SendRestartCommandItem)
	
	############################################################################
	# 13) Enable SQL Protocols Menu Item
	############################################################################
	$EnableSQLProtocolsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Enable SQL Protocols")
	$EnableSQLProtocolsItem.ToolTipText = "Enable TCP/IP, Named Pipes, and set static port for SQL Server on selected lane(s)."
	$EnableSQLProtocolsItem.Add_Click({
			$script:LastActivity = Get-Date
			Enable_SQL_Protocols_On_Selected_Lanes -StoreNumber $StoreNumber
		})
	[void]$ContextMenuLane.Items.Add($EnableSQLProtocolsItem)
	
	############################################################################
	# 14) Set the time on the lanes
	############################################################################
	$SetLaneTimeFromLocalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Set/Schedule Time on Lanes")
	$SetLaneTimeFromLocalItem.ToolTipText = "Synchronize or schedule time sync for selected lanes."
	$SetLaneTimeFromLocalItem.Add_Click({
			$script:LastActivity = Get-Date
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
			$script:LastActivity = Get-Date
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
	# 1) Repair BMS Service
	############################################################################
	$repairBMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair BMS Service")
	$repairBMSItem.ToolTipText = "Repairs the BMS service for scale deployment."
	$repairBMSItem.Add_Click({
			$script:LastActivity = Get-Date
			Repair_BMS
		})
	[void]$ContextMenuScale.Items.Add($repairBMSItem)
	
	############################################################################
	# 2) Troubleshoot ScaleCommApp
	############################################################################
	$troubleshootItem = New-Object System.Windows.Forms.ToolStripMenuItem("Troubleshoot ScaleCommApp")
	$troubleshootItem.ToolTipText = "Checks and fixes ScaleCommApp config (StoreName, First Scale IP, SQL instance/db, BatchSendFull=10000)."
	$troubleshootItem.Add_Click({
			$script:LastActivity = Get-Date
			# Auto-fix enabled; update ALL configs found; clear read-only if needed
			Troubleshoot_ScaleCommApp -AllMatches -Force
		})
	[void]$ContextMenuScale.Items.Add($troubleshootItem)
	
	############################################################################
	# 3) Reboot Scales
	############################################################################
	$Reboot_ScalesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Scales")
	$Reboot_ScalesItem.ToolTipText = "Reboot Scale/s."
	$Reboot_ScalesItem.Add_Click({
			$script:LastActivity = Get-Date
			Reboot_Nodes -StoreNumber "$StoreNumber" -NodeTypes Scale
		})
	[void]$ContextMenuScale.Items.Add($Reboot_ScalesItem)
	
	############################################################################
	# 4) Open Scale C$ Share(s)
	############################################################################
	$OpenScaleCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Scale C$ Share(s)")
	$OpenScaleCShareItem.ToolTipText = "Select scales and open their C$ administrative shares as 'bizuser' (bizerba/biyerba)."
	$OpenScaleCShareItem.Add_Click({
			$script:LastActivity = Get-Date
			Open_Selected_Node_C_Path -StoreNumber $StoreNumber -NodeTypes Scale
		})
	[void]$ContextMenuScale.Items.Add($OpenScaleCShareItem)
	
	############################################################################
	# 5) Edit_TBS_SCL_ver520_Table Menu Item
	############################################################################
	$OrganizeScaleTableItem = New-Object System.Windows.Forms.ToolStripMenuItem("Edit_TBS_SCL_ver520_Table")
	$OrganizeScaleTableItem.ToolTipText = "Organize the Scale SQL table (TBS_SCL_ver520)."
	$OrganizeScaleTableItem.Add_Click({
			$script:LastActivity = Get-Date
			Edit_TBS_SCL_ver520_Table
		})
	[void]$ContextMenuScale.Items.Add($OrganizeScaleTableItem)
	
	############################################################################
	# 6) Deploy Scale Currency Files
	############################################################################
	$DeployScaleCurrencyFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Deploy Scale Currency Files")
	$DeployScaleCurrencyFilesItem.ToolTipText = "Push currency-configured price files (.txt, .properties) to selected scales (Bizerba only)."
	$DeployScaleCurrencyFilesItem.Add_Click({
			$script:LastActivity = Get-Date
			Deploy_Scale_Currency_Files
		})
	[void]$ContextMenuScale.Items.Add($DeployScaleCurrencyFilesItem)
	
	############################################################################
	# 7) Edit Ishida SDPs (ALL configs)
	############################################################################
	$editIshidaSDPsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Edit Ishida SDPs (All Configs)")
	$editIshidaSDPsItem.ToolTipText = "Pick subdepartments from SDP_TAB (F04/F1022) and write to IshidaSDPs in ALL ScaleCommApp *.exe.config files."
	$editIshidaSDPsItem.Add_Click({
			# Bump activity timer (keeps your idle/auto-close logic happy)
			$script:LastActivity = Get-Date
			
			# Launch picker and update ALL configs in C:\ScaleCommApp and D:\ScaleCommApp top-level.
			# Add -Force if you want to auto-clear ReadOnly attributes.
			Pick_And_Update_IshidaSDPs -AllMatches
		})
	[void]$ContextMenuScale.Items.Add($editIshidaSDPsItem)
	
	############################################################################
	# 8) Update Scales Specials
	############################################################################
	$UpdateScalesSpecialsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Scales Specials")
	$UpdateScalesSpecialsItem.ToolTipText = "Update scale specials immediately or schedule as a daily 5AM task."
	$UpdateScalesSpecialsItem.Add_Click({
			$script:LastActivity = Get-Date
			Update_Scales_Specials_Interactive
		})
	[void]$ContextMenuScale.Items.Add($UpdateScalesSpecialsItem)
	
	############################################################################
	# 9) Update Scale Config and DB (F272 Upsert)
	############################################################################
	$UpdateScaleConfigAndDBItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Scale Config && DB (F272 Upsert)")
	$UpdateScaleConfigAndDBItem.ToolTipText = "Updates ScaleCommApp configs and upserts F272 in SCL_TAB for POS_TAB F82=1 in item range."
	$UpdateScaleConfigAndDBItem.Add_Click({
			$script:LastActivity = Get-Date
			Update_ScaleConfig_And_DB
		})
	[void]$ContextMenuScale.Items.Add($UpdateScaleConfigAndDBItem)
	
	############################################################################
	# 10) Configure Subdepartments & SdpDefault (OBJ.F16 / OBJ.F18 / OBJ.F17)
	############################################################################
	$sep_ConfigureSdpDefault = New-Object System.Windows.Forms.ToolStripSeparator
	#[void]$ContextMenuScale.Items.Add($sep_ConfigureSdpDefault)
	$ConfigureSdpDefaultItem = New-Object System.Windows.Forms.ToolStripMenuItem("Configure Sub-Departments for Scales")
	$ConfigureSdpDefaultItem.ToolTipText = "Populate subdepartments (Auto or FAM/RPC/CAT) and set SdpDefault to OBJ.F16 / OBJ.F18 / OBJ.F17, or restore SdpDefault=1."
	$ConfigureSdpDefaultItem.Add_Click({
			$script:LastActivity = Get-Date
			Configure_Scale_Sub-Departments
		})
	[void]$ContextMenuScale.Items.Add($ConfigureSdpDefaultItem)
	
	############################################################################
	# 11) Schedule Duplicate File Monitor
	############################################################################
	$ScheduleRemoveDupesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Remove duplicate files from (toBizerba)")
	$ScheduleRemoveDupesItem.ToolTipText = "Monitor for and auto-delete duplicate files in (toBizerba). Run now or schedule as SYSTEM."
	$ScheduleRemoveDupesItem.Add_Click({
			$script:LastActivity = Get-Date
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
	# General Tools Button
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
			$script:LastActivity = Get-Date
			Invoke_Secure_Script
		})
	[void]$contextMenuGeneral.Items.Add($activateItem)
	
	############################################################################
	# 2) Reboot System
	############################################################################
	$rebootItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot System")
	$rebootItem.ToolTipText = "Reboot the host system immediately."
	$rebootItem.Add_Click({
			$script:LastActivity = Get-Date
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
	# 3) Install Functions in SMS (One + Multi)  -- UPDATED
	############################################################################
	$Install_ONE_FUNCTION_Into_SMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Install 'DEPLOY_MULTI_FCT' in SMS")
	$Install_ONE_FUNCTION_Into_SMSItem.ToolTipText = "Updates DEPLOY_SYS.sql and installs DEPLOY_MULTI_FCT.sqm into SMS."
	$Install_ONE_FUNCTION_Into_SMSItem.Add_Click({
			$script:LastActivity = Get-Date
			Install_FUNCTIONS_Into_SMS -StoreNumber $StoreNumber -OfficePath $OfficePath
		})
	[void]$contextMenuGeneral.Items.Add($Install_ONE_FUNCTION_Into_SMSItem)
	
	############################################################################
	# 3b) Context menu item: Copy Files Between Nodes
	############################################################################
	$Copy_Files_Between_NodesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Files Between Nodes")
	$Copy_Files_Between_NodesItem.ToolTipText = "Copy (storeman) subfolders/files from Server or a Lane to selected lanes."
	$Copy_Files_Between_NodesItem.Add_Click({
			$script:LastActivity = Get-Date
			Copy_Files_Between_Nodes
		})
	[void]$contextMenuGeneral.Items.Add($Copy_Files_Between_NodesItem)
	
	############################################################################
	# 3c) Context menu item: Edit INIs (Setup.ini and others)
	############################################################################
	$INI_EditorItem = New-Object System.Windows.Forms.ToolStripMenuItem("INI_Editor")
	$INI_EditorItem.ToolTipText = "Edit \storeman\<relative>\*.ini (default: office\Setup.ini) on Server or a Lane, then optionally copy to other lanes."
	$INI_EditorItem.Add_Click({
			$script:LastActivity = Get-Date
			INI_Editor
		})
	[void]$contextMenuGeneral.Items.Add($INI_EditorItem)
	
	############################################################################
	# 5) Manual Repair
	############################################################################
	$manualRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Manual Repair")
	$manualRepairItem.ToolTipText = "Writes SQL repair scripts to the desktop."
	$manualRepairItem.Add_Click({
			$script:LastActivity = Get-Date
			Write_SQL_Scripts_To_Desktop -LaneSQL $script:LaneSQLFiltered -ServerSQL $script:ServerSQLScript
		})
	[void]$contextMenuGeneral.Items.Add($manualRepairItem)
	
	############################################################################
	# 6) Fix Journal
	############################################################################
	$fixJournalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Fix Journal")
	$fixJournalItem.ToolTipText = "Fix journal entries for the specified date."
	$fixJournalItem.Add_Click({
			$script:LastActivity = Get-Date
			Fix_Journal -StoreNumber $StoreNumber -OfficePath $OfficePath
		})
	[void]$contextMenuGeneral.Items.Add($fixJournalItem)
	
	############################################################################
	# 7) Reboot selected Backoffices
	############################################################################
	$RebootBackofficesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Backoffices")
	$RebootBackofficesItem.ToolTipText = "Reboot the selected Backoffice/s."
	$RebootBackofficesItem.Add_Click({
			$script:LastActivity = Get-Date
			Reboot_Nodes -StoreNumber $StoreNumber -NodeTypes Backoffice
		})
	[void]$contextMenuGeneral.Items.Add($RebootBackofficesItem)
	
	############################################################################
	# 8) Export All VNC Files
	############################################################################
	$ExportVNCFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Export All VNC Files")
	$ExportVNCFilesItem.ToolTipText = "Generate UltraVNC (.vnc) connection files for all lanes, scales, and backoffices."
	$ExportVNCFilesItem.Add_Click({
			$script:LastActivity = Get-Date
			Export_VNC_Files_For_All_Nodes `
										   -LaneNumToMachineName $script:FunctionResults['LaneNumToMachineName'] `
										   -ScaleCodeToIPInfo $script:FunctionResults['ScaleCodeToIPInfo'] `
										   -BackofficeNumToMachineName $script:FunctionResults['BackofficeNumToMachineName']`
										   -AllVNCPasswords $script:FunctionResults['AllVNCPasswords']
		})
	[void]$contextMenuGeneral.Items.Add($ExportVNCFilesItem)
	
	############################################################################
	# 9) Export Machines Hardware Info
	############################################################################
	$ExportMachineHardwareInfoItem = New-Object System.Windows.Forms.ToolStripMenuItem("Export Machines Hardware Info")
	$ExportMachineHardwareInfoItem.ToolTipText = "Collect and export manufacturer/model for all machines"
	$ExportMachineHardwareInfoItem.Add_Click({
			$script:LastActivity = Get-Date
			$didExport = Get_Remote_Machine_Info
		})
	[void]$contextMenuGeneral.Items.Add($ExportMachineHardwareInfoItem)
	
	############################################################################
	# 10) Remove Archive Bit
	############################################################################
	$RemoveArchiveBitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Remove Archive Bit")
	$RemoveArchiveBitItem.ToolTipText = "Remove archived bit from all lanes and server. Option to schedule as a repeating task."
	$RemoveArchiveBitItem.Add_Click({
			$script:LastActivity = Get-Date
			Remove_ArchiveBit_Interactive
		})
	[void]$contextMenuGeneral.Items.Add($RemoveArchiveBitItem)
	
	############################################################################
	# 11) Sync Hosts File for Selected Nodes
	############################################################################
	$SyncHostsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Sync Host Files")
	$SyncHostsItem.ToolTipText = "Update the host file with the nodes selected, then copy the local host file to the selected node."
	$SyncHostsItem.Add_Click({
			$script:LastActivity = Get-Date
			Sync_Selected_Node_Hosts -StoreNumber $StoreNumber
		})
	[void]$contextMenuGeneral.Items.Add($SyncHostsItem)
	
	############################################################################
	# 12) Insert Test Item
	############################################################################
	$InsertTestItem = New-Object System.Windows.Forms.ToolStripMenuItem("Insert Test Item")
	$InsertTestItem.ToolTipText = "Inserts or updates a test item (PLU 0020077700000 or alternatives) in the database."
	$InsertTestItem.Add_Click({
			$script:LastActivity = Get-Date
			Insert_Test_Item
		})
	[void]$ContextMenuGeneral.Items.Add($InsertTestItem)
	
	############################################################################
	# 13) Fix Deploy CHG
	############################################################################
	$FixDeployCHGItem = New-Object System.Windows.Forms.ToolStripMenuItem("Fix Deploy_CHG")
	$FixDeployCHGItem.ToolTipText = "Restores the deploy line to DEPLOY_CHG.sql for scale management."
	$FixDeployCHGItem.Add_Click({
			$script:LastActivity = Get-Date
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
	
	# -------- Bottom buttons + logbox layout (expand buttons, tight gap) --------
	$padSide = 50 # left/right margin for logbox & button row
	$padBottom = 10 # distance from buttons to bottom edge
	$gapAboveButtons = 8 # <-- vertical gap between logbox and buttons
	$btnH = 50 # button height (fixed)
	$btnMinW = 180 # minimum button width so labels never wrap
	$gapHPreferred = 18 # preferred horizontal gap between buttons
	$gapHMin = 10 # absolute minimum horizontal gap
	$minLogH = 140 # don't let the logbox collapse
	
	# keep WinForms from fighting our math
	$ServerToolsButton.Anchor = 'Bottom'
	$LaneToolsButton.Anchor = 'Bottom'
	$ScaleToolsButton.Anchor = 'Bottom'
	$GeneralToolsButton.Anchor = 'Bottom'
	$logBox.Anchor = 'Top,Left'
	
	$layoutBottom = {
		if (-not $form -or $form.IsDisposed) { return }
		
		# --- lay out the logbox first (full width minus side margins) -------------
		$rowW = [math]::Max(200, $form.ClientSize.Width - (2 * $padSide))
		$logBox.Left = $padSide
		$logBox.Top = 70
		$logBox.Width = $rowW
		
		# button row Y; then set logbox height so the vertical gap stays constant
		$btnY = $form.ClientSize.Height - $btnH - $padBottom
		$logBox.Height = [math]::Max($minLogH, $btnY - $logBox.Top - $gapAboveButtons)
		
		# --- compute dynamic button width so the four buttons fill the row --------
		# try with preferred gaps first
		$gapH = $gapHPreferred
		$btnW = [math]::Floor(($rowW - (3 * $gapH)) / 4)
		
		# if too small, clamp width and reduce gaps (not below gapHMin)
		if ($btnW -lt $btnMinW)
		{
			$btnW = $btnMinW
			$gapH = [math]::Floor(($rowW - (4 * $btnW)) / 3)
			if ($gapH -lt $gapHMin) { $gapH = $gapHMin }
		}
		else
		{
			# distribute any leftover pixels (from integer math) into gaps
			$used = (4 * $btnW) + (3 * $gapH)
			$extra = $rowW - $used
			if ($extra -gt 0) { $gapH += [math]::Floor($extra / 3) }
		}
		
		# --- position buttons left->right; group width == logbox width ------------
		$x = $padSide
		$ServerToolsButton.SetBounds($x, $btnY, $btnW, $btnH); $x += $btnW + $gapH
		$LaneToolsButton.SetBounds($x, $btnY, $btnW, $btnH); $x += $btnW + $gapH
		$ScaleToolsButton.SetBounds($x, $btnY, $btnW, $btnH); $x += $btnW + $gapH
		$GeneralToolsButton.SetBounds($x, $btnY, $btnW, $btnH)
		
		# keep right-side labels aligned with the logbox edge
		& $layoutTopStrip
	}
	
	# prevent shrinking so far that 4*minW + 3*minGap won't fit
	$minW = (2 * $padSide) + (4 * $btnMinW) + (3 * $gapHMin)
	if ($form.MinimumSize.Width -lt $minW)
	{
		$form.MinimumSize = New-Object System.Drawing.Size($minW, [math]::Max($form.MinimumSize.Height, 560))
	}
	
	# run now + on resize
	& $layoutBottom
	$form.add_Resize({ & $layoutBottom })
	
	# ========================= Global Activity Hooks (controls & menus) =========================
	$__activityControls = @(
		$form,
		$logBox,
		$ServerToolsButton, $LaneToolsButton, $ScaleToolsButton, $GeneralToolsButton,
		$storeNameLabel, $NodesBackoffices, $NodesStore, $scalesLabel, $smsVersionLabel, $bannerLabel,
		$ContextMenuServer, $ContextMenuLane, $ContextMenuScale, $ContextMenuGeneral,
		$global:ProtocolForm, $global:ProtocolGrid
	) | Where-Object { $_ }
	
	foreach ($__c in $__activityControls)
	{
		try
		{
			if ($__c -is [System.Windows.Forms.Control])
			{
				$__c.Add_MouseMove({ $script:LastActivity = Get-Date })
				$__c.Add_KeyDown({ $script:LastActivity = Get-Date })
				$__c.Add_Click({ $script:LastActivity = Get-Date })
			}
		}
		catch { }
	}
}

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

# Get the Lanes protocol info if it doesnt exist
Start_Lane_Protocol_Jobs -LaneNumToMachineName $LaneNumToMachineName -SqlModuleName $SqlModuleName

# Populate the hash table with results from various functions
$AliasToTable = Get_Table_Aliases

# Generate SQL scripts
Generate_SQL_Scripts -StoreNumber $StoreNumber -LanesqlFilePath $LanesqlFilePath -StoresqlFilePath $StoresqlFilePath

# Add all the Scales to the credential manager
Add_Scale_Credentials -ScaleCodeToIPInfo $script:FunctionResults['WindowsScales']

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
