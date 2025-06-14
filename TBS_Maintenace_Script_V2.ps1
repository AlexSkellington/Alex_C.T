<#
Param (
	[switch]$IsRelaunched
)
#>

# Write-Host "Script started. IsRelaunched: $IsRelaunched"
Write-Host "Script starting, pls wait..." -ForegroundColor Yellow
Write-Host "®Tecnica Bussiness System" -ForegroundColor Blue

# ===================================================================================================
#                                       SECTION: Parameters
# ---------------------------------------------------------------------------------------------------
# Description:
#   Defines the script parameters, allowing users to run the script in silent mode.
# ===================================================================================================
 
# Script build version (cunsult with Alex_C.T before changing this)
$VersionNumber = "2.2.0"

# Retrieve Major, Minor, Build, and Revision version numbers of PowerShell
$major = $PSVersionTable.PSVersion.Major
$minor = $PSVersionTable.PSVersion.Minor
$build = $PSVersionTable.PSVersion.Build
$revision = $PSVersionTable.PSVersion.Revision

# Combine them into a single version string
$PowerShellVersion = "$major.$minor.$build.$revision"

<# Determine if build version is considered too old
# Adjust the threshold as needed
$BuildThreshold = 15000
$IsOldBuild = $build -lt $BuildThreshold
#>

# Set Execution Policy to Bypass for the current process
# Set-ExecutionPolicy Bypass -Scope Process -Force

# Set Silent Mode based on the -Silent parameter
# $SilentMode = [bool]$Silent

# ===================================================================================================
#                                SECTION: Import Necessary Assemblies
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports required .NET assemblies for creating and managing Windows Forms and graphical components.
# ===================================================================================================

if (-not $SilentMode)
{
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
}

# ===================================================================================================
#                                   SECTION: Initialize Variables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Initializes all necessary variables required for the script's operation.
# ===================================================================================================

# Declare the script hash table to store results from functions
$script:FunctionResults = @{ }

# Initialize processed items lists with script scope
$script:ProcessedLanes = @()
$script:ProcessedStores = @()
$script:ProcessedServers = @()
$script:ProcessedHosts = @()

# Initialize counts
$NumberOfLanes = 0
$NumberOfStores = 0
$NumberOfServers = 0
$NumberOfHosts = 0

# "ANSI" on Western Windows systems
$ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)

# Create a UTF8 encoding instance without BOM
$utf8NoBOM = New-Object System.Text.UTF8Encoding($false)

# Initialize BasePath variable
$BasePath = $null

# 1) Look for local storeman* directories containing Startup.ini
$storemanDirs = Get-ChildItem -Path "$env:SystemDrive\" -Directory -Filter "*storeman*" -ErrorAction SilentlyContinue |
Where-Object { Test-Path -Path (Join-Path $_.FullName 'Startup.ini') }

if ($storemanDirs)
{
	if ($storemanDirs.Count -gt 1)
	{
		# Prefer one that is actually shared
		$shares = Get-SmbShare -ErrorAction SilentlyContinue
		foreach ($dir in $storemanDirs)
		{
			if ($shares.Path -contains $dir.FullName)
			{
				$BasePath = $dir.FullName
				break
			}
		}
		# If still none, pick the first
		if (-not $BasePath)
		{
			$BasePath = $storemanDirs[0].FullName
		}
	}
	else
	{
		# Only one candidate
		$BasePath = $storemanDirs[0].FullName
	}
}

# 2) If no local match, try UNC paths that contain Startup.ini
if (-not $BasePath)
{
	$uncCandidates = @(
		"\\localhost\storeman",
		"\\$env:COMPUTERNAME\storeman"
	)
	foreach ($path in $uncCandidates)
	{
		if (Test-Path -Path (Join-Path $path 'Startup.ini') -PathType Leaf)
		{
			$BasePath = $path
			break
		}
	}
}

# 3) Final fallback: C:\storeman only if it has Startup.ini
if (-not $BasePath)
{
	$fallback = "$env:SystemDrive\storeman"
	if (Test-Path -Path (Join-Path $fallback 'Startup.ini') -PathType Leaf)
	{
		$BasePath = $fallback
	}
	else
	{
		Throw "Could not locate a storeman folder containing Startup.ini."
	}
}

Write-Host "Selected (storeman) folder: '$BasePath'" -ForegroundColor Magenta

# Now define the rest of your paths
$OfficePath = Join-Path $BasePath "office"
$LoadPath = Join-Path $OfficePath "Load"
$StartupIniPath = Join-Path $BasePath "Startup.ini"
$SystemIniPath = Join-Path $OfficePath "system.ini"
$GasInboxPath = "$OfficePath\XchGAS\INBOX"

# Temp Directory
$TempDir = [System.IO.Path]::GetTempPath()

# SQI Location variables
$LanesqlFilePath = "$env:TEMP\Lane_Database_Maintenance.sqi"
$StoresqlFilePath = "$env:TEMP\Server_Database_Maintenance.sqi"

# Script Name
# $scriptName = Split-Path -Leaf $PSCommandPath

# Add the MailSlotSender to send MailSlot messages to the lanes
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
#                                   FUNCTION: Download-AndRelaunchSelf
# ---------------------------------------------------------------------------------------------------
# Description:
#   This function downloads a specified PowerShell script from a given URL, saves it to a designated
#   directory (defaulting to the system's temporary folder) with ANSI encoding, and relaunches the
#   downloaded script with elevated (Administrator) privileges in a hidden window. It includes
#   error handling to log any issues encountered during the download or relaunch processes. To
#   prevent infinite loops, an explicit relaunch indicator is used. If the download fails, the
#   function logs the error and allows the main script to continue executing without performing
#   further actions within the function.
# ===================================================================================================

function Download-AndRelaunchSelf
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$ScriptUrl,
		[Parameter(Mandatory = $false)]
		[string]$DestinationDirectory = "$env:TEMP",
		[Parameter(Mandatory = $false)]
		[string]$ScriptName = "TBS_Maintenance_Script.ps1",
		[switch]$IsRelaunched
	)
	
	Write-Host "Entering Download-AndRelaunchSelf. IsRelaunched: $IsRelaunched"
	
	# If the script has already been relaunched, do not proceed
	if ($IsRelaunched)
	{
		Write-Host "Script has already been relaunched. Exiting function."
		return
	}
	
	# Construct the full path to save the script
	$DestinationPath = Join-Path -Path $DestinationDirectory -ChildPath $ScriptName
	
	# Prevent infinite loop by checking if the script is already running from the destination path
	if ($MyInvocation.MyCommand.Path -ne $null)
	{
		try
		{
			$currentPath = (Resolve-Path $MyInvocation.MyCommand.Path).Path
			$targetPath = (Resolve-Path $DestinationPath).Path
			if ($currentPath -eq $targetPath)
			{
				# Script is already running from the destination path; do not proceed
				Write-Host "Script is already running from $DestinationPath. Exiting function."
				return
			}
		}
		catch
		{
			# If Resolve-Path fails, proceed to download
			Write-Warning "Resolve-Path failed. Proceeding to download."
		}
	}
	
	try
	{
		Write-Host "Attempting to download the script from $ScriptUrl"
		
		# Attempt to download the script content as a string
		$scriptContent = Invoke-RestMethod -Uri $ScriptUrl -UseBasicParsing
		
		# Save the script content with ANSI encoding
		Set-Content -Path $DestinationPath -Value $scriptContent -Encoding Default
		
		# Verify that the script was downloaded and saved successfully
		if (Test-Path $DestinationPath)
		{
			Write-Host "Script downloaded successfully to $DestinationPath with ANSI encoding."
		}
		else
		{
			Write-Error "Script was not downloaded successfully."
			return
		}
	}
	catch
	{
		# Log the error and exit the function without performing further actions
		Write-Error "Failed to download the script from $ScriptUrl. Error: $_"
		return
	}
	
	try
	{
		# Relaunch the downloaded script as Administrator in a hidden window
		
		# Prepare the arguments for the new PowerShell process, including the relaunch indicator
		$arguments = @(
			"-NoProfile"
			"-ExecutionPolicy"
			"Bypass"
			"-File"
			"`"$DestinationPath`""
			"-IsRelaunched"
		)
		
		Write-Host "Starting new process with arguments: $arguments"
		
		# Start the new process with elevated privileges
		Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs
		
		Write-Host "Process started successfully. Exiting current script."
		
		# Exit the current script to prevent multiple instances
		exit
	}
	catch
	{
		# Log any errors that occur during the relaunch process
		Write-Error "Failed to relaunch the script as Administrator. Error: $_"
	}
	finally
	{
		# Exit the current script regardless of success or failure
		Write-Host "Exiting the original script."
		exit
	}
}

# Rest of your script continues here
# Write-Host "Script is running with elevated privileges from $($MyInvocation.MyCommand.Path)"

# ===================================================================================================
#                              FUNCTION: Ensure Administrator Privileges
# ---------------------------------------------------------------------------------------------------
# Description:
#   Ensures that the script is running with administrative privileges. If not, it attempts to restart the script with elevated rights.
# ===================================================================================================

function Ensure-Administrator
{
	# Retrieve the current Windows identity
	$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
	# Create a WindowsPrincipal object with the current identity
	$principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
	
	# Check if the user is not in the Administrator role
	if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
	{
		try
		{
			# Build the argument list
			$arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
			if ($Silent)
			{
				$arguments += " -Silent"
			}
			
			# Create a ProcessStartInfo object
			$psi = New-Object System.Diagnostics.ProcessStartInfo
			$psi.FileName = (Get-Process -Id $PID).Path # Use the same PowerShell executable
			$psi.Arguments = $arguments
			$psi.Verb = 'runas' # Run as administrator
			$psi.UseShellExecute = $true
			$psi.WindowStyle = 'Normal' # Allow the console window to show (temporarily)
			
			# Start the new elevated process
			$process = [System.Diagnostics.Process]::Start($psi)
			exit # Exit the current process after starting the elevated one
		}
		catch
		{
			[System.Windows.Forms.MessageBox]::Show("Failed to elevate to administrator.`r`nError: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			exit 1
		}
	}
	else
	{
		# Elevated, continue execution
		# Optional: Display a message box if needed
		# [System.Windows.Forms.MessageBox]::Show("Running as Administrator.", "Info")
	}
}

# ===================================================================================================
#                                       FUNCTION: Write to Log
# ---------------------------------------------------------------------------------------------------
# Description:
#   Contains function to write messages to the log, either to a GUI log box or to a log file in silent mode.
# ===================================================================================================

function Write-Log
{
	param (
		[string]$Message,
		[string]$Color = "Black",
		[switch]$IncludeTimestamp = $true,
		[string]$LogFilePath = "$env:TEMP\ScriptLog.txt",
		[switch]$AppendToLogFile = $true
	)
	
	# Prepare timestamp if needed
	#	$timestamp = if ($IncludeTimestamp) { Get-Date -Format 'yyyy-MM-dd HH:mm:ss' }
	#	else { "" }
	
	if (-not $SilentMode)
	{
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
			if ($timestamp)
			{
				Write-Log "[$timestamp] $Message" -ForegroundColor $Color
			}
			else
			{
				Write-Log $Message -ForegroundColor $Color
			}
		}
	}
	else
	{
		# In silent mode, write to a log file
		$fullMessage = if ($timestamp) { "$timestamp - $Message" }
		else { "$Message" }
		
		if ($AppendToLogFile)
		{
			Add-Content -Path $LogFilePath -Value $fullMessage
		}
		else
		{
			Set-Content -Path $LogFilePath -Value $fullMessage
		}
	}
}

# ===================================================================================================
#                             FUNCTION: Get-ScriptOrExecutablePath
# ---------------------------------------------------------------------------------------------------
# Description:
#   Determines the full path of the current script or executable. This ensures that the script can
#   accurately locate its own directory whether it's running as a PowerShell script or has been
#   converted to an executable. This is essential for accessing resources relative to the script's
#   location.
# ===================================================================================================

function Get-ScriptOrExecutablePath
{
	try
	{
		Write-Log "Attempting to determine execution context..."  "Yellow"
		
		# 1. Check if running as a PowerShell script via MyInvocation
		if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path)
		{
			Write-Log "Detected execution as a PowerShell script via MyInvocation."  "Green"
			return $MyInvocation.MyCommand.Path
		}
		
		# 2. Check $PSCommandPath, available in PowerShell 3.0+
		if ($PSCommandPath)
		{
			Write-Log "Detected execution as a PowerShell script via `$PSCommandPath."  "Green"
			return $PSCommandPath
		}
		
		# 3. Try Process MainModule.FileName
		try
		{
			$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
			if ($exePath -and (Test-Path $exePath))
			{
				Write-Log "Detected execution as an executable via Process.MainModule.FileName: $exePath"  "Green"
				return $exePath
			}
			else
			{
				Write-Log "Process.MainModule.FileName returned invalid path: $exePath"  "Red"
			}
		}
		catch
		{
			Write-Log "Error using Process.MainModule.FileName: $_"  "Red"
		}
		
		# 4. Try GetExecutingAssembly().Location
		try
		{
			$exePath = [System.Reflection.Assembly]::GetExecutingAssembly().Location
			if ($exePath -and (Test-Path $exePath))
			{
				Write-Log "Detected execution as an executable via GetExecutingAssembly().Location: $exePath"  "Green"
				return $exePath
			}
			else
			{
				Write-Log "GetExecutingAssembly().Location returned invalid path: $exePath"  "Red"
			}
		}
		catch
		{
			Write-Log "Error using GetExecutingAssembly().Location: $_"  "Red"
		}
		
		# 5. Try GetEntryAssembly().Location
		try
		{
			$exePath = [System.Reflection.Assembly]::GetEntryAssembly().Location
			if ($exePath -and (Test-Path $exePath))
			{
				Write-Log "Detected execution as an executable via GetEntryAssembly().Location: $exePath"  "Green"
				return $exePath
			}
			else
			{
				Write-Log "GetEntryAssembly().Location returned invalid path: $exePath"  "Red"
			}
		}
		catch
		{
			Write-Log "Error using GetEntryAssembly().Location: $_"  "Red"
		}
		
		# 6. Check $PSScriptRoot, if available
		if ($PSScriptRoot)
		{
			Write-Log "Detected execution with PSScriptRoot: $PSScriptRoot"  "Green"
			return $PSScriptRoot
		}
		
		# If none of the above worked
		Write-Log "Unable to determine execution context."  "Red"
		return $null
	}
	catch
	{
		Write-Log "Error retrieving script or executable path: $_"  "Red"
		return $null
	}
}

#Determine the full path of the current script using the host
#$scriptPath = if ($HostInvocation -ne $null -and $HostInvocation.Path) { $HostInvocation.Path }
#elseif ($PSScriptRoot) { $PSScriptRoot }
#else { $null }

#Ensure that $scriptPath is not null
#if (-not $scriptPath)
#{
#	Write-Host "Unable to determine the script path. Ensure the script is executed from a file or an executable." -ForegroundColor Red
#	exit 1
#}


# ===================================================================================================
#                                       FUNCTION: Get-MemoryInfo
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the total system memory and calculates 25% of it in megabytes.
#   This information can be used for memory-related configurations and optimizations.
# ---------------------------------------------------------------------------------------------------
# Note:
#   Temporarily disabled due to [found a different way to get this].
# ===================================================================================================

function Get-MemoryInfo
{
	$TotalMemoryKB = (Get-CimInstance Win32_OperatingSystem).TotalVisibleMemorySize
	$TotalMemoryMB = [math]::Floor($TotalMemoryKB / 1024)
	$Memory25PercentMB = [math]::Floor($TotalMemoryMB * 0.25)
	return $Memory25PercentMB
}

# ===================================================================================================
#                               FUNCTION: Get-DatabaseConnectionString
# ---------------------------------------------------------------------------------------------------
# Description:
#   Searches for the Startup.ini file in specified locations, extracts the DBNAME value,
#   constructs the connection string, and stores it in a script-level hashtable.
# ===================================================================================================

function Get-DatabaseConnectionString
{
	# Ensure that the FunctionResults hashtable exists at the script level
	if (-not $script:FunctionResults)
	{
		$script:FunctionResults = @{ }
		Write-Log "Initialized script:FunctionResults hashtable." "green"
	}
	
	if ($StartupIniPath -ne $null)
	{
		#	Write-Log "Found Startup.ini at: $startupIniPath" "green"
	}
	
	if (-not $StartupIniPath)
	{
		Write-Log "Startup.ini file not found in any of the expected locations." "red"
		return
	}
	
	# Write-Log "Generating connection string..." "blue"
	
	# Read the Startup.ini file
	try
	{
		$content = Get-Content -Path $StartupIniPath -ErrorAction Stop
		
		# Extract DBSERVER
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }
		if ($dbServerLine)
		{
			$dbServer = $dbServerLine -replace '^DBSERVER=', ''
			$dbServer = $dbServer.Trim()
			if (-not $dbServer)
			{
				Write-Log "DBSERVER entry in Startup.ini is empty. Using 'localhost'." "yellow"
				$dbServer = "localhost"
			}
			else
			{
				#	Write-Log "Found DBSERVER in Startup.ini: $dbServer" "green"
			}
		}
		else
		{
			Write-Log "DBSERVER entry not found in Startup.ini. Using 'localhost'." "yellow"
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
				Write-Log "DBNAME entry in Startup.ini is empty." "red"
				return
			}
			else
			{
				#	Write-Log "Found DBNAME in Startup.ini: $dbName" "green"
			}
		}
		else
		{
			Write-Log "DBNAME entry not found in Startup.ini." "red"
			return
		}
	}
	catch
	{
		Write-Log "Failed to read Startup.ini: $_" "red"
		return
	}
	
	# Store DBSERVER and DBNAME in the FunctionResults hashtable
	$script:FunctionResults['DBSERVER'] = $dbServer
	# Write-Log "Stored DBSERVER in FunctionResults: $dbServer" "green"
	
	$script:FunctionResults['DBNAME'] = $dbName
	# Write-Log "Stored DBNAME in FunctionResults: $dbName" "green"
	
	# Build the connection string
	$ConnectionString = "Server=$dbServer;Database=$dbName;Integrated Security=True;"
	# Optionally, log the constructed connection string (be cautious with sensitive information)
	# Write-Log "Constructed connection string: $ConnectionString" "green"
	
	# Store the connection string in the FunctionResults hashtable
	$script:FunctionResults['ConnectionString'] = $ConnectionString
	
	# Write-Log "Variables ($ConnectionString) stored." "green"
}

# ===================================================================================================
#                                      FUNCTION: Get-StoreNumber
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store number via GUI prompts or configuration files.
#   Stores the result in $script:FunctionResults['StoreNumber'].
# ===================================================================================================

function Get-StoreNumber
{
	param (
		[string]$IniFilePath = "$StartupIniPath",
		[string]$BasePath = "$OfficePath"
	)
	
	# Initialize StoreNumber
	$script:FunctionResults['StoreNumber'] = "N/A"
	
	# Try to retrieve StoreNumber from the startup.ini file
	if (Test-Path $IniFilePath)
	{
		$storeNumber = Select-String -Path $IniFilePath -Pattern "^STORE=" | ForEach-Object {
			$_.Line.Split('=')[1].Trim()
		}
		if ($storeNumber)
		{
			$script:FunctionResults['StoreNumber'] = $storeNumber
			#	Write-Log "Store number found in startup.ini: $storeNumber" "green"
		}
		else
		{
			Write-Log "Store number not found in startup.ini." "yellow"
		}
	}
	else
	{
		Write-Log "INI file not found: $IniFilePath" "yellow"
	}
	
	# **Only proceed to check XF directories if StoreNumber was not found in INI**
	if ($script:FunctionResults['StoreNumber'] -eq "N/A")
	{
		if (Test-Path $BasePath)
		{
			$XFDirs = Get-ChildItem -Path $BasePath -Directory -Filter "XF*"
			foreach ($dir in $XFDirs)
			{
				if ($dir.Name -match "^XF(\d{3})")
				{
					$storeNumber = $Matches[1]
					if ($storeNumber -ne "999")
					{
						$script:FunctionResults['StoreNumber'] = $storeNumber
						Write-Log "Store number found from XF directory: $storeNumber" "green"
						break # Exit loop after finding the store number
					}
				}
			}
			if ($script:FunctionResults['StoreNumber'] -eq "N/A")
			{
				Write-Log "No valid XF directories found in $BasePath" "yellow"
			}
		}
		else
		{
			Write-Log "Base path not found: $BasePath" "yellow"
		}
	}
	
	# Update the storeNumberLabel in the GUI if store number was found without manual input
	if ($script:FunctionResults['StoreNumber'] -ne "")
	{
		if (-not $SilentMode -and $storeNumberLabel -ne $null)
		{
			$storeNumberLabel.Text = "Store Number: $($script:FunctionResults['StoreNumber'])"
			$form.Refresh()
			[System.Windows.Forms.Application]::DoEvents()
		}
		return # Exit function after successful retrieval and GUI update
	}
	
	# Prompt for manual input via GUI
	while (-not $script:FunctionResults['StoreNumber'])
	{
		$inputBox = New-Object System.Windows.Forms.Form
		$inputBox.Text = "Enter Store Number"
		$inputBox.Size = New-Object System.Drawing.Size(300, 150)
		$inputBox.StartPosition = "CenterParent"
		$inputBox.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$inputBox.MaximizeBox = $false
		$inputBox.MinimizeBox = $false
		$inputBox.TopMost = $true
		
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "Please enter the store number (e.g., 1, 12, 123):"
		$label.AutoSize = $true
		$label.Location = New-Object System.Drawing.Point(10, 20)
		$inputBox.Controls.Add($label)
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(10, 50)
		$textBox.Width = 260
		$inputBox.Controls.Add($textBox)
		
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(100, 80)
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$inputBox.AcceptButton = $okButton
		$inputBox.Controls.Add($okButton)
		
		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Text = "Cancel"
		$cancelButton.Location = New-Object System.Drawing.Point(180, 80)
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$inputBox.CancelButton = $cancelButton
		$inputBox.Controls.Add($cancelButton)
		
		$result = $inputBox.ShowDialog()
		
		if ($result -eq [System.Windows.Forms.DialogResult]::OK)
		{
			$input = $textBox.Text.Trim()
			if ($input -match "^\d{1,3}$" -and $input -ne "000")
			{
				# Pad the input with leading zeros to ensure it is 3 digits
				$paddedInput = $input.PadLeft(3, '0')
				$script:FunctionResults['StoreNumber'] = $paddedInput
				Write-Log "Store number entered by user: $paddedInput" "green"
				
				# Update the storeNumberLabel in the GUI
				if (-not $SilentMode -and $storeNumberLabel -ne $null)
				{
					$storeNumberLabel.Text = "Store Number: $input"
					$form.Refresh()
					[System.Windows.Forms.Application]::DoEvents()
				}
				
				break
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Store number must be 1 to 3 digits, numeric, and not '000'.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		}
		else
		{
			Write-Log "Store number input canceled by user." "red"
			exit 1
		}
	}
}

# ===================================================================================================
#                                        FUNCTION: Get-StoreName
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store name from the system.ini file.
#   Stores the result in $script:FunctionResults['StoreName'].
# ===================================================================================================

function Get-StoreName
{
	param (
		[string]$INIPath = "$SystemIniPath"
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
			Write-Log "Store name not found in system.ini." "yellow"
		}
	}
	else
	{
		Write-Log "INI file not found: $INIPath" "yellow"
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
#                                       SECTION: Mode Determination
# ---------------------------------------------------------------------------------------------------
# Description:
#   Determines whether the script is running in Host or Store mode based on the store number
#   and updates the GUI label accordingly.
# ===================================================================================================

function Determine-Mode
{
	param (
		[string]$StoreNumber
	)
	
	# Determine mode based on StoreNumber
	if ($StoreNumber -eq "999")
	{
		$Mode = "Host"
	}
	else
	{
		$Mode = "Store"
	}
	
	# Store the mode in FunctionResults
	$script:FunctionResults['Mode'] = $Mode
	
	# Update the modeLabel in the GUI
	if (-not $SilentMode -and $modeLabel -ne $null)
	{
		$modeLabel.Text = "Processing Mode: $Mode"
		$form.Refresh()
		[System.Windows.Forms.Application]::DoEvents()
	}
	
	# Return the mode
	return $Mode
}

# ===================================================================================================
#                                FUNCTION: Get-LaneDatabaseInfo
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the DB server names and database names for all lanes in TER_TAB.
#   The lane number starts with '0' and excludes '901' and '999'.
#   Constructs a unique connection string for each lane and stores it in $script:FunctionResults['LaneDatabaseInfo'].
# ===================================================================================================

function Get-LaneDatabaseInfo
{
	Write-Log "`r`n=== Starting Get-LaneDatabaseInfo Function ===" "blue"
	
	# Initialize the hashtable in FunctionResults to store the info
	$script:FunctionResults['LaneDatabaseInfo'] = @{ }
	
	# Retrieve the connection string from FunctionResults
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	# If connection string is not available, attempt to get it
	if (-not $ConnectionString)
	{
		Write-Log "Connection string not found. Attempting to generate it..." "yellow"
		Get-DatabaseConnectionString
		$ConnectionString = $script:FunctionResults['ConnectionString']
		if (-not $ConnectionString)
		{
			Write-Log "Unable to generate connection string. Cannot query TER_TAB." "red"
			return
		}
	}
	
	# Query the TER_TAB table to get machine names and lane numbers
	try
	{
		Write-Log "Querying TER_TAB to get machine names for lanes..." "blue"
		
		# Define the SQL query to get all lanes starting with '0' (excluding '901' and '999')
		$query = @"
SELECT F1057 AS LaneNumber,
       F1125 AS MachinePath
FROM TER_TAB
WHERE F1057 LIKE '0%' AND F1057 NOT IN ('8%', '9%')
"@
		
		# If StoreNumber is provided, filter by StoreNumber
		if ($StoreNumber)
		{
			$query += " AND F1056 = '$StoreNumber'"
		}
		
		# Execute the SQL query using Invoke-Sqlcmd
		$queryResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query -ErrorAction Stop
		
		if ($queryResult -ne $null -and $queryResult.Count -gt 0)
		{
			foreach ($row in $queryResult)
			{
				$laneNumber = $row.LaneNumber
				$machinePath = $row.MachinePath
				
				# Extract the machine name from the machine path
				if ($machinePath -match '\\\\([^\\]+)\\')
				{
					$machineName = $matches[1]
					
					# Access \\{MachineName}\storeman\Startup.ini
					$startupIniPath = "\\$machineName\storeman\Startup.ini"
					Write-Log "Attempting to access Startup.ini at $startupIniPath for Machine '$machineName'..." "blue"
					
					if (Test-Path $startupIniPath)
					{
						try
						{
							$content = Get-Content -Path $startupIniPath -ErrorAction Stop
							
							# Extract DBNAME
							$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' }
							if ($dbNameLine)
							{
								$dbName = $dbNameLine -replace '^DBNAME=', ''
								$dbName = $dbName.Trim()
							}
							else
							{
								Write-Log "DBNAME entry not found in Startup.ini for Machine '$machineName'." "red"
								continue
							}
							
							# Extract DBSERVER
							$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }
							if ($dbServerLine)
							{
								$dbServer = $dbServerLine -replace '^DBSERVER=', ''
								$dbServer = $dbServer.Trim()
								if (-not $dbServer)
								{
									Write-Log "DBSERVER entry is empty in Startup.ini for Machine '$machineName'. Defaulting to machine name." "yellow"
									$dbServer = $machineName
								}
							}
							else
							{
								Write-Log "DBSERVER entry not found in Startup.ini for Machine '$machineName'. Defaulting to machine name." "yellow"
								$dbServer = $machineName
							}
							
							# Build the connection string
							$laneConnectionString = "Server=$dbServer;Database=$dbName;Integrated Security=True;"
							
							Write-Log "Found DBNAME '$dbName' and DBSERVER '$dbServer' for Machine '$machineName'." "green"
							
							# Store in FunctionResults with LaneNumber as key
							$script:FunctionResults['LaneDatabaseInfo'][$laneNumber] = @{
								'MachineName'	   = $machineName
								'DBName'		   = $dbName
								'DBServer'		   = $dbServer
								'ConnectionString' = $laneConnectionString
							}
							
						}
						catch
						{
							Write-Log "Failed to read Startup.ini from Machine '$machineName': $_" "red"
						}
					}
					else
					{
						Write-Log "Startup.ini not found at $startupIniPath for Machine '$machineName'." "red"
					}
				}
				else
				{
					Write-Log "Lane #${laneNumber}: Invalid machine path '$machinePath'. Skipping." "yellow"
				}
			}
		}
		else
		{
			Write-Log "No lanes found in TER_TAB." "yellow"
			return
		}
	}
	catch
	{
		Write-Log "Failed to query TER_TAB. Error: $_" "red"
		return
	}
	
	Write-Log "`r`n=== Get-LaneDatabaseInfo Function Completed ===" "blue"
}

# ===================================================================================================
#                           FUNCTION: Retrieve-Nodes
# ---------------------------------------------------------------------------------------------------
# **Purpose:**
#   The `Retrieve-Nodes` function is designed to count various entities within a 
#   system, specifically **hosts**, **stores**, **lanes**, **servers**, and **scales**. It primarily retrieves 
#   these nodes from the `TER_TAB` database table and additional tables as needed. If database access fails, it gracefully falls 
#   back to a file system-based mechanism to obtain the counts. Additionally, the function updates 
#   GUI labels to reflect the current nodes and stores the results in a shared hashtable for use 
#   by other parts of the script. For scales, the function retrieves the IPNetwork information from the 
#   TBS_SCL_ver520 table.
#
# **Parameters:**
#   - `[string]$Mode` (Mandatory)
#       - **Description:** Determines the operational mode of the function.
#         - `"Host"`: Counts the number of hosts and stores.
#         - `"Store"`: Counts the number of servers, lanes, and scales (retrieving IPNetwork data for scales) within a specific store.
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
#       - `$SilentMode`: Determines whether the GUI should be updated.
#       - `$NodesHost`, `$NodesStore`, `$NodesScales`: GUI label controls displaying the counts.
#       - `$form`: GUI form that needs to be refreshed to display updated counts.
#
# **Workflow:**
#   1. **Retrieve Database Connection String:**
#      - Attempts to get the connection string from `FunctionResults`.
#      - If unavailable, calls `Get-DatabaseConnectionString` to generate it.
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
#   The `Retrieve-Nodes` function is a robust PowerShell utility that accurately counts system entities 
#   such as hosts, stores, lanes, servers, and scales. It prioritizes retrieving counts from a database to 
#   ensure accuracy and reliability but includes a fallback mechanism leveraging the file system for 
#   resilience. Additionally, it integrates with a GUI to display real-time counts, stores results 
#   for easy access by other script components, and retrieves IPNetwork information for scales from the TBS_SCL_ver520 table.
# ===================================================================================================

function Retrieve-Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Host", "Store")]
		[string]$Mode,
		[string]$StoreNumber
	)
	
	$HostPath = "$OfficePath"
	$NumberOfLanes = 0
	$NumberOfStores = 0
	$NumberOfHosts = 0
	$NumberOfServers = 0
	$NumberOfScales = 0 # NEW/UPDATED COUNTER FOR SCALES
	
	$LaneContents = @()
	$LaneMachines = @{ }
	$ScaleIPNetworks = @{ } # NEW: Hashtable to store IPNetwork info for scales
	
	# Retrieve the connection string from script variables
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	# If connection string is not available, attempt to get it
	if (-not $ConnectionString)
	{
		Write-Log "Connection string not found. Attempting to generate it..." "yellow"
		Get-DatabaseConnectionString
		$ConnectionString = $script:FunctionResults['ConnectionString']
		if (-not $ConnectionString)
		{
			Write-Log "Unable to generate connection string. Proceeding with fallback mechanism." "red"
			$NodesFromDatabase = $false
		}
		else
		{
			$NodesFromDatabase = $true
		}
	}
	else
	{
		$NodesFromDatabase = $true
	}
	
	# Initialize a flag to check if we successfully got Nodes from the database
	if ($NodesFromDatabase)
	{
		try
		{
			if ($Mode -eq "Host")
			{
				# -- Count Stores (excluding StoreNumber = '999') --
				$queryStores = "SELECT COUNT(DISTINCT F1056) AS StoreCount FROM TER_TAB WHERE F1056 <> '999'"
				try
				{
					$storeResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryStores -ErrorAction Stop
				}
				catch [System.Management.Automation.ParameterBindingException] {
					$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					$storeResult = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryStores -ErrorAction Stop
				}
				
				$NumberOfStores = $storeResult.StoreCount
				
				# -- Check if host exists --
				$queryHost = "SELECT COUNT(*) AS HostCount FROM TER_TAB WHERE F1056 = '999' AND F1057 = '901'"
				try
				{
					$hostResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryHost -ErrorAction Stop
				}
				catch [System.Management.Automation.ParameterBindingException] {
					$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					$hostResult = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryHost -ErrorAction Stop
				}
				
				$NumberOfHosts = if ($hostResult.HostCount -gt 0) { 1 }
				else { 0 }
			}
			elseif ($Mode -eq "Store")
			{
				if (-not $StoreNumber)
				{
					Write-Log "Store number is required in 'Store' mode." "red"
					return
				}
				
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
				
				# Extract machine names and store in hashtable
				foreach ($row in $laneContentsResult)
				{
					$laneNumber = $row.F1057
					$machinePath = $row.F1125
					
					if ($machinePath -match '\\\\([^\\]+)\\')
					{
						$machineName = $matches[1]
						$LaneMachines[$laneNumber] = $machineName
					}
					else
					{
						Write-Log "Lane #${laneNumber}: Invalid machine path '$machinePath'. Skipping machine name capture." "yellow"
					}
				}
				
				#--------------------------------------------------------------------------------
				# 2) Retrieve scales from TER_TAB (count only, no IPNetwork here)
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
				# 3) Retrieve additional scales from TBS_SCL_ver520 (with IPNetwork, IPDevice, ScaleName, ScaleCode)
				#--------------------------------------------------------------------------------
				$queryTbsSclScales = @"
SELECT ScaleCode, ScaleName, IPNetwork, IPDevice
FROM TBS_SCL_ver520
WHERE Active = 'Y'
"@
				try
				{
					$tbsSclScalesResult = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $queryTbsSclScales -ErrorAction Stop
				}
				catch [System.Management.Automation.ParameterBindingException] {
					$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					$tbsSclScalesResult = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryTbsSclScales -ErrorAction Stop
				}
				
				if ($tbsSclScalesResult)
				{
					# Increase the total count by the number of rows returned
					$NumberOfScales += $tbsSclScalesResult.Count
					foreach ($row in $tbsSclScalesResult)
					{
						# Build the full IP address by concatenating IPNetwork and IPDevice
						$fullIP = "$($row.IPNetwork)$($row.IPDevice)"
						# Build a custom object that includes the relevant fields
						$scaleObj = [PSCustomObject]@{
							ScaleCode = $row.ScaleCode
							ScaleName = $row.ScaleName
							FullIP    = $fullIP
							IPNetwork = $row.IPNetwork
							IPDevice  = $row.IPDevice
						}
						# Store the object in the hashtable (using ScaleCode as a unique key)
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
			}
		}
		catch
		{
			Write-Log "Failed to retrieve counts from the database: $_" "yellow"
			$NodesFromDatabase = $false
		}
	}
	
	#--------------------------------------------------------------------------------
	# Fallback: If counts from database failed, use directory-based logic
	#--------------------------------------------------------------------------------
	if (-not $NodesFromDatabase)
	{
		Write-Log "Using fallback mechanism to count items." "yellow"
		
		if ($Mode -eq "Host")
		{
			# Retrieve store directories matching the pattern, excluding XF999901
			$StoreFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF*901" | Where-Object { $_.Name -ne "XF999901" }
			$NumberOfStores = $StoreFolders.Count
			
			# Check for the host server directory
			$NumberOfHosts = if (Test-Path "$HostPath\XF999901") { 1 }
			else { 0 }
		}
		elseif ($Mode -eq "Store")
		{
			if (-not $StoreNumber)
			{
				Write-Log "Store number is required in 'Store' mode." "red"
				return
			}
			
			# Lanes (directories like XF<StoreNumber>0??)
			if (Test-Path $HostPath)
			{
				$LaneFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF${StoreNumber}0??"
				$NumberOfLanes = $LaneFolders.Count
			}
			
			# Scales (directories like XF<StoreNumber>8??)
			if (Test-Path $HostPath)
			{
				$ScaleFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF${StoreNumber}8??"
				$NumberOfScales = $ScaleFolders.Count
			}
			
			# Server
			$NumberOfServers = if (Test-Path "$HostPath\XF${StoreNumber}901") { 1 }
			else { 0 }
			
			# NOTE: If your fallback mechanism needs to read TBS_SCL_ver520 data from somewhere,
			#       add additional fallback logic here as needed.
		}
	}
	
	#--------------------------------------------------------------------------------
	# Final: Create a custom object with the counts
	#--------------------------------------------------------------------------------
	$Nodes = [PSCustomObject]@{
		NumberOfStores  = $NumberOfStores
		NumberOfHosts   = $NumberOfHosts
		NumberOfLanes   = $NumberOfLanes
		NumberOfServers = $NumberOfServers
		NumberOfScales  = $NumberOfScales
		LaneContents    = $LaneContents
		LaneMachines    = $LaneMachines
		ScaleIPNetworks = $ScaleIPNetworks # Contains IPNetwork info for scales from TBS_SCL_ver520
	}
	
	#--------------------------------------------------------------------------------
	# Store counts in FunctionResults
	#--------------------------------------------------------------------------------
	$script:FunctionResults['NumberOfStores'] = $NumberOfStores
	$script:FunctionResults['NumberOfHosts'] = $NumberOfHosts
	$script:FunctionResults['NumberOfLanes'] = $NumberOfLanes
	$script:FunctionResults['NumberOfServers'] = $NumberOfServers
	$script:FunctionResults['NumberOfScales'] = $NumberOfScales
	$script:FunctionResults['LaneContents'] = $LaneContents
	$script:FunctionResults['LaneMachines'] = $LaneMachines
	$script:FunctionResults['ScaleIPNetworks'] = $ScaleIPNetworks
	$script:FunctionResults['Nodes'] = $Nodes
	
	#--------------------------------------------------------------------------------
	# Update the GUI labels if not in silent mode
	#--------------------------------------------------------------------------------
	if (-not $SilentMode -and $NodesHost -ne $null -and $NodesStore -ne $null)
	{
		if ($Mode -eq "Host")
		{
			$NodesHost.Text = "Number of Hosts:  $NumberOfHosts"
			$NodesStore.Text = "Number of Stores: $NumberOfStores"
		}
		else
		{
			$NodesHost.Text = "Number of Servers: $NumberOfServers"
			$NodesStore.Text = "Number of Lanes:   $NumberOfLanes"
			$scalesLabel.Text = "Number of Scales: $NumberOfScales"
			
			# Update the Scales label if it exists
			if ($NodesScales -ne $null)
			{
				$NodesScales.Text = "Number of Scales: $NumberOfScales"
			}
		}
		
		# Refresh the form to display updates
		$form.Refresh()
	}
	
	# Return counts as a custom object
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

function Clear-XEFolder
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$folderPath = "$OfficePath\XE${StoreNumber}901",
		[Parameter(Mandatory = $false)]
		[int]$checkIntervalSeconds = 2
	)
	
	# Attempt to find a valid folder path if the UNC path doesn't work
	if (-not (Test-Path -Path $folderPath))
	{
		$localPaths = @(
			"C:\storeman\office\XE${StoreNumber}901",
			"D:\storeman\office\XE${StoreNumber}901"
		)
		
		$foundPath = $false
		foreach ($localPath in $localPaths)
		{
			if (Test-Path $localPath)
			{
				$folderPath = $localPath
				$foundPath = $true
				Write-Log "UNC path not accessible. Using local path: $localPath" "yellow"
				break
			}
		}
		
		if (-not $foundPath)
		{
			Write-Log "Folder 'XE${StoreNumber}901' was not found on UNC or local paths." "red"
			return
		}
	}
	
	# Function to determine if a file should be kept during initial clearing
	function ShouldKeepFileInitial($file)
	{
		# Do not keep FATAL* files
		if ($file.Name -like 'FATAL*')
		{
			return $false
		}
		
		# Check if it's an S*.??? file
		if ($file.Name -match '^S.*\.\w{3}$')
		{
			# Check file age (not older than 30 days)
			$currentTime = Get-Date
			if (($currentTime - $file.LastWriteTime).TotalDays -gt 30)
			{
				return $false
			}
			
			# Read file contents
			try
			{
				$content = Get-Content -Path $file.FullName -ErrorAction Stop
			}
			catch
			{
				# If we can't read the file for some reason, discard it
				return $false
			}
			
			$fromLine = $content | Where-Object { $_ -like 'From:*' }
			$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
			$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
			$lastRecordedStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
			
			# Check prerequisites:
			# From line: Extract store/lane
			if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
			{
				$fileStoreNumber = $Matches[1]
				$fileLaneNumber = $Matches[2]
			}
			else
			{
				return $false
			}
			
			if ($fileStoreNumber -ne $StoreNumber)
			{
				return $false
			}
			
			# From the original logic, $LaneNumber is derived from the filename. Let's extract it:
			if ($file.Name -match '^S.*\.(\d{3})$')
			{
				$LaneNumber = $Matches[1]
				# Confirm lane number matches that from the 'From' line
				if ($fileLaneNumber -ne $LaneNumber)
				{
					return $false
				}
			}
			else
			{
				return $false
			}
			
			# Subject must be Health
			if (-not ($subjectLine -match 'Subject:\s*(Health)'))
			{
				return $false
			}
			
			# MSG must be "This application is not running."
			if (-not ($msgLine -match 'MSG:\s*This application is not running\.'))
			{
				return $false
			}
			
			# Last recorded status must contain TRANS,<number>
			if (-not ($lastRecordedStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)'))
			{
				return $false
			}
			
			# If we reach this point, all conditions are met
			return $true
		}
		
		# If it doesn't match a qualifying S file, we remove it
		return $false
	}
	
	# Function to determine if a file should be kept during background monitoring
	function ShouldKeepFileBackground($file)
	{
		# Keep all FATAL* files
		if ($file.Name -like 'FATAL*')
		{
			return $true
		}
		
		# Check if it's an S*.??? file
		if ($file.Name -match '^S.*\.\w{3}$')
		{
			# Check file age (not older than 30 days)
			$currentTime = Get-Date
			if (($currentTime - $file.LastWriteTime).TotalDays -gt 30)
			{
				return $false
			}
			
			# Read file contents
			try
			{
				$content = Get-Content -Path $file.FullName -ErrorAction Stop
			}
			catch
			{
				# If we can't read the file for some reason, discard it
				return $false
			}
			
			$fromLine = $content | Where-Object { $_ -like 'From:*' }
			$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
			$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
			$lastRecordedStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
			
			# Check prerequisites:
			# From line: Extract store/lane
			if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
			{
				$fileStoreNumber = $Matches[1]
				$fileLaneNumber = $Matches[2]
			}
			else
			{
				return $false
			}
			
			if ($fileStoreNumber -ne $StoreNumber)
			{
				return $false
			}
			
			# From the original logic, $LaneNumber is derived from the filename. Let's extract it:
			if ($file.Name -match '^S.*\.(\d{3})$')
			{
				$LaneNumber = $Matches[1]
				# Confirm lane number matches that from the 'From' line
				if ($fileLaneNumber -ne $LaneNumber)
				{
					return $false
				}
			}
			else
			{
				return $false
			}
			
			# Subject must be Health
			if (-not ($subjectLine -match 'Subject:\s*(Health)'))
			{
				return $false
			}
			
			# MSG must be "This application is not running."
			if (-not ($msgLine -match 'MSG:\s*This application is not running\.'))
			{
				return $false
			}
			
			# Last recorded status must contain TRANS,<number>
			if (-not ($lastRecordedStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)'))
			{
				return $false
			}
			
			# If we reach this point, all conditions are met
			return $true
		}
		
		# If it doesn't match either FATAL* or a qualifying S file, we remove it
		return $false
	}
	
	# Initial clearing - Delete all files including FATAL*
	if (Test-Path -Path $folderPath)
	{
		try
		{
			Get-ChildItem -Path $folderPath -Recurse -Force | ForEach-Object {
				if (-not (ShouldKeepFileInitial $_))
				{
					Remove-Item -Path $_.FullName -Force -Recurse
				}
			}
			
			#	Write-Log "Folder 'XE${StoreNumber}901' initially cleaned, deleting all except valid (S*) files for transaction closing." "green"
		}
		catch
		{
			Write-Log "An error occurred during initial cleaning of 'XE${StoreNumber}901': $_" "red"
		}
	}
	else
	{
		Write-Log "Folder 'XE${StoreNumber}901' (Urgent Messages) does not exist." "red"
		return
	}
	
	# Continuous monitoring in a background job
	try
	{
		$job = Start-Job -Name "ClearXEFolderJob" -ScriptBlock {
			param ($folderPath,
				$checkIntervalSeconds,
				$StoreNumber,
				$OfficePath)
			
			function ShouldKeepFileBackground($file)
			{
				# Keep all FATAL* files
				if ($file.Name -like 'FATAL*')
				{
					return $true
				}
				
				if ($file.Name -match '^S.*\.\w{3}$')
				{
					$currentTime = Get-Date
					if (($currentTime - $file.LastWriteTime).TotalDays -gt 30)
					{
						return $false
					}
					
					try
					{
						$content = Get-Content -Path $file.FullName -ErrorAction Stop
					}
					catch
					{
						return $false
					}
					
					$fromLine = $content | Where-Object { $_ -like 'From:*' }
					$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
					$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
					$lastRecordedStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
					
					if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
					{
						$fileStoreNumber = $Matches[1]
						$fileLaneNumber = $Matches[2]
					}
					else
					{
						return $false
					}
					
					# Extract lane from filename
					if ($file.Name -match '^S.*\.(\d{3})$')
					{
						$LaneNumber = $Matches[1]
						if ($fileLaneNumber -ne $LaneNumber)
						{
							return $false
						}
					}
					else
					{
						return $false
					}
					
					if ($fileStoreNumber -ne $StoreNumber)
					{
						return $false
					}
					
					if (-not ($subjectLine -match 'Subject:\s*(Health)'))
					{
						return $false
					}
					
					if (-not ($msgLine -match 'MSG:\s*This application is not running\.'))
					{
						return $false
					}
					
					if (-not ($lastRecordedStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)'))
					{
						return $false
					}
					
					return $true
				}
				
				return $false
			}
			
			while ($true)
			{
				try
				{
					if (Test-Path -Path $folderPath)
					{
						Get-ChildItem -Path $folderPath -Recurse -Force | ForEach-Object {
							if (-not (ShouldKeepFileBackground $_))
							{
								Remove-Item -Path $_.FullName -Force -Recurse
							}
						}
					}
				}
				catch
				{
					# Suppress any errors
				}
				
				Start-Sleep -Seconds $checkIntervalSeconds
			}
		} -ArgumentList $folderPath, $checkIntervalSeconds, $StoreNumber, $OfficePath
		
		#	Write-Log "Background job 'ClearXEFolderJob' started to continuously monitor and clear 'XE${StoreNumber}901' folder, excluding FATAL* files." "green"
	}
	catch
	{
		Write-Log "Failed to start background job for 'XE${StoreNumber}901': $_" "red"
	}
	
	return $job
}

# ===================================================================================================
#                                  SECTION: Fix-Journal
# ---------------------------------------------------------------------------------------------------
# Description:
#   Processes EJ files within a ZX folder to correct specific lines based on a user-provided date.
#   - Prompts the user to select a date using a Windows Form.
#   - Constructs the ZX folder path.
#   - Identifies related EJ files based on the date/store data.
#   - Repairs lines in matching EJ files.
# ===================================================================================================

function Fix-Journal
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
	
	Write-Log "`r`n==================== Starting Fix-Journal Function ====================`r`n" "blue"
	
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
		Write-Log -Message "Date selection canceled by user. Exiting function." "yellow"
		return
	}
	
	# ---------------------------------------------------------------------------------------------
	# 3) Retrieve and format the selected date
	# ---------------------------------------------------------------------------------------------
	$snippetDate = $dateTimePicker.Value
	$formattedDate = $snippetDate.ToString('MMddyyyy') # MMDDYYYY format
	
	Write-Log -Message "Selected date: $formattedDate" "magenta"
	
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
		Write-Log -Message "ZX folder not found: $zxFolderPath." "red"
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
	
	Write-Log -Message "Looking for files named '$filePrefix.*' in $zxFolderPath..." "blue"
	
	# ---------------------------------------------------------------------------------------------
	# 7) Find matching EJ files in ZX folder: e.g., 41227001.*
	# ---------------------------------------------------------------------------------------------
	$searchPattern = "$filePrefix.*"
	$matchingFiles = Get-ChildItem -Path $zxFolderPath -Filter $searchPattern -File -ErrorAction SilentlyContinue
	
	if (-not $matchingFiles)
	{
		Write-Log -Message "No files matching '$searchPattern' found in $zxFolderPath." "yellow"
		return
	}
	
	Write-Log -Message "Found $($matchingFiles.Count) file(s) to fix." "green"
	
	# ---------------------------------------------------------------------------------------------
	# 8) For each matching EJ file, remove lines from <trs F10... up to <trs F1068...
	# ---------------------------------------------------------------------------------------------
	foreach ($file in $matchingFiles)
	{
		# [Optional] Skip files that have ".bak" anywhere in their name 
		# to avoid infinite backup loops:
		if ($file.Extension -eq ".bak")
		{
			Write-Log -Message "Skipping backup file: $($file.Name)" "yellow"
			continue
		}
		
		Write-Log -Message "Fixing lines in: $($file.FullName)" "yellow"
		
		# Read the file lines
		try
		{
			$originalLines = Get-Content -Path $file.FullName -ErrorAction Stop
		}
		catch
		{
			Write-Log -Message "Failed to read EJ file: $($file.FullName). Skipping." "red"
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
			Write-Log -Message "Backup created: $backupPath" "green"
		}
		catch
		{
			Write-Log -Message "Failed to create backup for: $($file.FullName). Skipping file edit." "red"
			continue
		}
		#>
		
		# -----------------------------------------------------------------------------------------
		# 11) Overwrite the original file with the fixed lines in ANSI encoding
		# -----------------------------------------------------------------------------------------
		try
		{
			$fixedLines | Set-Content -Path $file.FullName -Encoding Default -ErrorAction Stop
			#	Write-Log -Message "Successfully edited: $($file.FullName). Backup: $backupPath" "green"
			Write-Log -Message "Successfully edited: $($file.FullName)" "green"
		}
		catch
		{
			Write-Log -Message "Failed to write fixed content to: $($file.FullName)." "red"
			continue
		}
	}
	Write-Log "`r`n==================== Fix-Journal Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       SECTION: Generate SQL Scripts
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates SQL scripts for Lanes and Stores, including memory configuration and maintenance tasks.
# ===================================================================================================

function Generate-SQLScripts
{
	param (
		[string]$StoreNumber,
		[int]$Memory25PercentMB,
		[string]$LanesqlFilePath,
		[string]$StoresqlFilePath
	)
	
	# Ensure StoreNumber is properly formatted (e.g., '005')
	# $StoreNumber = $StoreNumber.PadLeft(3, '0')
	
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write-Log "Failed to retrieve the connection string." "red"
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
		#	Write-Log "Using DBNAME from FunctionResults: $dbName" "blue"
		$storeDbName = $dbName
	}
	else
	{
		Write-Log "No 'Database' in $script:FunctionResults. Defaulting to '$defaultStoreDbName'." "yellow"
		$storeDbName = $defaultStoreDbName
	}
	
	# Define replacements for SQL scripts
	# $storeDbName is now either the retrieved DBNAME or the default 'STORESQL'
	# $laneDbName remains as 'LANESQL' unless you wish to make it dynamic as well
	$laneDbName = $defaultLaneDbName # If LANESQL is also dynamic, you can retrieve it similarly
	
	# Write-Log "Generating SQL scripts using Store DB: '$storeDbName' and Lane DB: '$laneDbName'..." "blue"
	
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

/* Shrink database and log files */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'
DBCC SHRINKFILE ($laneDbName)
DBCC SHRINKFILE (${laneDbName}_Log)
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

/* Truncate PRICE_EVENT table for records older than 30 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-30,GETDATE());

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
	
	# Write-Log "SQL scripts generated successfully." "green"
}

# ===================================================================================================
#                                       FUNCTION: Get-TableAliases
# ---------------------------------------------------------------------------------------------------
# Description:
#   Parses SQL files in the specified target directory to extract table names and their corresponding
#   aliases from @CREATE statements. This function internally defines the list of base table names to search
#   for and uses regex to identify and capture the full table name and alias pairs. The results are returned
#   as a collection of objects containing details about each match, along with a hash table for quick
#   lookups in both directions (Table Name ? Alias).
# ===================================================================================================

function Get-TableAliases
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
			Write-Log "Could not find load files in either of the specified paths." "red"
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
	$script:FunctionResults['Get-TableAliases'] = @{
		Aliases   = $sortedResults
		AliasHash = $tableAliasHash
	}
}

# ===================================================================================================
#                                       SECTION: Create Scheduled Tasks
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates scheduled tasks to automate the execution of scripts at specified intervals.
#   Uses Write-Log to update the user via GUI after each command runs.
# ===================================================================================================

function Create-ScheduledTaskGUI
{
	param (
		[string]$ScriptPath = $null # Default to null; will be set inside the function if not provided
	)
	
	Write-Log "`r`n=== Starting Create-ScheduledTaskGUI Function ==="  "blue"
	
	# If ScriptPath is not provided, determine it using Get-ScriptOrExecutablePath
	if (-not $ScriptPath)
	{
		$ScriptPath = Get-ScriptOrExecutablePath
		Write-Log "Determined ScriptPath: $ScriptPath"  "blue"
	}
	
	# Ensure that $ScriptPath is not null
	if (-not $ScriptPath)
	{
		Write-Log "Unable to determine the script or executable path."  Red
		exit 1
	}
	
	# Load necessary assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Define the list of possible destination paths in order of preference
	$possiblePaths = @(
		"$BaseUNCPath\Scripts_by_Alex_C.T",
		"C:\Storeman\Scripts_by_Alex_C.T",
		"D:\Storeman\Scripts_by_Alex_C.T"
	)
	
	# Initialize $destinationPath as null
	$destinationPath = $null
	
	# Prompt the user via GUI
	$createTaskForm = New-Object System.Windows.Forms.Form
	$createTaskForm.Text = "Create Scheduled Task"
	$createTaskForm.Size = New-Object System.Drawing.Size(500, 200)
	$createTaskForm.StartPosition = "CenterScreen"
	$createTaskForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$createTaskForm.MaximizeBox = $false
	$createTaskForm.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Do you want to create a scheduled task to automatically run this script in silent mode?"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(20, 20)
	$createTaskForm.Controls.Add($label)
	
	$yesButton = New-Object System.Windows.Forms.Button
	$yesButton.Text = "Yes"
	$yesButton.Location = New-Object System.Drawing.Point(120, 70)
	$yesButton.Size = New-Object System.Drawing.Size(100, 30)
	$yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
	$createTaskForm.Controls.Add($yesButton)
	
	$noButton = New-Object System.Windows.Forms.Button
	$noButton.Text = "No"
	$noButton.Location = New-Object System.Drawing.Point(280, 70)
	$noButton.Size = New-Object System.Drawing.Size(100, 30)
	$noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
	$createTaskForm.Controls.Add($noButton)
	
	$createTaskForm.AcceptButton = $yesButton
	$createTaskForm.CancelButton = $noButton
	
	$result = $createTaskForm.ShowDialog()
	
	if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
	{
		# Iterate through each path and set $destinationPath to the first available one
		foreach ($path in $possiblePaths)
		{
			if (Test-Path -Path $path)
			{
				$destinationPath = $path
				Write-Log "Using existing path: $destinationPath" "blue"
				break
			}
		}
		
		# If none of the paths exist, attempt to create them in order
		if (-not $destinationPath)
		{
			foreach ($path in $possiblePaths)
			{
				try
				{
					# Attempt to create the directory
					New-Item -Path $path -ItemType Directory -Force | Out-Null
					$destinationPath = $path
					Write-Log "Created and using path: $destinationPath" "blue"
					break
				}
				catch
				{
					Write-Log "Failed to create directory '$path'. Error: $_" "red"
					continue
				}
			}
			
			# If still not set, show an error
			if (-not $destinationPath)
			{
				[System.Windows.Forms.MessageBox]::Show("Unable to create any of the specified destination directories.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return
			}
		}
		
		# Define the destination script path
		$scriptName = Split-Path -Leaf $ScriptPath
		$copiedScript = Join-Path -Path $destinationPath -ChildPath $scriptName
		
		# Copy the script to the destination directory
		try
		{
			Write-Log "Copying script to '$destinationPath' for scheduled execution..." "blue"
			Copy-Item -Path $ScriptPath -Destination $copiedScript -Force
			Write-Log "Script copied successfully." "green"
		}
		catch
		{
			Write-Log "Failed to copy script to '$destinationPath'. Error: $_" "red"
			[System.Windows.Forms.MessageBox]::Show("Failed to copy script to '$destinationPath'. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
		
		# Prompt for the interval in months via GUI
		$intervalForm = New-Object System.Windows.Forms.Form
		$intervalForm.Text = "Set Scheduled Task Interval"
		$intervalForm.Size = New-Object System.Drawing.Size(400, 180)
		$intervalForm.StartPosition = "CenterParent"
		$intervalForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$intervalForm.MaximizeBox = $false
		$intervalForm.MinimizeBox = $false
		
		$intervalLabel = New-Object System.Windows.Forms.Label
		$intervalLabel.Text = "Enter the interval (months) to run the script:"
		$intervalLabel.AutoSize = $true
		$intervalLabel.Location = New-Object System.Drawing.Point(20, 20)
		$intervalForm.Controls.Add($intervalLabel)
		
		$intervalTextBox = New-Object System.Windows.Forms.TextBox
		$intervalTextBox.Location = New-Object System.Drawing.Point(20, 50)
		$intervalTextBox.Width = 340
		$intervalForm.Controls.Add($intervalTextBox)
		
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(80, 100)
		$okButton.Size = New-Object System.Drawing.Size(100, 30)
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$intervalForm.Controls.Add($okButton)
		
		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Text = "Cancel"
		$cancelButton.Location = New-Object System.Drawing.Point(220, 100)
		$cancelButton.Size = New-Object System.Drawing.Size(100, 30)
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$intervalForm.Controls.Add($cancelButton)
		
		$intervalForm.AcceptButton = $okButton
		$intervalForm.CancelButton = $cancelButton
		
		$intervalResult = $intervalForm.ShowDialog()
		
		if ($intervalResult -eq [System.Windows.Forms.DialogResult]::OK)
		{
			$taskIntervalInput = $intervalTextBox.Text.Trim()
			if ($taskIntervalInput -match "^\d+$" -and [int]$taskIntervalInput -ge 1)
			{
				$taskInterval = [int]$taskIntervalInput
				Write-Log "Interval set to $taskInterval months." "green"
			}
			else
			{
				Write-Log "Invalid interval entered. Please enter a positive integer." "red"
				[System.Windows.Forms.MessageBox]::Show("Invalid interval entered. Please enter a positive integer.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return
			}
		}
		else
		{
			Write-Log "Scheduled task creation canceled by user." "yellow"
			return
		}
		
		# Define the scheduled task name
		$taskName = "Truncate_Tables_Task"
		
		# Check if the task already exists
		$existingTask = schtasks.exe /Query /TN $taskName /FO LIST /V 2>$null
		if ($LASTEXITCODE -eq 0)
		{
			$overwrite = [System.Windows.Forms.MessageBox]::Show("A task named '$taskName' already exists. Do you want to overwrite it?", "Task Exists", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
			if ($overwrite -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				Write-Log "Scheduled task creation aborted to prevent overwriting existing task." "yellow"
				return
			}
		}
		
		# Define the action to execute the script using PowerShell
		# Using -WindowStyle Hidden to run silently
		$powerShellPath = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
		if (-not (Test-Path -Path $powerShellPath))
		{
			Write-Log "PowerShell executable not found at '$powerShellPath'." "red"
			[System.Windows.Forms.MessageBox]::Show("PowerShell executable not found at '$powerShellPath'.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
		
		# Construct the schtasks.exe command
		$action = "`"$powerShellPath`" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$copiedScript`" -Silent"
		
		# Define the trigger using schtasks.exe syntax
		# Schedule the task to run monthly on the first day of every $taskInterval months at 1:00 AM
		$monthsInterval = $taskInterval
		$startTime = "01:00"
		
		# Prepare the schtasks.exe command parameters
		$schtasksParams = @(
			"/Create",
			"/TN", $taskName,
			"/TR", $action,
			"/SC", "MONTHLY",
			"/MO", $monthsInterval,
			"/D", "1",
			"/ST", $startTime,
			"/RL", "HIGHEST",
			"/F" # Force the task creation, overwriting if it exists
		)
		
		# Execute the schtasks.exe command
		try
		{
			Write-Log "Creating scheduled task..." "blue"
			schtasks.exe @schtasksParams | Out-Null
			if ($LASTEXITCODE -eq 0)
			{
				Write-Log "Scheduled task '$taskName' created successfully." "green"
				[System.Windows.Forms.MessageBox]::Show("Scheduled task '$taskName' created successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
			else
			{
				Write-Log "Failed to create scheduled task. schtasks.exe exited with code $LASTEXITCODE." "red"
				[System.Windows.Forms.MessageBox]::Show("Failed to create scheduled task. schtasks.exe exited with code $LASTEXITCODE.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		}
		catch
		{
			Write-Log "Failed to create scheduled task using schtasks.exe. Error: $_" "red"
			[System.Windows.Forms.MessageBox]::Show("Failed to create scheduled task using schtasks.exe. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
	}
}

# ===================================================================================================
#                                       FUNCTION: Execute-SQLLocallyGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Executes SQL scripts either from a script content variable or from a SQL file.
#   If the script content variable is present and not empty, it takes precedence.
#   Otherwise, it executes the SQL script from the specified file path.
#   Incorporates enhanced exception handling for ParameterBindingException.
# ===================================================================================================

function Execute-SQLLocallyGUI
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$SqlFilePath,
		[Parameter(Mandatory = $false)]
		[string[]]$SectionsToRun,
		# Switch: if used, show a form with checkboxes for each section
		[Parameter(Mandatory = $false)]
		[switch]$PromptForSections
	)
	
	# Make sure .NET WinForms assemblies are loaded
	[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
	
	# Configuration for retry mechanism
	$MaxRetries = 2
	$RetryDelaySeconds = 5
	$FailedCommandsPath = "$OfficePath\XF${StoreNumber}901\Failed_ServerSQLScript_Sections.sql"
	
	# Attempt to retrieve the SQL script
	$sqlScript = $script:ServerSQLScript
	$dbName = $script:FunctionResults['DBNAME']
	
	if (-not [string]::IsNullOrWhiteSpace($sqlScript))
	{
		Write-Log "Executing SQL script from variable..." "blue"
	}
	elseif ($SqlFilePath)
	{
		# If the script variable is empty, try reading from file
		if (-not (Test-Path $SqlFilePath))
		{
			Write-Log "SQL file not found: $SqlFilePath" "red"
			return
		}
		Write-Log "Executing SQL file: $SqlFilePath" "blue"
		try
		{
			$sqlScript = Get-Content -Path $SqlFilePath -Raw -ErrorAction Stop
		}
		catch
		{
			Write-Log "Failed to read SQL file: $_" "red"
			return
		}
	}
	else
	{
		Write-Log "No SQL script content or file path provided." "red"
		return
	}
	
	# Regex to capture sections: /* SectionName */ ...commands...
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$matches = [regex]::Matches($sqlScript, $sectionPattern)
	
	if ($matches.Count -eq 0)
	{
		Write-Log "No SQL sections found to execute." "red"
		return
	}
	
	# ------------------------------------------------------------------
	# If the user wants a GUI prompt for sections, show the form now.
	# ------------------------------------------------------------------
	if ($PromptForSections)
	{
		# Collect all section names
		$allSectionNames = $matches | ForEach-Object {
			$_.Groups['SectionName'].Value.Trim()
		}
		
		# Show the GUI form with checkboxes
		$SectionsToRun = Show-SectionSelectionForm -SectionNames $allSectionNames
		if (-not $SectionsToRun -or $SectionsToRun.Count -eq 0)
		{
			Write-Log "No sections selected or form was canceled. Aborting execution." "yellow"
			return
		}
	}
	
	# If user did NOT specify sections or prompt, run all by default
	$useSpecificSections = $false
	if ($SectionsToRun -and $SectionsToRun.Count -gt 0)
	{
		$useSpecificSections = $true
	}
	
	# Retrieve the connection string
	if (-not $script:FunctionResults)
	{
		$script:FunctionResults = @{ }
	}
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	# If connection string is not available, attempt to get it
	if (-not $ConnectionString)
	{
		Write-Log "Connection string not found. Attempting to generate it..." "yellow"
		$ConnectionString = Get-DatabaseConnectionString
		if (-not $ConnectionString)
		{
			Write-Log "Unable to generate connection string. Cannot execute SQL script." "red"
			return
		}
	}
	
	# Ensure the connection string contains Encrypt=True and TrustServerCertificate=True.
	if ($ConnectionString -notmatch '(?i)Encrypt\s*=')
	{
		$ConnectionString += ";Encrypt=True"
	}
	if ($ConnectionString -notmatch '(?i)TrustServerCertificate\s*=')
	{
		$ConnectionString += ";TrustServerCertificate=True"
	}
	
	# Optionally, log the connection string for debugging (remove sensitive info if necessary)
	Write-Log "Using connection string: $ConnectionString" "gray"
	
	# Determine if Invoke-Sqlcmd supports the -ConnectionString parameter
	$supportsConnectionString = $false
	try
	{
		$cmd = Get-Command Invoke-Sqlcmd -ErrorAction Stop
		$supportsConnectionString = $cmd.Parameters.Keys -contains 'ConnectionString'
	}
	catch
	{
		Write-Log "Invoke-Sqlcmd cmdlet not found: $_" "red"
		$supportsConnectionString = $false
	}
	
	# Initialize variables for retries
	$retryCount = 0
	$success = $false
	$failedSections = @()
	$failedCommands = @()
	
	while (-not $success -and $retryCount -lt $MaxRetries)
	{
		try
		{
			Write-Log "Starting execution of SQL script. Attempt $($retryCount + 1) of $MaxRetries." "blue"
			
			# Only execute failed sections after the first attempt
			$sectionsToExecute = if ($retryCount -eq 0) { $matches }
			else { $failedSections }
			$failedSections = @() # reset for this iteration
			
			foreach ($match in $sectionsToExecute)
			{
				$sectionName = $match.Groups['SectionName'].Value.Trim()
				$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()
				
				# If user specifically chose sections, skip any not in that selection
				if ($useSpecificSections -and ($SectionsToRun -notcontains $sectionName))
				{
					continue
				}
				
				if ([string]::IsNullOrWhiteSpace($sqlCommands))
				{
					Write-Log "Section '$sectionName' has no commands. Skipping..." "yellow"
					continue
				}
				
				Write-Log "`r`nExecuting section: '$sectionName'" "blue"
				Write-Log "--------------------------------------------------------------------------------"
				Write-Log "$sqlCommands" "orange"
				Write-Log "--------------------------------------------------------------------------------"
				
				try
				{
					if ($supportsConnectionString)
					{
						# Using the connection string that now includes Encrypt and TrustServerCertificate
						Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
					}
					else
					{
						# Parse ServerInstance and Database from ConnectionString
						$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
						$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
						
						# Check if Invoke-Sqlcmd supports the -TrustServerCertificate parameter
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
					Write-Log "Section '$sectionName' executed successfully." "green"
				}
				catch [System.Management.Automation.ParameterBindingException] {
					Write-Log "ParameterBindingException in section '$sectionName'. Attempting fallback." "yellow"
					try
					{
						$server = ($ConnectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
						$database = ($ConnectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
						# Try to use the TrustServerCertificate flag if available
						$cmdParams = (Get-Command Invoke-Sqlcmd).Parameters.Keys
						if ($cmdParams -contains 'TrustServerCertificate')
						{
							Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0 -TrustServerCertificate $true
						}
						else
						{
							Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
						}
						Write-Log "Section '$sectionName' executed successfully with fallback." "green"
					}
					catch
					{
						Write-Log "Error executing section '$sectionName' with fallback: $_" "red"
						$failedSections += $match
						if ($retryCount -eq $MaxRetries - 1)
						{
							$failedCommands += "/* $sectionName */`r`n$sqlCommands`r`n"
						}
					}
				}
				catch
				{
					Write-Log "Error executing section '$sectionName': $_" "red"
					$failedSections += $match
					if ($retryCount -eq $MaxRetries - 1)
					{
						$failedCommands += "/* $sectionName */`r`n$sqlCommands`r`n"
					}
				}
			}
			
			if ($failedSections.Count -eq 0)
			{
				Write-Log "`r`nAll SQL sections executed successfully." "green"
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
			Write-Log "Error during SQL script execution: $_" "red"
			if ($retryCount -lt $MaxRetries)
			{
				Write-Log "Retrying in $RetryDelaySeconds seconds..." "yellow"
				Start-Sleep -Seconds $RetryDelaySeconds
			}
		}
	}
	
	if (-not $success)
	{
		Write-Log "Max retries reached. SQL script execution failed." "red"
		$failedCommandsText = ($failedCommands -join "`r`n") + "`r`n"
		try
		{
			[System.IO.File]::WriteAllText($FailedCommandsPath, $failedCommandsText, $ansiPcEncoding)
			Write-Log "`r`nFailed SQL sections written to: $FailedCommandsPath" "yellow"
			Set-ItemProperty -Path $FailedCommandsPath -Name Attributes -Value (
				(Get-Item $FailedCommandsPath).Attributes -band (-bnot [System.IO.FileAttributes]::Archive)
			)
		}
		catch
		{
			Write-Log "Failed to write failed commands: $_" "red"
		}
	}
	else
	{
		Write-Log "SQL script executed successfully on '$dbName'." "green"
	}
}

# ===================================================================================================
#                              FUNCTION: Delete-Files
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

function Delete-Files
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
			Write-Log "The specified path '$Path' does not exist." "Red"
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
										Write-Log "Excluded: $($matchedItem.FullName)" "Yellow"
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
									Write-Log "Deleted: $($matchedItem.FullName)" "Green"
								}
								catch
								{
									Write-Log "Failed to delete $($matchedItem.FullName). Error: $_" "Red"
								}
							}
						}
					}
					else
					{
						Write-Log "No items matched the pattern: '$filePattern' in '$targetPath'." "Yellow"
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
								Write-Log "Excluded: $($item.FullName)" "Yellow"
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
							Write-Log "Deleted: $($item.FullName)" "Green"
						}
						catch
						{
							Write-Log "Failed to delete $($item.FullName). Error: $_" "Red"
						}
					}
				}
			}
			
			Write-Log "Total items deleted: $deletedCount" "Blue"
			return $deletedCount
		}
		catch
		{
			Write-Log "An error occurred during the deletion process. Error: $_" "Red"
			return $deletedCount
		}
	}
}

# ===================================================================================================
#                                       SECTION: Clean Temp Folder
# ---------------------------------------------------------------------------------------------------
# Description:
#   Cleans the temporary folder by deleting all files and directories within it.
# ===================================================================================================

function Clean-TempFolderGUI
{
	Write-Log "`nClearing temp folder..." "blue"
	$FilesAndDirsDeleted = 0
	
	try
	{
		# Get all items to delete
		$itemsToDelete = Get-ChildItem -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
		
		# Count the items
		$FilesAndDirsDeleted = $itemsToDelete.Count
		
		# Remove the items
		$itemsToDelete | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
		
		Write-Log "Temp folder cleared successfully." "green"
		Write-Log "Files and directories deleted: $FilesAndDirsDeleted" "green"
		
		# Return the count of deleted items
		return $FilesAndDirsDeleted
	}
	catch
	{
		Write-Log "Failed to clear temp folder: $_" "red"
		return $FilesAndDirsDeleted
	}
}

# ===================================================================================================
#                                       SECTION: Final Reporting
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates a final report summarizing processed items and cleanup actions.
# ===================================================================================================

function Final-ReportGUI
{
	param (
		[string]$Mode,
		[int]$FilesAndDirsDeletedfomTEMP
	)
	Write-Log "`r`n******************************************************" "blue"
	Write-Log "*                  Final Report                      *" "blue"
	Write-Log "******************************************************" "blue"
	
	if ($Mode -eq "Host")
	{
		$UniqueHostsProcessed = $script:ProcessedHosts | Select-Object -Unique
		$TotalHostsProcessed = $UniqueHostsProcessed.Count
		$UniqueStoresProcessed = $script:ProcessedStores | Select-Object -Unique
		$TotalStoresProcessed = $UniqueStoresProcessed.Count
		
		$StoresList = if ($UniqueStoresProcessed) { $UniqueStoresProcessed -join ', ' }
		else { 'None' }
		
		Write-Log "* Hosts Processed: $TotalHostsProcessed" "green"
		Write-Log "* Stores Processed: $TotalStoresProcessed" "green"
		Write-Log "* Stores Processed List: $StoresList" "green"
	}
	else
	{
		$UniqueServersProcessed = $script:ProcessedServers | Select-Object -Unique
		$TotalServersProcessed = $UniqueServersProcessed.Count
		$UniqueLanesProcessed = $script:ProcessedLanes | Select-Object -Unique
		$TotalLanesProcessed = $UniqueLanesProcessed.Count
		
		$LanesList = if ($UniqueLanesProcessed) { $UniqueLanesProcessed -join ', ' }
		else { 'None' }
		
		Write-Log "* Servers Processed: $TotalServersProcessed" "green"
		Write-Log "* Lanes Processed: $TotalLanesProcessed" "green"
		Write-Log "* Lanes Processed List: $LanesList" "green"
		# Write-Log "* Lanes Pumped List: $ProcessedLanes" "green"
	}
	
	Write-Log "* Files and directories deleted from Temp folder" "green"
	Write-Log "******************************************************" "blue"
	
	# Display a message box for final report if not in silent mode
	if (-not $SilentMode)
	{
		[System.Windows.Forms.MessageBox]::Show("Process completed. Check the log for details.", "Final Report", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
	}
}

# ===================================================================================================
#                                       FUNCTION: Process-HostGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Copies SQL files to the Host and executes the necessary SQL scripts for maintenance.
# ===================================================================================================

function Process-HostGUI
{
	param (
		[string]$StoresqlFilePath
	)
	
	Write-Log "`r`n=== Starting Host Database Repair ===" "blue"
	
	# Execute the SQL script
	Execute-SQLLocallyGUI -SqlFilePath $StoresqlFilePath
	
	# Add host to processed hosts if not already added
	if (-not ($script:ProcessedHosts -contains "localhost"))
	{
		$script:ProcessedHosts += "localhost"
	}
	
	Write-Log "Host processing completed." "green"
}

# ===================================================================================================
#                                       FUNCTION: Process-StoresGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Processes stores based on user selection obtained from Show-SelectionDialog.
#   Handles specific stores, a range of stores, or all stores.
# ===================================================================================================

function Process-StoresGUI
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoresqlFilePath
	)
	
	# Initialize the base path
	$HostPath = "$OfficePath"
	
	if (-not (Test-Path $HostPath))
	{
		Write-Log "Host path not found: $HostPath" "yellow"
		return
	}
	
	# Get the user's selection
	$selection = Show-SelectionDialog -Mode $Mode
	
	if ($selection -eq $null)
	{
		Write-Log "Store processing canceled by user." "yellow"
		return
	}
	
	$Type = $selection.Type
	$Stores = $selection.Stores
	
	switch ($Type)
	{
		'Specific' {
			Write-Log "`nProcessing Specific Store(s)..." "blue"
			foreach ($StoreNumber in $Stores)
			{
				Process-Store -StoreNumber $StoreNumber -StoresqlFilePath $StoresqlFilePath
			}
		}
		'Range' {
			Write-Log "`nProcessing Range of Stores..." "blue"
			foreach ($StoreNumber in $Stores)
			{
				Process-Store -StoreNumber $StoreNumber -StoresqlFilePath $StoresqlFilePath
				Start-Sleep -Seconds 1
			}
		}
		'All' {
			Write-Log "`nProcessing All Stores..." "blue"
			
			# Get all Store folders matching the pattern, excluding XF999901
			$StoreFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF*901" | Where-Object { $_.Name -ne "XF999901" }
			
			foreach ($folder in $StoreFolders)
			{
				$StoreNumber = $folder.Name.Substring(2, 3)
				Process-Store -StoreNumber $StoreNumber -StoresqlFilePath $StoresqlFilePath
			}
		}
		default {
			Write-Log "Unknown selection type." "red"
		}
	}
	
	Write-Log "`nTotal Stores processed: $($script:ProcessedStores.Count)" "green"
}

# Helper function to process a single store
function Process-Store
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$StoresqlFilePath
	)
	
	$StorePath = "$OfficePath\XF${StoreNumber}901"
	
	if (Test-Path $StorePath)
	{
		Write-Log "Processing Store #$StoreNumber..." "blue"
		Write-Log "Copying 'Server_Database_Maintenance.sqi' to $StorePath..." "blue"
		
		try
		{
			Copy-Item -Path $StoresqlFilePath -Destination "$StorePath\Server_Database_Maintenance.sqi" -Force
			Set-ItemProperty -Path "$StorePath\Server_Database_Maintenance.sqi" -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			Write-Log "Copied successfully to Store #$StoreNumber." "green"
			
			# Add store to processed stores if not already added
			if (-not ($script:ProcessedStores -contains $StoreNumber))
			{
				$script:ProcessedStores += $StoreNumber
			}
		}
		catch
		{
			Write-Log "Failed to copy to Store #{$StoreNumber}: $_" "red"
		}
	}
	else
	{
		Write-Log "Store #$StoreNumber not found at path: $StorePath" "yellow"
	}
}

# ===================================================================================================
#                                       FUNCTION: Process-AllStoresAndHostGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Copies SQL files to all Stores and the Host, then executes the necessary SQL scripts for maintenance.
# ===================================================================================================

function Process-AllStoresAndHostGUI
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoresqlFilePath
	)
	
	# Process all stores without prompting
	Process-AllStores -StoresqlFilePath $StoresqlFilePath
	
	# Process the host
	Process-HostGUI -ServerSQLScript $ServerSQLScrip -StoresqlFilePath $StoresqlFilePath
}

function Process-AllStores
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoresqlFilePath
	)
	
	Write-Log "`r`n=== Starting Process-AllStoresAndHostGUI function ===" "blue"
	
	# Initialize the base path
	$HostPath = "$OfficePath"
	
	if (-not (Test-Path $HostPath))
	{
		Write-Log "Host path not found: $HostPath" "yellow"
		return
	}
	
	# Get all Store folders matching the pattern, excluding XF999901 (the host)
	$StoreFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF*901" | Where-Object { $_.Name -ne "XF999901" }
	
	Write-Log "`nProcessing All Stores..." "blue"
	
	foreach ($folder in $StoreFolders)
	{
		$StoreNumber = $folder.Name.Substring(2, 3)
		Process-Store -StoreNumber $StoreNumber -StoresqlFilePath $StoresqlFilePath
	}
	
	Write-Log "`nTotal Stores processed: $($script:ProcessedStores.Count)" "green"
}

# ===================================================================================================
#                                       FUNCTION: Process-ServerGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Copies SQL files to the Server and executes the necessary SQL scripts for maintenance.
# ===================================================================================================

function Process-ServerGUI
{
	param (
		[string]$StoresqlFilePath
	)
	
	Write-Log "`r`n==================== Starting Server Database Repair ====================`r`n" "blue"
	
	# Execute the SQL script
	Execute-SQLLocallyGUI -SqlFilePath $StoresqlFilePath -PromptForSections
	
	# Add server to processed servers if not already added
	if (-not ($script:ProcessedServers -contains "localhost"))
	{
		$script:ProcessedServers += "localhost"
	}
	
	Write-Log "`r`n==================== Completed Server Database Repair ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Process-LanesGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Processes lanes based on user selection obtained from Show-SelectionDialog.
#   Handles specific lanes, a range of lanes, or all lanes.
#   When "All Lanes" is selected, it attempts to retrieve LaneContents.
#   If LaneContents retrieval fails, it uses the predefined NumberOfLanes variable.
# ===================================================================================================

function Process-LanesGUI
{
	param (
		[string]$LanesqlFilePath,
		[string]$StoreNumber,
		[switch]$ProcessAllLanes
	)
	
	Write-Log "`r`n==================== Starting Process-LanesGUI Function ====================`r`n" "blue"
	
	if (-not (Test-Path $OfficePath))
	{
		Write-Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Get the user's selection
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	
	if ($selection -eq $null)
	{
		Write-Log "Lane processing canceled by user." "yellow"
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
			#	Write-Log "User selected 'All Lanes'. Retrieving LaneContents..." "blue"
			$LaneContents = $script:FunctionResults['LaneContents']
			
			if ($LaneContents -and $LaneContents.Count -gt 0)
			{
				#    Write-Log "Successfully retrieved LaneContents. Processing all lanes." "green"
				$Lanes = $LaneContents
			}
			else
			{
				throw "LaneContents is empty or not available."
			}
		}
		catch
		{
			Write-Log "Failed to retrieve LaneContents: $_. Using NumberOfLanes: $script:FunctionResults['NumberOfLanes']." "yellow"
			# Use the predefined NumberOfLanes to generate lane numbers
			if ($script:FunctionResults['NumberOfLanes'] -gt 0)
			{
				Write-Log "Determined NumberOfLanes: $script:FunctionResults['NumberOfLanes']." "green"
				# Generate an array of lane numbers as zero-padded strings (e.g., '001', '002', ...)
				$Lanes = 1 .. $script:FunctionResults['NumberOfLanes'] | ForEach-Object { $_.ToString("D3") }
			}
			else
			{
				Write-Log "NumberOfLanes is not defined or is zero. Exiting Process-LanesGUI." "red"
				return
			}
		}
	}
	
	# Process lanes based on the type of selection
	switch ($Type)
	{
		'Specific' {
			Write-Log "`r`nProcessing Specific Lane(s)..." "blue"
			foreach ($Number in $Lanes)
			{
				Process-Lane -LaneNumber $Number -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
			}
		}
		'Range' {
			Write-Log "`r`nProcessing Range of Lanes..." "blue"
			foreach ($Number in $Lanes)
			{
				Process-Lane -LaneNumber $Number -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
				Start-Sleep -Seconds 1
			}
		}
		'All' {
			Write-Log "`r`nProcessing All Lanes..." "blue"
			
			foreach ($Number in $Lanes)
			{
				Process-Lane -LaneNumber $Number -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
			}
		}
		default {
			Write-Log "Unknown selection type." "red"
		}
	}
	
	Write-Log "`r`nTotal Lanes processed: $($script:ProcessedLanes.Count)" "green"
	if ($script:ProcessedLanes.Count -gt 0)
	{
		Write-Log "Processed Lanes: $($script:ProcessedLanes -join ', ')" "green"
	}
	
	Write-Log "`r`n==================== Process-LanesGUI Function Completed ====================" "blue"
}

# Helper function to process a single lane
function Process-Lane
{
	param (
		[string]$LaneNumber,
		[string]$LanesqlFilePath,
		[string]$StoreNumber
	)
	
	$LaneLocalPath = "$OfficePath\XF${StoreNumber}${LaneNumber}"
	
	if (Test-Path $LaneLocalPath)
	{
		Write-Log "`r`nProcessing Lane #${LaneNumber}..." "blue"
		Write-Log "Lane path found: $LaneLocalPath" "blue"
		Write-Log "Copying 'Lane_Database_Maintenance.sqi' to Lane..." "blue"
		
		try
		{
			# Copy-Item -Path $LanesqlFilePath -Destination "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Force
			# Write-Log "Copied successfully to Lane #${LaneNumber}." "green"
			Set-Content -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Value $LaneSQLScript -Encoding Ascii
			Set-ItemProperty -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			Write-Log "Created and wrote to file at Lane #${LaneNumber} successfully." "green"
			
			# Add lane to processed lanes if not already added
			if (-not ($script:ProcessedLanes -contains $LaneNumber))
			{
				$script:ProcessedLanes += $LaneNumber
			}
		}
		catch
		{
			# Write-Log "Failed to copy to Lane #${LaneNumber}: $_" "red"
			Write-Log "Failed to created and write to file at Lane #${LaneNumber} successfully." "green"
		}
	}
	else
	{
		Write-Log "`r`nLane #${LaneNumber} not found at path: $LaneLocalPath" "yellow"
	}
}

# ===================================================================================================
#                               FUNCTION: Process-LanesAndServerGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Processes all lanes and the server together.
#   Copies SQL files to all lanes and the server, then executes necessary scripts.
# ===================================================================================================

function Process-LanesAndServerGUI
{
	param (
		[string]$LanesqlFilePath,
		[string]$StoresqlFilePath,
		[string]$StoreNumber
	)
	
	# Process all lanes without prompting
	Process-AllLanes -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
	
	# Process the server
	Process-ServerGUI -StoresqlFilePath $StoresqlFilePath
}

function Process-AllLanes
{
	param (
		[string]$LanesqlFilePath,
		[string]$StoreNumber
	)
	
	Write-Log "`r`n=== Starting Process-LanesAndServerGUI Function ===" "blue"
	
	if (-not (Test-Path $OfficePath))
	{
		Write-Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Get all available lane numbers
	$laneFolders = Get-ChildItem -Path $OfficePath -Directory -Filter "XF${StoreNumber}0*"
	$allLanes = $laneFolders | ForEach-Object {
		$_.Name.Substring($_.Name.Length - 3, 3)
	}
	
	Write-Log "`nProcessing All Lanes..." "blue"
	
	foreach ($Number in $allLanes)
	{
		Process-Lane -LaneNumber $Number -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
	}
	
	Write-Log "`nTotal Lanes processed: $($script:ProcessedLanes.Count)" "green"
}

# ===================================================================================================
#                                       FUNCTION: Repair-Windows
# ---------------------------------------------------------------------------------------------------
# Description:
#   Performs various system maintenance tasks to repair Windows.
#   Updates Windows Defender signatures, runs a full scan, executes DISM commands,
#   runs System File Checker, performs disk cleanup, optimizes all fixed drives by trimming SSDs or defragmenting HDDs,
#   and schedules a disk check.
#   Uses Write-Log to provide updates after each command execution.
# ===================================================================================================

function Repair-Windows
{
	[CmdletBinding()]
	param ()
	
	Write-Log "`r`n==================== Starting Repair-Windows Function ====================`r`n" "blue"
	
	# Import necessary assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Create a confirmation dialog
	$confirmationResult = [System.Windows.Forms.MessageBox]::Show(
		"The Windows repair process will take a long time and will make significant changes to your system. Do you want to proceed?",
		"Confirmation Required",
		[System.Windows.Forms.MessageBoxButtons]::YesNo,
		[System.Windows.Forms.MessageBoxIcon]::Warning
	)
	
	# If the user selects 'No', exit the function
	if ($confirmationResult -ne [System.Windows.Forms.DialogResult]::Yes)
	{
		Write-Log "Windows repair process cancelled by the user." "yellow"
		return
	}
	
	Write-Log "Starting Windows repair process. This might take a while, please wait..." "blue"
	
	# Create a form for selecting operations
	$repairForm = New-Object System.Windows.Forms.Form
	$repairForm.Text = "Select Repair Operations"
	$repairForm.Size = New-Object System.Drawing.Size(400, 400)
	$repairForm.StartPosition = "CenterScreen"
	$repairForm.FormBorderStyle = 'FixedDialog'
	$repairForm.MaximizeBox = $false
	$repairForm.MinimizeBox = $false
	$repairForm.ShowInTaskbar = $false
	
	# Initialize an array to hold all operation checkboxes
	$operationCheckboxes = @()
	
	# Function to update the Run button's enabled state
	function Update-RunButtonState
	{
		$anyChecked = $operationCheckboxes | Where-Object { $_.Checked } | Measure-Object | Select-Object -ExpandProperty Count
		$runButton.Enabled = $anyChecked -gt 0
	}
	
	# Create and configure checkboxes for each operation
	$checkboxDefender = New-Object System.Windows.Forms.CheckBox
	$checkboxDefender.Text = "Windows Defender Update and Scan"
	$checkboxDefender.Location = New-Object System.Drawing.Point(20, 20)
	$checkboxDefender.Size = New-Object System.Drawing.Size(350, 25)
	$repairForm.Controls.Add($checkboxDefender)
	$operationCheckboxes += $checkboxDefender
	
	$checkboxDISM = New-Object System.Windows.Forms.CheckBox
	$checkboxDISM.Text = "Run DISM Commands"
	$checkboxDISM.Location = New-Object System.Drawing.Point(20, 60)
	$checkboxDISM.Size = New-Object System.Drawing.Size(350, 25)
	$repairForm.Controls.Add($checkboxDISM)
	$operationCheckboxes += $checkboxDISM
	
	$checkboxSFC = New-Object System.Windows.Forms.CheckBox
	$checkboxSFC.Text = "Run System File Checker (SFC)"
	$checkboxSFC.Location = New-Object System.Drawing.Point(20, 100)
	$checkboxSFC.Size = New-Object System.Drawing.Size(350, 25)
	$repairForm.Controls.Add($checkboxSFC)
	$operationCheckboxes += $checkboxSFC
	
	$checkboxDiskCleanup = New-Object System.Windows.Forms.CheckBox
	$checkboxDiskCleanup.Text = "Disk Cleanup"
	$checkboxDiskCleanup.Location = New-Object System.Drawing.Point(20, 140)
	$checkboxDiskCleanup.Size = New-Object System.Drawing.Size(350, 25)
	$repairForm.Controls.Add($checkboxDiskCleanup)
	$operationCheckboxes += $checkboxDiskCleanup
	
	$checkboxOptimizeDrives = New-Object System.Windows.Forms.CheckBox
	$checkboxOptimizeDrives.Text = "Optimize Drives"
	$checkboxOptimizeDrives.Location = New-Object System.Drawing.Point(20, 180)
	$checkboxOptimizeDrives.Size = New-Object System.Drawing.Size(350, 25)
	$repairForm.Controls.Add($checkboxOptimizeDrives)
	$operationCheckboxes += $checkboxOptimizeDrives
	
	$checkboxCheckDisk = New-Object System.Windows.Forms.CheckBox
	$checkboxCheckDisk.Text = "Schedule Check Disk"
	$checkboxCheckDisk.Location = New-Object System.Drawing.Point(20, 220)
	$checkboxCheckDisk.Size = New-Object System.Drawing.Size(350, 25)
	$repairForm.Controls.Add($checkboxCheckDisk)
	$operationCheckboxes += $checkboxCheckDisk
	
	# Create a checkbox to select all operations
	$checkboxSelectAll = New-Object System.Windows.Forms.CheckBox
	$checkboxSelectAll.Text = "Select All"
	$checkboxSelectAll.Location = New-Object System.Drawing.Point(20, 260)
	$checkboxSelectAll.Size = New-Object System.Drawing.Size(350, 25)
	$repairForm.Controls.Add($checkboxSelectAll)
	
	# Add event handler for Select All checkbox
	$checkboxSelectAll.Add_CheckedChanged({
			$checked = $checkboxSelectAll.Checked
			foreach ($cb in $operationCheckboxes)
			{
				$cb.Checked = $checked
			}
		})
	
	# Create the Run button
	$runButton = New-Object System.Windows.Forms.Button
	$runButton.Text = "Run"
	$runButton.Location = New-Object System.Drawing.Point(150, 310)
	$runButton.Size = New-Object System.Drawing.Size(100, 30)
	$runButton.Enabled = $false # Initially disabled
	$repairForm.Controls.Add($runButton)
	
	# Add event handlers for each operation checkbox to update Run button state
	foreach ($cb in $operationCheckboxes)
	{
		$cb.Add_CheckedChanged({ Update-RunButtonState })
	}
	
	# Add event handler for the Run button
	$runButton.Add_Click({
			# Determine which operations are selected
			$selectedParams = @{
				Defender	   = $checkboxDefender.Checked
				DISM		   = $checkboxDISM.Checked
				SFC		       = $checkboxSFC.Checked
				DiskCleanup    = $checkboxDiskCleanup.Checked
				OptimizeDrives = $checkboxOptimizeDrives.Checked
				CheckDisk	   = $checkboxCheckDisk.Checked
			}
			
			# Set DialogResult to OK and close the form
			$repairForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$repairForm.Close()
		})
	
	# Show the repair options form as a modal dialog
	$dialogResult = $repairForm.ShowDialog()
	
	# If the user closed the form without clicking Run, cancel the function
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write-Log "Windows repair process cancelled by the user." "yellow"
		return
	}
	
	# Retrieve selected parameters after the form is closed
	$selectedParams = @{
		Defender	   = $checkboxDefender.Checked
		DISM		   = $checkboxDISM.Checked
		SFC		       = $checkboxSFC.Checked
		DiskCleanup    = $checkboxDiskCleanup.Checked
		OptimizeDrives = $checkboxOptimizeDrives.Checked
		CheckDisk	   = $checkboxCheckDisk.Checked
	}
	
	Write-Log "Selected operations will be executed." "blue"
	
	# Update Windows Defender Signatures and run a full scan
	if ($selectedParams.Defender)
	{
		try
		{
			Write-Log "Updating Windows Defender signatures..." "blue"
			& "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -SignatureUpdate -ErrorAction Stop
			Write-Log "Windows Defender signatures updated successfully." "green"
			
			Write-Log "Running Windows Defender full scan..." "blue"
			& "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -Scan -ScanType 2 -ErrorAction Stop
			Write-Log "Windows Defender full scan completed." "green"
		}
		catch
		{
			Write-Log "An error occurred while updating or scanning with Windows Defender: $_" "red"
		}
	}
	else
	{
		Write-Log "Skipping Windows Defender update and scan as per user request." "yellow"
	}
	
	# Run DISM commands
	if ($selectedParams.DISM)
	{
		try
		{
			Write-Log "Running DISM StartComponentCleanup..." "blue"
			DISM /Online /Cleanup-Image /StartComponentCleanup /NoRestart
			Write-Log "DISM StartComponentCleanup completed." "green"
			
			Write-Log "Running DISM RestoreHealth..." "blue"
			DISM /Online /Cleanup-Image /RestoreHealth /NoRestart
			Write-Log "DISM RestoreHealth completed." "green"
		}
		catch
		{
			Write-Log "An error occurred while running DISM commands: $_" "red"
		}
	}
	else
	{
		Write-Log "Skipping DISM operations as per user request." "yellow"
	}
	
	# Run System File Checker
	if ($selectedParams.SFC)
	{
		try
		{
			Write-Log "Running System File Checker (SFC)..." "blue"
			SFC /scannow
			Write-Log "System File Checker completed." "green"
		}
		catch
		{
			Write-Log "An error occurred while running System File Checker: $_" "red"
		}
	}
	else
	{
		Write-Log "Skipping System File Checker as per user request." "yellow"
	}
	
	# Cleanup disk space
	if ($selectedParams.DiskCleanup)
	{
		try
		{
			Write-Log "Running Disk Cleanup..." "blue"
			# Ensure that a cleanup profile is set. You may need to configure /sageset:1 beforehand.
			Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait -ErrorAction Stop
			Write-Log "Disk Cleanup completed." "green"
		}
		catch
		{
			Write-Log "An error occurred while running Disk Cleanup: $_" "red"
		}
	}
	else
	{
		Write-Log "Skipping Disk Cleanup as per user request." "yellow"
	}
	
	# Optimize All Fixed Drives
	if ($selectedParams.OptimizeDrives)
	{
		try
		{
			Write-Log "Starting disk optimization for all fixed drives..." "blue"
			
			Get-Volume | Where-Object { $_.DriveType -eq 'Fixed' -and $_.DriveLetter } | ForEach-Object {
				Write-Log "Optimizing drive: $($_.DriveLetter)" "blue"
				Optimize-Volume -DriveLetter $_.DriveLetter -Verbose
			}
			
			Write-Log "Disk optimization for all fixed drives completed." "green"
		}
		catch
		{
			Write-Log "An error occurred while optimizing drives: $_" "red"
		}
	}
	else
	{
		Write-Log "Skipping disk optimization as per user request." "yellow"
	}
	
	# Schedule Check Disk
	if ($selectedParams.CheckDisk)
	{
		try
		{
			Write-Log "Scheduling Check Disk on C: drive..." "blue"
			# Automatically confirm the disk check and handle the need for a reboot
			Start-Process "cmd.exe" -ArgumentList "/c echo Y|chkdsk C: /f /r" -Verb RunAs -Wait -ErrorAction Stop
			Write-Log "Check Disk scheduled. A restart may be required to complete the process." "green"
		}
		catch
		{
			Write-Log "An error occurred while scheduling Check Disk: $_" "red"
		}
	}
	else
	{
		Write-Log "Skipping Check Disk scheduling as per user request." "yellow"
	}
	Write-Log "`r`n==================== Repair-Windows Function Completed ====================" "blue"
}

# ===================================================================================================
#                                     FUNCTION: Update-LaneFiles
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

function Update-LaneFiles
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Update-LaneFiles Function ====================" "blue"
	
	# Ensure $LoadPath exists
	if (-not (Test-Path $LoadPath))
	{
		Write-Log "`r`nLoad Base Path not found: $LoadPath" "yellow"
		return
	}
	
	# Get the user's lane selection
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	
	if ($selection -eq $null)
	{
		Write-Log "`r`nLane processing canceled by user." "yellow"
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
			Write-Log "Failed to retrieve LaneContents: $_. Falling back to user-selected lanes." "yellow"
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
			Write-Log "`r`nLane #$laneNumber not found at path: $laneFolderPath" "yellow"
			continue
		}
		
		$laneFolder = Get-Item -Path $laneFolderPath
		Write-Log "`r`nProcessing Lane #$laneNumber" "blue"
		
		$actionSummaries = @()
		
		# Determine Machine Name from LaneMachines
		try
		{
			Write-Log "Determining machine name for Lane #$laneNumber..." "blue"
			
			$MachineName = $script:FunctionResults['LaneMachines'][$laneNumber]
			
			if ($MachineName)
			{
				Write-Log "Lane #${laneNumber}: Retrieved machine name '$MachineName' from LaneMachines." "green"
			}
			else
			{
				Write-Log "Lane #${laneNumber}: Machine name not found in LaneMachines. Defaulting to 'POS${laneNumber}'." "yellow"
				$MachineName = "POS${laneNumber}"
			}
		}
		catch
		{
			Write-Log "Lane #${laneNumber}: Error retrieving machine name. Error: $_. Defaulting to 'POS${laneNumber}'." "red"
			$MachineName = "POS${laneNumber}"
		}
		
	<# 
		# Process each load SQL file (currently commented out; uncomment if needed)
		foreach ($file in $loadFiles) 
		{
		    Write-Log "Processing file '$($file.Name)' for Lane #$laneNumber..." "blue"
		    
		    # Read the original file content
		    try 
		    {
		        $originalContent = Get-Content -Path $file.FullName -ErrorAction Stop
		        Write-Log "Successfully read '$($file.Name)'." "green"
		    }
		    catch 
		    {
		        Write-Log "Failed to read '$($file.Name)'. Error: $_" "red"
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
		    
		            Write-Log "Successfully copied to '$destinationPath'." "green"
		            $actionSummaries += "Copied $($file.Name)"
		        }
		        catch 
		        {
		            Write-Log "Failed to copy '$($file.Name)' to '$destinationPath'. Error: $_" "red"
		            $actionSummaries += "Failed to copy $($file.Name)"
		        }
		    }
		    else 
		    {
		        Write-Log "No matching records found in '$($file.Name)' for Lane #$laneNumber." "yellow"
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
		Write-Log $summaryMessage "green"
		
		# Mark lane as processed
		if (-not ($script:ProcessedLanes -contains $laneNumber))
		{
			$script:ProcessedLanes += $laneNumber
		}
	}
	
	Write-Log "`r`n==================== Update-LaneFiles Function Completed ====================" "blue"
}

# ===================================================================================================
#                                 FUNCTION: Deploy_Load
# ---------------------------------------------------------------------------------------------------
# Description:
#   Lets you pick a lane (for @TER) using Show-SelectionDialog, then writes a ready-to-execute macro
#   for UD_DEPLOY_LOAD with ACTION always set to ADDRPL, all scenario and business logic blocks included.
#   Deploys the SQI macro file directly to XF<Store>901.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: The store number to process. (Mandatory)
# ---------------------------------------------------------------------------------------------------
# Requirements:
#   - Show-SelectionDialog function must exist and return lane (TER).
#   - $OfficePath must be defined.
#   - Write-Log function must be available.
# ===================================================================================================

function Deploy_Load
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Deploy_UD_DEPLOY_LOAD ====================`r`n" "blue"
	
	# ---- STEP 1: Pick lane (TER) ----
	$selection = Show-SelectionDialog -Mode "Store" -StoreNumber $StoreNumber
	if ($null -eq $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write-Log "No lane selected or operation cancelled." "yellow"
		Write-Log "`r`n==================== Deploy_UD_DEPLOY_LOAD Completed ====================" "blue"
		return
	}
	$TER = $selection.Lanes[0].PadLeft(3, '0')
	
	# ---- STEP 2: Build the macro file ----
	$MacroContent = @"
/* GET THE SCENARIO SWITCH */
@wizRpl(SCENARIO_SWITCH=@dbHot(INI,SAMPLES.INI,SWITCHES,DEPLOY_SCENARIO));

/* NOT PERMITTED MESSAGES */
@wizRpl(SCENARIO_MSG=A @WIZGET(SCENARIO_SWITCH) sample scenario is installed and does not allow using this script);
@fmt(CMP,'@dbHot(LANGUAGE)=ES','®wizRpl(SCENARIO_MSG=Se instala un escenario de muestra @WIZGET(SCENARIO_SWITCH) y no permite usar este script)');
@fmt(CMP,'@dbHot(LANGUAGE)=FR',"®wizRpl(SCENARIO_MSG=Un sample scenario @WIZGET(SCENARIO_SWITCH) est installé et ne permet pas d'utiliser ce script)");

/* CHECK TO SEE IF OPERATION IS PERMITTED BASED ON SCENARIO */
@fmt(CMP,'@STORE@wizGet(SCENARIO_SWITCH)=999STORE','®wizRpl(SCENARIO_EXIT=1)','®wizClr(SCENARIO_EXIT)®wizClr(SCENARIO_MSG)');
@FMT(CMP,'@dbHot(INI,SAMPLES.INI,SWITCHES,LOC_SUBSAMPLE)=LOC_BACK','®wizClr(SCENARIO_EXIT)®wizClr(SCENARIO_MSG)');

@fmt(CMP,'@wizGet(SCENARIO_EXIT)=1',"®wizRpl(OK=®tools(MESSAGEDLG,'!@wizGet(SCENARIO_MSG)',10,,OK))®fmt(CHR,27)");

/* REDIRECT THE BATCH EXECUTION BASED ON SCHEDULER */
@FMT(CMP,@WIZEXIST(BATCH_FILENAME)=1,"®WIZRPL(REDIRECT=®DBSELECT(SELECT LNK.F1056+LNK.F1057 FROM RUN_TAB RUN JOIN LNK_TAB LNK ON LNK.F1000=RUN.F1000 WHERE RUN.F1103 LIKE '%SQL=DEPLOY_CHG%' AND LNK.F1056+LNK.F1057<>'@STORE@TER'))");
@fmt(CMP,@wizIsBlank(REDIRECT)=1,'®wizClr(REDIRECT)','®wizRpl(SRC_PATH=®OFFICEXF®STORE®TER\®WIZGET(BATCH_FILENAME))®wizRpl(TAR_PATH=®OFFICEXF®wizGet(REDIRECT)\®wizGet(BATCH_FILENAME))®exec(XCH=COPYFILE)®fmt(CHR,26)');

/* Set Variables */
@WIZINIT;
@WIZTARGET(DEPLOYLOAD.STORE=$TER,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1","
SELECT DISTINCT STO.F1000,STO.F1018 
FROM LNK_TAB LN2 
JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000 
WHERE STO.F1181='1' AND STO.F1000<>'XAL' AND 
LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'"));
@WIZTARGET(TARGET_FILTER=$TER,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1","
SELECT DISTINCT STO.F1000,STO.F1018 
FROM LNK_TAB LN2 
JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000 
WHERE STO.F1181='1' AND STO.F1000<>'XAL' AND 
LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'"));
@WIZRPL(ACTION=ADDRPL(ACTION_CHOICE));
@WIZRPL(ONESQM=ALL);
@WIZDATE(DATE=@DSSF,TIME=@NOW);
@WIZRPL(DBASE_TIMEOUT=E);

/* TABLES WITH F1000 */
@FMT(CMP,@WIZGET(clk_load)=0,,®EXEC(SQM=clk_load));
@FMT(CMP,@WIZGET(clt_load)=0,,®EXEC(SQM=clt_load));
@FMT(CMP,@WIZGET(cll_load)=0,,®EXEC(SQM=cll_load));
@FMT(CMP,@WIZGET(pos_load)=0,,®EXEC(SQM=pos_load));
@FMT(CMP,@WIZGET(price_load)=0,,®EXEC(SQM=price_load));
@FMT(CMP,@WIZGET(cost_load)=0,,®EXEC(SQM=cost_load));
@FMT(CMP,@WIZGET(dsd_load)=0,,®EXEC(SQM=dsd_load));
@FMT(CMP,@WIZGET(ecl_load)=0,,®EXEC(SQM=ecl_load));
@FMT(CMP,@WIZGET(scl_load)=0,,®EXEC(SQM=scl_load));
@FMT(CMP,@WIZGET(scl_txt_load)=0,,®EXEC(SQM=scl_txt_load));
@FMT(CMP,@WIZGET(scl_nut_load)=0,,®EXEC(SQM=scl_nut_load));
@FMT(CMP,@WIZGET(scl_cpt_load)=0,,®EXEC(SQM=scl_cpt_load));
@FMT(CMP,@WIZGET(scl_cct_load)=0,,®EXEC(SQM=scl_cct_load));
@FMT(CMP,@WIZGET(scl_csl_load)=0,,®EXEC(SQM=scl_csl_load));
@FMT(CMP,@WIZGET(scl_ctx_load)=0,,®EXEC(SQM=scl_ctx_load));
@FMT(CMP,@WIZGET(scl_sto_load)=0,,®EXEC(SQM=scl_sto_load));
@FMT(CMP,@WIZGET(loc_load)=0,,®EXEC(SQM=loc_load));
@FMT(CMP,@WIZGET(alt_load)=0,,®EXEC(SQM=alt_load));
@FMT(CMP,@WIZGET(itz_load)=0,,®EXEC(SQM=itz_load));
@FMT(CMP,@WIZGET(itd_load)=0,,®EXEC(SQM=itd_load));
@FMT(CMP,@WIZGET(cls_load)=0,,®EXEC(SQM=cls_load));

@WIZRPL(TARGET=@WIZGET(DEPLOYLOAD.STORE));
@WIZRPL(TARGET_FILTER=@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE));

/* TABLES WITHOUT F1000 */
@FMT(CMP,@WIZGET(mix_load)=0,,®EXEC(SQM=mix_load));
@FMT(CMP,@WIZGET(bio_load)=0,,®EXEC(SQM=bio_load));
@FMT(CMP,@WIZGET(vendor_load)=0,,®EXEC(SQM=vendor_load));
@FMT(CMP,@WIZGET(obj_load)=0,,®EXEC(SQM=obj_load));
@FMT(CMP,@WIZGET(kit_load)=0,,®EXEC(SQM=kit_load));
@FMT(CMP,@WIZGET(clt_itm_load)=0,,®EXEC(SQM=clt_itm_load));
@FMT(CMP,@WIZGET(bmp_load)=0,,®EXEC(SQM=bmp_load));

 				/* DEPLOY_AUX */

/* SAVE ALL PARAMETERS */
@DBHOT(HOT_WIZ,PARAMTOLINE,PARAMSAV_DEPLOYAUX);

/* DEPLOY WITHOUT F1000 */
@FMT(CMP,@wizget(sdp_load)=0,,®EXEC(SQM=sdp_load));
@FMT(CMP,@wizget(dept_load)=0,,®EXEC(SQM=dept_load));
@FMT(CMP,@wizget(cat_load)=0,,®EXEC(SQM=cat_load));
@FMT(CMP,@wizget(fam_load)=0,,®EXEC(SQM=fam_load));
@FMT(CMP,@wizget(rpc_load)=0,,®EXEC(SQM=rpc_load));
@FMT(CMP,@wizget(btl_load)=0,,®EXEC(SQM=btl_load));
@FMT(CMP,@wizget(tar_load)=0,,®EXEC(SQM=tar_load));
@FMT(CMP,@wizget(lvl_load)=0,,®EXEC(SQM=lvl_load));
@FMT(CMP,@wizget(reason_load)=0,,®EXEC(SQM=reason_load));
@FMT(CMP,@wizget(res_load)=0,,®EXEC(SQM=res_load));
@FMT(CMP,@wizget(cpn_load)=0,,®EXEC(SQM=cpn_load));
@FMT(CMP,@wizget(clf_load)=0,,®EXEC(SQM=clf_load));
@FMT(CMP,@wizget(clg_load)=0,,®EXEC(SQM=clg_load));
@FMT(CMP,@wizget(clr_load)=0,,®EXEC(SQM=clr_load));
@FMT(CMP,@wizget(clf_sdp_load)=0,,®EXEC(SQM=clf_sdp_load));
@FMT(CMP,@wizget(mod_load)=0,,®EXEC(SQM=mod_load));
@FMT(CMP,@wizget(mod_itm_load)=0,,®EXEC(SQM=mod_itm_load));
@FMT(CMP,@wizget(like_load)=0,,®EXEC(SQM=like_load));
@FMT(CMP,@wizget(rcp_load)=0,,®EXEC(SQM=rcp_load));
@FMT(CMP,@wizget(rcp_det_load)=0,,®EXEC(SQM=rcp_det_load));
@FMT(CMP,@wizget(rcp_itm_load)=0,,®EXEC(SQM=rcp_itm_load));
@FMT(CMP,@wizget(label_tpl_load)=0,,®EXEC(SQM=Label_Tpl_load));
@FMT(CMP,@wizget(std_load)=0,,®EXEC(SQM=std_load));
@FMT(CMP,@wizget(unt_load)=0,,®EXEC(SQM=unt_load));
@FMT(CMP,@wizget(route_load)=0,,®EXEC(SQM=route_load));
@FMT(CMP,@wizget(cls_aux_load)=0,,®EXEC(SQM=cls_aux_load));

/* DEPLOY USING F1000 */
@WIZRESET; 
@DBHOT(HOT_WIZ,LINETOPARAM,PARAMSAV_DEPLOYAUX);
@WIZRPL(TARGET_FILTER=@WIZGET(TARGET));
@WIZCLR(TARGET);

@FMT(CMP,@wizget(cfg_load)=0,,®EXEC(SQM=cfg_load));
@FMT(CMP,@wizget(shf_load)=0,,®EXEC(SQM=shf_load));
@FMT(CMP,@wizget(delv_load)=0,,®EXEC(SQM=delv_load));

/* RESTORE TARGET */
@WIZRESET; 
@DBHOT(HOT_WIZ,LINETOPARAM,PARAMSAV_DEPLOYAUX);
@DBHOT(HOT_WIZ,CLR,PARAMSAV_DEPLOYAUX);

 				/* DEPLOY_SYS */

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

@FMT(CMP,@WIZGET(exe_activate_accept)=0,,®EXEC(SQM=exe_activate_accept));
@FMT(CMP,@wizget(exe_activate_accept_all)=0,,®EXEC(SQM=exe_activate_accept_aux));
@FMT(CMP,@wizget(exe_activate_accept_all)=0,,®EXEC(SQM=exe_activate_accept_sys));
@FMT(CMP,@wizget(exe_refresh_menu)=1,®EXEC(SQM=exe_refresh_menu));
@FMT(CMP,@WIZGET(exe_deploy_chg)=1,®EXEC(SQM=exe_deploy_chg));

/* KEEP THE LONG TIMEOUT FOR OTHER DEPLOY_LOAD */
@FMT(CMP,@WIZGET(UD_RUN)=0,®WIZCLR(DBASE_TIMEOUT));
"@
	
	# ---- STEP 3: Write file with forced CRLF and ANSI encoding ----
	$DeployPath = Join-Path -Path $OfficePath -ChildPath "XF${StoreNumber}901"
	if (-not (Test-Path $DeployPath))
	{
		Write-Log "Deploy path $DeployPath not found. Aborting." "red"
		return
	}
	$MacroFile = Join-Path -Path $DeployPath -ChildPath "UD_DEPLOY_LOAD.sqi"
	
	# CRLF normalization
	$MacroContentCRLF = $MacroContent -replace "`r?`n", "`r`n"
		
	[System.IO.File]::WriteAllText($MacroFile, $MacroContentCRLF, $ansiPcEncoding)
	Set-ItemProperty -Path $MacroFile -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
	
	Write-Log "Deployed UD_DEPLOY_LOAD macro for lane $TER in $DeployPath." "green"
	Write-Log "`r`n==================== Deploy_UD_DEPLOY_LOAD Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Pump-AllItems
# ---------------------------------------------------------------------------------------------------
# Description:
#   Extracts specified tables from the SQL Server and copies them to the selected lanes as a single
#   batch file named PUMP_ALL_ITEMS_TABLES.sql. Handles large tables efficiently while preserving
#   compatibility. This function now relies on Get-TableAliases to retrieve the table and alias
#   information, eliminating the need to define the table list within this function.
#
#   This version enforces:
#       - Windows-1252 ("ANSI") encoding (no BOM)
#       - CRLF line endings (`\r\n`)
# ===================================================================================================

function Pump-AllItems
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Pump-AllItems Function ====================`r`n" "blue"
	
	if (-not (Test-Path $OfficePath))
	{
		Write-Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Get the user's selection
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	
	if ($selection -eq $null)
	{
		Write-Log "Lane processing canceled by user." "yellow"
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
	# Fetch the alias data that Get-TableAliases stored
	# --------------------------------------------------------------------------------------------
	if ($script:FunctionResults.ContainsKey('Get-TableAliases'))
	{
		$aliasData = $script:FunctionResults['Get-TableAliases']
		$aliasResults = $aliasData.Aliases
		$aliasHash = $aliasData.AliasHash
	}
	else
	{
		Write-Log "Alias data not found. Ensure Get-TableAliases has been run." "red"
		return
	}
	
	if ($aliasResults.Count -eq 0)
	{
		Write-Log "No tables found to process. Exiting Pump-AllItems." "red"
		return
	}
	
	# --------------------------------------------------------------------------------------------
	# Get the SQL ConnectionString from script-scoped results
	# --------------------------------------------------------------------------------------------
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write-Log "Connection string not found. Cannot proceed with Pump-AllItems." "red"
		return
	}
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	# Open SQL connection
	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$sqlConnection.ConnectionString = $ConnectionString
	$sqlConnection.Open()
	
	# Lists to keep track of processed tables and files
	$generatedFiles = @()
	$copiedTables = @()
	$skippedTables = @()
	
	# --------------------------------------------------------------------------------------------
	# Process each table in $aliasResults
	# --------------------------------------------------------------------------------------------
	foreach ($aliasEntry in $aliasResults)
	{
		$table = $aliasEntry.Table # e.g. "XYZ_TAB"
		$tableAlias = $aliasEntry.Alias # e.g. "XYZ"
		
		if (-not $table -or -not $tableAlias)
		{
			Write-Log "Invalid table or alias: $($aliasEntry | ConvertTo-Json)" "yellow"
			continue
		}
		
		# Check if this table has data
		$dataCheckQuery = "SELECT COUNT(*) FROM [$table]"
		$cmdCheck = $sqlConnection.CreateCommand()
		$cmdCheck.CommandText = $dataCheckQuery
		
		try
		{
			$rowCount = $cmdCheck.ExecuteScalar()
		}
		catch
		{
			Write-Log "Error checking row count for '$table': $_" "red"
			continue
		}
		
		# Skip tables with zero rows
		if ($rowCount -eq 0)
		{
			$skippedTables += $table
			continue
		}
		
		Write-Log "Processing table '$table'..." "blue"
		
		# Remove "_TAB" suffix for the base name
		$baseTable = $table -replace '_TAB$', ''
		
		# File name for the extracted data
		$sqlFileName = "${baseTable}_Load.sql"
		$localTempPath = Join-Path $env:TEMP $sqlFileName
		
		# Check for a recent file in TEMP
		$useExistingFile = $false
		if (Test-Path $localTempPath)
		{
			$fileInfo = Get-Item $localTempPath
			$fileAge = (Get-Date) - $fileInfo.LastWriteTime
			if ($fileAge.TotalHours -le 1)
			{
				Write-Log "Recent SQL file found for '$table' in %TEMP%. Using existing file." "green"
				$useExistingFile = $true
			}
			else
			{
				Write-Log "SQL file for '$table' is older than 1 hour. Regenerating." "yellow"
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
				
				# --------------------------------------------------------------------------------
				# 1) Gather column data types
				# --------------------------------------------------------------------------------
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
				
				# --------------------------------------------------------------------------------
				# 2) Retrieve primary key columns
				# --------------------------------------------------------------------------------
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
				
				# --------------------------------------------------------------------------------
				# 3) Generate @CREATE, CREATE VIEW, and INSERT lines
				# --------------------------------------------------------------------------------
				$viewName = $baseTable.Substring(0, 1).ToUpper() + $baseTable.Substring(1).ToLower() + '_Load'
				$columnList = ($columnDataTypes.Keys) -join ','
				
				# Header (normalize any line breaks to CRLF)
				$header = @"
/* Set a long timeout so the entire script runs */
@WIZRPL(DBASE_TIMEOUT=E);

@CREATE($table,$tableAlias);
CREATE VIEW $viewName AS SELECT $columnList FROM $table;

INSERT INTO $viewName VALUES
"@
				# Normalize line endings to CRLF
				$header = $header -replace "(\r\n|\n|\r)", "`r`n"
				$streamWriter.WriteLine($header.TrimEnd()) # .WriteLine() uses CRLF from NewLine
				
				# --------------------------------------------------------------------------------
				# 4) Fetch data from the table
				# --------------------------------------------------------------------------------
				$dataQuery = "SELECT * FROM [$table]"
				$cmdData = $sqlConnection.CreateCommand()
				$cmdData.CommandText = $dataQuery
				$readerData = $cmdData.ExecuteReader()
				
				$firstRow = $true
				while ($readerData.Read())
				{
					# For each row, we gather column values
					if ($firstRow)
					{
						$firstRow = $false
					}
					else
					{
						# For subsequent rows => separate with comma & newline
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
						
						# For string-like columns
						if ($dataType -in @('char', 'nchar', 'varchar', 'nvarchar', 'text', 'ntext'))
						{
							$escapedVal = $val.ToString().Replace("'", "''")
							# Also remove/normalize any embedded newlines
							$escapedVal = $escapedVal -replace "(\r\n|\n|\r)", " "
							$values += "'$escapedVal'"
						}
						elseif ($dataType -in @('datetime', 'smalldatetime', 'date', 'datetime2'))
						{
							# Format as YYYYDDD HH:mm:ss
							$dayOfYear = $val.DayOfYear.ToString("D3")
							$formattedDate = "'{0}{1} {2}'" -f $val.Year, $dayOfYear, $val.ToString("HH:mm:ss")
							$values += $formattedDate
						}
						elseif ($dataType -eq 'bit')
						{
							$bitVal = if ($val) { "1" }
							else { "0" }
							$values += $bitVal
						}
						elseif ($dataType -in @('decimal', 'numeric', 'float', 'real', 'money', 'smallmoney'))
						{
							# Numeric: either integer or decimal
							if ([math]::Floor($val) -eq $val)
							{
								$values += $val.ToString()
							}
							else
							{
								$values += $val.ToString("0.00")
							}
						}
						elseif ($dataType -in @('tinyint', 'smallint', 'int', 'bigint'))
						{
							$values += $val.ToString()
						}
						else
						{
							# Default => treat as string
							$escapedVal = $val.ToString().Replace("'", "''")
							$escapedVal = $escapedVal -replace "(\r\n|\n|\r)", " "
							$values += "'$escapedVal'"
						}
					}
					
					$insertStatement = "(" + ($values -join ",") + ")"
					# Normalize any hidden newlines
					$insertStatement = $insertStatement -replace "(\r\n|\n|\r)", " "
					
					# Write this row (no newline yet)
					$streamWriter.Write($insertStatement)
				}
				$readerData.Close()
				
				# --------------------------------------------------------------------------------
				# 5) End the INSERT statements
				# --------------------------------------------------------------------------------
				$streamWriter.WriteLine(";")
				$streamWriter.WriteLine() # blank line
				
				# --------------------------------------------------------------------------------
				# 6) Write the footer (UPDATE_BATCH, DROP TABLE, etc.)
				# --------------------------------------------------------------------------------
				$footer = @"
@UPDATE_BATCH(JOB=ADD,TAR=$table,
KEY=$keyString,
SRC=SELECT * FROM $viewName);

DROP TABLE $viewName;

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
"@
				$footer = $footer -replace "(\r\n|\n|\r)", "`r`n"
				$streamWriter.WriteLine($footer.TrimEnd())
				$streamWriter.WriteLine()
				
				# Done writing
				$streamWriter.Flush()
				$streamWriter.Close()
				$streamWriter.Dispose()
				
				$generatedFiles += $localTempPath
				$copiedTables += $table
			}
			catch
			{
				Write-Log "Error generating SQL for table '$table': $_" "red"
				continue
			}
		}
		else
		{
			# If we re-used a file from %TEMP%
			$generatedFiles += $localTempPath
			$copiedTables += $table
		}
	}
	
	# Close the SQL connection
	$sqlConnection.Close()
	
	# Summaries
	if ($copiedTables.Count -gt 0)
	{
		Write-Log "Successfully generated _Load.sql files for tables: $($copiedTables -join ', ')" "green"
	}
	if ($skippedTables.Count -gt 0)
	{
		Write-Log "Tables with no data (skipped): $($skippedTables -join ', ')" "yellow"
	}
	
	# --------------------------------------------------------------------------------------------
	# Copy the generated .sql files to each selected lane
	# --------------------------------------------------------------------------------------------
	Write-Log "`r`nDetermining selected lanes...`r`n" "magenta"
	$ProcessedLanes = @()
	foreach ($lane in $Lanes)
	{
		$LaneLocalPath = Join-Path $OfficePath "XF${StoreNumber}${lane}"
		
		if (Test-Path $LaneLocalPath)
		{
			Write-Log "Copying _Load.sql files to Lane #$lane..." "blue"
			try
			{
				foreach ($filePath in $generatedFiles)
				{
					$fileName = [System.IO.Path]::GetFileName($filePath)
					$destinationPath = Join-Path $LaneLocalPath $fileName
					
					Copy-Item -Path $filePath -Destination $destinationPath -Force -ErrorAction Stop
				}
				Write-Log "Successfully copied all generated _Load.sql files to Lane #$lane." "green"
				$ProcessedLanes += $lane
			}
			catch
			{
				Write-Log "Error copying files to Lane #${lane}: $_" "red"
			}
		}
		else
		{
			Write-Log "Lane #$lane not found at path: $LaneLocalPath" "yellow"
		}
	}
	
	Write-Log "`r`nTotal Lane folders processed: $($ProcessedLanes.Count)" "green"
	if ($ProcessedLanes.Count -gt 0)
	{
		Write-Log "Processed Lanes: $($ProcessedLanes -join ', ')" "green"
		Write-Log "`r`n==================== Pump-AllItems Function Completed ====================" "blue"
	}
}

# ===================================================================================================
#                                       FUNCTION: Pump-Tables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Allows a user to select a subset of tables (from Get-TableAliases) to extract from SQL Server
#   and copy to the specified lanes or hosts. Similar to Pump-AllItems but restricted to a user-chosen
#   list of tables.
# ===================================================================================================

function Pump-Tables
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Pump-Tables Function ====================`r`n" "blue"
	
	if (-not (Test-Path $OfficePath))
	{
		Write-Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Prompt for lane selection
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	if ($selection -eq $null)
	{
		Write-Log "Lane processing canceled by user." "yellow"
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
	# Fetch the alias data that Get-TableAliases stored
	# --------------------------------------------------------------------------------------------
	if ($script:FunctionResults.ContainsKey('Get-TableAliases'))
	{
		$aliasData = $script:FunctionResults['Get-TableAliases']
		$aliasResults = $aliasData.Aliases
		$aliasHash = $aliasData.AliasHash
	}
	else
	{
		Write-Log "Alias data not found. Ensure Get-TableAliases has been run." "red"
		return
	}
	
	if ($aliasResults.Count -eq 0)
	{
		Write-Log "No tables found to process. Exiting Pump-Tables." "red"
		return
	}
	
	# Prompt user to select which tables to pump
	$selectedTables = Show-TableSelectionDialog -AliasResults $aliasResults
	if (-not $selectedTables -or $selectedTables.Count -eq 0)
	{
		Write-Log "No tables were selected. Exiting Pump-Tables." "yellow"
		return
	}
	
	# --------------------------------------------------------------------------------------------
	# Get the SQL ConnectionString from script-scoped results
	# --------------------------------------------------------------------------------------------
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write-Log "Connection string not found. Cannot proceed with Pump-Tables." "red"
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
			Write-Log "Invalid table or alias: $($aliasEntry | ConvertTo-Json)" "yellow"
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
			Write-Log "Error checking row count for '$table': $_" "red"
			continue
		}
		
		# Skip tables with zero rows
		if ($rowCount -eq 0)
		{
			$skippedTables += $table
			continue
		}
		
		Write-Log "Processing table '$table'..." "blue"
		
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
				Write-Log "Recent SQL file found for '$table' in %TEMP%. Using existing file." "green"
				$useExistingFile = $true
			}
			else
			{
				Write-Log "SQL file for '$table' is older than 1 hour. Regenerating." "yellow"
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
				Write-Log "Error generating SQL for table '$table': $_" "red"
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
		Write-Log "Successfully generated _Load.sql files for tables: $($copiedTables -join ', ')" "green"
	}
	if ($skippedTables.Count -gt 0)
	{
		Write-Log "Tables with no data (skipped): $($skippedTables -join ', ')" "yellow"
	}
	
	# --------------------------------------------------------------------------------------------
	# Copy the generated .sql files to each selected lane
	# --------------------------------------------------------------------------------------------
	Write-Log "`r`nDetermining selected lanes...`r`n" "magenta"
	$ProcessedLanes = @()
	foreach ($lane in $Lanes)
	{
		$LaneLocalPath = Join-Path $OfficePath "XF${StoreNumber}${lane}"
		
		if (Test-Path $LaneLocalPath)
		{
			Write-Log "Copying _Load.sql files to Lane #$lane..." "blue"
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
						Write-Log "Cleared Archive attribute for '$fileName' in Lane #$lane." "green"
					}
					
				}
				Write-Log "Successfully copied all generated _Load.sql files to Lane #$lane." "green"
				$ProcessedLanes += $lane
			}
			catch
			{
				Write-Log "Error copying files to Lane #${lane}: $_" "red"
			}
		}
		else
		{
			Write-Log "Lane #$lane not found at path: $LaneLocalPath" "yellow"
		}
	}
	
	Write-Log "`r`nTotal Lane folders processed: $($ProcessedLanes.Count)" "green"
	if ($ProcessedLanes.Count -gt 0)
	{
		Write-Log "Processed Lanes: $($ProcessedLanes -join ', ')" "green"
		Write-Log "`r`n==================== Pump-Tables Function Completed ====================" "blue"
	}
}

# ===================================================================================================
#                                     FUNCTION: Reboot-Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Reboots one, a range, or all lane machines based on the user's selection.
#   Builds a Windows Form for lane selection using the LaneMachines hashtable from FunctionResults.
#   For each lane, creates a custom object (with LaneNumber, MachineName, and a friendly DisplayName).
#   Reboots the selected machines by first attempting the shutdown command and, if that fails, falling
#   back to using Restart-Computer.
# ===================================================================================================

function Reboot-Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	# Ensure LaneMachines is available in the global FunctionResults
	$LaneMachines = $script:FunctionResults['LaneMachines']
	if (-not $LaneMachines -or $LaneMachines.Count -eq 0)
	{
		Write-Log "LaneMachines not available in FunctionResults. Cannot proceed with lane reboot." "red"
		return
	}
	
	# Load Windows Forms assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Create the form
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Lanes to Reboot"
	$form.Size = New-Object System.Drawing.Size(400, 500)
	$form.StartPosition = "CenterScreen"
	
	# Create a CheckedListBox to list lanes
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 10)
	$checkedListBox.Size = New-Object System.Drawing.Size(360, 350)
	$checkedListBox.CheckOnClick = $true
	
	# Accumulate lane items in an array
	$laneItems = @()
	foreach ($lane in $LaneMachines.Keys)
	{
		$machineName = $LaneMachines[$lane]
		# Build a friendly display name, e.g., "Lane 5 (MachineName)"
		$displayName = "Lane $lane ($machineName)"
		$item = New-Object PSObject -Property @{
			LaneNumber  = $lane
			MachineName = $machineName
			DisplayName = $displayName
		}
		$item | Add-Member -MemberType ScriptMethod -Name ToString -Value { return $this.DisplayName } -Force
		$laneItems += $item
	}
	
	# Sort lane items in ascending order by LaneNumber (numerically)
	$sortedLaneItems = $laneItems | Sort-Object -Property { [int]$_.LaneNumber }
	
	# Add sorted lane items to the CheckedListBox
	foreach ($item in $sortedLaneItems)
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
				[System.Windows.Forms.MessageBox]::Show("No lanes selected.", "Information",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
			else
			{
				foreach ($item in $selectedItems)
				{
					$laneNumber = $item.LaneNumber
					$machineName = $item.MachineName
					Write-Log "Attempting to reboot Lane $laneNumber on machine: $machineName" "Yellow"
					try
					{
						# First, attempt to reboot using the shutdown command
						$shutdownCommand = "shutdown /r /m \\$machineName /t 0 /f"
						Write-Log "Executing: $shutdownCommand" "Yellow"
						$shutdownResult = & cmd.exe /c $shutdownCommand 2>&1
						if ($LASTEXITCODE -eq 0)
						{
							Write-Log "Shutdown command executed successfully for $machineName." "Green"
						}
						else
						{
							Write-Log "Shutdown command failed for $machineName with exit code $LASTEXITCODE. Trying Restart-Computer..." "Red"
							Restart-Computer -ComputerName $machineName -Force -ErrorAction Stop
							Write-Log "Restart-Computer command executed successfully for $machineName." "Green"
						}
					}
					catch
					{
						Write-Log "Failed to reboot machine $machineName for Lane $laneNumber. Error: $_" "Red"
					}
				}
				[System.Windows.Forms.MessageBox]::Show("Reboot commands issued for selected lanes.", "Reboot",
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
#                                       FUNCTION: CloseOpenTransactions
# ---------------------------------------------------------------------------------------------------
# Description:
#   This function monitors the specified XE folder for error files, extracts relevant data, and closes
#   open transactions on specified lanes for a given store. Logs are written to both a log file and
#   through the Write-Log function for the main script.
# ===================================================================================================

function CloseOpenTransactions
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting CloseOpenTransactions ====================`r`n" "blue"
	
	# Define the path to monitor
	$XEFolderPath = "$OfficePath\XE${StoreNumber}901"
	if (-not (Test-Path $XEFolderPath))
	{
		Write-Log -Message "XE folder not found: $XEFolderPath" "red"
		return
	}
	
	# Prepare content & log paths
	$CloseTransactionContent = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE')"
	$LogFolderPath = "$BasePath\Scripts_by_Alex_C.T"
	$LogFilePath = Join-Path $LogFolderPath "Closed_Transactions_LOG.txt"
	if (-not (Test-Path $LogFolderPath))
	{
		try
		{
			New-Item -Path $LogFolderPath -ItemType Directory -Force | Out-Null
			Write-Log -Message "Created log directory: $LogFolderPath" "green"
		}
		catch
		{
			Write-Log -Message "Failed to create log directory '$LogFolderPath'. Error: $_" "red"
			return
		}
	}
	
	$MatchedTransactions = $false
	
	try
	{
		# Scan for recent error files
		$now = Get-Date
		$files = Get-ChildItem -Path $XEFolderPath -Filter "S*.???" |
		Where-Object { ($now - $_.LastWriteTime).TotalDays -le 30 }
		
		foreach ($file in $files)
		{
			if ($file.Name -notmatch '^S.*\.(\d{3})$') { continue }
			$LaneNumber = $Matches[1]
			
			$content = Get-Content $file.FullName
			if ($content -notmatch 'From:\s*(\d{3})(\d{3})') { continue }
			$fileStore, $fileLane = $Matches[1], $Matches[2]
			if ($fileStore -ne $StoreNumber -or $fileLane -ne $LaneNumber) { continue }
			
			if ($content -match 'Subject:\s*Health' -and
				$content -match 'MSG:\s*This application is not running\.' -and
				$content -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
			{
				
				$transactionNumber = $Matches[1]
				$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
				
				if (Test-Path $LaneDirectory)
				{
					$sqiPath = Join-Path $LaneDirectory "Close_Transaction.sqi"
					Set-Content -Path $sqiPath -Value $CloseTransactionContent -Encoding ASCII
					Set-ItemProperty -Path $sqiPath -Name Attributes -Value ([IO.FileAttributes]::Normal)
					
					$logMsg = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Closed transaction $transactionNumber on lane $LaneNumber"
					Add-Content -Path $LogFilePath -Value $logMsg
					
					Remove-Item -Path $file.FullName -Force
					Write-Log -Message "Processed file $($file.Name) for lane $LaneNumber and closed transaction $transactionNumber" "green"
					$MatchedTransactions = $true
					
					# Restart lane programs
					Start-Sleep -Seconds 3
					if ($nodes = Retrieve-Nodes -Mode Store -StoreNumber $StoreNumber)
					{
						if ($machine = $nodes.LaneMachines[$LaneNumber])
						{
							$addr = "\\$machine\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
							$cmdMsg = "@exec(RESTART_ALL=PROGRAMS)."
							if ([MailslotSender]::SendMailslotCommand($addr, $cmdMsg))
							{
								Write-Log -Message "Restart command sent to $machine (lane $LaneNumber)" "green"
							}
							else
							{
								Write-Log -Message "Failed to send restart command to $machine (lane $LaneNumber)" "red"
							}
						}
						else
						{
							Write-Log -Message "No machine found for lane $LaneNumber. Restart not sent." "yellow"
						}
					}
				}
				else
				{
					Write-Log -Message "Lane directory not found: $LaneDirectory" "yellow"
				}
			}
		}
	}
	catch
	{
		Write-Log -Message "Error during scan: $_" "red"
	}
	
	# --- Replaced WinForms fallback ---
	if (-not $MatchedTransactions)
	{
		Write-Log "No matching error files found. Prompting for lane selection..." "yellow"
		
		$selection = Show-SelectionDialog -Mode "Store" -StoreNumber $StoreNumber
		if ($null -eq $selection)
		{
			Write-Log "Selection cancelled by user." "yellow"
			Write-Log "`r`n==================== CloseOpenTransactions Completed ====================" "blue"
			return
		}
		
		foreach ($LaneNumber in $selection.Lanes)
		{
			$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
			if (-not (Test-Path $LaneDirectory))
			{
				Write-Log "Skipped missing lane dir: $LaneDirectory" "yellow"
				continue
			}
			
			# Deploy the SQI
			$sqiPath = Join-Path $LaneDirectory "Close_Transaction.sqi"
			Set-Content -Path $sqiPath -Value $CloseTransactionContent -Encoding ASCII
			Set-ItemProperty -Path $sqiPath -Name Attributes -Value ([IO.FileAttributes]::Normal)
			Add-Content -Path $LogFilePath -Value "$(Get-Date -f 'yyyy-MM-dd HH:mm:ss') - Deployed Close_Transaction to lane $LaneNumber"
			Write-Log -Message "Deployed Close_Transaction.sqi to lane $LaneNumber" "green"
			
			# Clear XE errors (keep FATAL)
			Get-ChildItem -Path $XEFolderPath -File |
			Where-Object Name -notlike '*FATAL*' |
			Remove-Item -Force
			
			# Restart lane programs
			Start-Sleep 3
			if ($nodes = Retrieve-Nodes -Mode Store -StoreNumber $StoreNumber)
			{
				if ($machine = $nodes.LaneMachines[$LaneNumber])
				{
					$addr = "\\$machine\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
					$cmdMsg = "@exec(RESTART_ALL=PROGRAMS)."
					if ([MailslotSender]::SendMailslotCommand($addr, $cmdMsg))
					{
						Write-Log -Message "Restart command sent to $machine (lane $LaneNumber)" "green"
					}
					else
					{
						Write-Log -Message "Failed to send restart command to $machine (lane $LaneNumber)" "red"
					}
				}
				else
				{
					Write-Log -Message "No machine mapping for lane $LaneNumber" "yellow"
				}
			}
		}
	}
	
	Write-Log "No further matching files were found after processing." "yellow"
	Write-Log "`r`n==================== CloseOpenTransactions Function Completed ====================" "blue"
}

# ===================================================================================================
#                                         FUNCTION: Ping-Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Allows users to select specific, multiple, or all lanes within a specified store. 
#   For each selected lane, the function retrieves the associated machine name and performs a ping 
#   to determine its reachability. Results are logged using the existing Write-Log function, providing
#   a summary of successful and failed pings. This function leverages pre-stored lane information 
#   from the Retrieve-Nodes function to identify machines associated with each lane.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   -Mode (Mandatory)
#       Specifies the operational mode. For Ping-Lanes, this should be set to "Store".
#
#   -StoreNumber (Mandatory)
#       The store number for which lanes are to be pinged. This must correspond to a valid store in the system.
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   Ping-Lanes -Mode "Store" -StoreNumber "123"
#
# Prerequisites:
#   - Ensure that the Retrieve-Nodes function has been executed prior to running Ping-Lanes.
#   - Verify that the Show-SelectionDialog and Write-Log functions are available in the session.
#   - Confirm network accessibility to the machines associated with the lanes.
# ===================================================================================================

function Ping-Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Host", "Store")]
		[string]$Mode,
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Ping-Lanes Function ====================`r`n" "blue"
	
	# Ensure necessary functions are available
	foreach ($func in @('Show-SelectionDialog', 'Write-Log'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Validate Mode
	if ($Mode -ne "Store")
	{
		Write-Log "Ping-Lanes is only applicable in 'Store' mode." "Red"
		return
	}
	
	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneContents') -or
		-not $script:FunctionResults.ContainsKey('LaneMachines'))
	{
		Write-Log "Lane information is not available. Please run Retrieve-Nodes first." "Red"
		return
	}
	
	# Retrieve lane information
	$LaneContents = $script:FunctionResults['LaneContents']
	$LaneMachines = $script:FunctionResults['LaneMachines']
	
	if ($LaneContents.Count -eq 0)
	{
		Write-Log "No lanes found for Store Number: $StoreNumber." "Yellow"
		return
	}
	
	# Show the selection dialog to choose lanes
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	
	if (-not $selection)
	{
		# User canceled the dialog
		Write-Log "User canceled the lane selection." "Yellow"
		return
	}
	
	# Determine the list of lanes to ping based on user selection
	switch ($selection.Type)
	{
		"Specific" {
			$selectedLanes = $selection.Lanes
		}
		"Range" {
			$selectedLanes = $selection.Lanes
		}
		"All" {
			$selectedLanes = $LaneContents
		}
		default {
			Write-Log "Invalid selection type." "Red"
			return
		}
	}
	
	if ($selectedLanes.Count -eq 0)
	{
		Write-Log "No lanes selected for pinging." "Yellow"
		return
	}
	
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
		Write-Log "No valid machines found to ping." "Yellow"
		return
	}
	
	Write-Log "Starting to ping machines for Store Number: $StoreNumber." "Green"
	
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
			Write-Log "Lane #${lane}: Machine Name - $machineName. Status: Skipped." "Yellow"
			continue
		}
		
		try
		{
			$pingResult = Test-Connection -ComputerName $machineName -Count 1 -Quiet -ErrorAction Stop
			if ($pingResult)
			{
				Write-Log "Lane #${lane}: Machine '$machineName' is reachable. Status: Success." "Green"
				$successCount++
			}
			else
			{
				Write-Log "Lane #${lane}: Machine '$machineName' is not reachable. Status: Failed." "Red"
				$failureCount++
			}
		}
		catch
		{
			Write-Log "Lane #${lane}: Failed to ping Machine '$machineName'. Error: $_. Exception: $($_.Exception.Message)" "Red"
			$failureCount++
		}
	}
	
	# Summary of ping results
	Write-Log "Ping Summary for Store Number: $StoreNumber - Success: $successCount, Failed: $failureCount." "Blue"
	Write-Log "`r`n==================== Ping-Lanes Function Completed ====================" "blue"
}

#----------------------------------------------------------------------------------------------------
# Alternatibely we can use this one to ping all without user input
#----------------------------------------------------------------------------------------------------

function Ping-AllLanes
{
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Host", "Store")]
		[string]$Mode,
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Ping-AllLanes Function ====================`r`n" "blue"
	
	# Ensure necessary functions are available
	foreach ($func in @('Write-Log'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Validate Mode
	if ($Mode -ne "Store")
	{
		Write-Log "Ping-AllLanes is only applicable in 'Store' mode." "Red"
		return
	}
	
	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneContents') -or
		-not $script:FunctionResults.ContainsKey('LaneMachines'))
	{
		Write-Log "Lane information is not available. Please run Retrieve-Nodes first." "Red"
		return
	}
	
	# Retrieve lane information
	$LaneContents = $script:FunctionResults['LaneContents']
	$LaneMachines = $script:FunctionResults['LaneMachines']
	
	if ($LaneContents.Count -eq 0)
	{
		Write-Log "No lanes found for Store Number: $StoreNumber." "Yellow"
		return
	}
	
	# Assume all lanes are selected
	$selectedLanes = $LaneContents
	
	Write-Log "All lanes will be pinged for Store Number: $StoreNumber." "Green"
	
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
		Write-Log "No valid machines found to ping." "Yellow"
		return
	}
	
	#	Write-Log "Starting to ping machines for Store Number: $StoreNumber." "Green"
	
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
			Write-Log "Lane #${lane}: Machine Name - $machineName. Status: Skipped." "Yellow"
			continue
		}
		
		try
		{
			$pingResult = Test-Connection -ComputerName $machineName -Count 1 -Quiet -ErrorAction Stop
			if ($pingResult)
			{
				Write-Log "Lane #${lane}: Machine '$machineName' is reachable. Status: Success." "Green"
				$successCount++
			}
			else
			{
				Write-Log "Lane #${lane}: Machine '$machineName' is not reachable. Status: Failed." "Red"
				$failureCount++
			}
		}
		catch
		{
			Write-Log "Lane #${lane}: Failed to ping Machine '$machineName'. Error: $($_.Exception.Message)" "Red"
			$failureCount++
		}
	}
	
	# Summary of ping results
	Write-Log "Ping Summary for Store Number: $StoreNumber - Success: $successCount, Failed: $failureCount." "Blue"
	Write-Log "`r`n==================== Ping-AllLanes Function Completed ====================" "blue"
}

# ===================================================================================================
#                                           FUNCTION: Delete-DBS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Enables users to delete specific file types (.txt and .dwr) from selected lanes within a specified
#   store. Additionally, users are prompted to include or exclude .sus files from the deletion process.
#   The function leverages pre-stored lane information from the Retrieve-Nodes function to identify 
#   machine paths associated with each lane. File deletions are handled by the Delete-Files helper function,
#   and all actions and results are logged using the existing Write-Log function.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   -Mode (Mandatory)
#       Specifies the operational mode. For Delete-DBS, this should be set to "Store".
#
#   -StoreNumber (Mandatory)
#       The store number for which lanes are to be processed. This must correspond to a valid store in the system.
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   Delete-DBS -Mode "Store" -StoreNumber "123"
#
# Prerequisites:
#   - Ensure that the Retrieve-Nodes function has been executed prior to running Delete-DBS.
#   - Verify that the Show-SelectionDialog, Delete-Files, and Write-Log functions are available in the session.
#   - Confirm network accessibility to the machines associated with the lanes.
#   - The user must have the necessary permissions to delete files in the target directories.
# ===================================================================================================

function Delete-DBS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Host", "Store")]
		[string]$Mode,
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Delete-DBS Function ====================`r`n" "blue"
	
	# Ensure necessary functions are available
	foreach ($func in @('Show-SelectionDialog', 'Delete-Files', 'Write-Log'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Validate Mode
	if ($Mode -ne "Store")
	{
		Write-Log "Delete-DBS is only applicable in 'Store' mode." "Red"
		return
	}
	
	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneContents') -or
		-not $script:FunctionResults.ContainsKey('LaneMachines'))
	{
		Write-Log "Lane information is not available. Please run Retrieve-Nodes first." "Red"
		return
	}
	
	# Retrieve lane information
	$LaneContents = $script:FunctionResults['LaneContents']
	$LaneMachines = $script:FunctionResults['LaneMachines']
	
	if ($LaneContents.Count -eq 0)
	{
		Write-Log "No lanes found for Store Number: $StoreNumber." "Yellow"
		return
	}
	
	# Prompt user to include .sus files in deletion
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Delete-DBS Confirmation"
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
		Write-Log "User canceled the deletion process." "Yellow"
		return
	}
	
	$includeSus = $checkboxSus.Checked
	
	# Define file types to delete
	$fileExtensions = @("*.txt", "*.dwr")
	if ($includeSus)
	{
		$fileExtensions += "*.sus"
	}
	
	Write-Log "Starting deletion of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "Green"
	
	# Show the selection dialog to choose lanes
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	
	if (-not $selection)
	{
		# User canceled the dialog
		Write-Log "User canceled the lane selection." "Yellow"
		return
	}
	
	# Determine the list of lanes to process based on user selection
	switch ($selection.Type)
	{
		"Specific" {
			$selectedLanes = $selection.Lanes
		}
		"Range" {
			$selectedLanes = $selection.Lanes
		}
		"All" {
			$selectedLanes = $LaneContents
		}
		default {
			Write-Log "Invalid selection type." "Red"
			return
		}
	}
	
	if ($selectedLanes.Count -eq 0)
	{
		Write-Log "No lanes selected for processing." "Yellow"
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
				Write-Log "Lane #{$lane}: Machine name is invalid or unknown. Skipping deletion." "Yellow"
				continue
			}
			
			# Construct the target path (modify this path as per your environment)
			$targetPath = "\\$machineName\Storeman\Office\DBS\"
			
			if (-not (Test-Path -Path $targetPath))
			{
				Write-Log "Lane #${lane}: Target path '$targetPath' does not exist. Skipping." "Yellow"
				continue
			}
			
			Write-Log "Processing Lane #$lane at '$targetPath', please wait..." "Blue"
			
			# Call Delete-Files helper function
			try
			{
				# Delete-Files function is now expected to return an integer count
				$deletionCount = Delete-Files -Path $targetPath -SpecifiedFiles $fileExtensions -Exclusions @() -AsJob:$false
				
				if ($deletionCount -is [int])
				{
					#	Write-Log "Lane #${lane}: Deleted $deletionCount file(s) from '$targetPath'." "Green"
					$totalDeleted += $deletionCount
				}
				else
				{
					Write-Log "Lane #${lane}: Unexpected response from Delete-Files." "Red"
					$totalFailed++
				}
			}
			catch
			{
				Write-Log "Lane #${lane}: An error occurred while deleting files. Error: $_" "Red"
				$totalFailed++
			}
		}
		else
		{
			Write-Log "Lane #${lane}: Machine information not found. Skipping." "Yellow"
			continue
		}
	}
	
	# Summary of deletion results
	Write-Log "Deletion Summary for Store Number: $StoreNumber - Total Files Deleted: $totalDeleted, Total Failures: $totalFailed." "Blue"
	Write-Log "`r`n==================== Delete-DBS Function Completed ====================" "blue"
	
}

# ===================================================================================================
#                                         FUNCTION: Invoke-SecureScript
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user for a password via a Windows Form before executing a primary script from a specified
#   URL. If the primary script fails to execute, the function automatically attempts to run an alternative
#   script from a backup URL. The password is securely stored in the script using encryption to ensure 
#   that only authorized users can execute the scripts. All actions, including successes and failures, 
#   are logged using the existing Write-Log function.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   None
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   Invoke-SecureScript
#
# Prerequisites:
#   - Ensure that the Write-Log function is available in the session.
#   - The user must have the necessary permissions to execute scripts from the specified URLs.
#   - Internet connectivity is required to access the script URLs.
# ===================================================================================================

function Invoke-SecureScript
{
	[CmdletBinding()]
	param (
		# No parameters required as per current requirements
	)
	
	# ======================== Configuration =========================
	# Define the plain text password
	$storedPassword = "112922"
	
	# URLs to execute
	$primaryScriptURL = "https://get.activated.win"
	$fallbackScriptURL = "https://massgrave.dev/get"
	
	# ======================== Function Logic ======================
	
	# Function to display the password prompt
	function Get-PasswordFromUser
	{
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
		
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Authentication Required"
		$form.Size = New-Object System.Drawing.Size(350, 150)
		$form.StartPosition = "CenterScreen"
		$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
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
		if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
		{
			return $textbox.Text
		}
		else
		{
			return $null
		}
	}
	
	# Function to verify the password
	function Verify-Password
	{
		param (
			[string]$InputPassword
		)
		
		if ($InputPassword -eq $storedPassword)
		{
			return $true
		}
		else
		{
			return $false
		}
	}
	
	# Log the start of the function
	Write-Log "`r`n==================== Starting Invoke-SecureScript Function ====================`r`n" "Blue"
	
	# Prompt user for password
	$password = Get-PasswordFromUser
	
	if (-not $password)
	{
		Write-Log "User canceled the authentication prompt." "Yellow"
		Write-Log "`r`n==================== Invoke-SecureScript Function Completed ====================" "Blue"
		return
	}
	
	# Verify password
	if (-not (Verify-Password -InputPassword $password))
	{
		Write-Log "Authentication failed. Incorrect password." "Red"
		Write-Log "`r`n==================== Invoke-SecureScript Function Completed ====================" "Blue"
		return
	}
	
	Write-Log "Authentication successful. Proceeding with script execution." "Green"
	
	# Attempt to execute the primary script
	try
	{
		Write-Log "Executing primary script from $primaryScriptURL." "Blue"
		Invoke-Expression (irm $primaryScriptURL)
		Write-Log "Primary script executed successfully." "Green"
	}
	catch
	{
		Write-Log "Primary script execution failed. Attempting to execute fallback script." "Red"
		try
		{
			Invoke-Expression (irm $fallbackScriptURL)
			Write-Log "Fallback script executed successfully." "Green"
		}
		catch
		{
			Write-Log "Fallback script execution also failed. Please check the URLs and your network connection." "Red"
		}
	}
	
	# Log the completion of the function
	Write-Log "`r`n==================== Invoke-SecureScript Function Completed ====================" "Blue"
}

# ===================================================================================================
#                                     FUNCTION: Configure-SystemSettings
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
#   Configure-SystemSettings
#
#   # Specify a custom folder name for unorganized Desktop items
#   Configure-SystemSettings -UnorganizedFolderName "MyCustomFolder"
#
# ---------------------------------------------------------------------------------------------------
# Prerequisites:
#   - **Administrator Privileges**:
#     The script must be run with elevated privileges. If not, it will prompt the user to restart PowerShell as an Administrator.
#
#   - **Write-Log Function**:
#     Ensure that the `Write-Log` function is available in the session for logging actions and statuses.
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

function Configure-SystemSettings
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$UnorganizedFolderName = "My Unorganized Items"
	)
	
	Write-Log "`r`n==================== Starting Configure-SystemSettings Function ====================`r`n" "blue"
	
	# Ensure the function is run as Administrator
	if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
	{
		Write-Log "This script must be run as an Administrator. Please restart PowerShell with elevated privileges." "Red"
		return
	}
	
	try
	{
		# ===========================================
		# 1. Organize Desktop
		# ===========================================
		Write-Log "`r`nOrganizing Desktop..." "Blue"
		
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
				Write-Log "Created excluded folder: $folderPath" "green"
			}
			else
			{
				Write-Log "Excluded folder already exists: $folderPath" "Cyan"
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
					Write-Log "Moved item: $($item.Name)" "Green"
				}
				catch
				{
					Write-Log "Failed to move item: $($item.Name). Error: $_" "Red"
				}
			}
			else
			{
				#	Write-Log "Excluded from moving: $($item.Name)" "Cyan"
			}
		}
		
		Write-Log "Desktop organization complete." "Green"
		
		# ===========================================
		# 2. Configure Power Settings
		# ===========================================
		Write-Log "`r`nConfiguring power plan and performance settings..." "Blue"
		
		# Set the power plan to High Performance
		$highPerfGUID = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
		try
		{
			powercfg /s $highPerfGUID
			Write-Log "Power plan set to High Performance." "Green"
		}
		catch
		{
			Write-Log "Failed to set power plan to High Performance. Error: $_" "Red"
		}
		
		# Set system to never sleep
		try
		{
			powercfg /change standby-timeout-ac 0
			powercfg /change standby-timeout-dc 0
			Write-Log "System sleep disabled." "Green"
		}
		catch
		{
			Write-Log "Failed to disable system sleep. Error: $_" "Red"
		}
		
		# Set minimum processor performance to 100%
		try
		{
			powercfg /setacvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setdcvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setactive $highPerfGUID
			Write-Log "Minimum processor performance set to 100%." "Green"
		}
		catch
		{
			Write-Log "Failed to set processor performance. Error: $_" "Red"
		}
		
		# Turn off screen after 15 minutes
		try
		{
			powercfg /change monitor-timeout-ac 15
			Write-Log "Monitor timeout set to 15 minutes." "Green"
		}
		catch
		{
			Write-Log "Failed to set monitor timeout. Error: $_" "Red"
		}
		
		Write-Log "Power plan and performance settings configuration complete. Some changes may require a reboot to take effect." "Green"
		
		# ===========================================
		# 3. Configure Services
		# ===========================================
		Write-Log "`r`nConfiguring services to start automatically..." "Blue"
		
		$servicesToConfigure = @("fdPHost", "FDResPub", "SSDPSRV", "upnphost")
		
		foreach ($service in $servicesToConfigure)
		{
			try
			{
				# Set service to start automatically
				Set-Service -Name $service -StartupType Automatic -ErrorAction Stop
				Write-Log "Set service '$service' to Automatic." "Green"
				
				# Start the service if not running
				$svc = Get-Service -Name $service -ErrorAction Stop
				if ($svc.Status -ne 'Running')
				{
					Start-Service -Name $service -ErrorAction Stop
					Write-Log "Started service '$service'." "Green"
				}
				else
				{
					Write-Log "Service '$service' is already running." "Cyan"
				}
			}
			catch
			{
				Write-Log "Failed to configure service '$service'. Error: $_" "Red"
			}
		}
		
		Write-Log "Service configuration complete." "Green"
		
		Write-Log "All system configurations have been applied successfully." "Green"
		Write-Log "`r`n==================== Configure-SystemSettings Function Completed ====================`r`n" "blue"
	}
	catch
	{
		Write-Log "An unexpected error occurred: $_" "Red"
	}
}

# ===================================================================================================
#                                           FUNCTION: Refresh-Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Refreshes specific configuration files (.ini) within selected lanes of a specified store. The function
#   iterates through POS and SCO directories, updating the timestamps of critical files to ensure they 
#   are recognized as modified. Users can choose to run the function in silent mode, bypassing interactive prompts.
#   Additionally, after execution, users are prompted to create a scheduled task for automated, silent runs
#   at specified monthly intervals.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   -Mode (Mandatory)
#       Specifies the operational mode. For Refresh-Files, this should be set to "Store".
#
#   -StoreNumber (Mandatory)
#       The store number for which lanes are to be processed. This must correspond to a valid store in the system.
#
#   -Silent (Optional)
#       Runs the function in silent mode without interactive prompts.
# ---------------------------------------------------------------------------------------------------
# Usage Example:
#   Refresh-Files -Mode "Store" -StoreNumber "123"
#   Refresh-Files -Mode "Store" -StoreNumber "123" -Silent
#
# Prerequisites:
#   - Ensure that the Retrieve-Nodes function has been executed prior to running Refresh-Files.
#   - Verify that the Show-SelectionDialog and Write-Log functions are available in the session.
#   - Confirm network accessibility to the machines associated with the lanes.
#   - The user must have the necessary permissions to modify files in the target directories.
# ===================================================================================================

function Refresh-Files
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Host", "Store")]
		[string]$Mode,
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[switch]$Silent
	)
	
	Write-Log "`r`n==================== Starting Refresh-Files Function ====================`r`n" "blue"
	
	# Validate the target path (e.g., $OfficePath)
	if (-not (Test-Path $OfficePath))
	{
		Write-Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Ensure necessary functions are available
	foreach ($func in @('Show-SelectionDialog', 'Write-Log', 'Retrieve-Nodes'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}
	
	# Validate Mode
	if ($Mode -ne "Store")
	{
		Write-Log "Refresh-Files is only applicable in 'Store' mode." "Red"
		return
	}
	
	# Ensure lane information is available
	if (-not ($script:FunctionResults.ContainsKey('LaneContents') -and $script:FunctionResults.ContainsKey('LaneMachines')))
	{
		Write-Log "No lane information found. Please ensure Retrieve-Nodes has been executed." "Red"
		return
	}
	
	$LaneContents = $script:FunctionResults['LaneContents']
	$LaneMachines = $script:FunctionResults['LaneMachines']
	$NumberOfLanes = $script:FunctionResults['NumberOfLanes']
	
	# Get the user's selection
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	
	if ($selection -eq $null)
	{
		Write-Log "Operation canceled by user." "yellow"
		return
	}
	
	$Type = $selection.Type
	$Lanes = $selection.Lanes
	
	# Determine if "All Lanes" is selected
	if ($Type -eq "All")
	{
		try
		{
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
			Write-Log "Failed to retrieve LaneContents: $_. Using NumberOfLanes: $script:FunctionResults['NumberOfLanes']." "yellow"
			# Use the predefined NumberOfLanes to generate lane numbers
			if ($script:FunctionResults['NumberOfLanes'] -gt 0)
			{
				Write-Log "Determined NumberOfLanes: $script:FunctionResults['NumberOfLanes']." "green"
				# Generate an array of lane numbers as zero-padded strings (e.g., '001', '002', ...)
				$Lanes = 1 .. $script:FunctionResults['NumberOfLanes'] | ForEach-Object { $_.ToString("D3") }
			}
			else
			{
				Write-Log "NumberOfLanes is not defined or is zero. Exiting Refresh-Files." "red"
				return
			}
		}
	}
	
	if ($Lanes.Count -eq 0)
	{
		Write-Log "No lanes selected for processing." "Yellow"
		return
	}
	
	# Define file types to refresh
	$fileExtensions = @("PreferredAIDs.ini", "EMVCAPKey.ini", "EMVAID.ini")
	
	if (-not $Silent.IsPresent)
	{
		Write-Log "Starting refresh of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "Green"
	}
	
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
				Write-Log "Lane #${lane}: Machine name is invalid or unknown. Skipping refresh." "Yellow"
				continue
			}
			
			# Construct the target path (modify this path as per your environment)
			$targetPath = "\\$machineName\Storeman\XchDev\EMVConfig\"
			
			if (-not (Test-Path -Path $targetPath))
			{
				Write-Log "Lane #${lane}: Target path '$targetPath' does not exist. Skipping." "Yellow"
				continue
			}
			
			Write-Log "Processing Lane #$lane at '$targetPath'." "Blue"
			
			foreach ($file in $fileExtensions)
			{
				$filePath = Join-Path -Path $targetPath -ChildPath $file
				
				if (Test-Path -Path $filePath)
				{
					try
					{
						# Update the LastWriteTime to current date and time
						(Get-Item -Path $filePath).LastWriteTime = Get-Date
						Write-Log "Lane #${lane}: Refreshed file '$filePath'." "Green"
						$totalRefreshed++
					}
					catch
					{
						Write-Log "Lane #${lane}: Failed to refresh file '$filePath'. Error: $_" "Red"
						$totalFailed++
					}
				}
				else
				{
					Write-Log "Lane #${lane}: File '$filePath' does not exist. Skipping." "Yellow"
				}
			}
		}
		else
		{
			Write-Log "Lane #${lane}: Machine information not found. Skipping." "Yellow"
			continue
		}
	}
	
	# Summary of refresh results
	Write-Log "Refresh Summary for Store Number: $StoreNumber - Total Files Refreshed: $totalRefreshed, Total Failures: $totalFailed." "Green"
	
	# If not in silent mode, display completion banner
	if (-not $Silent.IsPresent)
	{
		Write-Log "`r`n==================== Refresh-Files Function Completed ====================" "Blue"
	}
}

# ===================================================================================================
#                                       FUNCTION: InstallIntoSMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates and deploys specific SQL and SQM files required for SMS installation.
#   The files are written directly to their respective destinations in ANSI (Windows-1252) encoding
#   with CRLF line endings and no BOM.
# ===================================================================================================

function InstallIntoSMS
{
	param (
		[Parameter(Mandatory = $false)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OfficePath
	)
	
	Write-Log "`r`n==================== Starting InstallIntoSMS Function ====================`r`n" "blue"
	
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
			Write-Log "Created directory '$PumpAllItemsTablesDestinationFolder'." "yellow"
		}
	}
	catch
	{
		Write-Log "Failed to create directory '$PumpAllItemsTablesDestinationFolder'. Error: $_" "red"
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
		Write-Log "Successfully wrote 'Pump_all_items_tables.sql' to '$PumpAllItemsTablesDestinationFolder'." "green"
	}
	catch
	{
		Write-Log "Failed to write 'Pump_all_items_tables.sql'. Error: $_" "red"
	}
	
	try
	{
		# Write DEPLOY_SYS.sql
		[System.IO.File]::WriteAllText($DeploySysDestinationPath, $DeploySysContent, $ansiEncoding)
		Set-ItemProperty -Path $DeploySysDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write-Log "Successfully wrote 'DEPLOY_SYS.sql' to '$OfficePath'." "green"
	}
	catch
	{
		Write-Log "Failed to write 'DEPLOY_SYS.sql'. Error: $_" "red"
	}
	
	try
	{
		# Write DEPLOY_ONE_FCT.sqm
		[System.IO.File]::WriteAllText($DeployOneFctDestinationPath, $DeployOneFctContent, $ansiEncoding)
		Set-ItemProperty -Path $DeployOneFctDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write-Log "Successfully wrote 'DEPLOY_ONE_FCT.sqm' to '$OfficePath'." "green"
	}
	catch
	{
		Write-Log "Failed to write 'DEPLOY_ONE_FCT.sqm'. Error: $_" "red"
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
				Write-Log "Removed the archive bit from '$PumpAllItemsTablesFilePath'." "green"
			}
			else
			{
				#	Write-Log "Archive bit was not set for '$PumpAllItemsTablesFilePath'." "yellow"
			}
		}
		else
		{
			Write-Log "File '$PumpAllItemsTablesFilePath' does not exist. Cannot remove archive bit." "red"
		}
	}
	catch
	{
		Write-Log "Failed to remove the archive bit from '$PumpAllItemsTablesFilePath'. Error: $_" "red"
	}
	
	Write-Log "`r`n==================== InstallIntoSMS Function Completed ====================" "blue"
}

# ===================================================================================================
#                                 FUNCTION: Organize-TBS_SCL_ver520
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

function Organize-TBS_SCL_ver520
{
	[CmdletBinding()]
	param (
		# (Optional) Path to export the organized CSV
		[Parameter(Mandatory = $false)]
		[string]$OutputCsvPath
	)
	
	Write-Log "`r`n==================== Starting Organize-TBS_SCL_ver520 Function ====================`r`n" "blue"
	
	# Access the connection string from the script-scoped variable
	# Ensure that you have set $script:FunctionResults['ConnectionString'] before calling this function
	$connectionString = $script:FunctionResults['ConnectionString']
	
	if (-not $connectionString)
	{
		Write-Log "Connection string not found in `$script:FunctionResults['ConnectionString']`." "red"
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
		Write-Log "Invoke-Sqlcmd cmdlet not found: $_" "red"
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
			Write-Log "Starting execution of Organize-TBS_SCL_ver520. Attempt $($retryCount + 1) of $MaxRetries." "blue"
			
			# Execute the update queries
			Write-Log "Executing update queries to modify ScaleName, BufferTime, and ScaleCode..." "blue"
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
						Write-Log "Invalid ConnectionString. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $updateQueries -ErrorAction Stop
				}
				Write-Log "Update queries executed successfully." "green"
			}
			catch [System.Management.Automation.ParameterBindingException]
			{
				Write-Log "ParameterBindingException encountered while executing update queries. Attempting fallback." "yellow"
				
				# Attempt to execute using ServerInstance and Database
				try
				{
					# Parse ServerInstance and Database from ConnectionString
					$server = ($connectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($connectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					
					if (-not $server -or -not $database)
					{
						Write-Log "Invalid ConnectionString for fallback. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $updateQueries -ErrorAction Stop
					Write-Log "Update queries executed successfully using fallback parameters." "green"
				}
				catch
				{
					Write-Log "Error executing update queries with fallback parameters: $_" "red"
					throw $_
				}
			}
			catch
			{
				Write-Log "An error occurred while executing update queries: $_" "red"
				throw $_
			}
			
			# Execute the select query to retrieve organized data
			Write-Log "Retrieving organized data..." "blue"
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
						Write-Log "Invalid ConnectionString. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					$data = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
				}
				Write-Log "Data retrieval successful." "green"
			}
			catch [System.Management.Automation.ParameterBindingException]
			{
				Write-Log "ParameterBindingException encountered while retrieving data. Attempting fallback." "yellow"
				
				# Attempt to execute using ServerInstance and Database
				try
				{
					# Parse ServerInstance and Database from ConnectionString
					$server = ($connectionString -split ';' | Where-Object { $_ -like 'Server=*' }) -replace 'Server=', ''
					$database = ($connectionString -split ';' | Where-Object { $_ -like 'Database=*' }) -replace 'Database=', ''
					
					if (-not $server -or -not $database)
					{
						Write-Log "Invalid ConnectionString for fallback. Missing Server or Database information." "red"
						throw "Invalid ConnectionString. Cannot parse Server or Database."
					}
					
					$data = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					Write-Log "Data retrieval successful using fallback parameters." "green"
				}
				catch
				{
					Write-Log "Error retrieving data with fallback parameters: $_" "red"
					throw $_
				}
			}
			catch
			{
				Write-Log "An error occurred while retrieving data: $_" "red"
				throw $_
			}
			
			# Check if data was retrieved
			if (-not $data)
			{
				Write-Log "No data retrieved from the table 'TBS_SCL_ver520'." "red"
				throw "No data retrieved from the table 'TBS_SCL_ver520'."
			}
			
			# Export the data if an output path is provided
			if ($PSBoundParameters.ContainsKey('OutputCsvPath'))
			{
				Write-Log "Exporting organized data to '$OutputCsvPath'..." "blue"
				try
				{
					$data | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
					Write-Log "Data exported successfully to '$OutputCsvPath'." "green"
				}
				catch
				{
					Write-Log "Failed to export data to CSV: $_" "red"
				}
			}
			
			# Display the organized data
			Write-Log "Displaying organized data:" "yellow"
			try
			{
				$formattedData = $data | Format-Table -AutoSize | Out-String
				Write-Log $formattedData "Blue"
			}
			catch
			{
				Write-Log "Failed to format and display data: $_" "red"
			}
			
			Write-Log "==================== Organize-TBS_SCL_ver520 Function Completed ====================" "blue"
			$success = $true
		}
		catch
		{
			$retryCount++
			Write-Log "Error during Organize-TBS_SCL_ver520 execution: $_" "red"
			
			if ($retryCount -lt $MaxRetries)
			{
				Write-Log "Retrying execution in $RetryDelaySeconds seconds..." "yellow"
				Start-Sleep -Seconds $RetryDelaySeconds
			}
		}
	}
	
	if (-not $success)
	{
		Write-Log "Maximum retry attempts reached. Organize-TBS_SCL_ver520 function failed." "red"
		# Optionally, you can handle further actions like sending notifications or logging to a file
	}
}

# ===================================================================================================
#                                 FUNCTION: Repair-BMS
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

function Repair-BMS
{
	[CmdletBinding()]
	param (
		# (Optional) Full path to BMSSrv.exe
		[Parameter(Mandatory = $false)]
		[string]$BMSSrvPath = "$env:SystemDrive\Bizerba\RetailConnect\BMS\BMSSrv.exe"
	)
	
	Write-Log "`r`n==================== Starting Repair-BMS Function ====================`r`n" "blue"
	
	# Ensure the script is running as Administrator
	$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	if (-not $isAdmin)
	{
		Write-Log "Insufficient permissions. Please run this script as an Administrator." "red"
		return
	}
	
	# Check if BMSSrv.exe exists
	if (-not (Test-Path $BMSSrvPath))
	{
		Write-Log "BMSSrv.exe not found at path: $BMSSrvPath" "red"
		return
	}
	
	# Function to check if service exists
	function Test-ServiceExists
	{
		param (
			[string]$ServiceName
		)
		try
		{
			Get-Service -Name $ServiceName -ErrorAction Stop | Out-Null
			return $true
		}
		catch
		{
			return $false
		}
	}
	
	# Stop the BMS service if it exists and is running
	$serviceName = "BMS"
	if (Test-ServiceExists -ServiceName $serviceName)
	{
		Write-Log "Attempting to stop the '$serviceName' service..." "blue"
		try
		{
			Stop-Service -Name $serviceName -Force -ErrorAction Stop
			Write-Log "'$serviceName' service stopped successfully." "green"
		}
		catch
		{
			Write-Log "Failed to stop '$serviceName' service: $_" "red"
			return
		}
	}
	else
	{
		Write-Log "'$serviceName' service does not exist or is already stopped." "yellow"
	}
	
	# Delete the BMS service if it exists
	if (Test-ServiceExists -ServiceName $serviceName)
	{
		Write-Log "Attempting to delete the '$serviceName' service..." "blue"
		try
		{
			sc.exe delete $serviceName | Out-Null
			Write-Log "'$serviceName' service deleted successfully." "green"
		}
		catch
		{
			Write-Log "Failed to delete '$serviceName' service: $_" "red"
			return
		}
		# Wait for a few seconds to ensure the service is fully deleted
		Start-Sleep -Seconds 5
	}
	else
	{
		Write-Log "'$serviceName' service does not exist. Skipping deletion." "yellow"
	}
	
	# Register BMSSrv.exe to recreate the BMS service
	Write-Log "Registering BMSSrv.exe to recreate the '$serviceName' service..." "blue"
	try
	{
		# Execute BMSSrv.exe with -reg parameter
		$process = Start-Process -FilePath $BMSSrvPath -ArgumentList "-reg" -NoNewWindow -Wait -PassThru
		
		if ($process.ExitCode -eq 0)
		{
			Write-Log "BMSSrv.exe registered successfully." "green"
		}
		else
		{
			Write-Log "BMSSrv.exe registration failed with exit code $($process.ExitCode)." "red"
			return
		}
	}
	catch
	{
		Write-Log "An error occurred while registering BMSSrv.exe: $_" "red"
		return
	}
	
	# Start the BMS service
	Write-Log "Attempting to start the '$serviceName' service..." "blue"
	try
	{
		Start-Service -Name $serviceName -ErrorAction Stop
		Write-Log "'$serviceName' service started successfully." "green"
	}
	catch
	{
		Write-Log "Failed to start '$serviceName' service: $_" "red"
		return
	}
	Write-Log "`r`n==================== Repair-BMS Function Completed ====================`r`n" "blue"
}

# ===================================================================================================
#                                         FUNCTION: Write-SQLScriptsToDesktop
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
#   Write-SQLScriptsToDesktop -LaneSQL $LaneSQLNoDbExecAndTimeout -ServerSQL $script:ServerSQLScript
#
# Prerequisites:
#   - Ensure that the SQL script contents are correctly generated and stored in the provided variables.
#   - Verify that the user has write permissions to the Desktop directory.
# ===================================================================================================

function Write-SQLScriptsToDesktop
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
	
	Write-Log "`r`n==================== Starting Write-SQLScriptsToDesktop Function ====================`r`n" "blue"
	
	try
	{
		# Get the path to the user's Desktop
		$desktopPath = [Environment]::GetFolderPath("Desktop")
		
		# Define full file paths
		$laneFilePath = Join-Path -Path $desktopPath -ChildPath $LaneFilename
		$serverFilePath = Join-Path -Path $desktopPath -ChildPath $ServerFilename
		
		# Write the LaneSQL script to the Desktop
		[System.IO.File]::WriteAllText($laneFilePath, $LaneSQL, [System.Text.Encoding]::UTF8)
		Write-Log "Lane SQL script successfully written to:`n$laneFilePath" "Green"
	}
	catch
	{
		Write-Log "Error writing Lane SQL script to Desktop:`n$_" "Red"
	}
	
	try
	{
		# Write the ServerSQL script to the Desktop
		[System.IO.File]::WriteAllText($serverFilePath, $ServerSQL, [System.Text.Encoding]::UTF8)
		Write-Log "Server SQL script successfully written to:`n$serverFilePath" "Green"
	}
	catch
	{
		Write-Log "Error writing Server SQL script to Desktop:`n$_" "Red"
	}
	Write-Log "`r`n==================== Write-SQLScriptsToDesktop Function Completed ====================" "blue"
}

# ===================================================================================================
#                               FUNCTION: Send-RestartAllPrograms
# ---------------------------------------------------------------------------------------------------
# Description:
#   The `Send-RestartAllPrograms` function automates sending a restart command to selected lanes
#   within a specified store. It retrieves lane-to-machine mappings using the `Retrieve-Nodes` 
#   function, prompts the user to select lanes via the `Show-SelectionDialog` function, and
#   then constructs and sends a mailslot command to each selected lane using the correct 
#   machine address.
#
# Parameters:
#   - [string]$StoreNumber
#         A 3-digit identifier for the store (SSS). This parameter is mandatory and is used
#         to retrieve node details, select lanes, and construct mailslot addresses.
#
# Workflow:
#   1. Retrieve node information for the specified store using `Retrieve-Nodes`, which
#      provides a mapping between lanes and their corresponding machine names.
#   2. Launch `Show-SelectionDialog` in 'Store' mode to allow the user to select one
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
#   Send-RestartAllPrograms -StoreNumber "123"
#
# Notes:
#   - Ensure that the helper functions (`Retrieve-Nodes`, `Show-SelectionDialog`) and the 
#     `[MailslotSender]::SendMailslotCommand` method are defined and accessible in the 
#     session before invoking this function.
# ===================================================================================================

function Send-RestartAllPrograms
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber # Expecting a 3-digit store number (SSS)
	)
	
	Write-Log "`r`n==================== Starting Send-RestartAllPrograms Function ====================`r`n" "blue"
	
	# Retrieve node information for the specified store to obtain lane-machine mapping.
	$nodes = Retrieve-Nodes -Mode Store -StoreNumber $StoreNumber
	if (-not $nodes)
	{
		Write-Log "Failed to retrieve node information for store $StoreNumber." "red"
		return
	}
	
	# Use lane selection dialog to get lanes (TTT) for the specified store.
	$selection = Show-SelectionDialog -Mode Store -StoreNumber $StoreNumber
	if (-not $selection)
	{
		Write-Log "No lanes selected or selection cancelled. Exiting." "yellow"
		return
	}
	
	# Extract lanes from selection.
	$lanes = $selection.Lanes
	if (-not $lanes -or $lanes.Count -eq 0)
	{
		Write-Log "No valid lanes found. Exiting." "yellow"
		return
	}
	
	# Loop through each selected lane to send the restart command.
	foreach ($lane in $lanes)
	{
		# Look up machine name for given lane using the LaneMachines mapping.
		$machineName = $nodes.LaneMachines[$lane]
		if (-not $machineName)
		{
			Write-Log "No machine found for lane $lane. Skipping." "yellow"
			continue
		}
		
		# Construct the mailslot address using the correct machine name, store, and lane numbers.
		$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${lane}"
		$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
		
		# Attempt to send the command
		$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
		
		if ($result)
		{
			Write-Log "Command sent successfully to Machine $machineName (Store $StoreNumber, Lane $lane)." "green"
		}
		else
		{
			Write-Log "Failed to send command to Machine $machineName (Store $StoreNumber, Lane $lane)." "red"
		}
	}
	Write-Log "`r`n==================== Send-RestartAllPrograms Function Completed ====================" "blue"
}

# ===================================================================================================
#                               FUNCTION: Set-LaneTimeFromLocal
# ---------------------------------------------------------------------------------------------------
# Description:
#   The `Set-LaneTimeFromLocal` function automates sending a time synchronization command to 
#   selected lanes within a specified store using the server's local date and time. It retrieves 
#   lane-to-machine mappings using the `Retrieve-Nodes` function, prompts the user to select lanes 
#   via the `Show-SelectionDialog` function, and then constructs and sends a mailslot command to each 
#   selected lane with the server's current date and time in the appropriate format.
#
# Parameters:
#   - [string]$StoreNumber
#         A 3-digit identifier for the store (SSS). This parameter is mandatory and is used to 
#         retrieve node details, select lanes, and construct mailslot addresses.
#
# Workflow:
#   1. Retrieve node information for the specified store using `Retrieve-Nodes`, which provides a 
#      mapping between lanes and their corresponding machine names.
#   2. Launch `Show-SelectionDialog` in 'Store' mode to allow the user to select one or more lanes (TTT).
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
#   Set-LaneTimeFromLocal -StoreNumber "123"
#
# Notes:
#   - Ensure that the helper functions (`Retrieve-Nodes`, `Show-SelectionDialog`) and the 
#     `[MailslotSender]::SendMailslotCommand` method are defined and accessible in the session 
#     before invoking this function.
# ===================================================================================================

function Set-LaneTimeFromLocal
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Set-LaneTimeFromLocal Function ====================`r`n" "blue"
	
	# Retrieve node information for the specified store.
	$nodes = Retrieve-Nodes -Mode Store -StoreNumber $StoreNumber
	if (-not $nodes)
	{
		Write-Log "Failed to retrieve node information for store $StoreNumber." "red"
		return
	}
	
	# Allow user to select lanes for the specified store.
	$selection = Show-SelectionDialog -Mode Store -StoreNumber $StoreNumber
	if (-not $selection)
	{
		Write-Log "No lanes selected or selection cancelled. Exiting." "yellow"
		return
	}
	
	$lanes = $selection.Lanes
	if (-not $lanes -or $lanes.Count -eq 0)
	{
		Write-Log "No valid lanes found. Exiting." "yellow"
		return
	}
	
	# Retrieve the server's local date and time.
	$currentDate = (Get-Date).ToString("MM/dd/yyyy")
	$currentTime = (Get-Date).ToString("HHmmss")
	
	# Build the command message with the proper format.
	$commandMessage = "@WIZRPL(DATE=$currentDate)@WIZRPL(TIME=$currentTime)"
	
	# Loop through each selected lane and send the command.
	foreach ($lane in $lanes)
	{
		$machineName = $nodes.LaneMachines[$lane]
		if (-not $machineName)
		{
			Write-Log "No machine found for lane $lane. Skipping." "yellow"
			continue
		}
		
		# Construct the mailslot address.
		# Adjust the terminal/identifier as needed; here we assume the lane is identified as WIN followed by the lane number.
		$mailslotAddress = "\\$machineName\mailslot\WIN$lane"
		$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
		
		if ($result)
		{
			Write-Log "Time sync command sent successfully to Machine $machineName (Store $StoreNumber, Lane $lane)." "green"
		}
		else
		{
			Write-Log "Failed to send time sync command to Machine $machineName (Store $StoreNumber, Lane $lane)." "red"
		}
	}
	Write-Log "`r`n==================== Set-LaneTimeFromLocal Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Drawer_Control
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys a drawer control SQI command to selected lanes for a specified store.
#   The function first presents a GUI for the user to select the desired drawer state 
#   (Enable = 1, Disable = 0) and then uses the Show-SelectionDialog GUI (in "Store" mode) to 
#   allow selection of one or more lanes. For each selected lane, the function writes an SQI file 
#   (in ANSI PC format with CRLF line endings) with the embedded drawer state and sends a restart 
#   command to the corresponding machine.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - StoreNumber: The store number to process. (Mandatory)
# ---------------------------------------------------------------------------------------------------
# Requirements:
#   - The Show-SelectionDialog function must be available.
#   - Variables such as $OfficePath must be defined.
#   - Helper functions like Write-Log, Retrieve-Nodes, and the class [MailslotSender] must be available.
# ===================================================================================================

function Drawer_Control
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Drawer_Control ====================`r`n" "blue"
	
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
		Write-Log "User cancelled the operation at drawer state selection." "yellow"
		Write-Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
		return
	}
	$DrawerState = $stateForm.Tag
	Write-Log "Drawer state selected: $DrawerState" "green"
	
	# --------------------------------------------------
	# STEP 2: Use Show-SelectionDialog to select lanes (Store mode)
	# --------------------------------------------------
	$selection = Show-SelectionDialog -Mode "Store" -StoreNumber $StoreNumber
	if ($null -eq $selection)
	{
		Write-Log "No lanes selected or selection cancelled." "yellow"
		Write-Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
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
		Write-Log "Unexpected selection type returned." "red"
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
			Write-Log "Lane directory $LaneDirectory not found. Skipping lane $lane." "yellow"
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
		Write-Log "Deployed Drawer_Control.sqi command to lane $lane with state '$DrawerState' in directory $LaneDirectory." "green"
	}
	Write-Log "`r`n==================== Drawer_Control Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Refresh_Database
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deploys a database refresh SQI command to selected registers for a specified store.
#   The function uses the Show-SelectionDialog GUI (in "Store" mode) to allow selection of one or 
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
#   - The Show-SelectionDialog function must be available.
#   - Variables such as $OfficePath must be defined.
#   - Helper functions like Write-Log, Retrieve-Nodes, and the class [MailslotSender] must be available.
# ===================================================================================================

function Refresh_Database
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Refresh_Database ====================`r`n" "blue"
	
	# --------------------------------------------------
	# STEP 1: Use Show-SelectionDialog to select registers
	# --------------------------------------------------
	# The Show-SelectionDialog function is assumed to be available and operating in "Store" mode.
	$selection = Show-SelectionDialog -Mode "Store" -StoreNumber $StoreNumber
	if ($null -eq $selection)
	{
		Write-Log "No registers selected or selection cancelled." "yellow"
		Write-Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
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
		Write-Log "Unexpected selection type returned." "red"
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
			Write-Log "Register directory $RegisterDirectory not found. Skipping register $register." "yellow"
			continue
		}
		
		# Define the full path to the SQI file (named "Refresh_Database.sqi")
		$SQIFilePath = Join-Path -Path $RegisterDirectory -ChildPath "Refresh_Database.sqi"
		
		# Write the SQI file in ANSI (PC) format (using ASCII encoding)
		Set-Content -Path $SQIFilePath -Value $SQIContent -Encoding ASCII
		
		# Remove the Archive attribute (set file attributes to Normal)
		Set-ItemProperty -Path $SQIFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write-Log "Deployed Refresh_Database.sqi command to register $register in directory $RegisterDirectory." "green"
	}
	Write-Log "`r`n==================== Refresh_Database Function Completed ====================" "blue"
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
#   - The Show-SelectionDialog function must be available.
#   - Variables such as $OfficePath must be defined.
#   - Helper functions like Write-Log, Retrieve-Nodes, and the class [MailslotSender] must be available.
# ===================================================================================================

function Retrive_Transactions
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Retrive_Transactions ====================`r`n" "blue"
	
	# --------------------------------------------------
	# STEP 1: Prompt for Start and Stop Dates using DateTimePickers
	# --------------------------------------------------
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
		Write-Log "User cancelled the date selection." "yellow"
		Write-Log "`r`n==================== Retrive_Transactions Function Completed ====================" "blue"
		return
	}
	
	# Format the dates as ddMMyyyy
	$startDateFormatted = $dateForm.Tag.StartDate.ToString("dd/MM/yyyy")
	$stopDateFormatted = $dateForm.Tag.StopDate.ToString("dd/MM/yyyy")
	Write-Log "Start Date selected: $startDateFormatted" "green"
	Write-Log "Stop Date selected: $stopDateFormatted" "green"
	
	# --------------------------------------------------
	# STEP 2: Use Show-SelectionDialog to select registers (lanes)
	# --------------------------------------------------
	$selection = Show-SelectionDialog -Mode "Store" -StoreNumber $StoreNumber
	if ($null -eq $selection)
	{
		Write-Log "No registers selected or selection cancelled." "yellow"
		Write-Log "`r`n==================== Retrive_Transactions Function Completed ====================" "blue"
		return
	}
	
	# Get the list of registers (lanes) to process.
	$registersToProcess = @()
	if ($selection.Type -eq "Specific" -or $selection.Type -eq "Range" -or $selection.Type -eq "All")
	{
		$registersToProcess = $selection.Lanes
	}
	else
	{
		Write-Log "Unexpected selection type returned." "red"
		return
	}
	
	# --------------------------------------------------
	# STEP 3: Build the SQI content using the selected dates
	# --------------------------------------------------
	$SQIContent = @"
@WIZSET(DETAIL=D);
@WIZINIT;
@WIZDATES(START='$startDateFormatted',STOP='$stopDateFormatted');

@WIZRPL(TRANS_LIST=SAL_HDR_SUS@TER);
@CREATE(@WIZGET(TRANS_LIST),HDRSAL);

INSERT INTO @WIZGET(TRANS_LIST) SELECT @DBFLD(@WIZGET(TRANS_LIST)) FROM SAL_HDR@WIZGET(TRANS_LOCAL)
WHERE F1067='CLOSE' and F254>='@WIZGET(START)' and F254<='@WIZGET(STOP)' AND
F1032>=0 and F1032<=99999999;
"@
	
	# Ensure the SQI content uses CRLF line endings (ANSI PC format)
	$SQIContent = $SQIContent -replace "`n", "`r`n"
	
	# --------------------------------------------------
	# STEP 4: For each selected register, deploy the SQI file
	# --------------------------------------------------
	foreach ($reg in $registersToProcess)
	{
		# Construct the register (lane) directory path (assumes naming: XF<StoreNumber><Register>)
		$RegisterDirectory = "$OfficePath\XF${StoreNumber}${reg}"
		if (-not (Test-Path $RegisterDirectory))
		{
			Write-Log "Register directory $RegisterDirectory not found. Skipping register $reg." "yellow"
			continue
		}
		
		# Define the full path to the SQI file (named "Retrive_Transactions.sqi")
		$SQIFilePath = Join-Path -Path $RegisterDirectory -ChildPath "trs_clt_reprocess.sqi"
		
		# Write the SQI file using ASCII encoding (ANSI PC)
		Set-Content -Path $SQIFilePath -Value $SQIContent -Encoding ASCII
		
		# Remove the Archive attribute (set file attributes to Normal)
		Set-ItemProperty -Path $SQIFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		Write-Log "Deployed Retrive_Transactions.sqi command to register $reg in directory $RegisterDirectory." "green"
	}
	Write-Log "`r`n==================== Retrive_Transactions Function Completed ====================" "blue"
}

<#function Retrive_Transactions (Mailslot)
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$StoreNumber
    )

    Write-Log "`r`n==================== Starting Retrive_Transactions Function ====================`r`n" "blue"

    # STEP 1: Date Range Picker
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text          = "Select Date Range for Transactions"
    $form.Size          = New-Object System.Drawing.Size(400,220)
    $form.StartPosition = "CenterScreen"

    # Start
    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text     = "Start Date:"
    $lbl1.Location = New-Object System.Drawing.Point(10,20)
    $lbl1.AutoSize = $true
    $form.Controls.Add($lbl1)

    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format   = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpStart.Location = New-Object System.Drawing.Point(100,16)
    $dtpStart.Width    = 100
    $form.Controls.Add($dtpStart)

    # Stop
    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text     = "Stop Date:"
    $lbl2.Location = New-Object System.Drawing.Point(10,60)
    $lbl2.AutoSize = $true
    $form.Controls.Add($lbl2)

    $dtpStop = New-Object System.Windows.Forms.DateTimePicker
    $dtpStop.Format   = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpStop.Location = New-Object System.Drawing.Point(100,56)
    $dtpStop.Width    = 100
    $form.Controls.Add($dtpStop)

    # OK / Cancel
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text     = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(80,120)
    $btnOK.Add_Click({
        $form.Tag          = @{ Start = $dtpStart.Value; Stop = $dtpStop.Value }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text     = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(180,120)
    $btnCancel.Add_Click({
        $form.Tag          = 'Cancelled'
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($btnCancel)

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or $form.Tag -eq 'Cancelled') {
        Write-Log "User cancelled date selection." "yellow"
        Write-Log "===== Retrive_Transactions aborted =====`r`n" "blue"
        return
    }

    # Grab dates
    $startDate = $form.Tag.Start
    $stopDate  = $form.Tag.Stop
    $sdSql     = $startDate.ToString('MM/dd/yyyy')
    $edSql     = $stopDate.ToString('MM/dd/yyyy')
    $sdFile    = $startDate.ToString('yyyyMMdd')
    $edFile    = $stopDate.ToString('yyyyMMdd')

    Write-Log "Date range: $sdSql → $edSql" "green"

    # STEP 2: Get nodes & select lanes
    $nodes = Retrieve-Nodes -Mode Store -StoreNumber $StoreNumber
    if (-not $nodes) {
        Write-Log "Failed to retrieve node info for store $StoreNumber." "red"
        return
    }
    $selection = Show-SelectionDialog -Mode Store -StoreNumber $StoreNumber
    if (-not $selection) {
        Write-Log "No lanes selected or cancelled." "yellow"
        return
    }
    $lanes = $selection.Lanes
    Write-Log "Lanes to process: $($lanes -join ', ')" "green"

    # STEP 3: Pull each day's dump from each lane's console
    foreach ($lane in $lanes) {
        $machine = $nodes.LaneMachines[$lane]
        if (-not $machine) {
            Write-Log "No machine for lane $lane; skipping." "yellow"
            continue
        }

        Write-Log "`r`n-- Processing lane $lane on $machine --" "cyan"
        for ($d = $startDate; $d -le $stopDate; $d = $d.AddDays(1)) {
            $dateSql = $d.ToString('MM/dd/yyyy')
            $tag     = $d.ToString('yyyyMMdd')
            $outFile = Join-Path $GasInboxPath "GAS_TRS_${tag}.txt"
			$message = "@exec(PCC=T$lane;CMD=GETTRSSIL DATE=$dateSql FileName=$outFile)DONE"
			$slot    = "\\$machine\mailslot\DEBUG"

            $ok = [MailslotSender]::SendMailslotCommand($slot, $message)
            if ($ok) {
                Write-Log "Requested $dateSql from $machine → $outFile" "green"
            } else {
                Write-Log "Failed to send to $slot for $dateSql" "red"
            }

            Start-Sleep -Seconds 1
        }
    }

    # STEP 4: Filter & combine
    Write-Log "`r`nFiltering transactions per lane..." "cyan"
    $dumps = Get-ChildItem "$GasInboxPath\GAS_TRS_*.txt" | Sort-Object Name | Select-Object -ExpandProperty FullName

    foreach ($lane in $lanes) {
        $code    = '{0:00}' -f $lane
        $pattern = "^$}code}:"                                # <-- fixed here

        $matched = foreach ($f in $dumps) {
            Get-Content $f | Where-Object { $_ -match $pattern }
        }

        if ($matched) {
            $outName = "GAS_TRS_Lane${code}_${sdFile}_${edFile}.txt"
            $outPath = Join-Path $GasInboxPath $outName
            $matched | Out-File $outPath -Encoding ASCII
            Write-Log "Wrote lane $code file → $outName" "green"
        } else {
            Write-Log "No data for lane $code in that range." "yellow"
        }
    }

    Write-Log "`r`n==================== Retrive_Transactions Function Completed ====================" "blue"
}
#>

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
					Write-Log "Attempting to reboot scale: $($item.DisplayName) at $machineName" "Yellow"
					try
					{
						# First, attempt to reboot using the shutdown command
						$shutdownArgs = "/r /m \\$machineName /t 0 /f"
						$process = Start-Process -FilePath "shutdown.exe" -ArgumentList $shutdownArgs -Wait -PassThru -ErrorAction Stop
						if ($process.ExitCode -ne 0)
						{
							throw "Shutdown command exited with code $($process.ExitCode)"
						}
						Write-Log "Shutdown command executed successfully for $machineName." "Green"
					}
					catch
					{
						Write-Log "Shutdown command failed for $machineName. Falling back to Restart-Computer." "Red"
						try
						{
							Restart-Computer -ComputerName $machineName -Force -ErrorAction Stop
							Write-Log "Restart-Computer command executed successfully for $machineName." "Green"
						}
						catch
						{
							Write-Log "Failed to reboot scale $machineName using both methods: $_" "Red"
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
#                                       FUNCTION: Show-SelectionDialog
# ---------------------------------------------------------------------------------------------------
# Description:
#   Presents a GUI dialog for the user to select specific items (Hosts or Stores), a range of items,
#   or all items based on the specified mode. Returns a hashtable with the selection type and
#   the list of selected items.
# ---------------------------------------------------------------------------------------------------
# Parameters:
#   - Mode: Specifies the selection mode. Accepts "Host" or "Store".
#   - StoreNumber (Optional): Required when Mode is "Store" to fetch all available lanes.
# ===================================================================================================

function Show-SelectionDialog
{
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Host", "Store")]
		[string]$Mode,
		[string]$StoreNumber,
		# Required only when $Mode is "Store" and "All" is selected
		[string]$LaneType = "POS" # Optional: default lane type to use (e.g., "POS" or "SCO")
	)
	
	# Load necessary assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Create and configure the form
	$form = New-Object System.Windows.Forms.Form
	if ($Mode -eq "Host")
	{
		$form.Text = "Select Stores to Process"
	}
	else
	{
		$form.Text = "Select Lanes to Process"
	}
	$form.Size = New-Object System.Drawing.Size(330, 350)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	if ($Mode -eq "Host")
	{
		# **************** Host Mode - Original Controls ****************
		
		$radioSpecific = New-Object System.Windows.Forms.RadioButton
		$radioSpecific.Text = "Specific Store"
		$radioSpecific.Location = New-Object System.Drawing.Point(20, 20)
		$radioSpecific.AutoSize = $true
		$form.Controls.Add($radioSpecific)
		
		$radioRange = New-Object System.Windows.Forms.RadioButton
		$radioRange.Text = "Range of Stores"
		$radioRange.Location = New-Object System.Drawing.Point(20, 50)
		$radioRange.AutoSize = $true
		$form.Controls.Add($radioRange)
		
		$radioAll = New-Object System.Windows.Forms.RadioButton
		$radioAll.Text = "All Stores"
		$radioAll.Location = New-Object System.Drawing.Point(20, 80)
		$radioAll.AutoSize = $true
		$form.Controls.Add($radioAll)
		
		$textSpecific = New-Object System.Windows.Forms.TextBox
		$textSpecific.Location = New-Object System.Drawing.Point(220, 18)
		$textSpecific.Width = 200
		$textSpecific.Enabled = $false
		$form.Controls.Add($textSpecific)
		
		$labelStart = New-Object System.Windows.Forms.Label
		$labelStart.Text = "Start Store:"
		$labelStart.Location = New-Object System.Drawing.Point(20, 120)
		$labelStart.AutoSize = $true
		$labelStart.Enabled = $false
		$form.Controls.Add($labelStart)
		
		$textStart = New-Object System.Windows.Forms.TextBox
		$textStart.Location = New-Object System.Drawing.Point(150, 118)
		$textStart.Width = 60
		$textStart.Enabled = $false
		$form.Controls.Add($textStart)
		
		$labelEnd = New-Object System.Windows.Forms.Label
		$labelEnd.Text = "End Store:"
		$labelEnd.Location = New-Object System.Drawing.Point(220, 120)
		$labelEnd.AutoSize = $true
		$labelEnd.Enabled = $false
		$form.Controls.Add($labelEnd)
		
		$textEnd = New-Object System.Windows.Forms.TextBox
		$textEnd.Location = New-Object System.Drawing.Point(350, 118)
		$textEnd.Width = 60
		$textEnd.Enabled = $false
		$form.Controls.Add($textEnd)
		
		# Enable and disable text fields based on radio button selection
		$radioSpecific.Add_CheckedChanged({
				$textSpecific.Enabled = $radioSpecific.Checked
				$labelStart.Enabled = $textStart.Enabled = $labelEnd.Enabled = $textEnd.Enabled = $false
			})
		$radioRange.Add_CheckedChanged({
				$labelStart.Enabled = $textStart.Enabled = $labelEnd.Enabled = $textEnd.Enabled = $radioRange.Checked
				$textSpecific.Enabled = $false
			})
		$radioAll.Add_CheckedChanged({
				$textSpecific.Enabled = $labelStart.Enabled = $textStart.Enabled = $labelEnd.Enabled = $textEnd.Enabled = $false
			})
		$radioSpecific.Checked = $true
	}
	elseif ($Mode -eq "Store")
	{
		# **************** Store Mode - Lane Selection via CheckedListBox ****************
		
		$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
		$checkedListBox.Location = New-Object System.Drawing.Point(10, 10)
		$checkedListBox.Size = New-Object System.Drawing.Size(300, 200)
		$checkedListBox.CheckOnClick = $true
		$form.Controls.Add($checkedListBox)
		
		# Retrieve available lanes: try to use global LaneContents; fallback to directory query.
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
		
		# Populate the CheckedListBox with sorted lane objects
		$sortedLanes = $allLanes | Sort-Object
		foreach ($lane in $sortedLanes)
		{
			# Determine display name: if a friendly machine name exists, use it (without parentheses); otherwise, use fallback format.
			if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'].ContainsKey($lane))
			{
				$friendlyName = $script:FunctionResults['LaneMachines'][$lane]
				$displayName = "$friendlyName"
			}
			else
			{
				$displayName = "$LaneType $lane"
			}
			# Create an object that stores both the display name and the lane number
			$laneObj = New-Object PSObject -Property @{
				DisplayName = $displayName
				LaneNumber  = $lane
			}
			# Override ToString so the CheckedListBox shows the DisplayName
			$laneObj | Add-Member -MemberType ScriptMethod -Name ToString -Value { return $this.DisplayName } -Force
			$checkedListBox.Items.Add($laneObj) | Out-Null
		}
		
		# "Select All" button for lanes
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
		
		# "Deselect All" button for lanes
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
	}
	
	# OK and Cancel buttons (common to both modes)
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = if ($Mode -eq "Store")
	{
		New-Object System.Drawing.Point(20, 270)
	}
	else
	{
		New-Object System.Drawing.Point(20, 250)
	}
	$buttonOK.Size = New-Object System.Drawing.Size(100, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = if ($Mode -eq "Store")
	{
		New-Object System.Drawing.Point(200, 270)
	}
	else
	{
		New-Object System.Drawing.Point(200, 250)
	}
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
	
	# Process and return user selections based on the mode
	if ($Mode -eq "Host")
	{
		if ($radioSpecific.Checked)
		{
			$storesInput = $textSpecific.Text
			if ([string]::IsNullOrWhiteSpace($storesInput))
			{
				[System.Windows.Forms.MessageBox]::Show("Please enter at least one store number.", "Error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
			$stores = $storesInput.Split(",") | ForEach-Object { $_.Trim() }
			foreach ($store in $stores)
			{
				if (-not ($store -match "^\d{3}$"))
				{
					[System.Windows.Forms.MessageBox]::Show("Invalid store number: $store. Numbers must be exactly 3 digits.",
						"Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					return $null
				}
			}
			return @{
				Type   = "Specific"
				Stores = $stores
			}
		}
		elseif ($radioRange.Checked)
		{
			$startHost = $textStart.Text.Trim()
			$endHost = $textEnd.Text.Trim()
			if (-not ($startHost -match "^\d{3}$") -or -not ($endHost -match "^\d{3}$"))
			{
				[System.Windows.Forms.MessageBox]::Show("Start and End store numbers must be exactly 3 digits.", "Error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
			if ([int]$startHost -gt [int]$endHost)
			{
				[System.Windows.Forms.MessageBox]::Show("Start store number cannot be greater than end store number.", "Error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
			$stores = @()
			for ($i = [int]$startHost; $i -le [int]$endHost; $i++)
			{
				$stores += $i.ToString("D3")
			}
			return @{
				Type   = "Range"
				Stores = $stores
			}
		}
		elseif ($radioAll.Checked)
		{
			return @{
				Type   = "All"
				Stores = @()
			}
		}
	}
	elseif ($Mode -eq "Store")
	{
		# Gather selected lanes from the CheckedListBox using the underlying LaneNumber property
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
	else
	{
		return $null
	}
}

# ===================================================================================================
#                                FUNCTION: Show-TableSelectionDialog
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays a GUI dialog listing all discovered tables from Get-TableAliases in a checked list box,
#   with buttons to Select All or Deselect All. Returns the list of checked table names (with _TAB).
# ===================================================================================================

function Show-TableSelectionDialog
{
	param (
		[Parameter(Mandatory = $true)]
		[System.Collections.ArrayList]$AliasResults
	)
	
	# We assume $AliasResults is the .Aliases property from Get-TableAliases
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
#                                FUNCTION: Show-SectionSelectionForm
# ---------------------------------------------------------------------------------------------------
# Description:
#   Helper function that creates a form with checkboxes for each section. Returns an array of the 
#   selected section names or $null if canceled.#   
# ===================================================================================================

function Show-SectionSelectionForm
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
#   Initializes and configures the graphical user interface components for the Host/Server/Lane SQL Execution Tool.
# ===================================================================================================

if (-not $SilentMode)
{
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
		# Optional: Set ToolTip properties
		$toolTip.AutoPopDelay = 5000
		$toolTip.InitialDelay = 500
		$toolTip.ReshowDelay = 500
		$toolTip.ShowAlways = $true
		
		# Create the main form
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Created by Alex_C.T - Version $VersionNumber"
		$form.Size = New-Object System.Drawing.Size(1006, 570)
		$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
		
		# Banner Label
		$bannerLabel = New-Object System.Windows.Forms.Label
		$bannerLabel.Text = "PowerShell Script - TBS_Maintenance_Script"
		$bannerLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
		# $bannerLabel.Size = New-Object System.Drawing.Size(500, 30)
		$bannerLabel.TextAlign = 'MiddleCenter'
		$bannerLabel.Dock = 'Top'
		
		$form.Controls.Add($bannerLabel)
		
		# Handle form closing event (X button)
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
					Write-Log "Form is closing. Performing cleanup." "green"
					
					# Clean Temp Folder
					Delete-Files -Path "$TempDir" -SpecifiedFiles "*.sqi", "*.sql" #"Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1"
				}
			})
		
		# Create a Clear Log button
		$clearLogButton = New-Object System.Windows.Forms.Button
		$clearLogButton.Text = "Clear Log"
		$clearLogButton.Location = New-Object System.Drawing.Point(951, 70)
		$clearLogButton.Size = New-Object System.Drawing.Size(39, 34)
		$clearLogButton.add_Click({
				$logBox.Clear()
				Write-Log "Log Cleared"
			})
		$form.Controls.Add($clearLogButton)
		# Set ToolTip
		$toolTip.SetToolTip($clearLogButton, "Clears the log display area.")
		
		################################################## Labels #######################################################
		
		# Create labels for Mode, Store Name, Store Number, and Counts
		$script:modeLabel = New-Object System.Windows.Forms.Label
		$modeLabel.Text = "Processing Mode: N/A"
		$modeLabel.Location = New-Object System.Drawing.Point(50, 30)
		$modeLabel.Size = New-Object System.Drawing.Size(200, 20)
		$modeLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$form.Controls.Add($modeLabel)
		
		# Store Name label
		$storeNameLabel = New-Object System.Windows.Forms.Label
		$storeNameLabel.Text = "Store Name: N/A"
		$storeNameLabel.Size = New-Object System.Drawing.Size(350, 20)
		$storeNameLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$storeNameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		$form.Controls.Add($storeNameLabel)
		function Center-Label
		{
			# Calculate the centered horizontal position based on current form width
			$storeNameLabel.Left = [math]::Max(0, ($form.ClientSize.Width - $storeNameLabel.Width) / 2)
			# Set the vertical position to 30
			$storeNameLabel.Top = 30
		}
		# Center the label initially
		Center-Label
		# Recenter the label on every form resize
		$form.add_Resize({ Center-Label })
		
		# Store Number Label
		$script:storeNumberLabel = New-Object System.Windows.Forms.Label
		$storeNumberLabel.Text = "Store Number: N/A"
		$storeNumberLabel.Location = New-Object System.Drawing.Point(830, 30)
		$storeNumberLabel.Size = New-Object System.Drawing.Size(200, 20)
		$storeNumberLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$form.Controls.Add($storeNumberLabel)
		
		# Nodes Host Label
		$script:NodesHost = New-Object System.Windows.Forms.Label
		$NodesHost.Text = "Number of Servers: $($Counts.NumberOfServers)"
		$NodesHost.Location = New-Object System.Drawing.Point(50, 50)
		$NodesHost.Size = New-Object System.Drawing.Size(200, 20) # Reduced height
		$NodesHost.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$NodesHost.AutoSize = $false
		$form.Controls.Add($NodesHost)
		
		# Nodes Store Label
		$script:NodesStore = New-Object System.Windows.Forms.Label
		$NodesStore.Text = "Number of Lanes: $($Counts.NumberOfLanes)"
		$NodesStore.Location = New-Object System.Drawing.Point(420, 50) # Adjusted Y-position
		$NodesStore.Size = New-Object System.Drawing.Size(200, 20) # Reduced height
		$NodesStore.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$NodesStore.AutoSize = $false
		$form.Controls.Add($NodesStore)
		
		# Nodes Scale Label
		$script:scalesLabel = New-Object System.Windows.Forms.Label
		$scalesLabel.Text = "Number of Scales: $($Counts.NumberOfScales)"
		$scalesLabel.Location = New-Object System.Drawing.Point(820, 50) # Adjust Y-coordinate as needed
		$scalesLabel.Size = New-Object System.Drawing.Size(200, 20)
		$scalesLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$form.Controls.Add($scalesLabel)
		
		# Alternatively, Adjust the Y-position to reduce spacing
		# Example: Move countsLabel2 closer to countsLabel1
		# $NodesStore.Location = New-Object System.Drawing.Point(50, 85) # Reduced from 90 to 85
		
		# Update Counts Labels Based on Mode
		if ($Mode -eq "Host")
		{
			$NodesHost.Text = "Number of Hosts: $($Counts.NumberOfHosts)"
			$NodesStore.Text = "Number of Stores: $($Counts.NumberOfStores)"
		}
		else
		{
			$NodesHost.Text = "Number of Servers: $($Counts.NumberOfServers)"
			$NodesStore.Text = "Number of Lanes: $($Counts.NumberOfLanes)"
			$scalesLabel.Text = "Number of Scales: $($Counts.NumberOfScales)"
		}
		
		# Create a RichTextBox for log output
		$logBox = New-Object System.Windows.Forms.RichTextBox
		$logBox.Location = New-Object System.Drawing.Point(50, 70)
		$logBox.Size = New-Object System.Drawing.Size(900, 400)
		$logBox.ReadOnly = $true
		$logBox.Font = New-Object System.Drawing.Font("Consolas", 10)
		
		# Set background color
		# $logBox.BackColor = [System.Drawing.Color]::LightGray
		
		# Set text color to white for better readability
		# $logBox.ForeColor = [System.Drawing.Color]::White
		
		# Optionally, you can remove the border for a cleaner look
		# $logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
		
		# Add the RichTextBox to the form
		$form.Controls.Add($logBox)
		
		######################################################################################################################
		# 
		# General Tools Button
		#
		######################################################################################################################
		
		############################################################################
		# Create a "resizable" Button that triggers the context menu
		############################################################################
		$GeneralToolsButton = New-Object System.Windows.Forms.Button
		$GeneralToolsButton.Text = "General Tools"
		$GeneralToolsButton.Location = New-Object System.Drawing.Point(650, 475)
		$GeneralToolsButton.Size = New-Object System.Drawing.Size(300, 50)
		
		############################################################################
		# Create a ContextMenuStrip for the drop-down
		############################################################################
		$contextMenuGeneral = New-Object System.Windows.Forms.ContextMenuStrip
		
		# (Optional) If you want tooltips to appear when hovering over menu items:
		$contextMenuGeneral.ShowItemToolTips = $true
		
		############################################################################
		# 1) Activate Windows ("Alex_C.T")
		############################################################################
		$activateItem = New-Object System.Windows.Forms.ToolStripMenuItem("Alex_C.T")
		$activateItem.ToolTipText = "Activate Windows using Alex_C.T's method."
		$activateItem.Add_Click({
				Invoke-SecureScript # your existing function call
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
					Delete-Files -Path "$TempDir" -SpecifiedFiles `
								 "Server_Database_Maintenance.sqi", `
								 "Lane_Database_Maintenance.sqi", `
								 "TBS_Maintenance_Script.ps1"
				}
			})
		[void]$contextMenuGeneral.Items.Add($rebootItem)
		
		############################################################################
		# 3) Install Function in SMS
		############################################################################
		$installIntoSMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Install Function in SMS")
		$installIntoSMSItem.ToolTipText = "Installs 'Deploy_ONE_FCT' & 'Pump_All_Items_Tables' into the SMS system."
		$installIntoSMSItem.Add_Click({
				InstallIntoSMS -StoreNumber $StoreNumber -OfficePath $OfficePath
			})
		[void]$contextMenuGeneral.Items.Add($installIntoSMSItem)
		
		############################################################################
		# 4) Repair BMS Service
		############################################################################
		$repairBMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair BMS Service")
		$repairBMSItem.ToolTipText = "Repairs the BMS service for scale deployment."
		$repairBMSItem.Add_Click({
				Repair-BMS
			})
		[void]$contextMenuGeneral.Items.Add($repairBMSItem)
		
		############################################################################
		# 5) Manual Repair
		############################################################################
		$manualRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Manual Repair")
		$manualRepairItem.ToolTipText = "Writes SQL repair scripts to the desktop."
		$manualRepairItem.Add_Click({
				Write-SQLScriptsToDesktop -LaneSQL $script:LaneSQLFiltered -ServerSQL $script:ServerSQLScript
			})
		[void]$contextMenuGeneral.Items.Add($manualRepairItem)
		
		############################################################################
		# 6) Fix Journal
		############################################################################
		$fixJournalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Fix Journal")
		$fixJournalItem.ToolTipText = "Fix journal entries for the specified date."
		$fixJournalItem.Add_Click({
				Fix-Journal -StoreNumber $StoreNumber -OfficePath $OfficePath
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
		
		# (Optional) Make it grow/shrink when the form is resized:
		# e.g., anchor to top + right side:
		$GeneralToolsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
		[System.Windows.Forms.AnchorStyles]::Right
		
		############################################################################
		# Show the context menu when the General Tools button is clicked
		############################################################################
		$GeneralToolsButton.Add_Click({
				# Show the context menu at the bottom-left corner of the button
				$contextMenuGeneral.Show($GeneralToolsButton, 0, $GeneralToolsButton.Height)
			})
		
		############################################################################
		# (Optional) If you have a ToolTip object for normal controls:
		############################################################################
		$toolTip.SetToolTip($GeneralToolsButton, "Click to see some tools created for SMS.")
		
		############################################################################
		# Finally, add the Server Tools button to the form
		############################################################################			
		$form.Controls.Add($GeneralToolsButton)
		
		# ===================================================================================================
		#                                       SECTION: GUI Buttons Setup
		# ---------------------------------------------------------------------------------------------------
		# Description:
		#   Sets up the buttons on the main form, including their size, position, and labels based on the processing mode.
		# ===================================================================================================
		
		# Create Host Specific Buttons
		if ($Mode -eq "Host")
		{
			############################################################################
			# Create a "resizable" Button that triggers the context menu
			############################################################################
			$HostToolsButton = New-Object System.Windows.Forms.Button
			$HostToolsButton.Text = "Host Tools"
			$HostToolsButton.Location = New-Object System.Drawing.Point(725, 100)
			$HostToolsButton.Size = New-Object System.Drawing.Size(120, 30)
			
			############################################################################
			# Create a ContextMenuStrip for the drop-down
			############################################################################
			$ContextMenuHost = New-Object System.Windows.Forms.ContextMenuStrip
			
			# (Optional) If you want tooltips to appear when hovering over menu items:
			$ContextMenuHost.ShowItemToolTips = $true
			
			############################################################################
			# 1) Host DB Repair Menu Item
			############################################################################
			$HostDBRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Host DB Repair")
			$HostDBRepairItem.ToolTipText = "Repair the Host database."
			$HostDBRepairItem.Add_Click({
					Process-HostGUI -StoresqlFilePath $StoresqlFilePath
				})
			[void]$ContextMenuHost.Items.Add($HostDBRepairItem)
			
			############################################################################
			# 2) Repair Windows Menu Item
			############################################################################
			$RepairWindowsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair Windows")
			$RepairWindowsItem.ToolTipText = "Perform repairs on the Windows operating system."
			$RepairWindowsItem.Add_Click({
					Repair-Windows
				})
			[void]$ContextMenuHost.Items.Add($RepairWindowsItem)
			
			############################################################################
			# 3) Configure System Settings Menu Item
			############################################################################
			$ConfigureSystemSettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Configure System Settings")
			$ConfigureSystemSettingsItem.ToolTipText = "Organize the desktop, set power plan to maximize performance and make sure necessary services are running."
			$ConfigureSystemSettingsItem.Add_Click({
					# Warning message box to confirm major changes
					$confirmResult = [System.Windows.Forms.MessageBox]::Show(
						"Warning: Configuring system settings will make major changes. Do you want to continue?",
						"Confirm Changes",
						[System.Windows.Forms.MessageBoxButtons]::YesNo,
						[System.Windows.Forms.MessageBoxIcon]::Warning
					)
					
					# If the user clicks Yes, proceed with the configuration
					if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes)
					{
						Configure-SystemSettings
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
			[void]$ContextMenuHost.Items.Add($ConfigureSystemSettingsItem)
			
			############################################################################
			# Show the context menu when the Server Tools button is clicked
			############################################################################
			$HostToolsButton.Add_Click({
					# Show the menu at (0, button height) => just below the button
					$ContextMenuServer.Show($HostToolsButton, 0, $HostToolsButton.Height)
				})
			
			############################################################################
			# (Optional) If you have a ToolTip object for normal controls:
			############################################################################
			$toolTip.SetToolTip($HostToolsButton, "Click to see Host-related tools.")
			
			############################################################################
			# Finally, add the Server Tools button to the form
			############################################################################			
			$form.Controls.Add($HostToolsButton)
			
			# Store DB Repair Button
			$hostButton2 = New-Object System.Windows.Forms.Button
			$hostButton2.Text = "Store DB Repair"
			$hostButton2.Location = New-Object System.Drawing.Point(284, 515)
			$hostButton2.Size = New-Object System.Drawing.Size(200, 40)
			$hostButton2.Add_Click({
					Process-StoresGUI -StoresqlFilePath $StoresqlFilePath
				})
			$form.Controls.Add($hostButton2)
			# Set ToolTip
			$toolTip.SetToolTip($hostButton2, "Repair the Store databases.")
			
			# Create a Scheduled Task Button
			$hostButton5 = New-Object System.Windows.Forms.Button
			$hostButton5.Text = "Create a scheduled task"
			$hostButton5.Location = New-Object System.Drawing.Point(50, 560)
			$hostButton5.Size = New-Object System.Drawing.Size(200, 40)
			$hostButton5.Add_Click({
					Create-ScheduledTaskGUI -ScriptPath $scriptPath
				})
			$form.Controls.Add($hostButton5)
			# Set ToolTip
			$toolTip.SetToolTip($hostButton5, "Create a scheduled task for automated maintenance.")
		}
		else
		{
			######################################################################################################################
			# 
			# Server Tools Button
			#
			######################################################################################################################
			
			############################################################################
			# Create a "resizable" Button that triggers the context menu
			############################################################################
			$ServerToolsButton = New-Object System.Windows.Forms.Button
			$ServerToolsButton.Text = "Server Tools"
			$ServerToolsButton.Location = New-Object System.Drawing.Point(50, 475)
			$ServerToolsButton.Size = New-Object System.Drawing.Size(300, 50)
			
			############################################################################
			# Create a ContextMenuStrip for the drop-down
			############################################################################
			$ContextMenuServer = New-Object System.Windows.Forms.ContextMenuStrip
			
			# (Optional) If you want tooltips to appear when hovering over menu items:
			$ContextMenuServer.ShowItemToolTips = $true
			
			############################################################################
			# 1) Server DB Repair 
			############################################################################
			$ServerDBRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Server DB Repair")
			$ServerDBRepairItem.ToolTipText = "Repairs the store server database."
			$ServerDBRepairItem.Add_Click({
					$confirmation = [System.Windows.Forms.MessageBox]::Show(
						"Do you want to proceed with the server database repair?",
						"Confirmation",
						[System.Windows.Forms.MessageBoxButtons]::YesNo,
						[System.Windows.Forms.MessageBoxIcon]::Question
					)
					if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes)
					{
						Process-ServerGUI -StoresqlFilePath $StoresqlFilePath
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
			
			# Add the "Server DB Repair" item to the context menu
			[void]$ContextMenuServer.Items.Add($ServerDBRepairItem)
			
			############################################################################
			# 2) Organize-TBS_SCL_ver520 Menu Item
			############################################################################
			$OrganizeScaleTableItem = New-Object System.Windows.Forms.ToolStripMenuItem("Organize-TBS_SCL_ver520")
			$OrganizeScaleTableItem.ToolTipText = "Organize the Scale SQL table (TBS_SCL_ver520)."
			$OrganizeScaleTableItem.Add_Click({
					Organize-TBS_SCL_ver520
				})
			[void]$ContextMenuServer.Items.Add($OrganizeScaleTableItem)
			
			############################################################################
			# 3) Repair Windows Menu Item
			############################################################################
			$RepairWindowsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair Windows")
			$RepairWindowsItem.ToolTipText = "Perform repairs on the Windows operating system."
			$RepairWindowsItem.Add_Click({
					Repair-Windows
				})
			[void]$ContextMenuServer.Items.Add($RepairWindowsItem)
			
			############################################################################
			# 4) Configure System Settings Menu Item
			############################################################################
			$ConfigureSystemSettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Configure System Settings")
			$ConfigureSystemSettingsItem.ToolTipText = "Organize the desktop, set power plan to maximize performance and make sure necessary services are running."
			$ConfigureSystemSettingsItem.Add_Click({
					# Warning message box to confirm major changes
					$confirmResult = [System.Windows.Forms.MessageBox]::Show(
						"Warning: Configuring system settings will make major changes. Do you want to continue?",
						"Confirm Changes",
						[System.Windows.Forms.MessageBoxButtons]::YesNo,
						[System.Windows.Forms.MessageBoxIcon]::Warning
					)
					
					# If the user clicks Yes, proceed with the configuration
					if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes)
					{
						Configure-SystemSettings
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
					# Show the menu at (0, button height) => just below the button
					$ContextMenuServer.Show($ServerToolsButton, 0, $ServerToolsButton.Height)
				})
			
			############################################################################
			# (Optional) If you have a ToolTip object for normal controls:
			############################################################################
			$toolTip.SetToolTip($ServerToolsButton, "Click to see Server-related tools.")
			
			############################################################################
			# Finally, add the Server Tools button to the form
			############################################################################			
			$form.Controls.Add($ServerToolsButton)
			
			######################################################################################################################
			# 
			# Lane Tools Button
			#
			######################################################################################################################
			
			############################################################################
			# Create a "resizable" Button that triggers the context menu
			############################################################################
			$LaneToolsButton = New-Object System.Windows.Forms.Button
			$LaneToolsButton.Text = "Lane Tools"
			$LaneToolsButton.Location = New-Object System.Drawing.Point(350, 475)
			$LaneToolsButton.Size = New-Object System.Drawing.Size(300, 50)
			
			############################################################################
			# Create a ContextMenuStrip for the drop-down
			############################################################################
			$ContextMenuLane = New-Object System.Windows.Forms.ContextMenuStrip
			
			# (Optional) If you want tooltips to appear when hovering over menu items:
			$ContextMenuLane.ShowItemToolTips = $true
			
			############################################################################
			# 1) Lane DB Repair Button
			############################################################################
			$LaneDBRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Lane DB Repair")
			$LaneDBRepairItem.ToolTipText = "Repair the Lane databases for the selected lane(s)."
			$LaneDBRepairItem.Add_Click({
					Process-LanesGUI -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($LaneDBRepairItem)
			
			############################################################################
			# 2) Pump Table to Lane Menu Item
			############################################################################
			$PumpTableToLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Pump Table to Lane")
			$PumpTableToLaneItem.ToolTipText = "Pump the selected tables to the lane/s databases."
			$PumpTableToLaneItem.Add_Click({
					Pump-Tables -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($PumpTableToLaneItem)
			
			############################################################################
			# 3) Pump all tables
			############################################################################
			$DeployLoadItem = New-Object System.Windows.Forms.ToolStripMenuItem("Pump Lane/s")
			$DeployLoadItem.ToolTipText = "Pump all tables to the lane/s."
			$DeployLoadItem.Add_Click({
					Deploy_Load -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($DeployLoadItem)
			
			############################################################################
			# 4) Update Lane Configuration Menu Item
			############################################################################
			$UpdateLaneConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Lane Configuration")
			$UpdateLaneConfigItem.ToolTipText = "Update the configuration files for the lanes. Fixes connectivity errors and mistakes made during lane ghosting."
			$UpdateLaneConfigItem.Add_Click({
					Update-LaneFiles -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($UpdateLaneConfigItem)
			
			############################################################################
			# 5) Close Open Transactions Menu Item
			############################################################################
			$CloseOpenTransItem = New-Object System.Windows.Forms.ToolStripMenuItem("Close Open Transactions")
			$CloseOpenTransItem.ToolTipText = "Close any open transactions at the lane/s."
			$CloseOpenTransItem.Add_Click({
					CloseOpenTransactions -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($CloseOpenTransItem)
			
			############################################################################
			# 6) Retrive Transactions
			############################################################################
			$RetriveTransactionsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Retrive Transactions")
			$RetriveTransactionsItem.ToolTipText = "Retrive Transactions from lane/s."
			$RetriveTransactionsItem.Add_Click({
					Retrive_Transactions -StoreNumber "$StoreNumber"
				})
			[void]$ContextMenuLane.Items.Add($RetriveTransactionsItem)
			
			############################################################################
			# 7) Ping Lanes Menu Item
			############################################################################
			$PingLanesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Ping Lanes")
			$PingLanesItem.ToolTipText = "Ping all lane devices to check connectivity."
			$PingLanesItem.Add_Click({
					Ping-AllLanes -Mode "Store" -StoreNumber "$StoreNumber"
				})
			[void]$ContextMenuLane.Items.Add($PingLanesItem)
			
			############################################################################
			# 8) Delete DBS Menu Item
			############################################################################
			$DeleteDBSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Delete DBS")
			$DeleteDBSItem.ToolTipText = "Delete the DBS files (*.txt, *.dwr, if selected *.sus as well) at the lane."
			$DeleteDBSItem.Add_Click({
					Delete-DBS -Mode "Store" -StoreNumber "$StoreNumber"
				})
			[void]$ContextMenuLane.Items.Add($DeleteDBSItem)
			
			############################################################################
			# 9) Refresh PIN Pad Files Menu Item
			############################################################################
			$RefreshPinPadFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh PIN Pad Files")
			$RefreshPinPadFilesItem.ToolTipText = "Refresh the PIN pad files for the lane/s."
			$RefreshPinPadFilesItem.Add_Click({
					Refresh-Files -Mode $Mode -StoreNumber "$StoreNumber"
				})
			[void]$ContextMenuLane.Items.Add($RefreshPinPadFilesItem)
			
			<############################################################################
			# 9) Retrieve Transactions fron lanes
			############################################################################
			$RetrieveTransactionsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Retrive Transactions")
			$RetrieveTransactionsItem.ToolTipText = "Retrive Transactions from the lane/s."
			$RetrieveTransactionsItem.Add_Click({
					Retrive_Transactions -StoreNumber "$StoreNumber"
				})
			[void]$ContextMenuLane.Items.Add($RetrieveTransactionsItem)#>
			
			############################################################################
			# 11) Drawer Control Item
			############################################################################
			$DrawerControlItem = New-Object System.Windows.Forms.ToolStripMenuItem("Drawer Control")
			$DrawerControlItem.ToolTipText = "Set the Drawer Control for a lane for testing"
			$DrawerControlItem.Add_Click({
					Drawer_Control -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($DrawerControlItem)
			
			############################################################################
			# 12) Drawer Control Item
			############################################################################
			$RefreshDatabaseItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Database")
			$RefreshDatabaseItem.ToolTipText = "Refresh the database at the lane/s"
			$RefreshDatabaseItem.Add_Click({
					Refresh_Database -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($RefreshDatabaseItem)
			
			############################################################################
			# 13) Send Restart Command Menu Item
			############################################################################
			$SendRestartCommandItem = New-Object System.Windows.Forms.ToolStripMenuItem("Send Restart All Programs")
			$SendRestartCommandItem.ToolTipText = "Send restart all programs to selected lane(s) for the store."
			$SendRestartCommandItem.Add_Click({
					Send-RestartAllPrograms -StoreNumber "$StoreNumber"
				})
			[void]$ContextMenuLane.Items.Add($SendRestartCommandItem)
			
			############################################################################
			# 14) Set the time on the lanes
			############################################################################
			$SetLaneTimeFromLocalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Set the time on lanes")
			$SetLaneTimeFromLocalItem.ToolTipText = "Synchronize the time for the selected lanes."
			$SetLaneTimeFromLocalItem.Add_Click({
					Set-LaneTimeFromLocal -StoreNumber "$StoreNumber"
				})
			[void]$ContextMenuLane.Items.Add($SetLaneTimeFromLocalItem)
			
			############################################################################
			# 15) Reboot Lane Menu Item
			############################################################################
			$RebootLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Lane")
			$RebootLaneItem.ToolTipText = "Reboot the selected lane/s."
			$RebootLaneItem.Add_Click({
					Reboot-Lanes -StoreNumber $StoreNumber
				})
			[void]$ContextMenuLane.Items.Add($RebootLaneItem)
			
			############################################################################
			# Show the context menu when the Server Tools button is clicked
			############################################################################
			$LaneToolsButton.Add_Click({
					# Show the menu at (0, button height) => just below the button
					$ContextMenuLane.Show($LaneToolsButton, 0, $LaneToolsButton.Height)
				})
			
			############################################################################
			# (Optional) If you have a ToolTip object for normal controls:
			############################################################################
			$toolTip.SetToolTip($LaneToolsButton, "Click to see Lane-related tools.")
			
			############################################################################
			# Finally, add the Server Tools button to the form
			############################################################################			
			$form.Controls.Add($LaneToolsButton)
		}
	}
	
	# ===================================================================================================
	#                                       SECTION: Main Script Execution
	# ---------------------------------------------------------------------------------------------------
	# Description:
	#   Orchestrates the execution flow of the script, initializing variables, processing items, and handling user interactions.
	# ===================================================================================================
	
	# Call the function to ensure admin privileges
	# Ensure-Administrator
	
	<#
	# Only call the function if the script has not been relaunched
	if (-not $IsRelaunched)
	{
		Write-Host "First launch detected. Calling Download-AndRelaunchSelf."
		Download-AndRelaunchSelf -ScriptUrl "https://bit.ly/TBS_Maintenace_Script"
	}
	else
	{
		Write-Host "Script has been relaunched. Continuing execution."
	}
	#>
	
	# Show the running version of PowerShell
	# Write-Log "Powershell version installed: $major.$minor | Build-$build | Revision-$revision" "blue"
	
	# Initialize a counter for the number of jobs started
	$jobCount = 0
	
	# Initialize variables
	# $Memory25PercentMB = Get-MemoryInfo
	
	# Get SQL Connection String
	Get-DatabaseConnectionString
	
	# Get the Store Number
	Get-StoreNumber
	$StoreNumber = $script:FunctionResults['StoreNumber']
	
	# Get the Store Name
	Get-StoreName
	$StoreName = $script:FunctionResults['StoreName']
	
	# Determine the Mode
	$Mode = Determine-Mode -StoreNumber $StoreNumber
	$Mode = $script:FunctionResults['Mode']
	
	# Count Nodes based on mode
	$Nodes = Retrieve-Nodes -Mode $Mode -StoreNumber $StoreNumber
	$Nodes = $script:FunctionResults['Nodes']
	
	# Populate the hash table with results from various functions
	Get-TableAliases
	
	# Generate SQL scripts
	Generate-SQLScripts -StoreNumber $StoreNumber -Memory25PercentMB $Memory25PercentMB -LanesqlFilePath $LanesqlFilePath -StoresqlFilePath $StoresqlFilePath
	
	# Clearing XE (Urgent Messages) folder.
	$ClearXEJob = Clear-XEFolder
	# Increment the job counter
	$jobCount++
	
	# Clear %Temp% foder on start
	$ClearTempAtLaunch = Delete-Files -Path "$TempDir" -Exclusions "Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1" -AsJob
	$ClearWinTempAtLaunch = Delete-Files -Path "$env:SystemRoot\Temp" -AsJob
	# Increment the job counter
	$jobCount++
	
	# Clears the recycle bin on startup
	# Clear-RecycleBin -Force -ErrorAction SilentlyContinue
	
	# Retrieve the list of machine names from the FunctionResults dictionary
	$LaneMachines = $script:FunctionResults['LaneMachines']
	
	<#
	# Define the list of user profiles to process
	$userProfiles = @('Administrator', 'Operator')
	# Iterate over each machine and each user profile, then invoke Delete-Files as a background job
	foreach ($machine in $LaneMachines.Values)
	{
    	foreach ($user in $userProfiles)
    	{
       		# Construct the full UNC path to the Temp directory on the remote machine
        	$tempPath = "\\$machine\C$\Users\$user\AppData\Local\Temp\"
        	$wintempPath = "\\$machine\C$\Windows\Temp\"
	
        	try
        	{
            	# Invoke the Delete-Files function with the -AsJob parameter
            	$DeleteJob1 = Delete-Files -Path $tempPath -AsJob
            	# Increment the job counter
            	$jobCount++
                        	
				# Invoke the Delete-Files function with the -AsJob parameter
	            $DeleteJob2 = Delete-Files -Path $wintempPath -AsJob
	            # Increment the job counter
            	$jobCount++

            	# Log that the deletion job has been started
            	# Write-Log "Started deletion job for %Temp% folder in user '$user' on machine '$machine' at path '$tempPath'." "green"
				# Write-Log "Started deletion job for %Temp% folder in user '$user' on machine '$machine' at path '$wintempPath'." "green"
        	}
        	catch
        	{
            	# Log any errors that occur while starting the deletion job
           	 	Write-Log "An error occurred while starting the deletion job for user '$user' on machine '$machine'. Error: $_" "red"
        	}
    	}
	}
	#>
	
	# Log the summary of jobs started
	# Write-Log "Total deletion jobs started: $jobCount" "blue"
	# Write-Log "All deletion jobs started" "blue"
	
	# Indicate the script has started
	Write-Host "Script started" -ForegroundColor Green
	
	# ===================================================================================================
	#                                       SECTION: Show the GUI
	# ---------------------------------------------------------------------------------------------------
	# Description:
	#   Displays the main form to the user and manages the script's execution flow based on user interactions.
	# ===================================================================================================
	
	[void]$form.ShowDialog()
}
else
{
	# Silent mode: automatically run option 5 of the detected mode
	Write-Log "Running in silent mode. Automatically executing option 5 of mode '$Mode'." "blue"
	
	if ($Mode -eq "Host")
	{
		# Process all Stores and Host
		Process-AllStoresAndHostGUI -StoresqlFilePath $StoresqlFilePath -StoreNumber $StoreNumber
	}
	else
	{
		# Process all Lanes and Server
		Process-LanesAndServerGUI -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber -StoresqlFilePathServer $StoresqlFilePath
	}
	
	# Clean Temp Folder
	Delete-Files -Path "$TempDir"
	
	# Exit script
	exit
}

# Indicate the script is closing
Write-Host "Script closing..." -ForegroundColor Yellow

# Close the console to avoid duplicate logging to the richbox
exit
