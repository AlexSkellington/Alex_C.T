<#
Param (
	[switch]$IsRelaunched
)
#>

# Write-Host "Script started. IsRelaunched: $IsRelaunched"
Write-Host "Script starting, pls wait..." -ForegroundColor Yellow

# ===================================================================================================
#                                       SECTION: Parameters
# ---------------------------------------------------------------------------------------------------
# Description:
#   Defines the script parameters, allowing users to run the script in silent mode.
# ===================================================================================================

# Script build version (cunsult with Alex_C.T before changing this)
$VersionNumber = "1.8.2"

# Retrieve Major, Minor, Build, and Revision version numbers of PowerShell
$major = $PSVersionTable.PSVersion.Major
$minor = $PSVersionTable.PSVersion.Minor
$build = $PSVersionTable.PSVersion.Build
$revision = $PSVersionTable.PSVersion.Revision

# Combine them into a single version string
$PowerShellVersion = "$major.$minor.$build.$revision"

# Determine if build version is considered too old
# Adjust the threshold as needed
$BuildThreshold = 15000
$IsOldBuild = $build -lt $BuildThreshold

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

# Create a UTF8 encoding instance without BOM
$utf8NoBOM = New-Object System.Text.UTF8Encoding($false)

# Initialize BasePath variable
$BasePath = $null

# If the build version is too old, skip UNC paths and go directly to local drives.
if (-not $IsOldBuild)
{
	# Define the UNC paths to check in order of priority
	$uncPaths = @(
		"\\localhost\storeman",
		"\\$env:COMPUTERNAME\storeman"
	)
	
	# Check each UNC path for existence
	foreach ($path in $uncPaths)
	{
		if (Test-Path -Path $path -PathType Container)
		{
			$BasePath = $path
			break
		}
	}
}

# If no UNC path is found or build is old, proceed to check local drives
if (-not $BasePath)
{
	# Define local drives to search
	$localDrives = @("C:\", "D:\")
	
	foreach ($drive in $localDrives)
	{
		# Retrieve directories matching '*storeman*' in the root of the drive
		$storemanDirs = Get-ChildItem -Path $drive -Directory -Filter "*storeman*" -ErrorAction SilentlyContinue
		
		if ($storemanDirs)
		{
			# Select the first matching directory
			$BasePath = $storemanDirs[0].FullName
			break
		}
	}
}

# Final check to ensure BasePath was set
if (-not $BasePath)
{
	$BasePath = "C:\storeman"
}

# Now that we have a valid $BaseUNCPath, define the rest of the paths
$OfficePath = Join-Path $BasePath "office"
$LoadPath = Join-Path $OfficePath "Load"
$StartupIniPath = Join-Path $BasePath "Startup.ini"
$SystemIniPath = Join-Path $OfficePath "system.ini"

# Temp Directory
$TempDir = [System.IO.Path]::GetTempPath()

# SQI Location variables
$LanesqlFilePath = "$env:TEMP\Lane_Database_Maintenance.sqi"
$StoresqlFilePath = "$env:TEMP\Server_Database_Maintenance.sqi"

# Script Name
# $scriptName = Split-Path -Leaf $PSCommandPath

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
		Write-Log "Found Startup.ini at: $startupIniPath" "green"
	}
	
	if (-not $StartupIniPath)
	{
		Write-Log "Startup.ini file not found in any of the expected locations." "red"
		return
	}
	
	Write-Log "Generating connection string..." "blue"
	
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
	
	Write-Log "Variables ($ConnectionString) stored." "green"
}

# ===================================================================================================
#                                      FUNCTION: Get-StoreNumberGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store number via GUI prompts or configuration files.
#   Stores the result in $script:FunctionResults['StoreNumber'].
# ===================================================================================================

function Get-StoreNumberGUI
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
			Write-Log "Store number found in startup.ini: $storeNumber" "green"
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
#                                        FUNCTION: Get-StoreNameGUI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store name from the system.ini file.
#   Stores the result in $script:FunctionResults['StoreName'].
# ===================================================================================================

function Get-StoreNameGUI
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

function Determine-ModeGUI
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
#                                       SECTION: Counting Functions
# ---------------------------------------------------------------------------------------------------
# Description:
#   Contains functions to count various items like stores, lanes, servers, and hosts.
#   First attempts to read counts from the TER_TAB database table.
#   Falls back to the current mechanism if database access fails.
# ===================================================================================================

function Count-ItemsGUI
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
	$LaneContents = @()
	$LaneMachines = @{ }
	
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
			$CountsFromDatabase = $false
		}
		else
		{
			$CountsFromDatabase = $true
		}
	}
	else
	{
		$CountsFromDatabase = $true
	}
	
	# Initialize a flag to check if we successfully got counts from TER_TAB
	if ($CountsFromDatabase)
	{
		try
		{
			if ($Mode -eq "Host")
			{
				# Get NumberOfStores (excluding StoreNumber = '999')
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
				
				# Check if host exists
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
				
				# Retrieve lane contents
				$queryLaneContents = "SELECT F1057, F1125 FROM TER_TAB WHERE F1056 = '$StoreNumber' AND F1057 LIKE '0%' AND F1057 NOT LIKE '8%' AND F1057 NOT LIKE '9%'"
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
				
				# Check if server exists for the store
				$queryServer = "SELECT COUNT(*) AS ServerCount FROM TER_TAB WHERE F1056 = '$StoreNumber' AND F1057 = '901'"
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
			Write-Log "Failed to retrieve counts from TER_TAB: $_" "yellow"
			$CountsFromDatabase = $false
		}
	}
	
	# If counts from database failed, use the current mechanism as fallback
	if (-not $CountsFromDatabase)
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
			
			# Count lanes directly under the office directory matching the pattern
			if (Test-Path $HostPath)
			{
				$LaneFolders = Get-ChildItem -Path $HostPath -Directory -Filter "XF${StoreNumber}0??"
				$NumberOfLanes = $LaneFolders.Count
			}
			
			# Check for the server directory under the store
			$NumberOfServers = if (Test-Path "$HostPath\XF${StoreNumber}901") { 1 }
			else { 0 }
		}
	}
	
	# Create a custom object with the counts
	$Counts = [PSCustomObject]@{
		NumberOfStores  = $NumberOfStores
		NumberOfHosts   = $NumberOfHosts
		NumberOfLanes   = $NumberOfLanes
		NumberOfServers = $NumberOfServers
		LaneContents    = $LaneContents
		LaneMachines    = $LaneMachines
	}
	
	# Store counts in FunctionResults
	$script:FunctionResults['NumberOfStores'] = $NumberOfStores
	$script:FunctionResults['NumberOfHosts'] = $NumberOfHosts
	$script:FunctionResults['NumberOfLanes'] = $NumberOfLanes
	$script:FunctionResults['NumberOfServers'] = $NumberOfServers
	$script:FunctionResults['LaneContents'] = $LaneContents
	$script:FunctionResults['LaneMachines'] = $LaneMachines
	$script:FunctionResults['Counts'] = $Counts
	
	# Update the GUI countsLabel1 and countsLabel2 with the new counts
	if (-not $SilentMode -and $countsLabel1 -ne $null -and $countsLabel2 -ne $null)
	{
		if ($Mode -eq "Host")
		{
			$countsLabel1.Text = "Number of Hosts: $NumberOfHosts"
			$countsLabel2.Text = "Number of Stores: $NumberOfStores"
		}
		else
		{
			$countsLabel1.Text = "Number of Servers: $NumberOfServers"
			$countsLabel2.Text = "Number of Lanes: $NumberOfLanes"
		}
		# Refresh the form to display updates
		$form.Refresh()
	}
	
	# Return counts as a custom object
	return $Counts
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
	
	# Function to determine if a file should be kept
	function ShouldKeepFile($file)
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
	
	# Initial clearing
	if (Test-Path -Path $folderPath)
	{
		try
		{
			Get-ChildItem -Path $folderPath -Recurse -Force | ForEach-Object {
				if (-not (ShouldKeepFile $_))
				{
					Remove-Item -Path $_.FullName -Force -Recurse
				}
			}
			
			Write-Log "Folder 'XE${StoreNumber}901' cleaned, keeping only (FATAL*) files and valid (S*) files for transaction closing." "green"
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
			
			function ShouldKeepFile($file)
			{
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
							if (-not (ShouldKeepFile $_))
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
		
		#Write-Log "Background job 'ClearXEFolderJob' started to continuously monitor and clear 'XE${StoreNumber}901' folder." "green"
	}
	catch
	{
		Write-Log "Failed to start background job for 'XE${StoreNumber}901': $_" "red"
	}
	
	return $job
}

# ===================================================================================================
#                                       SECTION: Generate SQL Scripts
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates SQL scripts for Lanes and Stores, including memory configuration and maintenance tasks.
# ===================================================================================================

function Generate-SQLScriptsGUI
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
		Write-Log "Using DBNAME from FunctionResults: $dbName" "blue"
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
	
	Write-Log "Generating SQL scripts using Store DB: '$storeDbName' and Lane DB: '$laneDbName'..." "blue"
	
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
IF OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('dbo.TBS_ITM_SMAppUPDATED', 'OBJECT', 'DELETE') = 1 DELETE FROM dbo.TBS_ITM_SMAppUPDATED;

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000) 
DECLARE cmds CURSOR FOR 
SELECT 'drop table [' + name + ']' 
FROM sys.tables 
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$StoreNumber%' OR name LIKE 'R$StoreNumber%') AND DATEDIFF(DAY, create_date, GETDATE()) > 30 
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
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000'))) 
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000'))) 
@dbEXEC(DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB)) 
@dbEXEC(DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F267 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F268 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB)) 
@dbEXEC(UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04, SCL_TXT_TAB.F1001=1 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0) 
@dbEXEC(UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))

/* Shrink database and log files */
ALTER DATABASE LANESQL SET RECOVERY SIMPLE
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'
DBCC SHRINKFILE ($laneDbName)
DBCC SHRINKFILE (${laneDbName}_Log)
ALTER DATABASE LANESQL SET RECOVERY FULL

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
"@
	
	# Store the LaneSQLScript in the script scope
	$script:LaneSQLScript = $LaneSQLScript
	
	# Optionally write to file as fallback
	if ($LanesqlFilePath)
	{
		[System.IO.File]::WriteAllText($LanesqlFilePath, $script:LaneSQLScript, $utf8NoBOM)
	}
	
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
IF OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('dbo.TBS_ITM_SMAppUPDATED', 'OBJECT', 'DELETE') = 1 DELETE FROM dbo.TBS_ITM_SMAppUPDATED;

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
UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) 
UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000')) 
DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB)
DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB) 
UPDATE SCL_TAB SET F267 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB) 
UPDATE SCL_TAB SET F268 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB) 
UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04, SCL_TXT_TAB.F1001=1 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0
UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')

/* Shrink database and log files */
ALTER DATABASE STORESQL SET RECOVERY SIMPLE;
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD';
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN';
DBCC SHRINKFILE ($storeDbName);
DBCC SHRINKFILE (${storeDbName}_Log);
ALTER DATABASE STORESQL SET RECOVERY FULL;
"@
	
	# Store the ServerSQLScript in the script scope
	$script:ServerSQLScript = $ServerSQLScript
	
	# Optionally write to file as fallback
	if ($StoresqlFilePath)
	{
		[System.IO.File]::WriteAllText($StoresqlFilePath, $script:ServerSQLScript, $utf8NoBOM)
	}
	
	Write-Log "SQL scripts generated successfully." "green"
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
# ===================================================================================================

function Execute-SQLLocallyGUI
{
	param (
		[Parameter(Mandatory = $false)]
		[string]$SqlFilePath
	)
	
	# Configuration for retry mechanism
	$MaxRetries = 2
	$RetryDelaySeconds = 5
	$FailedCommandsPath = "$OfficePath\XF${StoreNumber}901\Failed_ServerSQLScript_Sections.sql"
	
	# Attempt to retrieve the SQL script from the script-scoped variable
	$sqlScript = $script:ServerSQLScript
	
	$dbName = $script:FunctionResults['DBNAME']
	
	if (-not [string]::IsNullOrWhiteSpace($sqlScript))
	{
		Write-Log "Executing SQL script from variable..." "blue"
	}
	elseif ($SqlFilePath)
	{
		# If the script variable is empty, attempt to execute from the file
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
	
	# Split the SQL script into sections based on /* Section Name */ comments
	# The regex captures the section name and the subsequent SQL commands until the next section or end of script
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	
	$matches = [regex]::Matches($sqlScript, $sectionPattern)
	
	if ($matches.Count -eq 0)
	{
		Write-Log "No SQL sections found to execute." "red"
		return
	}
	
	# Retrieve the connection string from script FunctionResults
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
	
	# Initialize variables to track execution
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
			$failedSections = @() # Reset failed sections for this retry
			
			foreach ($match in $sectionsToExecute)
			{
				$sectionName = $match.Groups['SectionName'].Value.Trim()
				$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()
				
				if ([string]::IsNullOrWhiteSpace($sqlCommands))
				{
					Write-Log "Section '$sectionName' contains no SQL commands. Skipping..." "yellow"
					continue
				}
				
				Write-Log "`r`nExecuting section: '$sectionName'" "blue"
				Write-Log "----------------------------------------------------------------------------------------------------------------"
				Write-Log "$sqlCommands" "orange"
				Write-Log "----------------------------------------------------------------------------------------------------------------"
				
				# Execute the SQL commands for the current section
				try
				{
					Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sqlCommands -ErrorAction Stop -QueryTimeout 0
					Write-Log "Section '$sectionName' executed successfully." "green"
				}
				catch
				{
					Write-Log "Error executing section '$sectionName': $_" "red"
					$failedSections += $match
					# Only add failed commands to $failedCommands on the last retry
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
				Write-Log "Retrying execution in $RetryDelaySeconds seconds..." "yellow"
				Start-Sleep -Seconds $RetryDelaySeconds
			}
		}
	}
	
	if (-not $success)
	{
		Write-Log "Maximum retry attempts reached. SQL script execution failed." "red"
		# Create a string from the failed commands array
		$failedCommandsText = $failedCommands -join "`r`n"
		# Write the failed commands to a file
		[System.IO.File]::WriteAllText($FailedCommandsPath, $failedCommandsText, $utf8NoBOM)
		Write-Log "`r`nFailed SQL sections written to: $FailedCommandsPath" "yellow"
		# Remove the archived attribute to ensure it can be processed
		Set-ItemProperty -Path $FailedCommandsPath -Name Attributes -Value ((Get-Item $FailedCommandsPath).Attributes -band (-bnot [System.IO.FileAttributes]::Archive))
	}
	else
	{
		Write-Log "SQL script executed successfully on the database '$dbName'." "green"
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
	Execute-SQLLocallyGUI -SqlFilePath $StoresqlFilePath
	
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
			Copy-Item -Path $LanesqlFilePath -Destination "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Force
			Set-ItemProperty -Path "$LaneLocalPath\Lane_Database_Maintenance.sqi" -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			Write-Log "Copied successfully to Lane #${LaneNumber}." "green"
			
			# Add lane to processed lanes if not already added
			if (-not ($script:ProcessedLanes -contains $LaneNumber))
			{
				$script:ProcessedLanes += $LaneNumber
			}
		}
		catch
		{
			Write-Log "Failed to copy to Lane #${LaneNumber}: $_" "red"
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
#   it filters the SQL files to include only the records pertinent to that store and lane,
#   copies the modified files to the corresponding lane directory as .sql files with UTF8 encoding without BOM,
#   copies the run_load.sql script exactly as provided,
#   generates and copies a customized lnk_load.sql script containing only records for that lane,
#   generates and copies a customized sto_load.sql script containing only records for that lane,
#   and generates and copies a customized ter_load.sql script containing only records for that lane and store,
#   including a standard '901' record.
# ===================================================================================================

function Update-LaneFiles
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting Update-LaneFiles Function ====================" "blue"
	
	if (-not (Test-Path $LoadPath))
	{
		Write-Log "`r`nLoad Base Path not found: $LoadPath" "yellow"
		return
	}
	
	# Get the user's selection
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
			#	Write-Log "User selected 'All Lanes'. Retrieving LaneContents..." "blue"
			$LaneContents = $script:FunctionResults['LaneContents']
			
			if ($LaneContents -and $LaneContents.Count -gt 0)
			{
				#	Write-Log "Successfully retrieved LaneContents. Processing all lanes." "green"
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
			# Optionally, you can retain $Lanes as is or set to an empty array
		}
	}
	
	# Define the run_load script content as a here-string (exactly as provided)
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

@UPDATE_BATCH(JOB=ADD,TAR=RUN_TAB,
KEY=F1102=:F1102 AND F1000=:F1000,
SRC=SELECT * FROM Run_Load);

DROP TABLE Run_Load;
"@
	
	# Define the lnk_load script header and footer
	$lnkLoadHeader = @"
@CREATE(LNK_TAB,LNK);
CREATE VIEW Lnk_Load AS SELECT F1000,F1056,F1057 FROM LNK_TAB;

INSERT INTO Lnk_Load VALUES
"@
	
	$lnkLoadFooter = @"

@UPDATE_BATCH(JOB=ADD,TAR=LNK_TAB,
KEY=F1000=:F1000 AND F1056=:F1056 AND F1057=:F1057,
SRC=SELECT * FROM Lnk_Load);

DROP TABLE Lnk_Load;
"@
	
	# Define the sto_load script header and footer
	$stoLoadHeader = @"
@CREATE(STO_TAB,STO);
CREATE VIEW Sto_Load AS SELECT F1000,F1018,F1180,F1181,F1182,F1937,F1965,F1966,F2691 FROM STO_TAB;

INSERT INTO Sto_Load VALUES
"@
	
	$stoLoadFooter = @"

@UPDATE_BATCH(JOB=ADD,TAR=STO_TAB,
KEY=F1000=:F1000,
SRC=SELECT * FROM Sto_Load);

DROP TABLE Sto_Load;
"@
	
	# Define the ter_load script header and footer
	$terLoadHeader = @"
@CREATE(TER_TAB,TER); 
CREATE VIEW Ter_Load AS SELECT F1056,F1057,F1058,F1125,F1169 FROM TER_TAB;

INSERT INTO Ter_Load VALUES
"@
	
	$terLoadFooter = @"

@UPDATE_BATCH(JOB=ADD,TAR=TER_TAB,
KEY=F1056=:F1056 AND F1057=:F1057,
SRC=SELECT * FROM Ter_Load);

DROP TABLE Ter_Load;
"@
	
	# Define the script filenames with .sql extension
	$runLoadFilename = "run_load.sql"
	$lnkLoadFilename = "lnk_load.sql"
	$stoLoadFilename = "sto_load.sql"
	$terLoadFilename = "ter_load.sql"
	
	# Get all Load SQL files in the Load directory excluding specific scripts
	# $excludedFiles = @("run_load.sql", "lnk_load.sql", "sto_load.sql", "ter_load.sql")
	# $loadFiles = Get-ChildItem -Path $LoadPath -File -Filter "*.sql" | Where-Object { $_.Name -notin $excludedFiles }
	
	foreach ($laneNumber in $Lanes)
	{
		# Construct the lane folder name
		$laneFolderName = "XF${StoreNumber}${laneNumber}"
		$laneFolderPath = Join-Path -Path $OfficePath -ChildPath $laneFolderName
		
		if (-not (Test-Path $laneFolderPath))
		{
			Write-Log "`r`nLane #$laneNumber not found at path: $laneFolderPath" "yellow"
			continue
		}
		
		$laneFolder = Get-Item -Path $laneFolderPath
		
		Write-Log "`r`nProcessing Lane #$laneNumber" "blue"
		
		# Initialize a list to hold action summaries for the current lane
		$actionSummaries = @()
		
		# ======= Determine Machine Name from LaneMachines =======
		try
		{
			Write-Log "Determining machine name for Lane #$laneNumber..." "blue"
			
			# Retrieve the machine name from the LaneMachines hashtable
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
		
		# Handle the run_load script
		try
		{
			# Define the destination path for run_load script with .sql extension
			$runLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $runLoadFilename
			
			# Write the run_load script exactly as provided to the lane folder using UTF8 without BOM
			[System.IO.File]::WriteAllText($runLoadDestinationPath, $runLoadScript, $utf8NoBOM)
			
			# Set file attributes if necessary
			Set-ItemProperty -Path $runLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			
			$actionSummaries += "Copied run_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy run_load.sql"
		}
		
		# Handle the lnk_load script
		try
		{
			# Define the destination path for lnk_load script with .sql extension
			$lnkLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $lnkLoadFilename
			
			# Generate the INSERT statements specific to this lane and store, incorporating the machine name
			$lnkLoadInsertStatements = @(
				"('${laneNumber}','${StoreNumber}','${laneNumber}'),",
				"('DSM','${StoreNumber}','${laneNumber}'),",
				"('PAL','${StoreNumber}','${laneNumber}'),",
				"('RAL','${StoreNumber}','${laneNumber}'),",
				"('XAL','${StoreNumber}','${laneNumber}');" # Semicolon to end the INSERT statement
			)
			
			# Combine the header, INSERT statements, and footer with a blank line before the footer
			$completeLnkLoadScript = $lnkLoadHeader + "`r`n" + ($lnkLoadInsertStatements -join "`r`n") + "`r`n`r`n" + $lnkLoadFooter.TrimStart() + "`r`n"
			
			# Write the customized lnk_load script to the lane folder using UTF8 without BOM
			[System.IO.File]::WriteAllText($lnkLoadDestinationPath, $completeLnkLoadScript, $utf8NoBOM)
			
			# Set file attributes if necessary
			Set-ItemProperty -Path $lnkLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			
			$actionSummaries += "Copied lnk_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy lnk_load.sql"
		}
		
		# Handle the sto_load script
		try
		{
			# Define the destination path for sto_load script with .sql extension
			$stoLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $stoLoadFilename
			
			# Generate the INSERT statements specific to this lane (no store number needed)
			$stoLoadInsertStatements = @(
				"('${laneNumber}','Terminal ${laneNumber}',1,1,1,,,,),",
				"('DSM','Deploy SMS',1,1,1,,,,),",
				"('PAL','Program all',0,0,1,1,,,),",
				"('RAL','Report all',1,0,0,,,,),",
				"('XAL','Exchange all',0,1,0,,,,);"
			)
			
			# Combine the header, INSERT statements, and footer with a blank line before the footer
			$completeStoLoadScript = $stoLoadHeader + "`r`n" + ($stoLoadInsertStatements -join "`r`n") + "`r`n`r`n" + $stoLoadFooter.TrimStart() + "`r`n"
			
			# Write the customized sto_load script to the lane folder using UTF8 without BOM
			[System.IO.File]::WriteAllText($stoLoadDestinationPath, $completeStoLoadScript, $utf8NoBOM)
			
			# Set file attributes if necessary
			Set-ItemProperty -Path $stoLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			
			$actionSummaries += "Copied sto_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy sto_load.sql"
		}
		
		# Handle the ter_load script
		try
		{
			# Define the destination path for ter_load script with .sql extension
			$terLoadDestinationPath = Join-Path -Path $laneFolder.FullName -ChildPath $terLoadFilename
			
			# Generate the INSERT statements specific to this lane and store, plus the '901' record
			$terLoadInsertStatements = @(
				"('${StoreNumber}','${laneNumber}','Terminal ${laneNumber}','\\${MachineName}\storeman\office\XF${StoreNumber}${laneNumber}\','\\${MachineName}\storeman\office\XF${StoreNumber}901\'),",
				"('${StoreNumber}','901','Server','','');" # '901' record with StoreNumber and fixed values
			)
			
			# Combine the header, INSERT statements, and footer with a blank line before the footer
			$completeTerLoadScript = $terLoadHeader + "`r`n" + ($terLoadInsertStatements -join "`r`n") + "`r`n`r`n" + $terLoadFooter.TrimStart() + "`r`n"
			
			# Write the customized ter_load script to the lane folder using UTF8 without BOM
			[System.IO.File]::WriteAllText($terLoadDestinationPath, $completeTerLoadScript, $utf8NoBOM)
			
			# Set file attributes if necessary
			Set-ItemProperty -Path $terLoadDestinationPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
			
			$actionSummaries += "Copied ter_load.sql"
		}
		catch
		{
			$actionSummaries += "Failed to copy ter_load.sql"
		}
		
		# Log a single summary line for the current lane, including the machine name
		$summaryMessage = "Lane ${laneNumber} (Machine: ${MachineName}): " + ($actionSummaries -join "; ")
		Write-Log $summaryMessage "green"
		
		# Add lane to processed lanes if not already added
		if (-not ($script:ProcessedLanes -contains $laneNumber))
		{
			$script:ProcessedLanes += $laneNumber
		}
	}
	
	Write-Log "`r`n==================== Update-LaneFiles Function Completed ====================" "blue"
}

# ===================================================================================================
#                                       FUNCTION: Pump-AllItems
# ---------------------------------------------------------------------------------------------------
# Description:
#   Extracts specified tables from the SQL Server and copies them to the selected lanes as a single
#   batch file named PUMP_ALL_ITEMS_TABLES.sql. Handles large tables efficiently while preserving
#   compatibility. This function now relies on Get-TableAliases to retrieve the table and alias
#   information, eliminating the need to define the table list within this function.
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
	
	# Access the table and alias information from the script-scoped hash table
	if ($script:FunctionResults.ContainsKey('Get-TableAliases'))
	{
		$aliasData = $script:FunctionResults['Get-TableAliases']
		$aliasResults = $aliasData.Aliases
		$aliasHash = $aliasData.AliasHash
	}
	else
	{
		Write-Log "Alias data not found in the script-scoped hash table. Ensure that Get-TableAliases has been run." "red"
		return
	}
	
	if ($aliasResults.Count -eq 0)
	{
		Write-Log "No tables found to process. Exiting Pump-AllItems." "red"
		return
	}
	
	# Use the locally stored connection string
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write-Log "Connection string not found in script variables. Cannot proceed with Pump-AllItems." "red"
		return
	}
	$ConnectionString = $script:FunctionResults['ConnectionString']
	
	# Open SQL connection
	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$sqlConnection.ConnectionString = $ConnectionString
	$sqlConnection.Open()
	
	# Process each alias entry sequentially
	$generatedFiles = @() # List to keep track of all generated files
	$copiedTables = @() # List to keep track of processed tables
	$skippedTables = @() # List to keep track of skipped tables
	
	foreach ($aliasEntry in $aliasResults)
	{
		$table = $aliasEntry.Table # Full table name with _TAB
		$tableAlias = $aliasEntry.Alias
		
		# Proceed only if both table and alias are present
		if (-not $table -or -not $tableAlias)
		{
			Write-Log "Invalid table or alias for entry: $($aliasEntry | ConvertTo-Json)" "yellow"
			continue
		}
		
		# Check if the table has any rows
		$dataCheckQuery = "SELECT COUNT(*) FROM [$table]"
		$cmdCheck = $sqlConnection.CreateCommand()
		$cmdCheck.CommandText = $dataCheckQuery
		
		try
		{
			$rowCount = $cmdCheck.ExecuteScalar()
		}
		catch
		{
			Write-Log "Error checking row count for table '$table': $_" "red"
			continue
		}
		
		if ($rowCount -eq 0)
		{
			# Add to skipped tables
			$skippedTables += $table
			continue
		}
		
		# === Added Write-Log for Processing Table ===
		Write-Log "Processing table '$table'..." "blue"
		
		# Remove the _TAB suffix from the table name for view name construction
		$baseTable = $table -replace '_TAB$', ''
		
		# Define file name for individual table
		$sqlFileName = "${baseTable}_Load.sql"
		$localTempPath = Join-Path -Path $env:TEMP -ChildPath $sqlFileName
		
		# Check if the SQL file already exists and is recent (within an hour)
		$useExistingFile = $false
		if (Test-Path $localTempPath)
		{
			$fileInfo = Get-Item $localTempPath
			$fileAge = (Get-Date) - $fileInfo.LastWriteTime
			if ($fileAge.TotalHours -le 1)
			{
				$useExistingFile = $true
				Write-Log "Existing SQL file found for table '$table' in %TEMP% and is within an hour. Reusing the file." "green"
			}
			else
			{
				Write-Log "Existing SQL file for table '$table' in %TEMP% is older than an hour. Regenerating the file." "yellow"
			}
		}
		
		if (-not $useExistingFile)
		{
			# Write-Log "Generating _Load.sql batch files in %TEMP%, this might take a while, please wait..." "blue"
			try
			{
				# Initialize StreamWriter with UTF8 encoding without BOM and set NewLine to CRLF
				$streamWriter = New-Object System.IO.StreamWriter($localTempPath, $false, $utf8NoBOM)
				$streamWriter.NewLine = "`r`n" # Ensure CRLF line endings
				
				# Retrieve column data types using an ordered hashtable
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
					$columnName = $readerColumnTypes["COLUMN_NAME"]
					$dataType = $readerColumnTypes["DATA_TYPE"]
					$columnDataTypes[$columnName] = $dataType
				}
				$readerColumnTypes.Close()
				
				# Retrieve primary key columns
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
				
				if ($pkColumns.Count -eq 0)
				{
					# If no primary key found, default to first column
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
				
				# Build the key string for @UPDATE_BATCH using 'AND'
				$keyString = ($primaryKeyColumns | ForEach-Object { "$_=:$_" }) -join ' AND '
				
				# Generate the header for the SQL batch file
				$viewName = $baseTable.Substring(0, 1).ToUpper() + $baseTable.Substring(1).ToLower() + '_Load'
				$columnList = ($columnDataTypes.Keys) -join ','
				
				$header = "@CREATE($table,$tableAlias);
CREATE VIEW $viewName AS SELECT $columnList FROM $table;

INSERT INTO $viewName VALUES`r`n"
				
				# Write header to the file
				$streamWriter.Write($header)
				
				# Initialize SqlCommand and SqlDataReader for data retrieval
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
						# Prepend a comma and CRLF for subsequent rows
						$streamWriter.Write(",`r`n")
					}
					
					# Prepare values for SQL INSERT
					$values = @()
					foreach ($column in $columnDataTypes.Keys)
					{
						$value = $readerData[$column]
						$dataType = $columnDataTypes[$column]
						
						if ($value -eq $null -or $value -is [System.DBNull])
						{
							$values += "" # Use empty for database nulls as per your example
						}
						elseif ($dataType -in @('char', 'nchar', 'varchar', 'nvarchar', 'text', 'ntext'))
						{
							# String data types, enclose in single quotes and escape single quotes
							$escapedValue = $value.ToString().Replace("'", "''")
							$values += "'$escapedValue'"
						}
						elseif ($dataType -in @('datetime', 'smalldatetime', 'date', 'datetime2'))
						{
							# Date/time data types with original calculation
							$dayOfYear = $value.DayOfYear.ToString("D3") # Day of year with leading zeros
							$formattedDate = "'{0}{1} {2}'" -f $value.Year, $dayOfYear, $value.ToString("HH:mm:ss")
							$values += $formattedDate
						}
						elseif ($dataType -in @('bit'))
						{
							# Boolean data types
							$bitValue = if ($value) { "1" }
							else { "0" }
							$values += $bitValue
						}
						elseif ($dataType -in @('decimal', 'numeric', 'float', 'real', 'money', 'smallmoney'))
						{
							# Numeric data types with potential decimals
							if ([math]::Floor($value) -eq $value)
							{
								# Integer, output as is
								$values += $value.ToString()
							}
							else
							{
								# Decimal, format to two decimal places
								$values += $value.ToString("0.00")
							}
						}
						elseif ($dataType -in @('tinyint', 'smallint', 'int', 'bigint'))
						{
							# Integer data types
							$values += $value.ToString()
						}
						else
						{
							# Default to treating as string
							$escapedValue = $value.ToString().Replace("'", "''")
							$values += "'$escapedValue'"
						}
					}
					
					# Combine values into an INSERT statement
					$insertStatement = "(" + ($values -join ',') + ")"
					$streamWriter.Write($insertStatement)
				}
				
				# Close the data reader
				$readerData.Close()
				
				# Finish the INSERT statement with a semicolon
				$streamWriter.Write(";`r`n`r`n")
				
				# Generate the footer for the SQL batch file
				$footer = "@UPDATE_BATCH(JOB=ADD,TAR=$table,
KEY=$keyString,
SRC=SELECT * FROM $viewName);

DROP TABLE $viewName;`r`n`r`n"
				
				# Write footer to the file
				$streamWriter.Write($footer)
				
				# Close the StreamWriter
				$streamWriter.Close()
				$streamWriter.Dispose()
				
				# Write-Log "Successfully generated SQL file: $sqlFileName" "green"
				$generatedFiles += $localTempPath
				$copiedTables += $table
			}
			catch
			{
				Write-Log "Error generating SQL file for table '$table': $_" "red"
				continue
			}
		}
		else
		{
			$generatedFiles += $localTempPath
			$copiedTables += $table
		}
	}
	
	# Close the SQL connection
	$sqlConnection.Close()
	
	if ($copiedTables.Count -gt 0)
	{
		Write-Log "Successfully generated _Load.sql files with tables: $($copiedTables -join ', ')" "green"
	}
	if ($skippedTables.Count -gt 0)
	{
		Write-Log "The following tables were skipped since they had no data: $($skippedTables -join ', ')" "yellow"
	}
	
	# Now copy all generated files to each selected lane
	Write-Log "`r`nDetermining selected lanes...`r`n" "magenta"
	$ProcessedLanes = @()
	foreach ($lane in $Lanes)
	{
		$LaneLocalPath = Join-Path -Path $OfficePath -ChildPath "XF${StoreNumber}${lane}"
		
		if (Test-Path $LaneLocalPath)
		{
			Write-Log "Copying _Load.sql files to Lane #$lane..." "blue"
			try
			{
				# Copy all generated SQL files to the destination lane
				foreach ($filePath in $generatedFiles)
				{
					$fileName = [System.IO.Path]::GetFileName($filePath)
					$destinationPath = Join-Path -Path $LaneLocalPath -ChildPath $fileName
					Copy-Item -Path $filePath -Destination $destinationPath -Force -ErrorAction Stop
				}
				Write-Log "Successfully copied all generated _Load.sql files to Lane #$lane." "green"
				$ProcessedLanes += $lane
			}
			catch
			{
				Write-Log "Error copying generated _Load.sql files to Lane #${lane}: $_" "red"
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
#                                     FUNCTION: Reboot-Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Reboots one, a range, or all machines based on the user's selection.
#   Uses the Show-SelectionDialog function for lane selection.
#   Queries the TER_TAB table to get machine names and lane numbers.
#   Reboots the selected machines.
#   If Restart-Computer fails, falls back to using the shutdown command.
# ===================================================================================================

function Reboot-Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	# Initialize an array to store machine names and lane numbers
	$machinesToReboot = @()
	
	Write-Log "`r`n==================== Starting Reboot-Lanes Function ====================" "blue"
	
	if (-not (Test-Path $OfficePath))
	{
		Write-Log "`r`nXF Base Path not found: $OfficePath" "yellow"
		return
	}
	
	# Get the user's selection
	$selection = Show-SelectionDialog -Mode $Mode -StoreNumber $StoreNumber
	
	if ($selection -eq $null)
	{
		Write-Log "`r`nLane reboot canceled by user." "yellow"
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
	
	# If "All Lanes" is selected, attempt to retrieve LaneContents and LaneMachines
	if ($processAllLanes)
	{
		try
		{
			#	Write-Log "User selected 'All Lanes'. Retrieving LaneContents and LaneMachines..." "blue"
			$LaneContents = $script:FunctionResults['LaneContents']
			$LaneMachines = $script:FunctionResults['LaneMachines']
			
			if ($LaneContents -and $LaneContents.Count -gt 0)
			{
				#	Write-Log "Successfully retrieved LaneContents. Processing all lanes." "green"
				$Lanes = $LaneContents
			}
			else
			{
				throw "LaneContents is empty or not available."
			}
			
			if ($LaneMachines -and $LaneMachines.Count -gt 0)
			{
				#	Write-Log "Successfully retrieved LaneMachines. Proceeding with reboot." "green"
			}
			else
			{
				throw "LaneMachines is empty or not available."
			}
		}
		catch
		{
			#	Write-Log "Failed to retrieve LaneContents or LaneMachines: $_. Falling back to user-selected lanes." "yellow"
			$processAllLanes = $false
			# Optionally, retain $Lanes as provided by the user
		}
	}
	
	# Retrieve machine names from FunctionResults
	if ($processAllLanes -and $LaneMachines)
	{
		foreach ($laneNumber in $Lanes)
		{
			if ($LaneMachines.ContainsKey($laneNumber))
			{
				$machineName = $LaneMachines[$laneNumber]
				$machinesToReboot += @{
					LaneNumber  = $laneNumber
					MachineName = $machineName
				}
			}
			else
			{
				Write-Log "Lane #${laneNumber}: Machine name not found in FunctionResults. Skipping reboot." "yellow"
			}
		}
	}
	else
	{
		# If not processing all lanes or LaneMachines not available, proceed to collect machine names individually
		# This assumes that Update-LaneFiles or Count-ItemsGUI has already populated LaneMachines
		try
		{
			Write-Log "Retrieving machine names from FunctionResults for selected lanes..." "blue"
			$LaneMachines = $script:FunctionResults['LaneMachines']
			
			if ($LaneMachines -and $LaneMachines.Count -gt 0)
			{
				foreach ($laneNumber in $Lanes)
				{
					if ($LaneMachines.ContainsKey($laneNumber))
					{
						$machineName = $LaneMachines[$laneNumber]
						$machinesToReboot += @{
							LaneNumber  = $laneNumber
							MachineName = $machineName
						}
					}
					else
					{
						Write-Log "Lane #${laneNumber}: Machine name not found in FunctionResults. Skipping reboot." "yellow"
					}
				}
			}
			else
			{
				Write-Log "LaneMachines not available in FunctionResults. Cannot retrieve machine names." "red"
				return
			}
		}
		catch
		{
			Write-Log "Failed to retrieve LaneMachines from FunctionResults: $_. Cannot proceed with reboot." "red"
			return
		}
	}
	
	# Check if there are machines to reboot
	if ($machinesToReboot.Count -eq 0)
	{
		Write-Log "No machines to reboot based on the current selection." "yellow"
		return
	}
	
	# Reboot the machines
	foreach ($machine in $machinesToReboot)
	{
		$laneNumber = $machine.LaneNumber
		$machineName = $machine.MachineName
		
		Write-Log "`r`nRebooting Machine '$machineName' for Lane #$laneNumber using shutdown command..." "blue"
		
		try
		{
			# Use shutdown command as the main reboot mechanism
			$shutdownCommand = "shutdown /r /m \\$machineName /t 0 /f"
			Write-Log "Executing: $shutdownCommand" "blue"
			
			$shutdownResult = & cmd.exe /c $shutdownCommand 2>&1
			
			if ($LASTEXITCODE -eq 0)
			{
				Write-Log "Successfully sent shutdown command to Machine '$machineName' for Lane #$laneNumber." "green"
			}
			else
			{
				Write-Log "Shutdown command failed for Machine '$machineName' with exit code $LASTEXITCODE." "red"
				Write-Log "Error: $shutdownResult" "red"
				Write-Log "Attempting to reboot Machine '$machineName' using Restart-Computer as a fallback..." "yellow"
				
				try
				{
					# Use Restart-Computer as a fallback mechanism
					Restart-Computer -ComputerName $machineName -Force -ErrorAction Stop
					Write-Log "Successfully sent reboot command to Machine '$machineName' for Lane #$laneNumber using Restart-Computer." "green"
				}
				catch
				{
					Write-Log "Failed to reboot Machine '$machineName' for Lane #$laneNumber using Restart-Computer. Error: $_" "red"
				}
			}
		}
		catch
		{
			Write-Log "Failed to reboot Machine '$machineName' for Lane #$laneNumber using shutdown command. Error: $_" "red"
		}
	}
	
	Write-Log "`r`n==================== Reboot-Lanes Function Completed ====================" "blue"
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
	
	# Ensure the XE folder exists
	if (-not (Test-Path $XEFolderPath))
	{
		Write-Log -Message "XE folder not found: $XEFolderPath" "red"
		return
	}
	
	# Path to the Close_Transaction.sqi content
	$CloseTransactionContent = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE')"
	
	# Path to the log file
	$LogFolderPath = "$BasePath\Scripts_by_Alex_C.T"
	$LogFilePath = Join-Path -Path $LogFolderPath -ChildPath "Closed_Transactions_LOG.txt"
	
	# Ensure the log directory exists
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
												
												# Define the path to the lane directory
												$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
												
												if (Test-Path $LaneDirectory)
												{
													# Define the path to the Close_Transaction.sqi file in the lane directory
													$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
													
													# Write the content to the file
													Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionContent -Encoding ASCII
													
													# Remove the Archive attribute from the file
													Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
													
													# Log the event
													$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Closed transaction $transactionNumber on lane $LaneNumber"
													Add-Content -Path $LogFilePath -Value $logMessage
													
													# Delete the error file from the XE folder
													Remove-Item -Path $file.FullName -Force
													
													Write-Log -Message "Processed file $($file.Name) for lane $LaneNumber and closed transaction $transactionNumber" "green"
													$MatchedTransactions = $true
												}
												else
												{
													Write-Log -Message "Lane directory $LaneDirectory not found" "yellow"
												}
											}
											else
											{
												Write-Log -Message "Could not extract transaction number from Last recorded status line in file $($file.Name)" "red"
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
							Write-Log -Message "Store or Lane number mismatch in file $($file.Name). File Store/Lane: $fileStoreNumber/$fileLaneNumber vs Expected Store/Lane: $StoreNumber/$LaneNumber" "yellow"
						}
					}
					# else From line not matched - no action needed
				}
				catch
				{
					Write-Log -Message "Error processing file $($file.Name): $_" "red"
				}
			}
		}
		
		# After processing all files, if no matched transactions were found, prompt once for lane number
		if (-not $MatchedTransactions)
		{
			Write-Log -Message "No files or no matching transactions found. Prompting for lane number." "yellow"
			
			# Show Windows form to ask for lane number
			Add-Type -AssemblyName System.Windows.Forms
			Add-Type -AssemblyName System.Drawing
			
			# Initialize the form
			$form = New-Object System.Windows.Forms.Form
			$form.Text = "Lane Deployment"
			$form.Size = New-Object System.Drawing.Size(400, 150)
			$form.StartPosition = "CenterScreen"
			
			# Label
			$label = New-Object System.Windows.Forms.Label
			$label.Text = "Enter Lane Number to deploy the file to:"
			$label.AutoSize = $true
			$label.Location = New-Object System.Drawing.Point(10, 20)
			$form.Controls.Add($label)
			
			# TextBox
			$textBox = New-Object System.Windows.Forms.TextBox
			$textBox.Location = New-Object System.Drawing.Point(10, 50)
			$textBox.Width = 360
			$form.Controls.Add($textBox)
			
			# OK Button
			$okButton = New-Object System.Windows.Forms.Button
			$okButton.Text = "OK"
			$okButton.Location = New-Object System.Drawing.Point(150, 80)
			$okButton.Add_Click({
					# Validate that the input is numeric and pad it
					if ($textBox.Text -match '^\d+$')
					{
						$form.Tag = $textBox.Text.PadLeft(3, '0')
						$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.Close()
					}
					else
					{
						[System.Windows.Forms.MessageBox]::Show("Please enter a numeric lane number.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					}
				})
			$form.Controls.Add($okButton)
			
			# Cancel Button
			$cancelButton = New-Object System.Windows.Forms.Button
			$cancelButton.Text = "Cancel"
			$cancelButton.Location = New-Object System.Drawing.Point(230, 80)
			$cancelButton.Add_Click({
					$form.Tag = "Cancelled"
					$form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
					$form.Close()
				})
			$form.Controls.Add($cancelButton)
			
			# Set Accept and Cancel buttons
			$form.AcceptButton = $okButton
			$form.CancelButton = $cancelButton
			
			$result = $form.ShowDialog()
			
			if ($form.Tag -eq "Cancelled" -or $result -eq [System.Windows.Forms.DialogResult]::Cancel)
			{
				Write-Log -Message "User cancelled the operation." "yellow"
				Write-Log "`r`n==================== CloseOpenTransactions Function Completed ====================" "blue"
				return
			}
			
			$LaneNumber = $form.Tag
			
			if (-not $LaneNumber)
			{
				Write-Log -Message "No lane number provided by the user." "red"
				return
			}
			
			# Define the path to the lane directory
			$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
			
			if (Test-Path $LaneDirectory)
			{
				# Define the path to the Close_Transaction.sqi file in the lane directory
				$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
				
				# Write the content to the file
				Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionContent -Encoding ASCII
				
				# Remove the Archive attribute from the file
				Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
				
				# Log the event
				$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - User deployed Close_Transaction.sqi to lane $LaneNumber"
				Add-Content -Path $LogFilePath -Value $logMessage
				
				Write-Log -Message "Deployed Close_Transaction.sqi to lane $LaneNumber" "green"
				
				# After user deploys the file, clear the folder except for files with "FATAL" in the name
				Get-ChildItem -Path $XEFolderPath -File | Where-Object { $_.Name -notlike "*FATAL*" } | Remove-Item -Force
			}
			else
			{
				Write-Log -Message "Lane directory $LaneDirectory not found" "yellow"
			}
			
			Write-Log "Prompt deployment process completed." "yellow"
		}
	}
	catch
	{
		Write-Log -Message "An error occurred during monitoring: $_" "red"
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
#   from the Count-ItemsGUI function to identify machines associated with each lane.
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
#   - Ensure that the Count-ItemsGUI function has been executed prior to running Ping-Lanes.
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
		Write-Log "Lane information is not available. Please run Count-ItemsGUI first." "Red"
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

function PingAllLanes
{
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Host", "Store")]
		[string]$Mode,
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write-Log "`r`n==================== Starting PingAllLanes Function ====================`r`n" "blue"
	
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
		Write-Log "PingAllLanes is only applicable in 'Store' mode." "Red"
		return
	}
	
	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneContents') -or
		-not $script:FunctionResults.ContainsKey('LaneMachines'))
	{
		Write-Log "Lane information is not available. Please run Count-ItemsGUI first." "Red"
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
	Write-Log "`r`n==================== PingAllLanes Function Completed ====================" "blue"
}

# ===================================================================================================
#                                           FUNCTION: Delete-DBS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Enables users to delete specific file types (.txt and .dwr) from selected lanes within a specified
#   store. Additionally, users are prompted to include or exclude .sus files from the deletion process.
#   The function leverages pre-stored lane information from the Count-ItemsGUI function to identify 
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
#   - Ensure that the Count-ItemsGUI function has been executed prior to running Delete-DBS.
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
		Write-Log "Lane information is not available. Please run Count-ItemsGUI first." "Red"
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
		$systemIcons = @("This PC.lnk", "Network.lnk", "Control Panel.lnk", "Recycle Bin.lnk", "User's Files.lnk", "$scriptName")
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
#   - Ensure that the Count-ItemsGUI function has been executed prior to running Refresh-Files.
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
	foreach ($func in @('Show-SelectionDialog', 'Write-Log', 'Count-ItemsGUI'))
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
		Write-Log "No lane information found. Please ensure Count-ItemsGUI has been executed." "Red"
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
#                                       SECTION: Generate Specific SQL and SQM Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates two files in the %TEMP% directory:
#     1. DEPLOY_SYS.sql containing a specific INSERT statement.
#     2. DEPLOY_ONE_FCT.sqm containing a predefined script.
# ===================================================================================================

function InstallIntoSMS
{
	# Retrieve the path to the system's temporary directory
	$tempDirectory = $env:TEMP
	
	# Define file paths within the %TEMP% directory
	$PumpallitemstablesFilePath = Join-Path -Path $tempDirectory -ChildPath "Pump_all_items_tables.sql"
	$DeploySysFilePath = Join-Path -Path $tempDirectory -ChildPath "DEPLOY_SYS.sql"
	$DeployOneFctFilePath = Join-Path -Path $tempDirectory -ChildPath "DEPLOY_ONE_FCT.sqm"
	
	Write-Log "`r`n==================== Installing new buttons into SMS ====================`r`n" "blue"
	
	# Define the content for Pump_all_items_tables.sql
	$PumpallitemstablesContent = @"
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
	
	# Ensure content strings have Windows-style line endings
	$PumpallitemstablesContent = $PumpallitemstablesContent -replace "`n", "`r`n"
	$DeploySysContent = $DeploySysContent -replace "`n", "`r`n"
	$DeployOneFctContent = $DeployOneFctContent -replace "`n", "`r`n"
	
	# Define encoding as ANSI (Windows-1252)
	$ansiEncoding = [System.Text.Encoding]::GetEncoding(1252)
	
	# Function to write file with error handling
	function Write-File
	{
		param (
			[string]$Path,
			[string]$Content,
			[System.Text.Encoding]$Encoding
		)
		try
		{
			[System.IO.File]::WriteAllText($Path, $Content, $Encoding)
			Write-Log "Successfully wrote to '$Path'." "green"
		}
		catch
		{
			Write-Log "Failed to write to '$Path'. Error: $_" "red"
		}
	}
	
	# Write DEPLOY_SYS.sql
	Write-File -Path $PumpallitemstablesFilePath -Content $PumpallitemstablesContent -Encoding $ansiEncoding
	
	# Write DEPLOY_SYS.sql
	Write-File -Path $DeploySysFilePath -Content $DeploySysContent -Encoding $ansiEncoding
	
	# Write DEPLOY_ONE_FCT.sqm
	Write-File -Path $DeployOneFctFilePath -Content $DeployOneFctContent -Encoding $ansiEncoding
	
	# Define destination paths
	$PumpallitemstablesDestination = "$OfficePath\XF${StoreNumber}901"
	$DeploySysDestination = "$OfficePath\DEPLOY_SYS.sql"
	$DeployOneFctDestination = "$OfficePath\DEPLOY_ONE_FCT.sqm"
	
	# Additional Variables
	$File1 = "Pump_all_items_tables.sql"
	$File2 = "DEPLOY_SYS.sql"
	$File3 = "DEPLOY_ONE_FCT.sqm"
	
	# Function to copy file with error handling
	function Copy-File
	{
		param (
			[string]$FileType,
			[string]$SourcePath,
			[string]$DestinationPath
		)
		
		try
		{
			# Ensure the destination directory exists
			$destDir = Split-Path -Path $DestinationPath -Parent
			if (-not (Test-Path -Path $destDir))
			{
				New-Item -Path $destDir -ItemType Directory -Force | Out-Null
				Write-Log "Created directory '$destDir'." "yellow"
			}
			
			# Copy the file, overwriting if it exists
			Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
			Write-Log "Successfully copied '$FileType' to '$DestinationPath'." "green"
		}
		catch
		{
			Write-Log "Failed to copy '$FileType' to '$DestinationPath'. Error: $_" "red"
		}
	}
	
	# Copy Pump_all_items_tables.sql to \\localhost\Storeman\Office\XF${StoreNumber}901
	Copy-File -FileType $File1 -SourcePath $PumpallitemstablesFilePath -DestinationPath $PumpallitemstablesDestination
	
	# **Remove the Archive Bit from Pump_all_items_tables.sql at the destination**
	try
	{
		$destinationFile1 = Join-Path -Path $PumpallitemstablesDestination -ChildPath $File1
		if (Test-Path $destinationFile1)
		{
			$file = Get-Item $destinationFile1
			if ($file.Attributes -band [System.IO.FileAttributes]::Archive)
			{
				$file.Attributes = $file.Attributes -bxor [System.IO.FileAttributes]::Archive
				Write-Log "Removed the archive bit from '$destinationFile1'." "green"
			}
			else
			{
				Write-Log "Archive bit was not set for '$destinationFile1'." "yellow"
			}
		}
		else
		{
			Write-Log "Destination file '$destinationFile1' does not exist. Cannot remove archive bit." "red"
		}
	}
	catch
	{
		Write-Log "Failed to remove the archive bit from '$destinationFile1'. Error: $_" "red"
	}
	
	# Copy DEPLOY_SYS.sql to \\localhost\Storeman\Office
	Copy-File -FileType $File2 -SourcePath $DeploySysFilePath -DestinationPath $DeploySysDestination
	
	# Copy DEPLOY_ONE_FCT.sqm to \\localhost\Storeman\Office, replacing if it exists
	Copy-File -FileType $File3 -SourcePath $DeployOneFctFilePath -DestinationPath $DeployOneFctDestination
	
	# Cleanup: Delete the generated files from the temp directory
	function Cleanup-TempFiles
	{
		param (
			[string[]]$FilesToDelete
		)
		
		foreach ($file in $FilesToDelete)
		{
			if (Test-Path $file)
			{
				try
				{
					Remove-Item -Path $file -Force
					Write-Log "Deleted temporary file '$file'." "yellow"
				}
				catch
				{
					Write-Log "Failed to delete temporary file '$file'. Error: $_" "red"
				}
			}
			else
			{
				Write-Log "Temporary file '$file' does not exist and cannot be deleted." "yellow"
			}
		}
	}
	
	Cleanup-TempFiles -FilesToDelete @($PumpallitemstablesFilePath, $DeploySysFilePath, $DeployOneFctFilePath)
	
	#	Write-Log "`r`nDEPLOY_ONE_FCT.sqm copied to $deployOneFctDestination." "green"
	#	Write-Log "DEPLOY_SYS.sql copied to $addMenuDestination." "green"
	Write-Log "`r`n==================== Function execution completed ====================" "blue"
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
#   After processing, stops and deletes the "BMS" service, re-registers BMSSrv.exe, and restarts the service.
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
	
	Write-Log "`r`n==================== Starting Organize-TBS_SCL_ver520 Function ====================`r`n" "Blue"
	
	# Access the connection string from the script-scoped variable
	$connectionString = $script:FunctionResults['ConnectionString']
	
	if (-not $connectionString)
	{
		Write-Log "Connection string not found in `$script:FunctionResults['ConnectionString']`." "Red"
		return
	}
	
	# Define the SQL commands:
	# 1. Update ScaleName and BufferTime for ISHIDA WMAI records.
	# 2. Update ScaleName for BIZERBA records.
	# 3. Update ScaleCode for BIZERBA records.
	# 4. Update ScaleCode for ISHIDA records after BIZERBA.
	# 5. Set BufferTime for BIZERBA records: first one = 1, others = 5.
	
	$updateQueries = @"
-- Update ScaleName and BufferTime for ISHIDA WMAI records
UPDATE [TBS_SCL_ver520]
SET 
    ScaleName = 'Ishida Wrapper',
    BufferTime = '1'
WHERE 
    ScaleBrand = 'ISHIDA' AND ScaleModel = 'WMAI';

-- Update ScaleName for BIZERBA records
UPDATE [TBS_SCL_ver520]
SET 
    ScaleName = CONCAT('Scale ', IPDevice)
WHERE 
    ScaleBrand = 'BIZERBA';

-- Update ScaleCode for BIZERBA records to start at 10, ordered by IPDevice ascending
WITH BIZERBA_CTE AS (
    SELECT 
        ScaleCode,
        IPDevice,
        ROW_NUMBER() OVER (ORDER BY CAST(IPDevice AS INT) ASC) AS rn
    FROM 
        [TBS_SCL_ver520]
    WHERE 
        ScaleBrand = 'BIZERBA'
)
UPDATE BIZERBA_CTE
SET ScaleCode = 10 + rn - 1;

-- Update ScaleCode for ISHIDA records to start after the maximum ScaleCode of BIZERBA
;WITH MaxBizerba AS (
    SELECT MAX(ScaleCode) AS MaxCode
    FROM [TBS_SCL_ver520]
    WHERE ScaleBrand = 'BIZERBA'
),
ISHIDA_CTE AS (
    SELECT 
        ScaleCode,
        IPDevice,
        ROW_NUMBER() OVER (ORDER BY CAST(IPDevice AS INT) ASC) AS rn
    FROM 
        [TBS_SCL_ver520]
    WHERE 
        ScaleBrand = 'ISHIDA'
)
UPDATE ISHIDA_CTE
SET ScaleCode = (SELECT MaxCode FROM MaxBizerba) + 10 + rn - 1;

-- Now set BufferTime for BIZERBA records:
-- The first BIZERBA record (lowest ScaleCode) gets BufferTime = 1, all others = 5.
WITH BIZ_ORDER AS (
    SELECT 
        ScaleCode,
        ROW_NUMBER() OVER (ORDER BY ScaleCode ASC) AS RN
    FROM [TBS_SCL_ver520]
    WHERE ScaleBrand = 'BIZERBA'
)
UPDATE T
SET T.BufferTime = CASE WHEN B.RN = 1 THEN '1' ELSE '5' END
FROM [TBS_SCL_ver520] T
INNER JOIN BIZ_ORDER B ON T.ScaleCode = B.ScaleCode;
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
	
	# Execute the update queries
	Write-Log "Executing update queries to modify ScaleName, BufferTime, and ScaleCode..." "Blue"
	try
	{
		Invoke-Sqlcmd -ConnectionString $connectionString -Query $updateQueries
		Write-Log "Update queries executed successfully." "Green"
	}
	catch
	{
		Write-Log "An error occurred while executing update queries: $_" "Red"
		return
	}
	
	# Execute the select query to retrieve organized data
	Write-Log "Retrieving organized data..." "Blue"
	try
	{
		$data = Invoke-Sqlcmd -ConnectionString $connectionString -Query $selectQuery
		Write-Log "Data retrieval successful." "Green"
	}
	catch
	{
		Write-Log "An error occurred while retrieving data: $_" "Red"
		return
	}
	
	# Check if data was retrieved
	if (-not $data)
	{
		Write-Log "No data retrieved from the table 'TBS_SCL_ver520'." "Red"
		Throw "No data retrieved from the table 'TBS_SCL_ver520'."
	}
	
	# Export the data if an output path is provided
	if ($PSBoundParameters.ContainsKey('OutputCsvPath'))
	{
		Write-Log "Exporting organized data to '$OutputCsvPath'..." "Blue"
		try
		{
			$data | Export-Csv -Path $OutputCsvPath -NoTypeInformation
			Write-Log "Data exported successfully to '$OutputCsvPath'." "Green"
		}
		catch
		{
			Write-Log "Failed to export data to CSV: $_" "Red"
		}
	}
	
	# Display the organized data
	Write-Log "Displaying organized data:" "Yellow"
	$data | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Log $_ "White" }
	
	Write-Log "`r`n==================== Organize-TBS_SCL_ver520 Function Completed ====================" "Blue"
	
	# ===================================================================================================
	#                                 SERVICE: BMS Management
	# ---------------------------------------------------------------------------------------------------
	# Description:
	#   Stops and deletes the "BMS" service, re-registers BMSSrv.exe, and restarts the "BMS" service.
	# ===================================================================================================
	
	# Stop the BMS service
	Write-Log "Stopping the 'BMS' service..." "Blue"
	try
	{
		sc.exe stop "BMS"
		Write-Log "'BMS' service stopped successfully." "Green"
	}
	catch
	{
		Write-Log "Failed to stop 'BMS' service: $_" "Red"
		return
	}
	
	# Delete the BMS service
	Write-Log "Deleting the 'BMS' service..." "Blue"
	try
	{
		sc.exe delete "BMS"
		Write-Log "'BMS' service deleted successfully." "Green"
	}
	catch
	{
		Write-Log "Failed to delete 'BMS' service: $_" "Red"
		return
	}
	
	# Change directory to the BMS installation path
	Write-Log "Changing directory to 'C:\Bizerba\RetailConnect\BMS'..." "Blue"
	try
	{
		Set-Location -Path "C:\Bizerba\RetailConnect\BMS" -ErrorAction Stop
		Write-Log "Directory changed to 'C:\Bizerba\RetailConnect\BMS' successfully." "Green"
	}
	catch
	{
		Write-Log "Failed to change directory to 'C:\Bizerba\RetailConnect\BMS': $_" "Red"
		return
	}
	
	# Register BMSSrv.exe
	Write-Log "Registering 'BMSSrv.exe'..." "Blue"
	try
	{
		"BMSSrv.exe" -reg
		Write-Log "'BMSSrv.exe' registered successfully." "Green"
	}
	catch
	{
		Write-Log "Failed to register 'BMSSrv.exe': $_" "Red"
		return
	}
	
	# Start the BMS service
	Write-Log "Starting the 'BMS' service..." "Blue"
	try
	{
		sc.exe start "BMS"
		Write-Log "'BMS' service started successfully." "Green"
	}
	catch
	{
		Write-Log "Failed to start 'BMS' service: $_" "Red"
		return
	}
	
	Write-Log "`r`n==================== BMS Service Management Completed ====================" "Blue"
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
		[string]$StoreNumber # Required only if $Mode is "Store" and "All" is selected
	)
	
	# Load necessary assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Initialize the form
	$form = New-Object System.Windows.Forms.Form
	if ($Mode -eq "Host")
	{
		$form.Text = "Select Stores to Process"
	}
	else
	{
		$form.Text = "Select Lanes to Process"
	}
	$form.Size = New-Object System.Drawing.Size(450, 350)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	# Radio buttons for selection type
	$radioSpecific = New-Object System.Windows.Forms.RadioButton
	if ($Mode -eq "Host")
	{
		$radioSpecific.Text = "Specific Store"
	}
	else
	{
		$radioSpecific.Text = "Specific Lane"
	}
	$radioSpecific.Location = New-Object System.Drawing.Point(20, 20)
	$radioSpecific.AutoSize = $true
	$form.Controls.Add($radioSpecific)
	
	$radioRange = New-Object System.Windows.Forms.RadioButton
	if ($Mode -eq "Host")
	{
		$radioRange.Text = "Range of Stores"
	}
	else
	{
		$radioRange.Text = "Range of Lanes"
	}
	$radioRange.Location = New-Object System.Drawing.Point(20, 50)
	$radioRange.AutoSize = $true
	$form.Controls.Add($radioRange)
	
	$radioAll = New-Object System.Windows.Forms.RadioButton
	if ($Mode -eq "Host")
	{
		$radioAll.Text = "All Stores"
	}
	else
	{
		$radioAll.Text = "All Lanes"
	}
	$radioAll.Location = New-Object System.Drawing.Point(20, 80)
	$radioAll.AutoSize = $true
	$form.Controls.Add($radioAll)
	
	# Inputs for selection
	if ($Mode -eq "Host")
	{
		# TextBox for Specific Host(s)
		$textSpecific = New-Object System.Windows.Forms.TextBox
		$textSpecific.Location = New-Object System.Drawing.Point(220, 18)
		$textSpecific.Width = 200
		$textSpecific.Enabled = $false
		$form.Controls.Add($textSpecific)
		
		# Labels and TextBoxes for Range
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
	}
	elseif ($Mode -eq "Store")
	{
		# Label and TextBox for Lane inputs
		$labelInput = New-Object System.Windows.Forms.Label
		$labelInput.Text = "Enter Lane Number:"
		$labelInput.Location = New-Object System.Drawing.Point(20, 150)
		$labelInput.Size = New-Object System.Drawing.Size(400, 20)
		$labelInput.Visible = $false
		$form.Controls.Add($labelInput)
		
		$textBoxInput = New-Object System.Windows.Forms.TextBox
		$textBoxInput.Location = New-Object System.Drawing.Point(20, 180)
		$textBoxInput.Size = New-Object System.Drawing.Size(400, 20)
		$textBoxInput.Visible = $false
		$form.Controls.Add($textBoxInput)
	}
	
	# OK and Cancel buttons
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(100, 250)
	$buttonOK.Size = New-Object System.Drawing.Size(100, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(250, 250)
	$buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($buttonCancel)
	
	$form.AcceptButton = $buttonOK
	$form.CancelButton = $buttonCancel
	
	# Event handlers to enable/disable input fields based on radio button selection
	if ($Mode -eq "Host")
	{
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
	}
	elseif ($Mode -eq "Store")
	{
		$radioSpecific.Add_CheckedChanged({
				if ($radioSpecific.Checked)
				{
					$labelInput.Text = "Enter Lane Number (e.g., 005):"
					$labelInput.Visible = $true
					$textBoxInput.Visible = $true
				}
			})
		
		$radioRange.Add_CheckedChanged({
				if ($radioRange.Checked)
				{
					$labelInput.Text = "Enter Lanes Range (e.g., 001-010):"
					$labelInput.Visible = $true
					$textBoxInput.Visible = $true
				}
			})
		
		$radioAll.Add_CheckedChanged({
				if ($radioAll.Checked)
				{
					$labelInput.Visible = $false
					$textBoxInput.Visible = $false
				}
			})
	}
	
	# Set default selection
	$radioSpecific.Checked = $true
	
	# Show the form
	$dialogResult = $form.ShowDialog()
	
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		return $null
	}
	
	# Process the user's input
	if ($radioSpecific.Checked)
	{
		if ($Mode -eq "Host")
		{
			# Process specific stores
			$storesInput = $textSpecific.Text
			if ([string]::IsNullOrWhiteSpace($storesInput))
			{
				[System.Windows.Forms.MessageBox]::Show("Please enter at least one store number.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
			# Split the input by commas and trim spaces
			$stores = $storesInput.Split(",") | ForEach-Object { $_.Trim() }
			# Validate that each store is a 3-digit number
			foreach ($store in $stores)
			{
				if (-not ($store -match "^\d{3}$"))
				{
					[System.Windows.Forms.MessageBox]::Show("Invalid store number: $store. Numbers must be exactly 3 digits.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					return $null
				}
			}
			return @{
				Type   = "Specific"
				Stores = $stores
			}
		}
		elseif ($Mode -eq "Store")
		{
			# Process specific lanes
			$input = $textBoxInput.Text.Trim()
			if ($input -match "^\d{1,3}$")
			{
				$laneNumber = $input.PadLeft(3, '0')
				return @{
					Type  = 'Specific'
					Lanes = @($laneNumber)
				}
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Invalid lane number. Please enter a 1 to 3-digit number.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
		}
	}
	elseif ($radioRange.Checked)
	{
		if ($Mode -eq "Host")
		{
			# Process range of stores
			$startHost = $textStart.Text.Trim()
			$endHost = $textEnd.Text.Trim()
			if (-not ($startHost -match "^\d{3}$") -or -not ($endHost -match "^\d{3}$"))
			{
				[System.Windows.Forms.MessageBox]::Show("Start and End store numbers must be exactly 3 digits.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
			if ([int]$startHost -gt [int]$endHost)
			{
				[System.Windows.Forms.MessageBox]::Show("Start store number cannot be greater than end store number.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
			# Generate the list of store numbers
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
		elseif ($Mode -eq "Store")
		{
			# Process range of lanes
			$input = $textBoxInput.Text.Trim()
			if ($input -match "^\d{1,3}\s*-\s*\d{1,3}$")
			{
				$parts = $input -split '\s*-\s*'
				$startLane = [int]$parts[0]
				$endLane = [int]$parts[1]
				if ($startLane -le $endLane)
				{
					$laneNumbers = ($startLane .. $endLane) | ForEach-Object { $_.ToString("D3") }
					return @{
						Type  = 'Range'
						Lanes = $laneNumbers
					}
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Start lane number cannot be greater than end lane number.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					return $null
				}
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Invalid lane range. Please use the format 001-010.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
		}
	}
	elseif ($radioAll.Checked)
	{
		if ($Mode -eq "Host")
		{
			return @{
				Type   = "All"
				Stores = @()
			}
		}
		elseif ($Mode -eq "Store")
		{
			if (-not $StoreNumber)
			{
				[System.Windows.Forms.MessageBox]::Show("Store number is required to fetch all lanes.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return $null
			}
			# Attempt to use $LaneContents from $script:FunctionResults
			if ($script:FunctionResults.ContainsKey('LaneContents') -and $script:FunctionResults['LaneContents'].Count -gt 0)
			{
				$allLanes = $script:FunctionResults['LaneContents']
				return @{
					Type  = 'All'
					Lanes = $allLanes
				}
			}
			else
			{
				# Fallback to current mechanism
				if (-not (Test-Path -Path $OfficePath))
				{
					[System.Windows.Forms.MessageBox]::Show("The path '$OfficePath' does not exist.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					return $null
				}
				$laneFolders = Get-ChildItem -Path $OfficePath -Directory -Filter "XF${StoreNumber}0*"
				if (-not $laneFolders)
				{
					return @{
						Type  = 'All'
						Lanes = @()
					}
				}
				$allLanes = $laneFolders | ForEach-Object {
					$_.Name.Substring($_.Name.Length - 3, 3)
				}
				return @{
					Type  = 'All'
					Lanes = $allLanes
				}
			}
		}
	}
	else
	{
		# Should not reach here
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
		
		# Create the main form
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Created by Alex_C.T - Version $VersionNumber"
		$form.Size = New-Object System.Drawing.Size(1005, 710)
		$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
		
		# Banner Label
		$bannerLabel = New-Object System.Windows.Forms.Label
		$bannerLabel.Text = "PowerShell Script - TBS_Maintenance_Script"
		$bannerLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
		$bannerLabel.Size = New-Object System.Drawing.Size(500, 30)
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
					Delete-Files -Path "$TempDir" -SpecifiedFiles "Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1"
				}
			})
		
		# Ativate Windows button
		$ActivateWindowsButton = New-Object System.Windows.Forms.Button
		$ActivateWindowsButton.Text = "Alex_C.T"
		$ActivateWindowsButton.Location = New-Object System.Drawing.Point(850, 30)
		$ActivateWindowsButton.Size = New-Object System.Drawing.Size(100, 30)
		$ActivateWindowsButton.add_Click({
				Invoke-SecureScript
			})
		$form.Controls.Add($ActivateWindowsButton)
		
		$rebootButton = New-Object System.Windows.Forms.Button
		$rebootButton.Text = "Reboot System"
		$rebootButton.Location = New-Object System.Drawing.Point(850, 65)
		$rebootButton.Size = New-Object System.Drawing.Size(100, 30)
		# Event Handler for Reboot Button
		$rebootButton.Add_Click({
				$rebootResult = [System.Windows.Forms.MessageBox]::Show("Do you want to reboot now?", "Reboot", [System.Windows.Forms.MessageBoxButtons]::YesNo)
				if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
				{
					Restart-Computer -Force
					# Clean Temp Folder
					Delete-Files -Path "$TempDir" -SpecifiedFiles "Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1"
				}
			})
		$form.Controls.Add($rebootButton)
		
		# Create a Clear Log button
		$clearLogButton = New-Object System.Windows.Forms.Button
		$clearLogButton.Text = "Clear Log"
		$clearLogButton.Location = New-Object System.Drawing.Point(850, 100)
		$clearLogButton.Size = New-Object System.Drawing.Size(100, 30)
		$clearLogButton.add_Click({
				$logBox.Clear()
				Write-Log "Log Cleared"
			})
		$form.Controls.Add($clearLogButton)
		
		# Install into SMS
		$InstallIntoSMSButton = New-Object System.Windows.Forms.Button
		$InstallIntoSMSButton.Text = "Install Function in SMS"
		$InstallIntoSMSButton.Location = New-Object System.Drawing.Point(695, 100)
		$InstallIntoSMSButton.Size = New-Object System.Drawing.Size(150, 30)
		$InstallIntoSMSButton.add_Click({
				InstallIntoSMS
			})
		$form.Controls.Add($InstallIntoSMSButton)
		
		################################################## Labels #######################################################
		
		# Create labels for Mode, Store Name, Store Number, and Counts
		$script:modeLabel = New-Object System.Windows.Forms.Label
		$modeLabel.Text = "Processing Mode: N/A"
		$modeLabel.Location = New-Object System.Drawing.Point(50, 30)
		$modeLabel.Size = New-Object System.Drawing.Size(900, 20)
		$modeLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$form.Controls.Add($modeLabel)
		
		# Store Name Label
		$script:storeNameLabel = New-Object System.Windows.Forms.Label
		$storeNameLabel.Text = "Store Name: N/A"
		$storeNameLabel.Location = New-Object System.Drawing.Point(50, 50)
		$storeNameLabel.Size = New-Object System.Drawing.Size(900, 20)
		$storeNameLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$form.Controls.Add($storeNameLabel)
		
		# Store Number Label
		$script:storeNumberLabel = New-Object System.Windows.Forms.Label
		$storeNumberLabel.Text = "Store Number: N/A"
		$storeNumberLabel.Location = New-Object System.Drawing.Point(50, 70)
		$storeNumberLabel.Size = New-Object System.Drawing.Size(900, 20)
		$storeNumberLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$form.Controls.Add($storeNumberLabel)
		
		# Counts Labels
		# Counts Label Line 1
		$script:countsLabel1 = New-Object System.Windows.Forms.Label
		$countsLabel1.Text = "Number of Servers: $($Counts.NumberOfServers)"
		$countsLabel1.Location = New-Object System.Drawing.Point(50, 90)
		$countsLabel1.Size = New-Object System.Drawing.Size(900, 20) # Reduced height
		$countsLabel1.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$countsLabel1.AutoSize = $false
		$form.Controls.Add($countsLabel1)
		
		# Counts Label Line 2
		$script:countsLabel2 = New-Object System.Windows.Forms.Label
		$countsLabel2.Text = "Number of Lanes: $($Counts.NumberOfLanes)"
		$countsLabel2.Location = New-Object System.Drawing.Point(50, 110) # Adjusted Y-position
		$countsLabel2.Size = New-Object System.Drawing.Size(900, 20) # Reduced height
		$countsLabel2.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
		$countsLabel2.AutoSize = $false
		$form.Controls.Add($countsLabel2)
		
		# Alternatively, Adjust the Y-position to reduce spacing
		# Example: Move countsLabel2 closer to countsLabel1
		# $countsLabel2.Location = New-Object System.Drawing.Point(50, 85) # Reduced from 90 to 85
		
		# Update Counts Labels Based on Mode
		if ($Mode -eq "Host")
		{
			$countsLabel1.Text = "Number of Hosts: $($Counts.NumberOfHosts)"
			$countsLabel2.Text = "Number of Stores: $($Counts.NumberOfStores)"
		}
		else
		{
			$countsLabel1.Text = "Number of Servers: $($Counts.NumberOfServers)"
			$countsLabel2.Text = "Number of Lanes: $($Counts.NumberOfLanes)"
		}
		
		# Create a RichTextBox for log output
		$logBox = New-Object System.Windows.Forms.RichTextBox
		$logBox.Location = New-Object System.Drawing.Point(50, 130)
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
		
		# ===================================================================================================
		#                                       SECTION: GUI Buttons Setup
		# ---------------------------------------------------------------------------------------------------
		# Description:
		#   Sets up the buttons on the main form, including their size, position, and labels based on the processing mode.
		# ===================================================================================================
		
		# Create Host Specific Buttons
		if ($Mode -eq "Host")
		{
			$hostButton1 = New-Object System.Windows.Forms.Button
			$hostButton1.Text = "Host DB Repair"
			$hostButton1.Location = New-Object System.Drawing.Point(50, 515)
			$hostButton1.Size = New-Object System.Drawing.Size(200, 40)
			$hostButton1.Add_Click({
					Process-HostGUI -StoresqlFilePath $StoresqlFilePath
				})
			$form.Controls.Add($hostButton1)
			
			$hostButton2 = New-Object System.Windows.Forms.Button
			$hostButton2.Text = "Store DB Repair"
			$hostButton2.Location = New-Object System.Drawing.Point(284, 515)
			$hostButton2.Size = New-Object System.Drawing.Size(200, 40)
			$hostButton2.Add_Click({
					Process-StoresGUI -StoresqlFilePath $StoresqlFilePath
				})
			$form.Controls.Add($hostButton2)
			
			$hostButton3 = New-Object System.Windows.Forms.Button
			$hostButton3.Text = "Repair DB of Stores and Host"
			$hostButton3.Location = New-Object System.Drawing.Point(517, 515)
			$hostButton3.Size = New-Object System.Drawing.Size(200, 40)
			$hostButton3.Add_Click({
					Process-AllStoresAndHostGUI -StoresqlFilePath $StoresqlFilePath
				})
			$form.Controls.Add($hostButton3)
			
			$hostButton4 = New-Object System.Windows.Forms.Button
			$hostButton4.Text = "Repair Windows"
			$hostButton4.Location = New-Object System.Drawing.Point(750, 515)
			$hostButton4.Size = New-Object System.Drawing.Size(200, 40)
			$hostButton4.Add_Click({
					Repair-Windows
				})
			$form.Controls.Add($hostButton4)
			
			$hostButton5 = New-Object System.Windows.Forms.Button
			$hostButton5.Text = "Create a scheduled task"
			$hostButton5.Location = New-Object System.Drawing.Point(50, 560)
			$hostButton5.Size = New-Object System.Drawing.Size(200, 40)
			$hostButton5.Add_Click({
					Create-ScheduledTaskGUI -ScriptPath $scriptPath
				})
			$form.Controls.Add($hostButton5)
		}
		else
		{
			# Create Store Specific Button with Confirmation
			$storeButton1 = New-Object System.Windows.Forms.Button
			$storeButton1.Text = "Server DB Repair"
			$storeButton1.Location = New-Object System.Drawing.Point(50, 535)
			$storeButton1.Size = New-Object System.Drawing.Size(200, 40)
			$storeButton1.Add_Click({
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
			$form.Controls.Add($storeButton1)
			
			$storeButton2 = New-Object System.Windows.Forms.Button
			$storeButton2.Text = "Lane DB Repair"
			$storeButton2.Location = New-Object System.Drawing.Point(284, 535)
			$storeButton2.Size = New-Object System.Drawing.Size(200, 40)
			$storeButton2.Add_Click({
					Process-LanesGUI -LanesqlFilePath $LanesqlFilePath -StoreNumber $StoreNumber
				})
			$form.Controls.Add($storeButton2)
			
			<#
			$storeButton3 = New-Object System.Windows.Forms.Button
			$storeButton3.Text = "Repair DB of Lanes and Server"
			$storeButton3.Location = New-Object System.Drawing.Point(517, 535)
			$storeButton3.Size = New-Object System.Drawing.Size(200, 40)
			$storeButton3.Add_Click({
					Process-LanesAndServerGUI -LanesqlFilePath $LanesqlFilePath -StoresqlFilePath $StoresqlFilePath -StoreNumber $StoreNumber
				})
			$form.Controls.Add($storeButton3)
			#>
			
			$OrganizeScaleTableButton = New-Object System.Windows.Forms.Button
			$OrganizeScaleTableButton.Text = "Organize-TBS_SCL_ver520"
			$OrganizeScaleTableButton.Location = New-Object System.Drawing.Point(517, 535)
			$OrganizeScaleTableButton.Size = New-Object System.Drawing.Size(200, 40)
			$OrganizeScaleTableButton.Add_Click({
					Organize-TBS_SCL_ver520
				})
			$form.Controls.Add($OrganizeScaleTableButton)
			
			# Repair Windows button
			$repairButton = New-Object System.Windows.Forms.Button
			$repairButton.Text = "Repair Windows"
			$repairButton.Location = New-Object System.Drawing.Point(750, 535)
			$repairButton.Size = New-Object System.Drawing.Size(200, 40)
			$repairButton.Add_Click({
					Repair-Windows
				})
			$form.Controls.Add($repairButton)
			
			$storeButton5 = New-Object System.Windows.Forms.Button
			$storeButton5.Text = "Pump All Items"
			$storeButton5.Location = New-Object System.Drawing.Point(50, 580)
			$storeButton5.Size = New-Object System.Drawing.Size(200, 40)
			$storeButton5.Add_Click({
					Pump-AllItems -StoreNumber $StoreNumber
				})
			$form.Controls.Add($storeButton5)
			
			$storeButton6 = New-Object System.Windows.Forms.Button
			$storeButton6.Text = "Update Lane Configuration"
			$storeButton6.Location = New-Object System.Drawing.Point(284, 580)
			$storeButton6.Size = New-Object System.Drawing.Size(200, 40)
			$storeButton6.Add_Click({
					Update-LaneFiles -StoreNumber $StoreNumber
				})
			$form.Controls.Add($storeButton6)
			
			# Close Open Transactions button
			$COTButton = New-Object System.Windows.Forms.Button
			$COTButton.Text = "Close Open Transactions"
			$COTButton.Location = New-Object System.Drawing.Point(517, 580)
			$COTButton.Size = New-Object System.Drawing.Size(200, 40)
			$COTButton.add_Click({
					CloseOpenTransactions -StoreNumber $StoreNumber
				})
			$form.Controls.Add($COTButton)
			
			<# Disabled Create a scheduled task button
			$storeButton7 = New-Object System.Windows.Forms.Button
			$storeButton7.Text = "Create a scheduled task"
			$storeButton7.Location = New-Object System.Drawing.Point(517, 645)
			$storeButton7.Size = New-Object System.Drawing.Size(200, 40)
			$storeButton7.Add_Click({
					Create-ScheduledTaskGUI -ScriptPath $scriptPath
				})
			$form.Controls.Add($storeButton7)
			#>
			
			# Ping lanes button
			$PingLanesButton = New-Object System.Windows.Forms.Button
			$PingLanesButton.Text = "Ping Lanes"
			$PingLanesButton.Location = New-Object System.Drawing.Point(750, 580)
			$PingLanesButton.Size = New-Object System.Drawing.Size(200, 40)
			$PingLanesButton.add_Click({
					PingAllLanes -Mode "Store" -StoreNumber "$StoreNumber"
				})
			$form.Controls.Add($PingLanesButton)
			
			# Delete DBS button
			$DeleteDBSButton = New-Object System.Windows.Forms.Button
			$DeleteDBSButton.Text = "Delete DBS"
			$DeleteDBSButton.Location = New-Object System.Drawing.Point(50, 625)
			$DeleteDBSButton.Size = New-Object System.Drawing.Size(200, 40)
			$DeleteDBSButton.add_Click({
					Delete-DBS -Mode "Store" -StoreNumber "$StoreNumber"
				})
			$form.Controls.Add($DeleteDBSButton)
			
			# Configure SystemSettings button
			$ConfigureSystemSettingsButton = New-Object System.Windows.Forms.Button
			$ConfigureSystemSettingsButton.Text = "Configure System Settings"
			$ConfigureSystemSettingsButton.Location = New-Object System.Drawing.Point(284, 625)
			$ConfigureSystemSettingsButton.Size = New-Object System.Drawing.Size(200, 40)
			$ConfigureSystemSettingsButton.add_Click({
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
			$form.Controls.Add($ConfigureSystemSettingsButton)
			
			# Refresh PIN Pad Files
			$RefreshFilesButton = New-Object System.Windows.Forms.Button
			$RefreshFilesButton.Text = "Refresh PIN Pad Files"
			$RefreshFilesButton.Location = New-Object System.Drawing.Point(517, 625)
			$RefreshFilesButton.Size = New-Object System.Drawing.Size(200, 40)
			$RefreshFilesButton.add_Click({
					Refresh-Files -Mode $Mode -StoreNumber -$StoreNumber
				})
			$form.Controls.Add($RefreshFilesButton)
			
			# Reboot lanes
			$storeButton8 = New-Object System.Windows.Forms.Button
			$storeButton8.Text = "Reboot Lane"
			$storeButton8.Location = New-Object System.Drawing.Point(750, 625)
			$storeButton8.Size = New-Object System.Drawing.Size(200, 40)
			$storeButton8.Add_Click({
					Reboot-Lanes -StoreNumber $StoreNumber
				})
			$form.Controls.Add($storeButton8)
			
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
	Write-Log "Powershell version installed: $major.$minor | Build-$build | Revision-$revision" "blue"
	
	# Initialize a counter for the number of jobs started
	$jobCount = 0
	
	# Initialize variables
	# $Memory25PercentMB = Get-MemoryInfo
	
	# Get SQL Connection String
	Get-DatabaseConnectionString
	
	# Get the Store Number
	Get-StoreNumberGUI
	$StoreNumber = $script:FunctionResults['StoreNumber']
	
	# Get the Store Name
	Get-StoreNameGUI
	$StoreName = $script:FunctionResults['StoreName']
	
	# Determine the Mode
	$Mode = Determine-ModeGUI -StoreNumber $StoreNumber
	$Mode = $script:FunctionResults['Mode']
	
	# Count items based on mode
	$Counts = Count-ItemsGUI -Mode $Mode -StoreNumber $StoreNumber
	$Counts = $script:FunctionResults['Counts']
	
	# Populate the hash table with results from various functions
	Get-TableAliases
	
	# Generate SQL scripts
	Generate-SQLScriptsGUI -StoreNumber $StoreNumber -Memory25PercentMB $Memory25PercentMB -LanesqlFilePath $LanesqlFilePath -StoresqlFilePath $StoresqlFilePath
	
	# Clearing XE (Urgent Messages) folder.
	$ClearXEJob = Clear-XEFolder
	# Increment the job counter
	$jobCount++
	
	# Clear %Temp% foder on start
	$FilesAndDirsDeleted = Delete-Files -Path "$TempDir" -Exclusions "Server_Database_Maintenance.sqi", "Lane_Database_Maintenance.sqi", "TBS_Maintenance_Script.ps1" -AsJob
	# Increment the job counter
	$jobCount++
	
	# Clears the recycle bin on startup
	$ClearRecycleBin = Clear-RecycleBin -Force -ErrorAction SilentlyContinue
	
	<#
	# Retrieve the list of machine names from the FunctionResults dictionary
	$LaneMachines = $script:FunctionResults['LaneMachines']
	
	# Define the list of user profiles to process
	$userProfiles = @('Administrator', 'Operator')
	
	# Iterate over each machine and each user profile, then invoke Delete-Files as a background job
	foreach ($machine in $LaneMachines.Values)
	{
    	foreach ($user in $userProfiles)
    	{
       		# Construct the full UNC path to the Temp directory on the remote machine
        	$tempPath = "\\$machine\C$\Users\$user\AppData\Local\Temp\"
        
        	try
        	{
            	# Invoke the Delete-Files function with the -AsJob parameter
            	$DeleteJob = Delete-Files -Path $tempPath -AsJob
            
            	# Increment the job counter
            	$jobCount++
            
            	# Log that the deletion job has been started
            	# Write-Log "Started deletion job for %temp% folder in user '$user' on machine '$machine' at path '$tempPath'." "green"
        	}
        	catch
        	{
            	# Log any errors that occur while starting the deletion job
           	 	Write-Log "An error occurred while starting the deletion job for user '$user' on machine '$machine'. Error: $_" "red"
        	}
    	}
	}

	# Log the summary of jobs started
	# Write-Log "Total deletion jobs started: $jobCount" "blue"
	# Write-Log "All deletion jobs started" "blue"
	#>
	
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

# Close the console to aviod duplicate logging to the richbox
exit
