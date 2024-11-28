# ===================================================================================================
#                                       SECTION: Import Modules
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports necessary PowerShell modules required for the script's operation.
# ===================================================================================================

# Import necessary modules
Import-Module -Name Microsoft.PowerShell.Utility

# ===================================================================================================
#                                       SECTION: Ensure Administrator Privileges
# ---------------------------------------------------------------------------------------------------
# Description:
#   Ensures that the script is running with administrative privileges. If not, it attempts to restart the script with elevated rights.
# ===================================================================================================

function Ensure-Administrator {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        try {
            # Reconstruct the original arguments without adding -Silent
            $originalArgs = $MyInvocation.UnboundArguments -join ' '
            Start-Process -FilePath "PowerShell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $originalArgs" -Verb RunAs
            exit
        } catch {
            Write-Error "Failed to elevate to administrator."
            exit 1
        }
    }
}

# ===================================================================================================
#                                       SECTION: Import Necessary Assemblies
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports required .NET assemblies for creating and managing Windows Forms and graphical components.
# ===================================================================================================

# Add necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===================================================================================================
#                                       SECTION: Initialize GUI Components
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and initializes the main graphical user interface (GUI) form and its components.
# ===================================================================================================

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mini Ghost PowerShell Script"
$form.Size = New-Object System.Drawing.Size(505, 370)
$form.StartPosition = "CenterScreen"

# Banner Label
$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Text = "PowerShell Script - Mini Ghost"
$bannerLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
$bannerLabel.Size = New-Object System.Drawing.Size(500, 40)
$bannerLabel.TextAlign = 'MiddleCenter'
$bannerLabel.Dock = 'Top'

$form.Controls.Add($bannerLabel)

# Created by Label
$createdByLabel = New-Object System.Windows.Forms.Label
$createdByLabel.Text = "Created by Alex_C.T"
$createdByLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
$createdByLabel.Size = New-Object System.Drawing.Size(500, 25)
$createdByLabel.TextAlign = 'MiddleCenter'
$createdByLabel.Dock = 'Top'

$form.Controls.Add($createdByLabel)

# ===================================================================================================
#                                       SECTION: Script Variables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Initializes all necessary variables required for the script's operation.
# ===================================================================================================

# Ensure script is running as administrator
# Ensure-Administrator

# Initialize script-scoped variables for new store number and new machine name
$script:newStoreNumber = $null
$script:newMachineName = $null

# Get the current machine name
$currentMachineName = $env:COMPUTERNAME

# Define paths
$startupIniPath = "\\localhost\storeman\startup.ini"
$baseDirectory = "\\localhost\storeman\office"  # Set the base directory for folder retrieval

# ===================================================================================================
#                                       FUNCTION: Get-StoreNameFromINI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store name from the system INI file. Returns "N/A" if not found.
# ===================================================================================================

function Get-StoreNameFromINI {
    param (
        [string]$iniPath = '\\localhost\storeman\office\system.ini'
    )

    $storeName = 'N/A'

    if (Test-Path $iniPath) {
        $iniContent = Get-Content $iniPath
        foreach ($line in $iniContent) {
            if ($line -match '^NAME=(.*)') {
                $storeName = $matches[1].Trim()
                break
            }
        }
    } else {
        $storeName = "INI file not found at $iniPath"
    }

    return $storeName
}

$storeName = Get-StoreNameFromINI

# ===================================================================================================
#                                       FUNCTION: Get-CurrentStoreNumber
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the current store number from the startup.ini file or XF directories without prompting the user.
# ===================================================================================================

function Get-CurrentStoreNumber {
    param (
        [string]$startupIniPath,
        [string]$baseDirectory
    )
    
    # Define helper functions
    function Format-StoreNumber {
        param (
            [string]$storeNumber
        )
        return $storeNumber.PadLeft(3, '0')
    }
    
    function Get-StoreNumberFromINI {   
        if (Test-Path $startupIniPath) {
            $iniContent = Get-Content $startupIniPath
            foreach ($line in $iniContent) {
                if ($line -match "^STORE=(\d{3})") {
                    return $matches[1]  # Return store number found in the .ini file
                }
            }
        }
        return $null
    }
    
    $storeNumber = Get-StoreNumberFromINI

    if ($storeNumber -eq $null) {
        # If store number not found in startup.ini, check the base directory for XF folders    
        if (-Not (Test-Path $baseDirectory)) {
            $storeNumber = "Unknown"
        } else {
            # Retrieve the folders starting with "XF" to extract the store number
            $folders = Get-ChildItem -Path $baseDirectory -Directory | Where-Object { $_.Name -match "^XF\d{6}" }

            # Initialize an array to store unique store numbers
            $storeNumbers = @()

            # Extract the store numbers from folder names
            foreach ($folder in $folders) {
                $folderName = $folder.Name
                $folderStoreNumber = $folderName.Substring(2, 3)  # Extract characters 2-4 as the store number

                # Add store number to the list if it's not already included
                if (-not $storeNumbers.Contains($folderStoreNumber)) {
                    $storeNumbers += $folderStoreNumber
                }
            }

            # Determine store number based on retrieval results
            if ($storeNumbers.Count -eq 1) {
                # Only one store number found
                $storeNumber = Format-StoreNumber $storeNumbers[0]  # Ensure it's 3 digits
            } elseif ($storeNumbers.Count -gt 1) {
                # Multiple store numbers found
                $storeNumber = "Multiple store numbers found: $($storeNumbers -join ', ')"
            } else {
                # No store numbers found
                $storeNumber = "Unknown"
            }
        }
    }

    return $storeNumber
}

# Get the current store number without prompting
$currentStoreNumber = Get-CurrentStoreNumber -startupIniPath $startupIniPath -baseDirectory $baseDirectory

# ===================================================================================================
#                                       FUNCTION: Get-ActiveIPConfig
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the active IP configuration for network adapters that are up and have valid IPv4 addresses.
# ===================================================================================================

function Get-ActiveIPConfig {
    $ipConfig = Get-NetIPConfiguration | Where-Object {
        $_.NetAdapter.Status -eq 'Up' -and $_.IPv4Address.IPAddress -notlike '169.254*' -and $_.IPv4Address.IPAddress -ne $null
    }

    if ($ipConfig) {
        # Return all valid configuration objects
        return $ipConfig
    } else {
        return $null
    }
}

# Get current IP configuration
$currentConfigs = Get-ActiveIPConfig

# ===================================================================================================
#                                       SECTION: Display Labels
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and displays labels on the GUI form to show current machine information.
# ===================================================================================================

# Display current machine name and store number in labels
$machineNameLabel = New-Object System.Windows.Forms.Label
$machineNameLabel.Text = "Current Machine Name: $env:COMPUTERNAME"
$machineNameLabel.Location = New-Object System.Drawing.Point(10, 70)
$machineNameLabel.Size = New-Object System.Drawing.Size(480, 20)

$storeNameLabel = New-Object System.Windows.Forms.Label
$storeNameLabel.Text = "Store Name: $storeName"
$storeNameLabel.Location = New-Object System.Drawing.Point(10, 95)
$storeNameLabel.Size = New-Object System.Drawing.Size(480, 20)

$storeNumberLabel = New-Object System.Windows.Forms.Label
$storeNumberLabel.Text = "Store Number: $currentStoreNumber"
$storeNumberLabel.Location = New-Object System.Drawing.Point(10, 120)
$storeNumberLabel.Size = New-Object System.Drawing.Size(480, 20)

# Display current IP address in a label
$currentIP = if ($currentConfigs -and $currentConfigs.Count -gt 0) {
    $currentConfigs[0].IPv4Address.IPAddress
} else {
    "IP Not Found"
}

$ipAddressLabel = New-Object System.Windows.Forms.Label
$ipAddressLabel.Text = "Current IP Address: $currentIP"
$ipAddressLabel.Location = New-Object System.Drawing.Point(10, 145)
$ipAddressLabel.Size = New-Object System.Drawing.Size(480, 20)

$form.Controls.AddRange(@($machineNameLabel, $storeNameLabel, $storeNumberLabel, $ipAddressLabel))

# ===================================================================================================
#                                       SECTION: Operation Tracking
# ---------------------------------------------------------------------------------------------------
# Description:
#   Initializes a hashtable to track the status of each operation performed by the script.
# ===================================================================================================

# Initialize a hashtable to track the status of each operation
$operationStatus = @{
    "StoreNumberChange"        = @{ Status = "Pending"; Message = ""; Details = "" }
    "MachineNameChange"        = @{ Status = "Pending"; Message = ""; Details = "" }
    "OldXFoldersDeletion"      = @{ Status = "Pending"; Message = ""; Details = "" }
    "StartupIniUpdate"         = @{ Status = "Pending"; Message = ""; Details = "" }
    "IPConfiguration"          = @{ Status = "Pending"; Message = ""; Details = "" }
    "TableTruncation"          = @{ Status = "Pending"; Message = ""; Details = "" }
    "DatabaseRepair"           = @{ Status = "Pending"; Message = ""; Details = "" }
    "RegistryCleanup"          = @{ Status = "Pending"; Message = ""; Details = "" }
    "SQLDatabaseUpdate"        = @{ Status = "Pending"; Message = ""; Details = "" }
    "ConfigurePowerSettings"   = @{ Status = "Pending"; Message = ""; Details = "" } 
    "ConfigureServices"        = @{ Status = "Pending"; Message = ""; Details = "" } 
    "ConfigureAdvancedSettings"= @{ Status = "Pending"; Message = ""; Details = "" }
}

# ===================================================================================================
#                                       SECTION: Function Definitions
# ---------------------------------------------------------------------------------------------------
# Description:
#   Defines all the functions used within the script to perform various operations.
# ===================================================================================================

# ===================================================================================================
#                                       FUNCTION: Get-MemoryInfo
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the total system memory and calculates 25% of it in megabytes.
#   This information can be used for memory-related configurations and optimizations.
# ===================================================================================================

# function Get-MemoryInfo {
#     $TotalMemoryKB = (Get-CimInstance Win32_OperatingSystem).TotalVisibleMemorySize
#     $TotalMemoryMB = [math]::Floor($TotalMemoryKB / 1024)
#     $Memory25PercentMB = [math]::Floor($TotalMemoryMB * 0.25)
#     return $Memory25PercentMB
# }

# ===================================================================================================
#                                       FUNCTION: Configure-PowerSettings
# ---------------------------------------------------------------------------------------------------
# Description:
#   Configures the system's power settings to optimize performance, including setting the power plan to High Performance,
#   disabling sleep modes, setting processor performance, and disabling USB selective suspend.
# ===================================================================================================

function Configure-PowerSettings {
    try {
        Write-Output "Starting configuration of power plan and performance settings..."

        # Step 1: Set the power plan to High Performance
        Write-Output "Setting power scheme to High Performance..."
        powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
        Write-Output "Power plan set to High Performance."

        # Step 2: Set system to never sleep
        Write-Output "Disabling standby timeout for AC power..."
        powercfg /change standby-timeout-ac 0
        Write-Output "standby-timeout-ac set to 0."

        Write-Output "Disabling standby timeout for DC power..."
        powercfg /change standby-timeout-dc 0
        Write-Output "standby-timeout-dc set to 0."

        # Step 3: Set minimum processor performance to 100%
        Write-Output "Setting minimum processor state to 100% for AC power..."
        powercfg /setacvalueindex 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
        Write-Output "Processor minimum state for AC set to 100%."

        Write-Output "Setting minimum processor state to 100% for DC power..."
        powercfg /setdcvalueindex 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
        Write-Output "Processor minimum state for DC set to 100%."

        Write-Output "Activating High Performance power scheme..."
        powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
        Write-Output "High Performance power scheme activated."

        # Step 4: Turn off screen never
        Write-Output "Setting monitor timeout to never for AC power..."
        powercfg /change monitor-timeout-ac 0
        Write-Output "monitor-timeout-ac set to never."

        # Step 6: Disable USB selective suspend via registry
        Write-Output "Disabling USB selective suspend using registry..."
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\USB\Parameters"
        $valueName = "DisableSelectiveSuspend"

        if (-not (Test-Path $regPath)) {
            Write-Output "Registry path not found. Creating registry path..."
            New-Item -Path $regPath -Force | Out-Null
            Write-Output "Registry path created."
        } else {
            Write-Output "Registry path exists."
        }

        # Set DisableSelectiveSuspend to 1
        Write-Output "Setting DisableSelectiveSuspend to 1..."
        Set-ItemProperty -Path $regPath -Name $valueName -Value 1 -Type DWord -Force
        Write-Output "USB selective suspend registry setting applied."

        Write-Output "Power plan and performance settings configuration complete. Some changes may require a reboot to take effect."
        [System.Windows.Forms.MessageBox]::Show("Power settings configured successfully. A reboot may be required for all changes to take effect.", "Configure Power Settings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Update operationStatus
        $operationStatus["ConfigurePowerSettings"].Status = "Successful"
        $operationStatus["ConfigurePowerSettings"].Message = "Power settings configured successfully."
        $operationStatus["ConfigurePowerSettings"].Details = "Power plan set to High Performance, sleep settings disabled, processor performance set to 100%, screen timeout set to never, and USB selective suspend disabled."
    }
    catch {
        Write-Error "Error configuring power settings: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to configure power settings. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        # Update operationStatus
        $operationStatus["ConfigurePowerSettings"].Status = "Failed"
        $operationStatus["ConfigurePowerSettings"].Message = "Failed to configure power settings."
        $operationStatus["ConfigurePowerSettings"].Details = $_.Exception.Message
    }
}

# ===================================================================================================
#                                       FUNCTION: Configure-Services
# ---------------------------------------------------------------------------------------------------
# Description:
#   Configures specified services to start automatically and ensures they are running.
# ===================================================================================================

function Configure-Services {
    try {
        Write-Output "Configuring services to start automatically..."

        # Define services to configure
        $services = @("fdPHost", "FDResPub", "SSDPSRV", "upnphost")

        foreach ($service in $services) {
            # Set service to start automatically
            Set-Service -Name $service -StartupType Automatic -ErrorAction Stop
            Write-Output "Service '$service' set to start automatically."

            # Start service if not running
            $serviceStatus = Get-Service -Name $service
            if ($serviceStatus.Status -ne 'Running') {
                Start-Service -Name $service -ErrorAction Stop
                Write-Output "Service '$service' started."
            }
            else {
                Write-Output "Service '$service' is already running."
            }
        }

        Write-Output "Service configuration complete."
        [System.Windows.Forms.MessageBox]::Show("Services configured to start automatically and are running.", "Configure Services", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Update operationStatus
        $operationStatus["ConfigureServices"].Status = "Successful"
        $operationStatus["ConfigureServices"].Message = "Services configured successfully."
        $operationStatus["ConfigureServices"].Details = "Services set to start automatically and verified running status."
    }
    catch {
        Write-Error "Error configuring services: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to configure services. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        # Update operationStatus
        $operationStatus["ConfigureServices"].Status = "Failed"
        $operationStatus["ConfigureServices"].Message = "Failed to configure services."
        $operationStatus["ConfigureServices"].Details = $_.Exception.Message
    }
}

# ===================================================================================================
#                                       FUNCTION: Configure-AdvancedSettings
# ---------------------------------------------------------------------------------------------------
# Description:
#   Configures advanced system settings, including visual effects and ClearType font smoothing.
# ===================================================================================================

function Configure-AdvancedSettings {
    try {
        Write-Output "Configuring Advanced System Settings..."

        # Set visual effects to "Adjust for best performance"
        $visualEffectsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
        Set-ItemProperty -Path $visualEffectsPath -Name VisualFXSetting -Value 2 -Type DWord -Force
        Write-Output "Visual effects set to 'Adjust for best performance'."

        # Set UserPreferencesMask to disable all visual effects
        $desktopPath = "HKCU:\Control Panel\Desktop"
        Set-ItemProperty -Path $desktopPath -Name UserPreferencesMask -Value ([byte[]](0x90,0x12,0x00,0x00)) -Type Binary -Force
        Write-Output "UserPreferencesMask set to disable all visual effects."

        # Enable ClearType font smoothing
        Set-ItemProperty -Path $desktopPath -Name FontSmoothing -Value "2" -Type String -Force
        Set-ItemProperty -Path $desktopPath -Name FontSmoothingType -Value 2 -Type DWord -Force
        Set-ItemProperty -Path $desktopPath -Name FontSmoothingGamma -Value 0x00000578 -Type DWord -Force
        Write-Output "ClearType font smoothing enabled."

        Write-Output "Advanced System Settings configuration complete."
        [System.Windows.Forms.MessageBox]::Show("Advanced system settings configured successfully.", "Configure Advanced Settings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Update operationStatus
        $operationStatus["ConfigureAdvancedSettings"].Status = "Successful"
        $operationStatus["ConfigureAdvancedSettings"].Message = "Advanced system settings configured successfully."
        $operationStatus["ConfigureAdvancedSettings"].Details = "Visual effects set to best performance, UserPreferencesMask updated, and ClearType font smoothing enabled."
    }
    catch {
        Write-Error "Error configuring advanced system settings: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to configure advanced system settings. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        # Update operationStatus
        $operationStatus["ConfigureAdvancedSettings"].Status = "Failed"
        $operationStatus["ConfigureAdvancedSettings"].Message = "Failed to configure advanced system settings."
        $operationStatus["ConfigureAdvancedSettings"].Details = $_.Exception.Message
    }
}

# ===================================================================================================
#                                       FUNCTION: Remove-OldXFolders
# ---------------------------------------------------------------------------------------------------
# Description:
#   Removes old XF and XW folders based on the provided machine name and store number.
# ===================================================================================================

function Remove-OldXFolders {
    param (
        [string]$oldMachineName,
        [string]$oldStoreNumber
    )

    # Define folder types to process
    $folderTypes = @("XF", "XW")

    # Initialize results
    $deletedFolders = @()
    # $keptFolders = @()  # Removed to prevent displaying kept folders
    $failedToDeleteFolders = @()

    # Extract new machine number from new machine name
    if (-not [string]::IsNullOrEmpty($script:newMachineName) -and $script:newMachineName.Length -ge 6) {
        $newMachineNumber = $script:newMachineName.Substring(3,3)
    } else {
        $newMachineNumber = ""
    }

    # Iterate through each folder type
    foreach ($folderType in $folderTypes) {
        # Define the path to the folder type directory
        $folderTypePath = "\\$oldMachineName\storeman\office"

        # Check if the folder type path exists
        if (-not (Test-Path $folderTypePath)) {
            $operationStatus["OldXFoldersDeletion"].Status = "Failed"
            $operationStatus["OldXFoldersDeletion"].Message = "Folder path '$folderTypePath' does not exist."
            $operationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
            return
        }

        # Get all folders starting with the folder type
        $folders = Get-ChildItem -Path $folderTypePath -Directory | Where-Object { $_.Name -like "$folderType*" }

        foreach ($folder in $folders) {
            $folderName = $folder.Name

            # Extract StoreNumber and MachineNumber
            if ($folderName.Length -ge 6) {
                $storeNumber = $folderName.Substring(2,3)
                $machineNumber = $folderName.Substring(5,3)
            } else {
                # Invalid folder name format, skip
                continue
            }

            # Determine if the folder should be kept
            if (($machineNumber -eq "901") -or ($machineNumber -eq $newMachineNumber)) {
                # Keep the folder
                # $keptFolders += $folderName  # Removed to prevent displaying kept folders
            } else {
                # Delete the folder
                try {
                    Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
                    $deletedFolders += $folderName
                } catch {
                    $failedToDeleteFolders += $folderName
                }
            }
        }
    }

    # Build the deletion result
    $resultMessage = ""
    if ($deletedFolders.Count -gt 0) {
        $resultMessage += "Deleted folders:`n$($deletedFolders -join "`n")`n"
    }
    # if ($keptFolders.Count -gt 0) {
    #     $resultMessage += "Kept folders:`n$($keptFolders -join "`n")`n"
    # }
    if ($failedToDeleteFolders.Count -gt 0) {
        $resultMessage += "Failed to delete folders:`n$($failedToDeleteFolders -join "`n")`n"
    }

    # Update operationStatus
    if ($failedToDeleteFolders.Count -eq 0) {
        $operationStatus["OldXFoldersDeletion"].Status = "Successful"
        $operationStatus["OldXFoldersDeletion"].Message = "Old XF and XW folders deleted successfully."
        $operationStatus["OldXFoldersDeletion"].Details = $resultMessage
    } elseif ($deletedFolders.Count -gt 0) {
        $operationStatus["OldXFoldersDeletion"].Status = "Partial Failure"
        $operationStatus["OldXFoldersDeletion"].Message = "Some old XF and XW folders could not be deleted."
        $operationStatus["OldXFoldersDeletion"].Details = $resultMessage
    } else {
        $operationStatus["OldXFoldersDeletion"].Status = "Failed"
        $operationStatus["OldXFoldersDeletion"].Message = "Failed to delete any old XF and XW folders."
        $operationStatus["OldXFoldersDeletion"].Details = $resultMessage
    }

    return
}

# ===================================================================================================
#                                       FUNCTION: Execute-SqlCommand
# ---------------------------------------------------------------------------------------------------
# Description:
#   Executes a given SQL command using the provided connection string.
# ===================================================================================================

function Execute-SqlCommand {
    param (
        [string]$commandText
    )
    
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $sqlConnection.ConnectionString = $connectionString
    $sqlCommand = $sqlConnection.CreateCommand()
    $sqlCommand.CommandText = $commandText

    try {
        $sqlConnection.Open()
        $sqlCommand.ExecuteNonQuery() | Out-Null
        return $true  # Indicate success
    } catch {
        return $false  # Indicate failure
    } finally {
        $sqlConnection.Close()
    }
}

# ===================================================================================================
#                                       FUNCTION: Get-SQLInstanceName
# ---------------------------------------------------------------------------------------------------
# Description:
#   Determines whether a named SQL Server instance exists and returns the appropriate server name.
# ===================================================================================================

function Get-SQLInstanceName {
    $computerName = $env:COMPUTERNAME  # Get machine name
    $namedInstance = "$computerName\SQLEXPRESS"  # Named instance
    $defaultInstance = $computerName  # Default instance (localhost)

    # Check if we can connect to the named instance
    $namedInstanceConnectionString = "Server=$namedInstance;Database=master;Integrated Security=True;"
    $defaultInstanceConnectionString = "Server=$defaultInstance;Database=master;Integrated Security=True;"
    
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection

    # Try to connect to named instance
    $sqlConnection.ConnectionString = $namedInstanceConnectionString
    try {
        $sqlConnection.Open()
        $sqlConnection.Close()
        return $namedInstance  # Named instance is available
    } catch {
        # If named instance fails, return default instance
    }

    return $defaultInstance
}

# ===================================================================================================
#                                       FUNCTION: Get-ValidStoreNumber
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user via a GUI to enter a valid store number (1-3 digits) and returns it padded to 3 digits.
# ===================================================================================================

function Get-ValidStoreNumber {
    while ($true) {
        $storeNumberForm = New-Object System.Windows.Forms.Form
        $storeNumberForm.Text = "Enter New Store Number"
        $storeNumberForm.Size = New-Object System.Drawing.Size(350, 180)
        $storeNumberForm.StartPosition = "CenterParent"

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "New Store Number (1-3 digits):"
        $label.Location = New-Object System.Drawing.Point(10, 20)
        $label.Size = New-Object System.Drawing.Size(315, 20)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10, 50)
        $textBox.Size = New-Object System.Drawing.Size(320, 20)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(85, 90)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(175, 90)
        $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $storeNumberForm.AcceptButton = $okButton
        $storeNumberForm.CancelButton = $cancelButton

        $storeNumberForm.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))

        $dialogResult = $storeNumberForm.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $newStoreNumberInput = $textBox.Text.Trim()

            # Validate store number: 1 to 3 digits
            if ($newStoreNumberInput -match "^\d{1,3}$") {
                # Pad the store number with leading zeros to make it 3 digits
                $paddedStoreNumber = $newStoreNumberInput.PadLeft(3, '0')
                return $paddedStoreNumber
            } else {
                [System.Windows.Forms.MessageBox]::Show("Invalid store number. Please enter 1 to 3 digits.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } elseif ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
            return $null
        }
    }
}

# ===================================================================================================
#                                       FUNCTION: Update-StoreNumberInINI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates the store number in the startup.ini file with the new store number provided.
# ===================================================================================================

function Update-StoreNumberInINI {
    param (
        [string]$newStoreNumber,
        [string]$startupIniPath
    )    
    if (Test-Path $startupIniPath) {
        $iniContent = Get-Content $startupIniPath
        $updatedContent = $iniContent -replace "^STORE=\d{3}", "STORE=$newStoreNumber"
        Set-Content $startupIniPath $updatedContent
        return $true
    } else {
        return $false
    }
}

# ===================================================================================================
#                                       FUNCTION: Get-StoreNumberFromINI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store number from the startup.ini file.
# ===================================================================================================

function Get-StoreNumberFromINI {   
    if (Test-Path $startupIniPath) {
        $iniContent = Get-Content $startupIniPath
        foreach ($line in $iniContent) {
            if ($line -match "^STORE=(\d{3})") {
                return $matches[1]  # Return store number found in the .ini file
            }
        }
    }
    return $null
}

# ===================================================================================================
#                                       FUNCTION: Update-SQLTablesForStoreNumberChange
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates STD_TAB in the SQL database after store number change.
# ===================================================================================================

function Update-SQLTablesForStoreNumberChange {
    param (
        [string]$storeNumber
    )

    # Variables
    $stdTableName = "STD_TAB"

    # SQL commands for updating STD_TAB
    $createViewCommandStd = @"
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
"@

    $updateStdTabCommand = @"
UPDATE $stdTableName 
SET F1056 = '$storeNumber';
"@

    $dropViewCommandStd = "DROP VIEW Std_Load;"

    # Execute the SQL commands
    $sqlCommands = @(
        $createViewCommandStd,
        $updateStdTabCommand,
        $dropViewCommandStd
    )

    $allSqlSuccessful = $true
    $failedSqlCommands = @()

    foreach ($command in $sqlCommands) {
        if (-not (Execute-SqlCommand -commandText $command)) {
            $allSqlSuccessful = $false
            $failedSqlCommands += $command
        }
    }

    # Return the result
    return @{
        Success = $allSqlSuccessful
        FailedCommands = $failedSqlCommands
    }
}

# ===================================================================================================
#                                       FUNCTION: Get-ValidMachineName
# ---------------------------------------------------------------------------------------------------
# Description:
#   Prompts the user via a GUI to enter a valid machine name (POS/SCO followed by 3 digits) and returns it in uppercase.
# ===================================================================================================

function Get-ValidMachineName {
    while ($true) {
        # Create the form
        $machineNameForm = New-Object System.Windows.Forms.Form
        $machineNameForm.Text = "Enter New Machine Name"
        $machineNameForm.Size = New-Object System.Drawing.Size(350, 150)
        $machineNameForm.StartPosition = "CenterParent"

        # Create and add the label
        $label = New-Object System.Windows.Forms.Label
        $label.Text = "New Machine Name (POS/SCO + 3 digits):"
        $label.Location = New-Object System.Drawing.Point(10, 20)
        $label.Size = New-Object System.Drawing.Size(320, 20)
        $label.AutoSize = $true

        # Create and add the text box
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10, 50)
        $textBox.Size = New-Object System.Drawing.Size(315, 20)

        # Create and add the OK button
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(85, 80)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

        # Create and add the Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(175, 80)
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        # Set the Accept and Cancel buttons for the form
        $machineNameForm.AcceptButton = $okButton
        $machineNameForm.CancelButton = $cancelButton

        # Add all controls to the form
        $machineNameForm.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))

        # Show the form and get the result
        $dialogResult = $machineNameForm.ShowDialog()

        # If user presses OK
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $newMachineNameInput = $textBox.Text.Trim().ToUpper()

            # Ensure the user did not just press OK with an empty name
            if ([string]::IsNullOrEmpty($newMachineNameInput)) {
                [System.Windows.Forms.MessageBox]::Show("Machine name cannot be empty. Please enter a valid machine name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                continue
            }

            # Validate the format of the machine name
            if ($newMachineNameInput -match "^(POS|SCO)\d{3}$") {
                
                # Check if the machine name is already in use
                if ($newMachineNameInput -eq $env:COMPUTERNAME) {
                    [System.Windows.Forms.MessageBox]::Show("The new machine name is the same as the current one. Please enter a different machine name.", "Duplicate Machine Name", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    continue
                }
                
                # If validation passes and it's not a duplicate, return the valid machine name
                return $newMachineNameInput
            } else {
                [System.Windows.Forms.MessageBox]::Show("Invalid machine name. Please enter a name in the format POS or SCO followed by exactly 3 digits.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }

        } elseif ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
            # If user presses Cancel, return null
            return $null
        } else {
            # Handle other dialog results if necessary
            return $null
        }

        # Dispose of the form to free resources
        $machineNameForm.Dispose()
    }
}

# ===================================================================================================
#                                       FUNCTION: Update-SQLTablesForMachineNameChange
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates STO_TAB, TER_TAB, LNK_TAB, and RUN_TAB in the SQL database after machine name change.
# ===================================================================================================

function Update-SQLTablesForMachineNameChange {
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

    foreach ($command in $sqlCommands) {
        if (-not (Execute-SqlCommand -commandText $command)) {
            $allSqlSuccessful = $false
            $failedSqlCommands += $command
        }
    }

    # Return the result
    return @{
        Success = $allSqlSuccessful
        FailedCommands = $failedSqlCommands
    }
}

# ===================================================================================================
#                                       FUNCTION: Truncate-Tables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Truncates the specified list of tables in the SQL database.
# ===================================================================================================

function Truncate-Tables {
    param (
        [string[]]$tables
    )

    # Initialize an array to store failed truncate commands
    $failedTruncateTables = @()

    # Prepare and execute Truncate Commands
    foreach ($table in $tables) {
        $command = "TRUNCATE TABLE $table;"
        if (-not (Execute-SqlCommand -commandText $command)) {
            $failedTruncateTables += $table  # Add failed table to the array
        }
    }

    # Return the list of failed tables
    return $failedTruncateTables
}

# ===================================================================================================
#                                       FUNCTION: Repair-Database
# ---------------------------------------------------------------------------------------------------
# Description:
#   Performs various SQL database repair operations, including configuration changes, table truncations,
#   and index rebuilding.
# ===================================================================================================

function Repair-Database {
    # Initialize an array to store failed additional commands
    $failedAdditionalCommands = @()

    # Additional SQL Operations excluding recovery and shrink commands
    $additionalCommands = @(
        @"
-- Declare a variable to hold 25% of total physical memory in MB
DECLARE @Memory25PercentMB BIGINT;

-- Calculate 25% of total physical memory and assign it to the variable
SELECT @Memory25PercentMB = (total_physical_memory_kb / 1024) * 25 / 100 
FROM sys.dm_os_sys_memory;

-- Set memory configuration
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', @Memory25PercentMB;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;
"@,

        # Truncate unnecessary tables
        "IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;",
        "IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;",
        "IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;",
        "IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;",
        "IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;",
        "IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;",
        "IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;",
        "IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;",
        "IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;",
        "IF OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('dbo.TBS_ITM_SMAppUPDATED', 'OBJECT', 'DELETE') = 1 DELETE FROM dbo.TBS_ITM_SMAppUPDATED;",

        # Drop temporary tables
        @"
DECLARE @cmd varchar(4000);
DECLARE cmds CURSOR FOR
SELECT 'DROP TABLE [' + Table_Name + ']' 
FROM INFORMATION_SCHEMA.TABLES 
WHERE Table_Name LIKE 'TMP_%';
OPEN cmds;
WHILE 1 = 1
BEGIN
    FETCH cmds INTO @cmd;
    IF @@fetch_status != 0 BREAK;
    EXEC(@cmd);
END;
CLOSE cmds;
DEALLOCATE cmds;
"@,

        # Drop specific tables older than 30 days
        @"
DECLARE @cmd1 varchar(4000);
DECLARE cmds CURSOR FOR
SELECT 'DROP TABLE [' + name + ']' 
FROM sys.tables 
WHERE (name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$currentStoreNumber%') 
  AND DATEDIFF(DAY, create_date, GETDATE()) > 30;
OPEN cmds;
WHILE 1 = 1
BEGIN
    FETCH cmds INTO @cmd1;
    IF @@fetch_status != 0 BREAK;
    EXEC(@cmd1);
END;
CLOSE cmds;
DEALLOCATE cmds;
"@,

        # Cleaning HEADER_SAV
        "IF OBJECT_ID('HEADER_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('HEADER_SAV', 'OBJECT', 'DELETE') = 1 
    DELETE FROM HEADER_SAV 
    WHERE (F903 = 'SVHOST' OR F903 = 'MPHOST' OR F903 = CONCAT('M', '$currentStoreNumber', '901')) 
      AND (DATEDIFF(DAY, F907, GETDATE()) > 30 OR DATEDIFF(DAY, F909, GETDATE()) > 30);",

        # Delete bad SMS items
        "IF OBJECT_ID('OBJ_TAB', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_TAB', 'OBJECT', 'DELETE') = 1
    DELETE FROM OBJ_TAB 
    WHERE F01='0020000000000' 
        OR F01 LIKE '% %' 
        OR LEN(F01)<>13 
        OR (SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000') 
        OR (SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000');",

        "IF OBJECT_ID('POS_TAB', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_TAB', 'OBJECT', 'DELETE') = 1
    DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);",

        "IF OBJECT_ID('PRICE_TAB', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_TAB', 'OBJECT', 'DELETE') = 1
    DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);",

        "IF OBJECT_ID('COST_TAB', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_TAB', 'OBJECT', 'DELETE') = 1
    DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);",

        # SCL_TAB operations
        @"
IF OBJECT_ID('SCL_TAB', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SCL_TAB', 'OBJECT', 'DELETE, UPDATE') = 1
BEGIN
    DELETE FROM SCL_TAB 
    WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB) 
        OR SUBSTRING(F01,1,3) <> '002' 
        OR (SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000');
    
    UPDATE SCL_TAB 
    SET F267 = SCL_TXT.F267, F1001 = 1 
    FROM SCL_TAB SCL 
    JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002', FORMAT(SCL_TXT.F267, '00000'), '00000'));
    
    UPDATE SCL_TAB 
    SET F268 = SCL_NUT.F268, F1001 = 1 
    FROM SCL_TAB SCL 
    JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002', FORMAT(SCL_NUT.F268, '00000'), '00000'));
    
    UPDATE SCL_TAB 
    SET F267 = NULL, F1001 = 1 
    WHERE F01 NOT IN (SELECT CONCAT('002', FORMAT(F267, '00000'), '00000') FROM SCL_TXT_TAB);
    
    UPDATE SCL_TAB 
    SET F268 = NULL, F1001 = 1 
    WHERE F01 NOT IN (SELECT CONCAT('002', FORMAT(F268, '00000'), '00000') FROM SCL_NUT_TAB);
    
    UPDATE SCL_TXT_TAB 
        SET F04 = POS.F04, F1001 = 1 
        FROM SCL_TXT_TAB SCL_TXT 
        JOIN POS_TAB POS ON (POS.F01=CONCAT('002', FORMAT(SCL_TXT.F267, '00000'), '00000')) 
        WHERE ISNUMERIC(SCL_TXT.F04)=0;
    
    UPDATE SCL_TAB 
        SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' '),
            F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' '),
            F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' '),
            F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' ');
END
"@,

        # SCL_TXT_TAB operations
        @"
IF OBJECT_ID('SCL_TXT_TAB', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SCL_TXT_TAB', 'OBJECT', 'DELETE, UPDATE') = 1
BEGIN
    DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB);
    
    UPDATE SCL_TXT_TAB 
        SET F04 = POS.F04, F1001 = 1 
        FROM SCL_TXT_TAB SCL_TXT 
        JOIN POS_TAB POS ON POS.F01=CONCAT('002', FORMAT(SCL_TXT.F267, '00000'), '00000') 
        WHERE ISNUMERIC(SCL_TXT.F04) = 0;
    
    UPDATE SCL_TXT_TAB 
        SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13), ' '), CHAR(10), ' '), CHAR(9), ' ');
END
"@,

        # SCL_NUT_TAB operations
        "IF OBJECT_ID('SCL_NUT_TAB', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SCL_NUT_TAB', 'OBJECT', 'DELETE') = 1
    DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB);"
    )

    # Execute Additional Commands excluding recovery and shrink commands
    foreach ($command in $additionalCommands) {
        if (-not (Execute-SqlCommand -commandText $command)) {
            $failedAdditionalCommands += $command  # Add failed command to the array
        }
    }

    # Final SQL Operations: Recovery and Shrink
    $finalCommands = @(
        "ALTER DATABASE LANESQL SET RECOVERY SIMPLE;",
        "DBCC CHECKDB('LANESQL');",
        "EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD';",
        "EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN';",
        "DBCC SHRINKFILE (LANESQL);",
        "DBCC SHRINKFILE (LANESQL_Log);",
        "ALTER DATABASE LANESQL SET RECOVERY FULL;"
    )

    # Execute Final Commands
    foreach ($command in $finalCommands) {
        if (-not (Execute-SqlCommand -commandText $command)) {
            $failedAdditionalCommands += $command  # Add failed command to the array
        }
    }

    # Return the list of failed commands
    return $failedAdditionalCommands
}

# ===================================================================================================
#                                       FUNCTION: Remove-GTRegistryValues
# ---------------------------------------------------------------------------------------------------
# Description:
#   Removes all registry values starting with 'GT' from specified registry paths.
# ===================================================================================================

function Remove-GTRegistryValues {
    # Define registry paths for 32-bit and 64-bit
    $regPath32 = "HKLM:\SOFTWARE\Store Management\Counters"
    $regPath64 = "HKLM:\SOFTWARE\Wow6432Node\Store Management\Counters"

    # Check system architecture
    $is64bit = [System.Environment]::Is64BitOperatingSystem

    # Initialize total deleted count
    $totalDeletedCount = 0
    $success = $true
    $message = ""
    $status = "Successful"  # Default status

    # Function to delete values starting with GT
    function Delete-GTValuesInPath($path) {
        try {
            # Get all values in the registry path
            $values = Get-Item -Path $path -ErrorAction Stop | Get-ItemProperty

            # Filter values that start with "GT"
            $gtValues = $values.PSObject.Properties | Where-Object { $_.Name -like "GT*" }

            # Count the number of values found
            $valueCount = $gtValues.Count

            # Loop through the GT values and delete them
            foreach ($value in $gtValues) {
                try {
                    Remove-ItemProperty -Path $path -Name $value.Name -ErrorAction Stop
                    $totalDeletedCount++
                } catch {
                    # Handle individual deletion errors
                }
            }

            return $valueCount  # Return the number of values deleted
        } catch {
            $success = $false
            $status = "Failed"
            $message = "Error accessing registry path: $path. Error: $_"
            return 0  # Indicate access failed
        }
    }

    # Check which path to use based on environment
    if ($is64bit -and (Test-Path $regPath64)) {
        Delete-GTValuesInPath -path $regPath64 | Out-Null
    } elseif (Test-Path $regPath32) {
        Delete-GTValuesInPath -path $regPath32 | Out-Null
    } else {
        $success = $false
        $status = "Failed"
        $message = "No valid registry paths found for the current environment."
    }

    # Return an object with success status and deleted count
    return @{
        Success = $success
        Status = $status
        DeletedCount = $totalDeletedCount
        Message = $message
    }
}

# ===================================================================================================
#                                       SECTION: SQL Variables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Sets up variables related to SQL Server connections and database configurations.
# ===================================================================================================

$serverName = Get-SQLInstanceName # Get SQL Server name automatically
$databaseName = "lanesql"  # You can also prompt for this if needed
$connectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True;"
$oldMachineName = $currentMachineName # Set the old machine name variable

# ===================================================================================================
#                                       SECTION: GUI Buttons
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and configures buttons on the GUI form for various operations.
# ===================================================================================================

# Buttons for various operations
$updateStoreNumberButton = New-Object System.Windows.Forms.Button
$updateStoreNumberButton.Text = "Update Store Number"
$updateStoreNumberButton.Location = New-Object System.Drawing.Point(10, 170)
$updateStoreNumberButton.Size = New-Object System.Drawing.Size(150, 35)

$changeMachineNameButton = New-Object System.Windows.Forms.Button
$changeMachineNameButton.Text = "Change Machine Name"
$changeMachineNameButton.Location = New-Object System.Drawing.Point(170, 170)
$changeMachineNameButton.Size = New-Object System.Drawing.Size(150, 35)

$configureNetworkButton = New-Object System.Windows.Forms.Button
$configureNetworkButton.Text = "Configure Network"
$configureNetworkButton.Location = New-Object System.Drawing.Point(330, 170)
$configureNetworkButton.Size = New-Object System.Drawing.Size(150, 35)

$truncateTablesButton = New-Object System.Windows.Forms.Button
$truncateTablesButton.Text = "Truncate Tables"
$truncateTablesButton.Location = New-Object System.Drawing.Point(10, 210)
$truncateTablesButton.Size = New-Object System.Drawing.Size(150, 35)

$repairDatabaseButton = New-Object System.Windows.Forms.Button
$repairDatabaseButton.Text = "Repair Database"
$repairDatabaseButton.Location = New-Object System.Drawing.Point(170, 210)
$repairDatabaseButton.Size = New-Object System.Drawing.Size(150, 35)

$registryCleanupButton = New-Object System.Windows.Forms.Button
$registryCleanupButton.Text = "Registry Cleanup"
$registryCleanupButton.Location = New-Object System.Drawing.Point(330, 210)
$registryCleanupButton.Size = New-Object System.Drawing.Size(150, 35)

$configurePowerButton = New-Object System.Windows.Forms.Button
$configurePowerButton.Text = "Configure Power Settings"
$configurePowerButton.Location = New-Object System.Drawing.Point(10, 250)
$configurePowerButton.Size = New-Object System.Drawing.Size(150, 35)

$configureServicesButton = New-Object System.Windows.Forms.Button
$configureServicesButton.Text = "Configure Services"
$configureServicesButton.Location = New-Object System.Drawing.Point(170, 250)
$configureServicesButton.Size = New-Object System.Drawing.Size(150, 35)

$configureAdvancedButton = New-Object System.Windows.Forms.Button
$configureAdvancedButton.Text = "Configure Advanced Settings"
$configureAdvancedButton.Location = New-Object System.Drawing.Point(330, 250)
$configureAdvancedButton.Size = New-Object System.Drawing.Size(150, 35)

$updateSQLDatabaseButton = New-Object System.Windows.Forms.Button
$updateSQLDatabaseButton.Text = "Update SQL Database"
$updateSQLDatabaseButton.Location = New-Object System.Drawing.Point(10, 290)
$updateSQLDatabaseButton.Size = New-Object System.Drawing.Size(150, 35)

$summaryButton = New-Object System.Windows.Forms.Button
$summaryButton.Text = "Show Summary"
$summaryButton.Location = New-Object System.Drawing.Point(170, 290)
$summaryButton.Size = New-Object System.Drawing.Size(150, 35)

$rebootButton = New-Object System.Windows.Forms.Button
$rebootButton.Text = "Reboot System"
$rebootButton.Location = New-Object System.Drawing.Point(330, 290)
$rebootButton.Size = New-Object System.Drawing.Size(150, 35)

# Add all buttons to the form
$form.Controls.AddRange(@(
    $updateStoreNumberButton, $changeMachineNameButton, $configureNetworkButton,
    $truncateTablesButton, $repairDatabaseButton, $registryCleanupButton,
    $configurePowerButton, $configureServicesButton, $configureAdvancedButton,
    $updateSQLDatabaseButton, $summaryButton, $rebootButton
))

# ===================================================================================================
#                                       SECTION: Event Handlers for Buttons
# ---------------------------------------------------------------------------------------------------
# Description:
#   Defines the actions to be taken when each GUI button is clicked.
# ===================================================================================================

# Event Handler for Update Store Number Button
$updateStoreNumberButton.Add_Click({
    # Get the old store number from startup.ini
    $oldStoreNumber = Get-StoreNumberFromINI

    if ($oldStoreNumber -ne $null) {
        # Prompt for new store number
        $newStoreNumberInput = Get-ValidStoreNumber
        if ($newStoreNumberInput -ne $null) {
            # Show warning before updating
            $warningResult = [System.Windows.Forms.MessageBox]::Show("You are about to change the store number from '$oldStoreNumber' to '$newStoreNumberInput'. Do you want to proceed?", "Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

            if ($warningResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Update startup.ini
                if (Test-Path $startupIniPath) {
                    $updateSuccess = Update-StoreNumberInINI -newStoreNumber $newStoreNumberInput -startupIniPath $startupIniPath
                    if ($updateSuccess) {
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
                        $sqlUpdateResult = Update-SQLTablesForStoreNumberChange -storeNumber $script:newStoreNumber

                        if ($sqlUpdateResult.Success) {
                            $operationStatus["SQLDatabaseUpdate"].Status = "Successful"
                            $operationStatus["SQLDatabaseUpdate"].Message = "STD_TAB updated successfully after store number change."
                            $operationStatus["SQLDatabaseUpdate"].Details = "STD_TAB updated with new store number."
                        } else {
                            $operationStatus["SQLDatabaseUpdate"].Status = "Failed"
                            $operationStatus["SQLDatabaseUpdate"].Message = "Failed to update STD_TAB after store number change."
                            $operationStatus["SQLDatabaseUpdate"].Details = "Failed commands: $($sqlUpdateResult.FailedCommands -join ', ')"
                        }

                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Failed to update startup.ini.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        $operationStatus["StoreNumberChange"].Status = "Failed"
                        $operationStatus["StoreNumberChange"].Message = "Failed to update store number."
                        $operationStatus["StoreNumberChange"].Details = "Error updating startup.ini."
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("startup.ini not found at $startupIniPath.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    $operationStatus["StoreNumberChange"].Status = "Failed"
                    $operationStatus["StoreNumberChange"].Message = "Failed to update store number."
                    $operationStatus["StoreNumberChange"].Details = "startup.ini not found."
                }
            } else {
                $operationStatus["StoreNumberChange"].Status = "Cancelled"
                $operationStatus["StoreNumberChange"].Message = "Store number change was cancelled by the user."
                $operationStatus["StoreNumberChange"].Details = "Old store number remains '$oldStoreNumber'."
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Store number not found in startup.ini.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $operationStatus["StoreNumberChange"].Status = "Failed"
        $operationStatus["StoreNumberChange"].Message = "Store number not found."
        $operationStatus["StoreNumberChange"].Details = "startup.ini not found or store number not defined."
    }
})

# Event Handler for Change Machine Name Button
$changeMachineNameButton.Add_Click({
    # Ensure the script is running with administrative privileges
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        [System.Windows.Forms.MessageBox]::Show("This script must be run as an administrator.", "Insufficient Privileges", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Get the new machine name from the user
    $newMachineNameInput = Get-ValidMachineName

    if ($newMachineNameInput -ne $null) {
        # Confirm the change
        $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to change the machine name to '$newMachineNameInput'?", "Confirm Machine Name Change", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Proceed to change machine name
            try {
                Rename-Computer -NewName $newMachineNameInput -Force -ErrorAction Stop

                # Assign to script-level variable
                $script:newMachineName = $newMachineNameInput

                # Update the machine name label (ensure $machineNameLabel is defined)
                $machineNameLabel.Text = "The machine name will change from: $env:COMPUTERNAME to $script:newMachineName"

                # Update operation status (ensure $operationStatus is initialized)
                $operationStatus["MachineNameChange"].Status = "Successful"
                $operationStatus["MachineNameChange"].Message = "Machine name changed successfully."
                $operationStatus["MachineNameChange"].Details = "Machine name changed to '$script:newMachineName'."

                # Call Remove-OldXFolders (ensure this function and variables are defined)
                Remove-OldXFolders -oldMachineName $oldMachineName -oldStoreNumber $currentStoreNumber

                # Update startup.ini file after changing machine name
                $startupIniPath = "\\localhost\storeman\startup.ini"
                $newDbServerName = $script:newMachineName

                $terValue = "TER=$($newDbServerName.Substring(3))"
                $dbServerValue = "DBSERVER=$($newDbServerName)\$($serverName.Split('\')[1])"  # Ensure $serverName is defined

                if (Test-Path $startupIniPath) {
                    $content = Get-Content $startupIniPath
                    $updatedContent = $content -replace "(?i)TER=\d{3}", $terValue -replace "(?i)DBSERVER=.*", $dbServerValue
                    Set-Content $startupIniPath $updatedContent

                    $operationStatus["StartupIniUpdate"].Status = "Successful"
                    $operationStatus["StartupIniUpdate"].Message = "startup.ini updated successfully."
                    $operationStatus["StartupIniUpdate"].Details = "Updated TER to '$terValue' and DBSERVER to '$dbServerValue'."
                } else {
                    $operationStatus["StartupIniUpdate"].Status = "Failed"
                    $operationStatus["StartupIniUpdate"].Message = "startup.ini file not found."
                    $operationStatus["StartupIniUpdate"].Details = "File not found at $startupIniPath."
                }

                # Call the SQL update function
                # Determine store number and machine number
                $storeNumber = Get-StoreNumberFromINI
                if ($script:newMachineName.Length -ge 6) {
                    $machineNumber = $script:newMachineName.Substring(3,3)
                } else {
                    $machineNumber = ""
                }

                $sqlUpdateResult = Update-SQLTablesForMachineNameChange -storeNumber $storeNumber -machineName $script:newMachineName -machineNumber $machineNumber

                if ($sqlUpdateResult.Success) {
                    $operationStatus["SQLDatabaseUpdate"].Status = "Successful"
                    $operationStatus["SQLDatabaseUpdate"].Message = "SQL tables updated successfully after machine name change."
                    $operationStatus["SQLDatabaseUpdate"].Details = "STO_TAB, TER_TAB, LNK_TAB, and RUN_TAB updated."
                } else {
                    $operationStatus["SQLDatabaseUpdate"].Status = "Failed"
                    $operationStatus["SQLDatabaseUpdate"].Message = "Failed to update SQL tables after machine name change."
                    $operationStatus["SQLDatabaseUpdate"].Details = "Failed commands: $($sqlUpdateResult.FailedCommands -join ', ')"
                }

                # Inform the user about the reboot
                $rebootResult = [System.Windows.Forms.MessageBox]::Show("Machine name changed successfully to '$script:newMachineName'. The system will need to reboot for changes to take effect. Do you want to reboot now?", "Success", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)

                if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Restart-Computer -Force
                }

            } catch {
                $errorMessage = $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show("Error changing machine name: $errorMessage", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $operationStatus["MachineNameChange"].Status = "Failed"
                $operationStatus["MachineNameChange"].Message = "Error changing machine name."
                $operationStatus["MachineNameChange"].Details = "Error: $errorMessage"
            }
        }
    } else {
        # Handle cancellation
        $operationStatus["MachineNameChange"].Status = "Cancelled"
        $operationStatus["MachineNameChange"].Message = "Machine name change was cancelled by the user."
        [System.Windows.Forms.MessageBox]::Show("Machine name change was cancelled.", "Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Event Handler for Configure Network Button
$configureNetworkButton.Add_Click({
    # Implement the ConfigureNetwork function with GUI elements

    if ($currentConfigs -eq $null -or $currentConfigs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No valid active IP configuration found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # If multiple adapters are found, ask the user to select one
    if ($currentConfigs.Count -gt 1) {
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

        for ($i = 0; $i -lt $currentConfigs.Count; $i++) {
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
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $currentConfigs.Count) {
            $currentConfig = $currentConfigs[$selectedIndex]
        } else {
            [System.Windows.Forms.MessageBox]::Show("No network adapter selected.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    } else {
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

    if ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Configure DHCP
        try {
            netsh interface ip set address name="$adapterName" source=dhcp
            netsh interface ip set dns name="$adapterName" source=dhcp
            [System.Windows.Forms.MessageBox]::Show("DHCP configuration applied.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $operationStatus["IPConfiguration"].Status = "Successful"
            $operationStatus["IPConfiguration"].Message = "DHCP configuration applied."
            $operationStatus["IPConfiguration"].Details = "Adapter: $adapterName"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to configure DHCP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $operationStatus["IPConfiguration"].Status = "Failed"
            $operationStatus["IPConfiguration"].Message = "Failed to configure DHCP."
            $operationStatus["IPConfiguration"].Details = "Error: $_"
        }

    } elseif ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::No) {
        # Configure Static IP
        # Loop until valid input or user cancels
        $validInput = $false
        while (-not $validInput) {
            # Ask for last octet
            $gatewayParts = $currentGateway.Split('.')
            if ($gatewayParts.Length -lt 4) {
                [System.Windows.Forms.MessageBox]::Show("Invalid gateway format: $currentGateway", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $operationStatus["IPConfiguration"].Status = "Failed"
                $operationStatus["IPConfiguration"].Message = "Invalid gateway format."
                $operationStatus["IPConfiguration"].Details = "Gateway: $currentGateway"
                return
            }
            $gatewayPrefix = $gatewayParts[0..2] -join '.'
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

            if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
                $lastOctet = $textBox.Text.Trim()
                # Validate last octet
                if ($lastOctet -match '^\d{1,3}$' -and [int]$lastOctet -ge 1 -and [int]$lastOctet -le 254) {
                    $ipAddress = "$gatewayPrefix.$lastOctet"

                    # Check if IP is in use
                    $pingResult = Test-Connection -ComputerName $ipAddress -Count 1 -Quiet
                    if ($pingResult) {
                        [System.Windows.Forms.MessageBox]::Show("The IP address '$ipAddress' is already in use. Please choose a different one.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    } else {
                        try {
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
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show("Failed to configure static IP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            $operationStatus["IPConfiguration"].Status = "Failed"
                            $operationStatus["IPConfiguration"].Message = "Failed to configure static IP."
                            $operationStatus["IPConfiguration"].Details = "Error: $_"
                            $validInput = $true  # Exit loop since an error occurred
                        }
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Invalid input. Please enter a valid number between 1 and 254.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            } else {
                # User canceled the dialog
                $operationStatus["IPConfiguration"].Status = "Skipped"
                $operationStatus["IPConfiguration"].Message = "User canceled static IP configuration."
                $operationStatus["IPConfiguration"].Details = ""  # Empty details
                $validInput = $true  # Exit loop
            }
        }

    } else {
        $operationStatus["IPConfiguration"].Status = "Skipped"
        $operationStatus["IPConfiguration"].Message = "User chose not to configure network settings."
    }
})

# Event Handler for Truncate Tables Button
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
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Truncate tables
        $failedTruncateTables = Truncate-Tables -tables $tablesToTruncate

        if ($failedTruncateTables.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Tables truncated successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $operationStatus["TableTruncation"].Status = "Successful"
            $operationStatus["TableTruncation"].Message = "Tables truncated successfully."
            $operationStatus["TableTruncation"].Details = "All tables were truncated successfully."
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to truncate some tables: $($failedTruncateTables -join ', ')", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $operationStatus["TableTruncation"].Status = "Failed"
            $operationStatus["TableTruncation"].Message = "Failed to truncate some tables."
            $operationStatus["TableTruncation"].Details = "Failed tables: $($failedTruncateTables -join ', ')"
        }
    } else {
        $operationStatus["TableTruncation"].Status = "Skipped"
        $operationStatus["TableTruncation"].Message = "User chose not to truncate tables."
    }
})

# Event Handler for Repair Database Button
$repairDatabaseButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Do you want to repair the database?", "Repair Database", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $failedRepairCommands = Repair-Database
        if ($failedRepairCommands.Count -eq 0) {
            $operationStatus["DatabaseRepair"].Status = "Successful"
            $operationStatus["DatabaseRepair"].Message = "Database repaired successfully."
            $operationStatus["DatabaseRepair"].Details = "All repair operations ran successfully."
            [System.Windows.Forms.MessageBox]::Show("Database repaired successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            $operationStatus["DatabaseRepair"].Status = "Failed"
            $operationStatus["DatabaseRepair"].Message = "Failed to execute some repair operations."
            $operationStatus["DatabaseRepair"].Details = "Failed operations: $($failedRepairCommands -join ', ')"
            [System.Windows.Forms.MessageBox]::Show("Failed to execute some SQL commands.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        $operationStatus["DatabaseRepair"].Status = "Skipped"
        $operationStatus["DatabaseRepair"].Message = "User chose not to repair the database."
    }
})

# Event Handler for Registry Cleanup Button
$registryCleanupButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Do you want to delete all registry values starting with 'GT'?", "Registry Cleanup", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $gtRegistryCleanupResult = Remove-GTRegistryValues

        $registryStatus = $gtRegistryCleanupResult.Status

        if ($registryStatus -eq 'Successful') {
            $operationStatus["RegistryCleanup"].Status = "Successful"
            $operationStatus["RegistryCleanup"].Message = "GT registry values removed successfully."
            $operationStatus["RegistryCleanup"].Details = "$($gtRegistryCleanupResult.DeletedCount) 'GT' registry keys were deleted."
            [System.Windows.Forms.MessageBox]::Show("GT registry values removed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } elseif ($registryStatus -eq 'Skipped') {
            $operationStatus["RegistryCleanup"].Status = "Skipped"
            $operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
            $operationStatus["RegistryCleanup"].Details = ""
        } else {
            $operationStatus["RegistryCleanup"].Status = "Failed"
            $operationStatus["RegistryCleanup"].Message = "Failed to remove GT registry values."
            $operationStatus["RegistryCleanup"].Details = $gtRegistryCleanupResult.Message
            [System.Windows.Forms.MessageBox]::Show("Failed to remove GT registry values.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        $operationStatus["RegistryCleanup"].Status = "Skipped"
        $operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
    }
})

# Event Handler for Configure Power Settings Button
$configurePowerButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Do you want to configure the power settings?", "Configure Power Settings", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Configure-PowerSettings
    } else {
        $operationStatus["ConfigurePowerSettings"].Status = "Skipped"
        $operationStatus["ConfigurePowerSettings"].Message = "User chose not to configure power settings."
        $operationStatus["ConfigurePowerSettings"].Details = ""
    }
})

# Event Handler for Configure Services Button
$configureServicesButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Do you want to configure the services?", "Configure Services", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Configure-Services
    } else {
        $operationStatus["ConfigureServices"].Status = "Skipped"
        $operationStatus["ConfigureServices"].Message = "User chose not to configure services."
        $operationStatus["ConfigureServices"].Details = ""
    }
})

# Event Handler for Configure Advanced Settings Button
$configureAdvancedButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Do you want to configure the advanced system settings?", "Configure Advanced Settings", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Configure-AdvancedSettings
    } else {
        $operationStatus["ConfigureAdvancedSettings"].Status = "Skipped"
        $operationStatus["ConfigureAdvancedSettings"].Message = "User chose not to configure advanced system settings."
        $operationStatus["ConfigureAdvancedSettings"].Details = ""
    }
})

# Event Handler for Update SQL Database Button
$updateSQLDatabaseButton.Add_Click({
    # Read the store number directly from startup.ini
    $storeNumberFromINI = Get-StoreNumberFromINI

    if ($storeNumberFromINI -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Store number not found in startup.ini.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $operationStatus["SQLDatabaseUpdate"].Status = "Failed"
        $operationStatus["SQLDatabaseUpdate"].Message = "Store number not found in startup.ini."
        $operationStatus["SQLDatabaseUpdate"].Details = ""
        return
    }

    $storeNumber = $storeNumberFromINI

    # Determine the machine name to use
    if (-not [string]::IsNullOrEmpty($script:newMachineName)) {
        $machineName = $script:newMachineName
    } else {
        $machineName = $currentMachineName
    }

    # Extract the machine number from machine name
    if ($machineName.Length -ge 6) {
        $machineNumber = $machineName.Substring(3,3)
    } else {
        [System.Windows.Forms.MessageBox]::Show("Machine name '$machineName' is invalid. Cannot extract machine number.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $operationStatus["SQLDatabaseUpdate"].Status = "Failed"
        $operationStatus["SQLDatabaseUpdate"].Message = "Invalid machine name."
        $operationStatus["SQLDatabaseUpdate"].Details = "Cannot extract machine number."
        return
    }

    # Proceed with SQL update code, using $storeNumber, $machineName, and $machineNumber
    # Variables
    $terTableName = "TER_TAB"
    $runTableName = "RUN_TAB"
    $stoTableName = "STO_TAB"
    $lnkTableName = "LNK_TAB"
    $stdTableName = "STD_TAB"

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

    # STD_TAB commands
    $createViewCommandStd = @"
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
"@

    $updateStdTabCommand = @"
UPDATE $stdTableName 
SET F1056 = '$storeNumber';
"@

    $dropViewCommandStd = "DROP VIEW Std_Load;"

    # Now execute the SQL commands
    $allSqlSuccessful = $true
    $failedSqlCommands = @()

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
        $dropViewCommandLnk,

        # STD_TAB commands
        $createViewCommandStd,
        $updateStdTabCommand,
        $dropViewCommandStd
    )

    foreach ($command in $sqlCommands) {
        if (-not (Execute-SqlCommand -commandText $command)) {
            $allSqlSuccessful = $false
            $failedSqlCommands += $command
        }
    }

    if ($allSqlSuccessful) {
        $operationStatus["SQLDatabaseUpdate"].Status = "Successful"
        $operationStatus["SQLDatabaseUpdate"].Message = "SQL database updated successfully."
        $operationStatus["SQLDatabaseUpdate"].Details = "All SQL commands executed successfully."
        [System.Windows.Forms.MessageBox]::Show("SQL database updated successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        $operationStatus["SQLDatabaseUpdate"].Status = "Failed"
        $operationStatus["SQLDatabaseUpdate"].Message = "Failed to execute some SQL commands."
        $operationStatus["SQLDatabaseUpdate"].Details = "Failed SQL commands are listed below."
        [System.Windows.Forms.MessageBox]::Show("Failed to execute some SQL commands.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

        # Optionally, display failed SQL commands in a new form
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
        $failedCommandsForm.ShowDialog()
    }
})

# Event Handler for Show Summary Button
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
    foreach ($operationKey in $operationStatus.Keys) {
        $operation = $operationStatus[$operationKey]
        $status = $operation.Status
        $message = $operation.Message
        $details = $operation.Details

        $summaryText += "${operationKey}: $status`r`n"
        if ($message -ne "") {
            $summaryText += "  Message: $message`r`n"
        }
        if ($details -ne "") {
            $summaryText += "  Details: $details`r`n"
        }
        $summaryText += "`r`n"
    }

    $textBox.Text = $summaryText
    $summaryForm.Controls.Add($textBox)
    $summaryForm.ShowDialog()
})

# Event Handler for Reboot Button
$rebootButton.Add_Click({
    $rebootResult = [System.Windows.Forms.MessageBox]::Show("Do you want to reboot now?", "Reboot", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Restart-Computer -Force
    }
})

# Add a handler for when the user clicks the 'X' to close the form
$form.add_FormClosing({
    param($sender, $e)
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to exit?", 
        "Confirm Exit", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        $e.Cancel = $true  # Cancel the closing event
    }
    # Else, allow the form to close naturally
})

# ===================================================================================================
#                                       SECTION: Show the Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays the main GUI form to the user.
# ===================================================================================================

$form.ShowDialog()

# Explicitly exit the script after the form is closed
exit
