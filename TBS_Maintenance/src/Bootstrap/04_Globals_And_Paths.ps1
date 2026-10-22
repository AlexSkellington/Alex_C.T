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
$script:AutoExportVncOnLaunch = $true
$script:AutoExportRemoteMachineInfoOnLaunch = $true
$script:AutoQuarterlySystemHealthMaintenanceOnLaunch = $true
$script:VncExportTask = $null
$script:VncExportReaper = $null
$script:RemoteMachineInfoTask = $null
$script:RemoteMachineInfoReaper = $null
$script:LaneDbSizeMap = @{ }
$script:LaneDbSizeJobs = @{ }
$script:LaneDbSizePool = $null
$script:LaneDbSizeRefreshSeconds = 5
$script:ProtocolPopupMode = 'Lane'
$script:StartupWarmupStarted = $false
$script:StartupWarmupTimer = $null
$script:StartupWarmupSteps = @()
$script:StartupWarmupIndex = 0
$script:StartupWarmupStatus = $null
$script:StartupScaleCredStarted = $false
$script:StartupScaleCredTimer = $null
$script:PopupThemeTimer = $null
$script:StyledPopupForms = @{ }
$script:LaneSqlProtocolEnableTaskState = @{ }
$script:LaneSqlProtocolEnableJobs = @{ }

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

# === Central Path Configuration ===
$script:AlexRoot = "C:\Tecnica_Systems\Alex_C.T"
$script:AlexProgramDataRoot = Join-Path $env:ProgramData "Tecnica_Systems\Alex_C.T"
$script:StoremanRoots = @("C:\Storeman", "D:\Storeman")
$script:StoremanShareRoot = "\\localhost\storeman"
$script:PrimaryStoremanRoot = $script:StoremanRoots[0]
$script:StoremanOfficeRoots = @($script:StoremanRoots | ForEach-Object { [System.IO.Path]::Combine($_, "Office") })
$script:PrimaryStoremanOfficePath = $script:StoremanOfficeRoots[0]
$script:StoremanOfficeDbsFallbacks = @($script:StoremanOfficeRoots | ForEach-Object { [System.IO.Path]::Combine($_, "Dbs") }) + @([System.IO.Path]::Combine($script:StoremanShareRoot, "office\dbs"))
$script:LocalStoremanOfficePath = [System.IO.Path]::Combine($script:StoremanShareRoot, "office")

$script:BackupRoot = ((Join-Path $script:AlexRoot "Backups").Trim().TrimEnd('\') + '\')
$script:ScriptsFolder = ((Join-Path $script:AlexRoot "Scripts").Trim().TrimEnd('\') + '\')
$script:ScriptsLog = ((Join-Path $script:AlexRoot "Log").Trim().TrimEnd('\') + '\')
$script:SetupFiles = Join-Path $script:AlexRoot "Setup_Files"
$script:ToolsDir = Join-Path $script:AlexRoot "Tools"
$script:SystemHealthMaintenanceStatePath = Join-Path $script:SetupFiles "Quarterly_System_Health_Maintenance_State.json"
$script:SystemHealthMaintenanceLockPath = Join-Path $script:SetupFiles "Quarterly_System_Health_Maintenance.lock"
$script:SystemHealthMaintenanceWorkerPath = Join-Path $script:ScriptsFolder "Quarterly_System_Health_Maintenance_Worker.ps1"
$script:SystemHealthMaintenanceSummaryLogPath = Join-Path $script:ScriptsLog "Quarterly_System_Health_Maintenance_SUMMARY.txt"
$script:SystemHealthMaintenanceRunLogPrefix = "Quarterly_System_Health_Maintenance_"
$script:SystemHealthMaintenanceRunLogPattern = "Quarterly_System_Health_Maintenance_*.log"
$script:SystemHealthMaintenanceIntervalMonths = 3
$script:SystemHealthMaintenanceRetentionCount = 8
$script:SystemHealthMaintenanceSummaryRetentionLines = 250
$script:ProtocolResultsFile = Join-Path $script:SetupFiles "Protocol_Results.txt"
$script:NssmDefaultPath = Join-Path $script:ToolsDir "NSSM\nssm.exe"

$script:ScaleCommRoots = @("C:\ScaleCommApp", "D:\ScaleCommApp")
$script:PrimaryScaleCommRoot = $script:ScaleCommRoots[0]
$script:ScaleManagementAppExeName = "ScaleManagementApp.exe"
$script:ScaleManagementAppFastDeployExeName = "ScaleManagementApp_FastDEPLOY.exe"
$script:ScaleManagementAppUpdateSpecialsExeName = "ScaleManagementAppUpdateSpecials.exe"
$script:ScaleManagementAppExePath = Join-Path $script:PrimaryScaleCommRoot $script:ScaleManagementAppExeName
$script:ScaleManagementAppFastDeployPath = Join-Path $script:PrimaryScaleCommRoot $script:ScaleManagementAppFastDeployExeName
$script:ScaleManagementAppUpdateSpecialsPath = Join-Path $script:PrimaryScaleCommRoot $script:ScaleManagementAppUpdateSpecialsExeName
$script:ScaleManagementAppPathCandidates = @($script:ScaleCommRoots | ForEach-Object { [System.IO.Path]::Combine($_, $script:ScaleManagementAppExeName) })
$script:ScaleManagementAppFastDeployPathCandidates = @($script:ScaleCommRoots | ForEach-Object { [System.IO.Path]::Combine($_, $script:ScaleManagementAppFastDeployExeName) })
$script:ScaleManagementAppUpdateSpecialsPathCandidates = @($script:ScaleCommRoots | ForEach-Object { [System.IO.Path]::Combine($_, $script:ScaleManagementAppUpdateSpecialsExeName) })

$scaleCommDeployExecPatterns = @()
foreach ($root in @($script:ScaleCommRoots))
{
	if ([string]::IsNullOrWhiteSpace($root)) { continue }
	$scaleCommDeployExecPatterns += [regex]::Escape([System.IO.Path]::Combine($root, $script:ScaleManagementAppExeName))
	$scaleCommDeployExecPatterns += [regex]::Escape([System.IO.Path]::Combine($root, $script:ScaleManagementAppFastDeployExeName))
}
$scaleCommDeployExecPatterns = @($scaleCommDeployExecPatterns | Select-Object -Unique)
if ($scaleCommDeployExecPatterns.Count -eq 0)
{
	$script:ScaleCommDeployExecRegex = "(?i)@EXEC\(RUN='[A-Z]:\\ScaleCommApp\\ScaleManagementApp(_FastDEPLOY)?\.exe'\);"
}
else
{
	$script:ScaleCommDeployExecRegex = "(?i)@EXEC\(RUN='(?:" + ($scaleCommDeployExecPatterns -join '|') + ")'\);"
}

$script:BmsRootPath = "C:\Bizerba\RetailConnect\BMS"
$script:BmssrvExePath = Join-Path $script:BmsRootPath "BMSSrv.exe"
$script:BmsExePath = Join-Path $script:BmsRootPath "BMS.exe"
$script:BmsToBizerbaPath = Join-Path $script:BmsRootPath "toBizerba"
$script:BmsTerminalsPath = Join-Path $script:BmsRootPath "terminals"
$script:BmsScalesDeploymentRoot = Join-Path $script:BmsRootPath "Scales_Deployment"
$script:BmsScalesDeploymentInputPath = Join-Path $script:BmsScalesDeploymentRoot "toBizerba"
$script:DuplicateFileMonitorServiceRoot = Join-Path $script:AlexProgramDataRoot "Duplicate_File_Monitor_Service"
$script:BizStoreCardGenericInputPath = "C:\bizstorecard\bizerba\_fileIO\generic_data\in"
$script:UltraVncFolder = "C:\Users\Administrator\Documents\UltraVNC"

foreach ($corePath in @($script:BackupRoot, $script:ScriptsFolder, $script:ScriptsLog, $script:SetupFiles, $script:ToolsDir))
{
	if (-not (Test-Path -LiteralPath $corePath))
	{
		New-Item -Path $corePath -ItemType Directory -Force | Out-Null
	}
}

# === SQL Backup/Automation Credentials ===
$script:BackupSqlUser = "Tecnica"
$script:BackupSqlPass = "TB`$upp0rT"
