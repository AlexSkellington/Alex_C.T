# ===================================================================================================
#                                SECTION: Assemblies and Global State
# ---------------------------------------------------------------------------------------------------
# Description:
#   Loads WinForms/Drawing and initialises every script-wide variable: result/state bags,
#   launch-time feature switches, encodings, scale credentials and the central path layout
#   (Tecnica_Systems folders, Storeman/ScaleCommApp/BMS roots). Creates the Tecnica_Systems
#   working folders if they are missing.
#
# Notes:
#   - Paths that other code derives from live here only; change a root once and everything
#     that uses it follows.
#   - tests\Smoke.ps1 patches the AlexRoot / AlexProgramDataRoot / StoremanRoots lines and the
#     three Auto*OnLaunch switches by exact text - keep those lines as they are.
# ===================================================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------------------------------
# Result bag + background-work state (filled by the functions and the GUI timers)
# ---------------------------------------------------------------------------------------------------
$script:FunctionResults = @{ }
$script:ProcessedLanes = @()
$script:LaneProtocols = @{ }
$script:LaneProtocolJobs = @{ }
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

# Launch-time background tasks (each one is a warmup step in Main\02_Startup_Warmup)
$script:AutoExportVncOnLaunch = $true
$script:AutoExportRemoteMachineInfoOnLaunch = $true
$script:AutoServerMaintenanceTasksOnLaunch = $true

# Node counts (set by Retrieve_Nodes)
$NumberOfLanes = 0
$NumberOfServers = 0
$NumberOfScales = 0
$NumberOfBackoffices = 0

# ---------------------------------------------------------------------------------------------------
# Encodings
# ---------------------------------------------------------------------------------------------------
$script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)  # Windows-1252 legacy store files
$script:utf8NoBOM = New-Object System.Text.UTF8Encoding($false)     # UTF-8 without BOM (generated files)

# ---------------------------------------------------------------------------------------------------
# Bizerba scale credentials (the two factory passwords; Add_Scale_Credentials tries both)
# ---------------------------------------------------------------------------------------------------
$script:credBizerba = New-Object System.Management.Automation.PSCredential ('bizuser', (ConvertTo-SecureString 'bizerba' -AsPlainText -Force))
$script:credBiyerba = New-Object System.Management.Automation.PSCredential ('bizuser', (ConvertTo-SecureString 'biyerba' -AsPlainText -Force))

# ---------------------------------------------------------------------------------------------------
# Tecnica_Systems working folders
# ---------------------------------------------------------------------------------------------------
$script:AlexRoot = "C:\Tecnica_Systems\Alex_C.T"
$script:AlexProgramDataRoot = Join-Path $env:ProgramData "Tecnica_Systems\Alex_C.T"

$script:BackupRoot = ((Join-Path $script:AlexRoot "Backups").Trim().TrimEnd('\') + '\')
$script:ScriptsFolder = ((Join-Path $script:AlexRoot "Scripts").Trim().TrimEnd('\') + '\')
$script:ScriptsLog = ((Join-Path $script:AlexRoot "Log").Trim().TrimEnd('\') + '\')
$script:SetupFiles = Join-Path $script:AlexRoot "Setup_Files"
$script:ToolsDir = Join-Path $script:AlexRoot "Tools"
$script:ProtocolResultsFile = Join-Path $script:SetupFiles "Protocol_Results.txt"
$script:NssmDefaultPath = Join-Path $script:ToolsDir "NSSM\nssm.exe"
$script:DuplicateFileMonitorServiceRoot = Join-Path $script:AlexProgramDataRoot "Duplicate_File_Monitor_Service"

# Server maintenance tasks (Install_Server_Maintenance_Tasks, Repair_Windows)
$script:ServerMaintenanceWeeklyTime = '02:30'     # every Sunday: drive optimize
$script:ServerMaintenanceMonthlyTime = '03:00'    # first Sunday of the month: DISM, SFC, check disk

foreach ($corePath in @($script:BackupRoot, $script:ScriptsFolder, $script:ScriptsLog, $script:SetupFiles, $script:ToolsDir))
{
	if (-not (Test-Path -LiteralPath $corePath))
	{
		$null = New-Item -Path $corePath -ItemType Directory -Force
	}
}

# ---------------------------------------------------------------------------------------------------
# Storeman roots (the real one is resolved in 05_Storeman_Detection)
# ---------------------------------------------------------------------------------------------------
$script:StoremanRoots = @("C:\Storeman", "D:\Storeman")
$script:PrimaryStoremanRoot = $script:StoremanRoots[0]
$script:StoremanOfficeRoots = @($script:StoremanRoots | ForEach-Object { [System.IO.Path]::Combine($_, "Office") })
$script:PrimaryStoremanOfficePath = $script:StoremanOfficeRoots[0]
$script:LocalStoremanOfficePath = "\\localhost\storeman\office"
$script:StoremanOfficeDbsFallbacks = @($script:StoremanOfficeRoots | ForEach-Object { [System.IO.Path]::Combine($_, "Dbs") }) + @("\\localhost\storeman\office\dbs")

# ---------------------------------------------------------------------------------------------------
# ScaleCommApp (Ishida / ScaleManagementApp)
# ---------------------------------------------------------------------------------------------------
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

# Matches the deploy line in DEPLOY_CHG.sql, e.g. @EXEC(RUN='C:\ScaleCommApp\ScaleManagementApp.exe');
$scaleCommDeployExecPatterns = @(foreach ($root in $script:ScaleCommRoots)
	{
		[regex]::Escape([System.IO.Path]::Combine($root, $script:ScaleManagementAppExeName))
		[regex]::Escape([System.IO.Path]::Combine($root, $script:ScaleManagementAppFastDeployExeName))
	}) | Select-Object -Unique
$script:ScaleCommDeployExecRegex = "(?i)@EXEC\(RUN='(?:" + (@($scaleCommDeployExecPatterns) -join '|') + ")'\);"

# ---------------------------------------------------------------------------------------------------
# Bizerba BMS / RetailConnect, BizStoreCard, UltraVNC
# ---------------------------------------------------------------------------------------------------
$script:BmsRootPath = "C:\Bizerba\RetailConnect\BMS"
$script:BmssrvExePath = Join-Path $script:BmsRootPath "BMSSrv.exe"
$script:BmsExePath = Join-Path $script:BmsRootPath "BMS.exe"
$script:BmsToBizerbaPath = Join-Path $script:BmsRootPath "toBizerba"
$script:BmsTerminalsPath = Join-Path $script:BmsRootPath "terminals"
$script:BizStoreCardGenericInputPath = "C:\bizstorecard\bizerba\_fileIO\generic_data\in"
$script:UltraVncFolder = "C:\Users\Administrator\Documents\UltraVNC"

# ---------------------------------------------------------------------------------------------------
# SQL login used by the scheduled backup / automation tasks
# ---------------------------------------------------------------------------------------------------
$script:BackupSqlUser = "Tecnica"
$script:BackupSqlPass = "TB`$upp0rT"
