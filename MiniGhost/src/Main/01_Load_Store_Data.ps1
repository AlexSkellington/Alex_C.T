# ===================================================================================================
#                                       SECTION: Main Script Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Orchestrates the execution flow of the script, initializing variables, processing items, and handling user interactions.
# ===================================================================================================

# Get current IP configuration
$currentConfigs = Get_Active_IP_Config

# Get the store number
$currentStoreNumber = Get_Store_Number_From_INI -UpdateLabel
if (-not $currentStoreNumber) { $currentStoreNumber = $script:FunctionResults['StoreNumber'] }

# Get the store name
Get_Store_Name_From_INI
$storeName = $script:FunctionResults['StoreName']

# Detect the machine role from DBNAME before showing the UI.
$dbNameForRole = $null
try
{
	if (Test-Path $startupIniPath)
	{
		$dbNameLine = Get-Content -LiteralPath $startupIniPath -ErrorAction Stop | Where-Object { $_ -match '^\s*(?i:DBNAME)\s*=' } | Select-Object -First 1
		if ($dbNameLine -and ($dbNameLine -match '^\s*(?i:DBNAME)\s*=\s*(.+?)\s*$'))
		{
			$dbNameForRole = $matches[1].Trim().ToUpper()
			$script:FunctionResults['DBNAME'] = $dbNameForRole
		}
	}
}
catch { }

$script:NodeRole = "Unknown"
if ($dbNameForRole -match '^(?i)LANESQL$')
{
	$script:NodeRole = "Lane"
}
elseif ($dbNameForRole -match '^(?i)STORESQL$')
{
	$script:NodeRole = "StoreServer"
}
elseif ($dbNameForRole -match '^(?i)HOSTSQL$')
{
	$script:NodeRole = "HostServer"
}

# Database connection details remain lazy; DBNAME is read here only for node-role detection.
$connectionString = $null

# Set the old machine name variable
$oldMachineName = $currentMachineName

# Clear %Temp% foder on start
# $FilesAndDirsDeleted = Delete_Files_And_Folders -Path "$TempDir" -Exclusions "MiniGhost.ps1" -AsJob

# Indicate the script has started
Write-Host "Script started" -ForegroundColor Green
