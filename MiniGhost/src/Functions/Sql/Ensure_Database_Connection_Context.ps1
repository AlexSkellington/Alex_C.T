# ===================================================================================================
#                          FUNCTION: Ensure_Database_Connection_Context
# ---------------------------------------------------------------------------------------------------
# Description:
# Lazily refreshes DBSERVER / DBNAME / ConnectionString from Startup.ini so SQL-specific work can
# initialize only when needed instead of front-loading that work during GUI startup.
# ===================================================================================================

function Ensure_Database_Connection_Context
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	if (-not $script:FunctionResults)
	{
		$script:FunctionResults = @{ }
	}
	
	$connectionStringValue = $null
	if ($script:FunctionResults.ContainsKey('ConnectionString') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConnectionString']))
	{
		$connectionStringValue = [string]$script:FunctionResults['ConnectionString']
	}
	
	$dbNameValue = $null
	if ($script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBNAME']))
	{
		$dbNameValue = [string]$script:FunctionResults['DBNAME']
	}
	
	$dbServerValue = $null
	if ($script:FunctionResults.ContainsKey('DBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBSERVER']))
	{
		$dbServerValue = [string]$script:FunctionResults['DBSERVER']
	}
	
	if ([string]::IsNullOrWhiteSpace($connectionStringValue) -or [string]::IsNullOrWhiteSpace($dbNameValue) -or [string]::IsNullOrWhiteSpace($dbServerValue))
	{
		Get_Database_Connection_String -StartupIniPath $StartupIniPath
		
		if ($script:FunctionResults.ContainsKey('ConnectionString') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConnectionString']))
		{
			$connectionStringValue = [string]$script:FunctionResults['ConnectionString']
		}
	}
	
	$script:connectionString = $connectionStringValue
	return $connectionStringValue
}
