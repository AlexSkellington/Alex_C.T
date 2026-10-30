# ===================================================================================================
#                               FUNCTION: Get_Database_Connection_String
# ---------------------------------------------------------------------------------------------------
# Description:
#   Searches for the Startup.ini file in specified locations, extracts the DBNAME value,
#   constructs the connection string, and stores it in a script-level hashtable.
# ===================================================================================================

function Get_Database_Connection_String
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
	
	$effectiveStartupIniPath = $StartupIniPath
	if ([string]::IsNullOrWhiteSpace($effectiveStartupIniPath))
	{
		$effectiveStartupIniPath = $startupIniPath
	}
	
	if (-not $effectiveStartupIniPath)
	{
		return
	}
	
	# Read the Startup.ini file
	try
	{
		$content = Get-Content -Path $effectiveStartupIniPath -ErrorAction Stop
		
		# Extract DBSERVER
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }
		if ($dbServerLine)
		{
			$dbServer = $dbServerLine -replace '^DBSERVER=', ''
			$dbServer = $dbServer.Trim()
			if (-not $dbServer)
			{
				$dbServer = "localhost"
			}
		}
		else
		{
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
				return
			}
		}
		else
		{
			return
		}
	}
	catch
	{
		return
	}
	
	$configuredDbServer = $dbServer
	$resolvedRuntimeDbServer = $configuredDbServer
	$connectionString = "Server=$configuredDbServer;Database=$dbName;Integrated Security=True;Application Name=MiniGhost;"
	$candidateServers = New-Object System.Collections.Generic.List[string]
	$addCandidate = {
		param (
			[string]$Candidate
		)
		
		if ([string]::IsNullOrWhiteSpace($Candidate))
		{
			return
		}
		
		$normalizedCandidate = $Candidate.Trim()
		if ([string]::IsNullOrWhiteSpace($normalizedCandidate))
		{
			return
		}
		
		foreach ($existingCandidate in $candidateServers)
		{
			if ([string]::Equals($existingCandidate, $normalizedCandidate, [System.StringComparison]::OrdinalIgnoreCase))
			{
				return
			}
		}
		
		[void]$candidateServers.Add($normalizedCandidate)
	}
	
	$configuredHost = $configuredDbServer
	$instanceName = $null
	if ($configuredDbServer -match '\\')
	{
		$configuredParts = $configuredDbServer.Split('\', 2)
		$configuredHost = $configuredParts[0].Trim()
		if ($configuredParts.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($configuredParts[1]))
		{
			$instanceName = $configuredParts[1].Trim()
		}
	}
	
	$localAliases = @('localhost', '.', '(local)')
	if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME))
	{
		$localAliases += $env:COMPUTERNAME
	}
	
	$isConfiguredLocal = $false
	foreach ($alias in $localAliases)
	{
		if ([string]::Equals($configuredHost, $alias, [System.StringComparison]::OrdinalIgnoreCase))
		{
			$isConfiguredLocal = $true
			break
		}
	}
	
	if ($isConfiguredLocal)
	{
		& $addCandidate $configuredDbServer
	}
	
	if ($instanceName)
	{
		foreach ($alias in $localAliases)
		{
			& $addCandidate ("{0}\{1}" -f $alias, $instanceName)
		}
	}
	else
	{
		foreach ($alias in $localAliases)
		{
			& $addCandidate $alias
		}
		
		foreach ($registryPath in @(
				'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
				'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'
			))
		{
			if (-not (Test-Path $registryPath)) { continue }
			
			try
			{
				$itemProps = Get-ItemProperty -Path $registryPath -ErrorAction Stop
				foreach ($prop in $itemProps.PSObject.Properties)
				{
					if ($prop.Name -in @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider'))
					{
						continue
					}
					
					if ([string]::IsNullOrWhiteSpace($prop.Name))
					{
						continue
					}
					
					$instanceCandidate = $prop.Name.Trim()
					foreach ($alias in $localAliases)
					{
						& $addCandidate ("{0}\{1}" -f $alias, $instanceCandidate)
					}
				}
			}
			catch { }
		}
	}
	
	& $addCandidate $configuredDbServer
	
	foreach ($candidateServer in $candidateServers)
	{
		$testConnection = $null
		try
		{
			$testConnectionString = "Server=$candidateServer;Database=$dbName;Integrated Security=True;Connect Timeout=3;Application Name=MiniGhost;"
			$testConnection = New-Object System.Data.SqlClient.SqlConnection($testConnectionString)
			$testConnection.Open()
			
			$resolvedRuntimeDbServer = $candidateServer
			$connectionString = "Server=$candidateServer;Database=$dbName;Integrated Security=True;Application Name=MiniGhost;"
			break
		}
		catch { }
		finally
		{
			if ($testConnection)
			{
				$testConnection.Close()
				$testConnection.Dispose()
			}
		}
	}
	
	# Store configured and runtime database context in the FunctionResults hashtable
	$script:FunctionResults['ConfiguredDBSERVER'] = $configuredDbServer
	$script:FunctionResults['DBSERVER'] = $configuredDbServer
	$script:FunctionResults['RuntimeDBSERVER'] = $resolvedRuntimeDbServer
	$script:FunctionResults['DBNAME'] = $dbName
	$script:FunctionResults['ConnectionString'] = $connectionString
}
