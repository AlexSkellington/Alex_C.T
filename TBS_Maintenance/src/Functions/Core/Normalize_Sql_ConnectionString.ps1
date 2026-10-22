# ===================================================================================================
#                              FUNCTION: Normalize_Sql_ConnectionString
# ---------------------------------------------------------------------------------------------------
# Description:
#   Turns a stored connection string (or the Server/Database/User/Pwd parts) into one clean
#   SqlClient connection string. Treats 'N/A' placeholders as empty and always adds
#   TrustServerCertificate=True so older Windows builds can reach newer SQL servers.
#
# Parameters:
#   - RawConnectionString  [any]: value from FunctionResults (string, hashtable or object with .ConnectionString)
#   - Server               [string]: SQL server name (used when no usable connection string is stored)
#   - Database             [string]: database name (used when no usable connection string is stored)
#   - SqlUser              [string]: SQL login (used with Server/Database)
#   - SqlPwd               [string]: SQL login password (used with Server/Database)
#
# Returns:
#   The normalized connection string, or $null when nothing usable was given.
# ===================================================================================================

function Normalize_Sql_ConnectionString
{
	param (
		[Parameter(Mandatory = $false)]
		$RawConnectionString,
		[string]$Server,
		[string]$Database,
		[string]$SqlUser,
		[string]$SqlPwd
	)

	$connStr = $null
	try
	{
		if ($RawConnectionString -is [hashtable] -and $RawConnectionString.ContainsKey('ConnectionString'))
		{
			$connStr = [string]$RawConnectionString['ConnectionString']
		}
		elseif ($null -ne $RawConnectionString -and $RawConnectionString.PSObject -and $RawConnectionString.PSObject.Properties['ConnectionString'])
		{
			$connStr = [string]$RawConnectionString.ConnectionString
		}
		else
		{
			$connStr = [string]$RawConnectionString
		}
	}
	catch
	{
		try { $connStr = [string]$RawConnectionString } catch { $connStr = $null }
	}

	if (-not [string]::IsNullOrWhiteSpace($connStr))
	{
		$connStr = $connStr.Trim().Trim('"')
		if ($connStr -match '^\s*ConnectionString\s*=\s*(.+?)\s*$')
		{
			$connStr = $matches[1].Trim()
		}
	}

	if ([string]::IsNullOrWhiteSpace($connStr) -or $connStr -eq 'N/A' -or $connStr -eq 'System.Collections.Hashtable')
	{
		$connStr = $null
	}

	# Treat the 'N/A' placeholder the same as empty - otherwise this builds a
	# real connection string pointing at a SQL server literally named "N/A",
	# and every caller pays a full DNS/connect timeout (~15s at startup).
	$serverNorm = $Server
	$databaseNorm = $Database
	try
	{
		if ([string]$Server -eq 'N/A' -or [string]$Server -eq 'System.Collections.Hashtable') { $serverNorm = '' }
		if ([string]$Database -eq 'N/A' -or [string]$Database -eq 'System.Collections.Hashtable') { $databaseNorm = '' }
	}
	catch { }
	
	if (-not $connStr -and -not [string]::IsNullOrWhiteSpace($serverNorm) -and -not [string]::IsNullOrWhiteSpace($databaseNorm))
	{
		if (-not [string]::IsNullOrWhiteSpace($SqlUser) -and -not [string]::IsNullOrWhiteSpace($SqlPwd))
		{
			$connStr = "Server=${serverNorm};Database=${databaseNorm};User ID=$SqlUser;Password=$SqlPwd;TrustServerCertificate=True;"
		}
		else
		{
			$connStr = "Server=${serverNorm};Database=${databaseNorm};Integrated Security=SSPI;TrustServerCertificate=True;"
		}
	}

	if (-not [string]::IsNullOrWhiteSpace($connStr))
	{
		try
		{
			$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
			$csb.ConnectionString = $connStr
			if (-not $csb.ContainsKey('TrustServerCertificate'))
			{
				$csb['TrustServerCertificate'] = 'True'
			}
			$connStr = $csb.ConnectionString
		}
		catch { }
	}

	return $connStr
}
