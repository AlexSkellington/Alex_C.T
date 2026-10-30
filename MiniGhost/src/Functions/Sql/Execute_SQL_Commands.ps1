# ===================================================================================================
#                                       FUNCTION: Execute_SQL_Commands
# ---------------------------------------------------------------------------------------------------
# Description:
#   Executes a given SQL command using the provided connection string.
# ===================================================================================================

function Execute_SQL_Commands
{
	param (
		[string]$commandText
	)
	
	$script:LastSqlExecutionError = $null
	$resolvedConnectionString = Ensure_Database_Connection_Context
	if ([string]::IsNullOrWhiteSpace($resolvedConnectionString))
	{
		$script:LastSqlExecutionError = "Database connection string is not available."
		return $false
	}
	
	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$sqlConnection.ConnectionString = $resolvedConnectionString
	$sqlCommand = $sqlConnection.CreateCommand()
	$sqlCommand.CommandText = $commandText
	
	try
	{
		$sqlConnection.Open()
		$sqlCommand.ExecuteNonQuery() | Out-Null
		return $true # Indicate success
	}
	catch
	{
		$script:LastSqlExecutionError = $_.Exception.Message
		return $false # Indicate failure
	}
	finally
	{
		$sqlConnection.Close()
	}
}
