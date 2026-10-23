# ===================================================================================================
#                              FUNCTION: Insert_Test_Item
# ---------------------------------------------------------------------------------------------------
# Description:
#   Inserts a deterministic test item into the store database: picks one of three test PLUs
#   (the first that is unused or already a test item), deletes any existing rows for it, then
#   inserts fixed rows into SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB and SCL_TXT_TAB.
#   Uses System.Data.SqlClient (works without the SqlServer / SQLPS module) over one connection,
#   and does the deletes + inserts in one transaction: a failure leaves the database unchanged
#   instead of half a test item.
#
# Parameters:
#   - ConnectionString         [string]: defaults to the FunctionResults connection string
#   - Silent                   [switch]: suppress Write_Log output (startup job: no Write_Log there)
#   - QueryTimeoutSeconds      [int]: SQL query timeout in seconds (default 15)
#   - ConnectionTimeoutSeconds [int]: SQL connect timeout in seconds (default 5)
#
# Variables:
#   - $script:FunctionResults: reads ConnectionString, DBSERVER, DBNAME
#
# Notes:
#   - Self-contained: the startup warmup runs it in a plain Start-Job with only its own body.
#
# Writes:
#   - store database (SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB, SCL_TXT_TAB)
# ===================================================================================================

function Insert_Test_Item
{
	[CmdletBinding()]
	param (
		[string]$ConnectionString = $script:FunctionResults['ConnectionString'],
		[switch]$Silent,
		[int]$QueryTimeoutSeconds = 15,
		[int]$ConnectionTimeoutSeconds = 5
	)

	$writeLog = {
		param (
			[string]$Message,
			[string]$Color = "white"
		)
		if (-not $Silent) { Write_Log $Message $Color }
	}

	# No stored connection string: Windows login to DBSERVER / DBNAME
	if ([string]::IsNullOrWhiteSpace($ConnectionString) -or $ConnectionString -eq 'N/A')
	{
		$dbServer = [string]$script:FunctionResults['DBSERVER']
		$dbName = [string]$script:FunctionResults['DBNAME']
		$ConnectionString = $null
		if ($dbServer -and $dbName -and $dbServer -ne 'N/A' -and $dbName -ne 'N/A')
		{
			$ConnectionString = "Server=$dbServer;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"
		}
	}
	if (-not $ConnectionString)
	{
		& $writeLog "Insert_Test_Item: no connection string and no DBSERVER / DBNAME. Run Get_Store_And_Database_Info first." "red"
		return
	}
	try
	{
		$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder $ConnectionString
		$csb.PSBase.ConnectTimeout = $ConnectionTimeoutSeconds
		$ConnectionString = $csb.ConnectionString
	}
	catch
	{
		if ($ConnectionString -notmatch '(?i)\bConnect\s*Timeout\s*=') { $ConnectionString = $ConnectionString.Trim().TrimEnd(';') + ";Connect Timeout=$ConnectionTimeoutSeconds" }
	}

	& $writeLog "`r`n==================== Starting Insert_Test_Item ====================`r`n" "blue"

	$conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
	try { $conn.Open() }
	catch
	{
		& $writeLog "Insert_Test_Item: cannot connect to the store database: $($_.Exception.Message)" "red"
		& $writeLog "`r`n==================== Insert_Test_Item Aborted (No Connectivity) ====================`r`n" "blue"
		$conn.Dispose()
		return
	}

	try
	{
		$tx = $null
		$runSql = {
			param ([string]$Sql, [switch]$Scalar)
			$cmd = $conn.CreateCommand()
			$cmd.CommandText = $Sql
			$cmd.CommandTimeout = $QueryTimeoutSeconds
			if ($tx) { $cmd.Transaction = $tx }
			try
			{
				if ($Scalar)
				{
					$v = $cmd.ExecuteScalar()
					if ($v -is [DBNull]) { return $null }
					return $v
				}
				[void]$cmd.ExecuteNonQuery()
			}
			finally { $cmd.Dispose() }
		}

		# ------------------------------------------------------------------------------------------
		# Test PLU: the first candidate that is unused or already a test item (description
		# contains "test" / "tecnica"). A candidate whose lookup FAILS is skipped - never assume
		# a PLU is free when it could not be checked (it may be a real item).
		# ------------------------------------------------------------------------------------------
		$PLU = $null
		$TestF267 = $null
		foreach ($candidate in @(
				@{ Plu = '0020077700000'; F267 = 777; Kind = 'preferred' },
				@{ Plu = '0020777700000'; F267 = 7777; Kind = 'alternative' },
				@{ Plu = '0027777700000'; F267 = 77777; Kind = 'fallback' }))
		{
			try
			{
				$posDesc = [string](& $runSql "SELECT TOP 1 F02 FROM POS_TAB WHERE F01 = '$($candidate.Plu)'" -Scalar)
				$objDesc = [string](& $runSql "SELECT TOP 1 F29 FROM OBJ_TAB WHERE F01 = '$($candidate.Plu)'" -Scalar)
				$usable = ($posDesc -match '(?i)test|tecnica') -or ($objDesc -match '(?i)test|tecnica') -or
				([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
			}
			catch
			{
				& $writeLog ("Could not check PLU {0}: {1} - skipping it." -f $candidate.Plu, $_.Exception.Message) "yellow"
				$usable = $false
			}
			if ($usable)
			{
				$PLU = $candidate.Plu
				$TestF267 = $candidate.F267
				& $writeLog ("Using {0} PLU: {1} with F267: {2}" -f $candidate.Kind, $PLU, $TestF267) "green"
				break
			}
		}
		if (-not $PLU)
		{
			& $writeLog "`r`n==================== Insert_Test_Item Completed (No Data) ====================`r`n" "blue"
			return
		}

		$now = Get-Date
		$nowFull = $now.ToString("yyyy-MM-dd HH:mm:ss.fff")
		$nowDate = $now.ToString("yyyy-MM-dd 00:00:00.000")
		$F29 = 'Tecnica Test Item'

		# Old rows first, so the inserts are deterministic. The scale text row is only removed
		# when no other item uses that F267 (it could be a real one).
		$steps = @(
			@('delete SCL_TAB', "DELETE FROM SCL_TAB     WHERE F01 = '$PLU'"),
			@('delete OBJ_TAB', "DELETE FROM OBJ_TAB     WHERE F01 = '$PLU'"),
			@('delete POS_TAB', "DELETE FROM POS_TAB     WHERE F01 = '$PLU'"),
			@('delete PRICE_TAB', "DELETE FROM PRICE_TAB   WHERE F01 = '$PLU'"),
			@('delete SCL_TXT_TAB', "DELETE FROM SCL_TXT_TAB WHERE F267 = $TestF267 AND F267 NOT IN (SELECT F267 FROM SCL_TAB WHERE F01 <> '$PLU' AND F267 IS NOT NULL)"),
			@('SCL_TAB', "INSERT INTO SCL_TAB (F01, F1000, F902, F1001, F253, F258, F264, F267, F1952, F1964, F2581, F2582) VALUES ('$PLU', 'PAL', 'MANUAL', 1, '$nowFull', 10, 7, $TestF267, 'Test Descriptor 2', '001', 'Test Descriptor 3', 'Test Descriptor 4')"),
			@('OBJ_TAB', "INSERT INTO OBJ_TAB (F01, F902, F1001, F21, F29, F270, F1118, F1959) VALUES ('$PLU', '00001153', 0, 1, '$F29', 123.45, '001', '001')"),
			@('POS_TAB', "INSERT INTO POS_TAB (F01, F1000, F902, F1001, F02, F09, F79, F80, F82, F104, F115, F176, F178, F217, F1964, F2119) VALUES ('$PLU', 'PAL', 'MANUAL', 0, 'Tecnica Test Item', '$nowDate', '1', '1', '1', '0', '0', '1', '1', 1.0, '001', '1')"),
			@('PRICE_TAB', "INSERT INTO PRICE_TAB (F01, F1000, F126, F902, F1001, F21, F30, F31, F113, F1006, F1007, F1008, F1009, F1803) VALUES ('$PLU', 'PAL', 1, 'MANUAL', 0, 1, 777.77, 1, 'REG', 1, 777.77, '$nowDate', '1858', 1.0)"),
			@('SCL_TXT_TAB', "INSERT INTO SCL_TXT_TAB (F267, F1000, F253, F297, F902, F1001, F1836) VALUES ($TestF267, 'PAL', '$nowFull', 'Ingredients Test', 'MANUAL', 0, 'Tecnica Test Item')")
		)

		$tx = $conn.BeginTransaction()
		$failed = $null
		foreach ($s in $steps)
		{
			try { & $runSql $s[1] }
			catch
			{
				$failed = "{0}: {1}" -f $s[0], $_.Exception.Message
				break
			}
		}
		if ($failed)
		{
			try { $tx.Rollback() }
			catch { }
			& $writeLog "Test item NOT inserted - $failed (rolled back, the database is unchanged)." "red"
			& $writeLog "`r`n==================== Insert_Test_Item Aborted ====================`r`n" "blue"
			return
		}
		$tx.Commit()
		& $writeLog "Test item $PLU inserted into SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB and SCL_TXT_TAB." "green"
		& $writeLog "`r`n==================== Insert_Test_Item Completed ====================`r`n" "blue"
	}
	catch
	{
		& $writeLog "Insert_Test_Item failed: $($_.Exception.Message)" "red"
	}
	finally
	{
		$conn.Dispose()
	}
}
