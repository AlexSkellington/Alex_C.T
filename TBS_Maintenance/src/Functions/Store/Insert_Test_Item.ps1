# ===================================================================================================
#                              FUNCTION: Insert_Test_Item
# ---------------------------------------------------------------------------------------------------
# Description:
#   Inserts a deterministic test item into the store database: picks one of three test PLUs,
#   deletes any existing rows for it, then inserts fixed rows into SCL_TAB, OBJ_TAB, POS_TAB,
#   PRICE_TAB and SCL_TXT_TAB. Uses Invoke-Sqlcmd with the connection string from
#   Get_Store_And_Database_Info, falling back to -ServerInstance/-Database.
#
# Parameters:
#   - ConnectionString         [string]: defaults to the FunctionResults connection string
#   - Silent                   [switch]: suppress Write_Log output
#   - QueryTimeoutSeconds      [int]: SQL query timeout in seconds (default 15)
#   - ConnectionTimeoutSeconds [int]: SQL connect timeout in seconds (default 5)
#
# Variables:
#   - $script:FunctionResults: reads ConnectionString, DBSERVER, DBNAME, SqlModuleName
#
# Writes:
#   - store database (SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB, SCL_TXT_TAB)
# ===================================================================================================

function Insert_Test_Item
{
	[CmdletBinding()]
	param (
		# If omitted, we use the ConnectionString assembled by Get_Store_And_Database_Info.
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

	if ($ConnectionString)
	{
		try
		{
			$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder $ConnectionString
			$csb.PSBase.ConnectTimeout = $ConnectionTimeoutSeconds
			$ConnectionString = $csb.ConnectionString
		}
		catch
		{
			if ($ConnectionString -notmatch '(?i)\bConnect\s*Timeout\s*=')
			{
				$ConnectionString = $ConnectionString.Trim().TrimEnd(';') + ";Connect Timeout=$ConnectionTimeoutSeconds"
			}
		}
	}

	# ----------------------------------------------------------------------------------------------
	# Guard: Need connection context from Get_Store_And_Database_Info
	# (DBSERVER/DBNAME and/or ConnectionString should be available)
	# ----------------------------------------------------------------------------------------------
	if (-not $ConnectionString -and
		(-not $script:FunctionResults['DBSERVER'] -or -not $script:FunctionResults['DBNAME']))
	{
		& $writeLog "Insert_Test_Item: No ConnectionString nor DBSERVER/DBNAME available. Run Get_Store_And_Database_Info first." "red"
		return
	}

	& $writeLog "`r`n==================== Starting Insert_Test_Item ====================`r`n" "blue"

	# ----------------------------------------------------------------------------------------------
	# Ensure Invoke-Sqlcmd is available (prefer SqlServer, fallback to SQLPS)
	# We honor the detection already done in Get_Store_And_Database_Info.
	# ----------------------------------------------------------------------------------------------
	$sqlModule = $script:FunctionResults['SqlModuleName']
	if (-not (Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue))
	{
		try
		{
			if ($sqlModule -eq 'SqlServer')
			{
				Import-Module SqlServer -ErrorAction Stop
			}
			elseif ($sqlModule -eq 'SQLPS')
			{
				Import-Module SQLPS -DisableNameChecking -ErrorAction Stop
			}
		}
		catch
		{
			# If import fails, we'll check again below; if still missing, we abort.
		}
	}

	if (-not (Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue))
	{
		& $writeLog "Insert_Test_Item: Invoke-Sqlcmd not available (SqlServer/SQLPS modules not loaded/installed)." "red"
		& $writeLog "`r`n==================== Insert_Test_Item Aborted (No SQL Module) ====================`r`n" "blue"
		return
	}

	$invokeSqlcmdCommand = $null
	$supportsConnTimeout = $false
	try
	{
		$invokeSqlcmdCommand = Get-Command Invoke-Sqlcmd -ErrorAction Stop
		if ($invokeSqlcmdCommand -and $invokeSqlcmdCommand.Parameters)
		{
			$supportsConnTimeout = $invokeSqlcmdCommand.Parameters.ContainsKey('ConnectionTimeout')
		}
	}
	catch { }

	# ----------------------------------------------------------------------------------------------
	# Decide primary vs. fallback path:
	#   - If SqlServer module is detected, we ATTEMPT -ConnectionString (modern path).
	#   - Otherwise (SQLPS/None), we go straight to fallback (-ServerInstance/-Database).
	#   - Even with SqlServer, if the probe SELECT 1 fails, we flip to fallback.
	# ----------------------------------------------------------------------------------------------
	$useConnectionString = $false # Default to fallback unless we validate modern path
	$fallbackParams = @{ } # Will be used with -ServerInstance/-Database

	# Pull server/database from FunctionResults (preferred over parsing)
	$serverInstanceFromInfo = $script:FunctionResults['DBSERVER']
	$databaseFromInfo = $script:FunctionResults['DBNAME']

	# Compose fallback param set from known values
	if ($serverInstanceFromInfo) { $fallbackParams['ServerInstance'] = $serverInstanceFromInfo }
	if ($databaseFromInfo) { $fallbackParams['Database'] = $databaseFromInfo }

	# If we have SqlServer module, try modern path first.
	if ($sqlModule -eq 'SqlServer' -and $ConnectionString)
	{
		try
		{
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query "SELECT 1" -QueryTimeout $QueryTimeoutSeconds -ErrorAction Stop | Out-Null
			$useConnectionString = $true
			& $writeLog "SQL connectivity via -ConnectionString verified (SqlServer module)." "green"
		}
		catch
		{
			& $writeLog "Modern path (-ConnectionString) failed: $($_.Exception.Message)`r`nFalling back to -ServerInstance/-Database..." "yellow"
			$useConnectionString = $false
		}
	}
	else
	{
		# Either SQLPS or no module indicated from detection - use fallback path.
		$useConnectionString = $false
	}

	# If planning to use fallback, validate it can connect
	if (-not $useConnectionString)
	{
		# If DBSERVER/DBNAME were missing, try to salvage from the connection string
		if (-not $fallbackParams['ServerInstance'] -or -not $fallbackParams['Database'])
		{
			# Last-resort: parse minimal fields from ConnectionString (only if needed)
			if ($ConnectionString -match '(?i)(?:Data\s*Source|Server)\s*=\s*([^;]+)')
			{
				$fallbackParams['ServerInstance'] = $matches[1].Trim()
			}
			if ($ConnectionString -match '(?i)(?:Initial\s*Catalog|Database)\s*=\s*([^;]+)')
			{
				$fallbackParams['Database'] = $matches[1].Trim()
			}
		}

		try
		{
			$fallbackProbeParams = $fallbackParams.Clone()
			$fallbackProbeParams['Query'] = 'SELECT 1'
			$fallbackProbeParams['QueryTimeout'] = $QueryTimeoutSeconds
			$fallbackProbeParams['ErrorAction'] = 'Stop'
			if ($supportsConnTimeout) { $fallbackProbeParams['ConnectionTimeout'] = $ConnectionTimeoutSeconds }
			Invoke-Sqlcmd @fallbackProbeParams | Out-Null
			$dbTxt = if ($fallbackParams['Database']) { " (DB=$($fallbackParams['Database']))" }
			else { "" }
			& $writeLog "SQL connectivity via fallback -ServerInstance verified: $($fallbackParams['ServerInstance'])$dbTxt" "green"
		}
		catch
		{
			& $writeLog "Fallback (-ServerInstance/-Database) failed: $($_.Exception.Message)" "red"
			& $writeLog "`r`n==================== Insert_Test_Item Aborted (No Connectivity) ====================`r`n" "blue"
			return
		}
	}

	# ----------------------------------------------------------------------------------------------
	# Centralized SQL runner - ALWAYS use this for queries below so fallback is automatic.
	# (CHANGE: kept as a scriptblock variable - not a nested *function* - to honor 'no nested functions')
	# ----------------------------------------------------------------------------------------------
	$RunSql = {
		param ([Parameter(Mandatory = $true)]
			[string]$Query)
		if ($useConnectionString)
		{
			return Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $Query -QueryTimeout $QueryTimeoutSeconds -ErrorAction Stop
		}
		else
		{
			$p = $fallbackParams.Clone()
			$p['Query'] = $Query
			$p['QueryTimeout'] = $QueryTimeoutSeconds
			$p['ErrorAction'] = 'Stop'
			if ($supportsConnTimeout) { $p['ConnectionTimeout'] = $ConnectionTimeoutSeconds }
			return Invoke-Sqlcmd @p
		}
	}

	# ----------------------------------------------------------------------------------------------
	# Business logic: choose appropriate PLU (prefer the "test" one), then delete + insert rows
	# (CHANGE: removed nested helper function; inlined the PLU test logic 3x for clarity/compliance)
	# ----------------------------------------------------------------------------------------------
	$now = Get-Date
	$nowFull = $now.ToString("yyyy-MM-dd HH:mm:ss.fff")
	$nowDate = $now.ToString("yyyy-MM-dd 00:00:00.000")

	$preferredPLU = '0020077700000'
	$alternativePLU = '0020777700000'
	$fallbackPLU = '0027777700000'

	$PLU = $null
	$TestF267 = 777
	$doInsert = $false

	# ---- Check preferred PLU (INLINE)
	$okUse = $false
	try
	{
		$pos = & $RunSql "SELECT F02 FROM POS_TAB WHERE F01 = '$preferredPLU'"
		$obj = & $RunSql "SELECT F29 FROM OBJ_TAB WHERE F01 = '$preferredPLU'"
		$posDesc = if ($pos) { [string]$pos.F02 }
		else { "" }
		$objDesc = if ($obj) { [string]$obj.F29 }
		else { "" }
		$okUse = ($posDesc -match '(?i)test|tecnica') -or
		($objDesc -match '(?i)test|tecnica') -or
		([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
	}
	catch { $okUse = $true } # conservative default if lookups fail
	if ($okUse)
	{
		$PLU = $preferredPLU
		$TestF267 = 777
		$doInsert = $true
		& $writeLog "Using preferred PLU: $PLU with F267: $TestF267" "green"
	}

	# ---- If not chosen yet, check alternative PLU (INLINE)
	if (-not $doInsert)
	{
		$okUse = $false
		try
		{
			$pos = & $RunSql "SELECT F02 FROM POS_TAB WHERE F01 = '$alternativePLU'"
			$obj = & $RunSql "SELECT F29 FROM OBJ_TAB WHERE F01 = '$alternativePLU'"
			$posDesc = if ($pos) { [string]$pos.F02 }
			else { "" }
			$objDesc = if ($obj) { [string]$obj.F29 }
			else { "" }
			$okUse = ($posDesc -match '(?i)test|tecnica') -or
			($objDesc -match '(?i)test|tecnica') -or
			([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
		}
		catch { $okUse = $true }
		if ($okUse)
		{
			$PLU = $alternativePLU
			$TestF267 = 7777
			$doInsert = $true
			& $writeLog "Using alternative PLU: $PLU with F267: $TestF267" "green"
		}
	}

	# ---- If still not chosen, check fallback PLU (INLINE)
	if (-not $doInsert)
	{
		$okUse = $false
		try
		{
			$pos = & $RunSql "SELECT F02 FROM POS_TAB WHERE F01 = '$fallbackPLU'"
			$obj = & $RunSql "SELECT F29 FROM OBJ_TAB WHERE F01 = '$fallbackPLU'"
			$posDesc = if ($pos) { [string]$pos.F02 }
			else { "" }
			$objDesc = if ($obj) { [string]$obj.F29 }
			else { "" }
			$okUse = ($posDesc -match '(?i)test|tecnica') -or
			($objDesc -match '(?i)test|tecnica') -or
			([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
		}
		catch { $okUse = $true }
		if ($okUse)
		{
			$PLU = $fallbackPLU
			$TestF267 = 77777
			$doInsert = $true
			& $writeLog "Using fallback PLU: $PLU with F267: $TestF267" "green"
		}
	}

	if ($doInsert -and $PLU)
	{
		& $writeLog "Deleting existing records for PLU: $PLU and F267: $TestF267" "yellow"

		# --- Always delete old rows for the chosen PLU/F267 to keep inserts deterministic
		$deleteQueries = @(
			"DELETE FROM SCL_TAB     WHERE F01 = '$PLU'",
			"DELETE FROM OBJ_TAB     WHERE F01 = '$PLU'",
			"DELETE FROM POS_TAB     WHERE F01 = '$PLU'",
			"DELETE FROM PRICE_TAB   WHERE F01 = '$PLU'",
			"DELETE FROM SCL_TXT_TAB WHERE F267 = $TestF267"
		)
		foreach ($q in $deleteQueries)
		{
			try { & $RunSql $q | Out-Null }
			catch { & $writeLog "Error during deletion: $($_.Exception.Message)" "red" }
		}

		# --- SCL_TAB
		& $writeLog "Inserting into SCL_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO SCL_TAB (F01, F1000, F902, F1001, F253, F258, F264, F267, F1952, F1964, F2581, F2582)
VALUES
('$PLU', 'PAL', 'MANUAL', 1, '$nowFull', 10, 7, $TestF267, 'Test Descriptor 2', '001', 'Test Descriptor 3', 'Test Descriptor 4')
"@ | Out-Null
			& $writeLog "SCL_TAB insertion successful" "green"
		}
		catch { & $writeLog "Error inserting into SCL_TAB: $($_.Exception.Message)" "red" }

		# --- OBJ_TAB
		& $writeLog "Inserting into OBJ_TAB..." "yellow"
		try
		{
			$F29 = 'Tecnica Test Item'
			if ($F29.Length -gt 60) { $F29 = $F29.Substring(0, 60) }
			& $RunSql @"
INSERT INTO OBJ_TAB (F01, F902, F1001, F21, F29, F270, F1118, F1959)
VALUES
('$PLU', '00001153', 0, 1, '$F29', 123.45, '001', '001')
"@ | Out-Null
			& $writeLog "OBJ_TAB insertion successful" "green"
		}
		catch { & $writeLog "Error inserting into OBJ_TAB: $($_.Exception.Message)" "red" }

		# --- POS_TAB
		& $writeLog "Inserting into POS_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO POS_TAB (F01, F1000, F902, F1001, F02, F09, F79, F80, F82, F104, F115, F176, F178, F217, F1964, F2119)
VALUES
('$PLU', 'PAL', 'MANUAL', 0, 'Tecnica Test Item', '$nowDate', '1', '1', '1', '0', '0', '1', '1', 1.0, '001', '1')
"@ | Out-Null
			& $writeLog "POS_TAB insertion successful" "green"
		}
		catch { & $writeLog "Error inserting into POS_TAB: $($_.Exception.Message)" "red" }

		# --- PRICE_TAB
		& $writeLog "Inserting into PRICE_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO PRICE_TAB (F01, F1000, F126, F902, F1001, F21, F30, F31, F113, F1006, F1007, F1008, F1009, F1803)
VALUES
('$PLU', 'PAL', 1, 'MANUAL', 0, 1, 777.77, 1, 'REG', 1, 777.77, '$nowDate', '1858', 1.0)
"@ | Out-Null
			& $writeLog "PRICE_TAB insertion successful" "green"
		}
		catch { & $writeLog "Error inserting into PRICE_TAB: $($_.Exception.Message)" "red" }

		# --- SCL_TXT_TAB
		& $writeLog "Inserting into SCL_TXT_TAB..." "yellow"
		try
		{
			& $RunSql @"
INSERT INTO SCL_TXT_TAB (F267, F1000, F253, F297, F902, F1001, F1836)
VALUES
($TestF267, 'PAL', '$nowFull', 'Ingredients Test', 'MANUAL', 0, 'Tecnica Test Item')
"@ | Out-Null
			& $writeLog "SCL_TXT_TAB insertion successful" "green"
		}
		catch { & $writeLog "Error inserting into SCL_TXT_TAB: $($_.Exception.Message)" "red" }

		& $writeLog "`r`n==================== Insert_Test_Item Completed ====================`r`n" "blue"
	}
	else
	{
		& $writeLog "`r`n==================== Insert_Test_Item Completed (No Data) ====================`r`n" "blue"
	}
}
