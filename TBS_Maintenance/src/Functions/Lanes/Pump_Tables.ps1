# ===================================================================================================
#                              FUNCTION: Pump_Tables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Copies the selected tables (from Get_Table_Aliases) from the store SQL Server to the selected
#   lanes as an UPSERT: rows the lane already has (same primary key) are updated when a value
#   differs, missing rows are added, nothing is truncated or deleted and the lane's table schema
#   is kept (a missing table is created with the server's columns and primary key).
#   - Lanes reachable over SQL (cached $script:LaneProtocols TCP / Named Pipes) are pumped in
#     parallel (up to 6 at once, runspace pool), one transaction per table; the window stays
#     responsive and shows each lane's progress.
#   - Lanes on the File protocol, and tables whose SQL pump failed, get the SMS file drop
#     (@UPDATE_BATCH ADDRPL = update-or-add too), tables of up to 1000 rows. Each file is
#     generated once and copied to every lane that needs it.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName / MachineNameToLaneNum (input)
#   - $script:LaneProtocols: cached per-lane protocol (TCP / Named Pipes / File)
#   - $OfficePath: Storeman Office root (XF lane folders)
# ===================================================================================================

function Pump_Tables
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	# Log function start
	Write_Log "`r`n==================== Starting Pump_Tables Function ====================`r`n" "blue"

	# ------------------------------------------------------------------------------------------------
	# FIX: Initialize tracking arrays locally (prevents null/type issues)
	# ------------------------------------------------------------------------------------------------
	$protocolLanes = @()
	$fileCopyLanes = @()
	$ProcessedLanes = @()

	# ------------------------------------------------------------------------------------------------
	# FIX: Per-run file generation cache (generate each selected table ONCE, then copy to many lanes)
	#      Key: table name -> temp file path
	# ------------------------------------------------------------------------------------------------
	$generatedSqlFiles = @{ }

	# Ensure OfficePath is present and valid
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "Office path not found (cannot locate XF folders): $OfficePath" "yellow"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 1: Load and validate lane mappings from Retrieve_Nodes (and allow for all flexible lookups)
	# ------------------------------------------------------------------------------------------------
	if (-not ($script:FunctionResults.ContainsKey('LaneNumToMachineName')))
	{
		Write_Log "Lane mappings not found. Please run Retrieve_Nodes first." "red"
		return
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$MachineNameToLaneNum = $script:FunctionResults['MachineNameToLaneNum']
	if (-not $MachineNameToLaneNum) { $MachineNameToLaneNum = @{ } }

	# ------------------------------------------------------------------------------------------------
	# STEP 2: Prompt user for lane selection via GUI
	# ------------------------------------------------------------------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($null -eq $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "Lane processing canceled by user." "yellow"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# Accept all supported forms (lane number, machine name, POSxxx, etc)
	$Lanes = $selection.Lanes

	# FIX: Normalize selection to 3-digit lane numbers using Retrieve_Nodes mappings
	$Lanes = $Lanes | ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		elseif ($_ -match '^\d{3}$') { $_ }
		elseif ($MachineNameToLaneNum.ContainsKey($_)) { $MachineNameToLaneNum[$_] }
		elseif ($_ -match '(\d{1,3})$') { $Matches[1].PadLeft(3, '0') }
		else { $_ }
	}
	$Lanes = @($Lanes | Where-Object { $LaneNumToMachineName.ContainsKey($_) } | Select-Object -Unique)

	if (-not $Lanes -or $Lanes.Count -eq 0)
	{
		Write_Log "No valid lanes to process. Exiting Pump_Tables." "yellow"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 3: Retrieve table aliases and prompt for tables to process
	# ------------------------------------------------------------------------------------------------
	if ($script:FunctionResults.ContainsKey('Get_Table_Aliases'))
	{
		$aliasData = $script:FunctionResults['Get_Table_Aliases']
		$aliasResults = $aliasData.Aliases
		$aliasHash = $aliasData.AliasHash

		if (($aliasHash -isnot [System.Collections.IDictionary]) -and $aliasHash -and $aliasHash.PSObject -and $aliasHash.PSObject.Properties)
		{
			$tmpAliasHash = @{ }
			foreach ($p in $aliasHash.PSObject.Properties) { $tmpAliasHash[$p.Name] = [string]$p.Value }
			$aliasHash = $tmpAliasHash
		}

		if ((-not $aliasResults -or $aliasResults.Count -eq 0) -and ($aliasHash -is [System.Collections.IDictionary]) -and $aliasHash.Count -gt 0)
		{
			$aliasResults = New-Object System.Collections.ArrayList
			foreach ($tableName in ($aliasHash.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique))
			{
				$matchingAliases = @($aliasHash.GetEnumerator() | Where-Object { $_.Value -eq $tableName } | Select-Object -ExpandProperty Key)
				$aliasLabel = if ($matchingAliases.Count -gt 0) { ($matchingAliases | Sort-Object)[0] }
				else { $null }
				$null = $aliasResults.Add([pscustomobject]@{
						Table = [string]$tableName
						Alias = [string]$aliasLabel
					})
			}
		}
	}
	else
	{
		Write_Log "Alias data not found. Ensure Get_Table_Aliases has been run." "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}
	if (-not $aliasResults -or $aliasResults.Count -eq 0)
	{
		Write_Log "No tables found to process. Exiting Pump_Tables." "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# Prompt user to select which tables to pump (Show_Table_Selection_Form returns TABLE NAMES)
	$selectedTables = Show_Table_Selection_Form -AliasResults $aliasResults
	if (-not $selectedTables -or $selectedTables.Count -eq 0)
	{
		Write_Log "No tables were selected. Exiting Pump_Tables." "yellow"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# FIX: Build the actual alias entries to loop (was missing in previous version)
	$selectedTables = @($selectedTables | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
	$filteredAliasEntries = @($aliasResults | Where-Object { $selectedTables -contains $_.Table })
	if (-not $filteredAliasEntries -or $filteredAliasEntries.Count -eq 0)
	{
		Write_Log "Selected tables did not match alias list. Nothing to pump." "yellow"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 4: Resolve LOCAL source ConnectionString (where we read tables FROM)
	# ------------------------------------------------------------------------------------------------
	$ConnectionString = $null

	# 1) Preferred: already in FunctionResults
	if ($script:FunctionResults.ContainsKey('ConnectionString'))
	{
		$ConnectionString = $script:FunctionResults['ConnectionString']
	}

	# 2) Fallback: pull from cached JSON in SetupFiles (Get_All_Lanes_Database_Info.txt)
	if ([string]::IsNullOrWhiteSpace($ConnectionString))
	{
		try
		{
			$cachePath = Join-Path $script:SetupFiles "Get_All_Lanes_Database_Info.txt"
			if (Test-Path $cachePath)
			{
				$cacheObj = (Get-Content -Path $cachePath -Raw -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
				if ($cacheObj -and $cacheObj.LaneDatabaseInfo)
				{
					$laneDbInfo = $cacheObj.LaneDatabaseInfo
					$localMachine = ($env:COMPUTERNAME).ToUpper()
					$match = $null
					foreach ($p in $laneDbInfo.PSObject.Properties)
					{
						if ($p.Value -and $p.Value.MachineName -and ($p.Value.MachineName.ToString().ToUpper() -eq $localMachine))
						{
							$match = $p.Value
							break
						}
					}
					# If ONLY ONE lane selected, use that lane's cached connstr as last resort
					if (-not $match -and $Lanes -and $Lanes.Count -eq 1)
					{
						$laneKey = [string]$Lanes[0]
						$prop = $laneDbInfo.PSObject.Properties | Where-Object { $_.Name -eq $laneKey } | Select-Object -First 1
						if ($prop) { $match = $prop.Value }
					}
					if ($match)
					{
						if (-not [string]::IsNullOrWhiteSpace([string]$match.ConnectionString))
						{
							$ConnectionString = [string]$match.ConnectionString
						}
						elseif (-not [string]::IsNullOrWhiteSpace([string]$match.TcpConnStr))
						{
							$ConnectionString = [string]$match.TcpConnStr
						}
					}
				}
			}
		}
		catch
		{
			# Silent here; we handle final failure below
		}
	}

	# Final: must have source connection
	if ([string]::IsNullOrWhiteSpace($ConnectionString))
	{
		Write_Log "Connection string not found (FunctionResults + ${script:SetupFiles}\Get_All_Lanes_Database_Info.txt). Cannot proceed with Pump_Tables." "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# Normalize + cache it so nothing else needs to change
	$ConnectionString = $ConnectionString.Trim().Trim('"')
	$script:FunctionResults['ConnectionString'] = $ConnectionString

	# ------------------------------------------------------------------------------------------------
	# STEP 4b: Open SOURCE SQL connection ONCE (was missing before)
	# ------------------------------------------------------------------------------------------------
	$srcSqlConnection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
	try
	{
		$srcSqlConnection.Open()
		if ($srcSqlConnection.State -ne 'Open')
		{
			Write_Log "Source SQL connection did not open. Cannot proceed." "red"
			Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
			return
		}
	}
	catch
	{
		Write_Log "Source SQL connection failed: $_" "red"
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 5: Read each selected table's shape from the SOURCE once (shared by every lane)
	# ------------------------------------------------------------------------------------------------
	$tableRowCounts = @{ }
	$tablePlans = New-Object System.Collections.Generic.List[object]
	foreach ($aliasEntry in $filteredAliasEntries)
	{
		$tableName = [string]$aliasEntry.Table
		if ([string]::IsNullOrWhiteSpace($tableName) -or $tableRowCounts.ContainsKey($tableName)) { continue }
		# Table-name sanity (prevents weird alias entries from breaking queries)
		if ($tableName -notmatch '^[A-Za-z0-9_]+$')
		{
			Write_Log "Skipping invalid table name (bad characters): '$tableName'" "yellow"
			continue
		}
		try
		{
			$cntCmd = $srcSqlConnection.CreateCommand()
			$cntCmd.CommandTimeout = 60
			$cntCmd.CommandText = "SELECT COUNT_BIG(1) FROM [dbo].[$tableName]"
			$tableRowCounts[$tableName] = [long]$cntCmd.ExecuteScalar()

			$metaCmd = $srcSqlConnection.CreateCommand()
			$metaCmd.CommandTimeout = 60
			$metaCmd.CommandText = @"
SELECT c.name AS ColumnName, t.name AS DataType, c.max_length AS MaxLength, c.precision AS [Precision],
	c.scale AS [Scale], c.is_nullable AS IsNullable, c.is_identity AS IsIdentity, c.is_computed AS IsComputed,
	ic.seed_value AS IdentitySeed, ic.increment_value AS IdentityIncrement
FROM sys.columns c
JOIN sys.types t ON c.user_type_id = t.user_type_id
LEFT JOIN sys.identity_columns ic ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE c.object_id = OBJECT_ID(@obj)
ORDER BY c.column_id;
"@
			$param = $metaCmd.Parameters.Add("@obj", [System.Data.SqlDbType]::NVarChar, 256)
			$param.Value = "dbo.$tableName"
			$metaRows = New-Object System.Collections.Generic.List[object]
			$metaReader = $metaCmd.ExecuteReader()
			try
			{
				while ($metaReader.Read())
				{
					$metaRows.Add([pscustomobject]@{
							ColumnName = [string]$metaReader['ColumnName']
							DataType   = [string]$metaReader['DataType']
							MaxLength  = [int]$metaReader['MaxLength']
							Precision  = [int]$metaReader['Precision']
							Scale	   = [int]$metaReader['Scale']
							IsNullable = [bool]$metaReader['IsNullable']
							IsIdentity = [bool]$metaReader['IsIdentity']
							IsComputed = [bool]$metaReader['IsComputed']
							IdentitySeed = $metaReader['IdentitySeed']
							IdentityIncrement = $metaReader['IdentityIncrement']
						})
				}
			}
			finally { $metaReader.Close() }
			if ($metaRows.Count -eq 0) { throw "source table metadata not found" }

			$pkCmd = $srcSqlConnection.CreateCommand()
			$pkCmd.CommandTimeout = 60
			$pkCmd.CommandText = @"
SELECT c.name FROM sys.indexes i
JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE i.object_id = OBJECT_ID(@obj) AND i.is_primary_key = 1
ORDER BY ic.key_ordinal;
"@
			$p2 = $pkCmd.Parameters.Add("@obj", [System.Data.SqlDbType]::NVarChar, 256)
			$p2.Value = "dbo.$tableName"
			$srcPk = New-Object System.Collections.Generic.List[string]
			$pkReader = $pkCmd.ExecuteReader()
			try { while ($pkReader.Read()) { $srcPk.Add([string]$pkReader[0]) } }
			finally { $pkReader.Close() }

			# CREATE TABLE text for lanes that do not have the table yet (same shape as the source)
			$table = $tableName
			$hasComputed = @($metaRows | Where-Object { $_.IsComputed }).Count -gt 0
			$createSql = $null
			if (-not $hasComputed)
			{
				# Build CREATE TABLE from source metadata
				$colDefs = @()
				foreach ($c in $metaRows)
				{
					$colName = $c.ColumnName
					$dataType = $c.DataType
					$maxLen = [int]$c.MaxLength
					$prec = [int]$c.Precision
					$scale = [int]$c.Scale
					$isNull = [bool]$c.IsNullable
					$isIdent = [bool]$c.IsIdentity

					# Skip computed + rowversion for CREATE when destination is missing
					if ($c.IsComputed -or $dataType -in @('timestamp', 'rowversion')) { continue }

					# Type formatting
					$typeText = $dataType
					switch ($dataType)
					{
						'nvarchar' {
							if ($maxLen -eq -1) { $typeText = 'nvarchar(max)' }
							else { $typeText = "nvarchar($([int]($maxLen/2)))" }
						}
						'varchar'  {
							if ($maxLen -eq -1) { $typeText = 'varchar(max)' }
							else { $typeText = "varchar($maxLen)" }
						}
						'nchar'    { $typeText = "nchar($([int]($maxLen/2)))" }
						'char'     { $typeText = "char($maxLen)" }
						'varbinary' {
							if ($maxLen -eq -1) { $typeText = 'varbinary(max)' }
							else { $typeText = "varbinary($maxLen)" }
						}
						'binary'   { $typeText = "binary($maxLen)" }
						'decimal'  { $typeText = "decimal($prec,$scale)" }
						'numeric'  { $typeText = "numeric($prec,$scale)" }
						'datetime2' {
							$typeText = if ($scale -ge 0) { "datetime2($scale)" }
							else { 'datetime2' }
						}
						'time' {
							$typeText = if ($scale -ge 0) { "time($scale)" }
							else { 'time' }
						}
						'datetimeoffset' {
							$typeText = if ($scale -ge 0) { "datetimeoffset($scale)" }
							else { 'datetimeoffset' }
						}
						'float' {
							$typeText = if ($prec -and $prec -ne 53) { "float($prec)" }
							else { 'float' }
						}
					}

					# Identity
					$identityText = ''
					if ($isIdent)
					{
						$seed = $c.IdentitySeed
						$inc = $c.IdentityIncrement
						if ($seed -ne $null -and $seed -ne [DBNull]::Value -and $inc -ne $null -and $inc -ne [DBNull]::Value)
						{
							$identityText = " IDENTITY($seed,$inc)"
						}
						else
						{
							$identityText = " IDENTITY(1,1)"
						}
					}

					$nullText = if ($isNull) { 'NULL' }
					else { 'NOT NULL' }
					$colDefs += "[$colName] $typeText$identityText $nullText"
				}

				# Same primary key as the server, so later pumps can match rows
				if ($colDefs.Count -gt 0 -and $srcPk.Count -gt 0) { $colDefs += ("CONSTRAINT [PK_" + $table + "] PRIMARY KEY (" + (($srcPk | ForEach-Object { "[$_]" }) -join ', ') + ")") }
				if ($colDefs.Count -gt 0) { $createSql = "CREATE TABLE [dbo].[$table] (" + ($colDefs -join ', ') + ")" }
			}

			$tablePlans.Add([pscustomobject]@{
					Table     = $tableName
					Columns   = [string[]]@($metaRows | Where-Object { -not $_.IsComputed -and $_.DataType -notin @('timestamp', 'rowversion') } | ForEach-Object { $_.ColumnName })
					SourcePk  = [string[]]$srcPk.ToArray()
					CreateSql = $createSql
					Rows      = $tableRowCounts[$tableName]
				})
		}
		catch
		{
			Write_Log "Could not read table '$tableName' from the SOURCE - skipping it: $($_.Exception.Message)" "yellow"
			$tableRowCounts[$tableName] = $null
		}
	}
	if ($tablePlans.Count -eq 0)
	{
		Write_Log "None of the selected tables could be read from the source database." "red"
		try { $srcSqlConnection.Dispose() } catch { }
		Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 6: Sort lanes: SQL-reachable (TCP / Named Pipes) or file drop
	# ------------------------------------------------------------------------------------------------
	$sqlWork = New-Object System.Collections.Generic.List[object]
	$fileWork = New-Object System.Collections.Generic.List[object]
	foreach ($laneNum in $Lanes)
	{
		$laneNum = [string]$laneNum
		if ($laneNum -match '^\d+$') { $laneNum = $laneNum.PadLeft(3, '0') }
		$machineName = $LaneNumToMachineName[$laneNum]
		if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
		{
			Write_Log "Lane #${laneNum}: Machine name is invalid or unknown. Skipping." "yellow"
			continue
		}

		$protocolType = $null
		if ($script:LaneProtocols)
		{
			foreach ($k in @($laneNum, $machineName, $machineName.ToLower()))
			{
				if ($script:LaneProtocols.ContainsKey($k)) { $protocolType = $script:LaneProtocols[$k]; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($protocolType)) { $protocolType = 'File' }

		$laneConnStr = $null
		if ($protocolType -eq 'TCP' -or $protocolType -eq 'Named Pipes')
		{
			$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNum
			if ($laneInfo -and ($laneInfo -isnot [System.Collections.IDictionary]))
			{
				$tmp = @{ }
				foreach ($p in $laneInfo.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
				$laneInfo = $tmp
			}
			if ($laneInfo)
			{
				$laneConnStr = if ($protocolType -eq 'Named Pipes') { [string]$laneInfo['NamedPipesConnStr'] } else { [string]$laneInfo['TcpConnStr'] }
			}
			if ([string]::IsNullOrWhiteSpace($laneConnStr) -or $laneConnStr -notmatch '^\s*Server\s*=' -or $laneConnStr -notmatch 'Database\s*=')
			{
				Write_Log "Lane $laneNum ($machineName) protocol cached as '$protocolType' but connection string is invalid. Using file fallback." "yellow"
				$laneConnStr = $null
			}
		}

		if ($laneConnStr) { $sqlWork.Add([pscustomobject]@{ Lane = $laneNum; Machine = $machineName; Protocol = $protocolType; ConnStr = $laneConnStr }) }
		else { $fileWork.Add([pscustomobject]@{ Lane = $laneNum; Machine = $machineName; Tables = @($filteredAliasEntries) }) }
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 7: SQL lanes - all at once (up to 6 in parallel), each table UPSERTED in one transaction:
	#   - the lane's own table is kept (schema, and rows the server does not have)
	#   - server rows are bulk-copied into a #staging table shaped like the LANE's table
	#   - rows whose key exists are UPDATED (only when a value differs), missing rows are INSERTED
	#   - nothing is truncated or deleted; a failed table rolls back and goes to the file drop
	# ------------------------------------------------------------------------------------------------
	if ($sqlWork.Count -gt 0)
	{
		$laneUpsertWorker = @'
param ($SrcConnStr, $LaneConnStr, $Lane, $Machine, $Protocol, $Plans, $Log)
$say = { param ([string]$Text, [string]$Color) $Log.Enqueue([pscustomobject]@{ Text = $Text; Color = $Color }) }
$failed = New-Object System.Collections.Generic.List[string]
$src = $null
$dst = $null
try
{
	$src = New-Object System.Data.SqlClient.SqlConnection $SrcConnStr
	$src.Open()
	$dst = New-Object System.Data.SqlClient.SqlConnection $LaneConnStr
	$dst.Open()
}
catch
{
	& $say ("Lane {0} ({1}): SQL connection failed ({2}) - using the file drop. {3}" -f $Lane, $Machine, $Protocol, $_.Exception.Message) 'yellow'
	try { if ($src) { $src.Dispose() } } catch { }
	try { if ($dst) { $dst.Dispose() } } catch { }
	return [pscustomobject]@{ Lane = $Lane; Connected = $false; Failed = @($Plans | ForEach-Object { $_.Table }) }
}
& $say ("Lane {0} ({1}): connected over {2}, pumping {3} table(s)..." -f $Lane, $Machine, $Protocol, @($Plans).Count) 'blue'
$noCompare = @('text', 'ntext', 'image', 'xml', 'sql_variant', 'geography', 'geometry', 'hierarchyid')
foreach ($plan in $Plans)
{
	$table = $plan.Table
	$tx = $null
	try
	{
		$tx = $dst.BeginTransaction()
		$run = {
			param ([string]$Sql, [int]$Timeout = 120)
			$c = $dst.CreateCommand()
			$c.Transaction = $tx
			$c.CommandTimeout = $Timeout
			$c.CommandText = $Sql
			return $c
		}

		$objId = (& $run "SELECT OBJECT_ID(N'[dbo].[$table]', N'U')").ExecuteScalar()
		if ($null -eq $objId -or $objId -is [DBNull])
		{
			if (-not $plan.CreateSql) { throw "the lane has no $table table and it cannot be created from the source (computed columns)" }
			$null = (& $run $plan.CreateSql).ExecuteNonQuery()
			& $say ("Lane {0}: created table {1} (same columns as the server)" -f $Lane, $table) 'green'
		}

		# The LANE's columns decide what is written - its schema is never changed
		$destCols = @{ }
		$rd = (& $run "SELECT c.name, t.name, c.is_identity, c.is_computed FROM sys.columns c JOIN sys.types t ON c.user_type_id = t.user_type_id WHERE c.object_id = OBJECT_ID(N'[dbo].[$table]')").ExecuteReader()
		try { while ($rd.Read()) { $destCols[[string]$rd[0]] = [pscustomobject]@{ Type = [string]$rd[1]; Identity = [bool]$rd[2]; Skip = ([bool]$rd[3] -or ([string]$rd[1]) -in @('timestamp', 'rowversion')) } } }
		finally { $rd.Close() }
		$cols = @($plan.Columns | Where-Object { $destCols.ContainsKey($_) -and -not $destCols[$_].Skip })
		if ($cols.Count -eq 0) { throw "no columns in common with the server table" }

		# Match rows on the lane's primary key (the server's when the lane has none)
		$pk = New-Object System.Collections.Generic.List[string]
		$rd = (& $run "SELECT c.name FROM sys.indexes i JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id WHERE i.object_id = OBJECT_ID(N'[dbo].[$table]') AND i.is_primary_key = 1 ORDER BY ic.key_ordinal").ExecuteReader()
		try { while ($rd.Read()) { $pk.Add([string]$rd[0]) } }
		finally { $rd.Close() }
		if ($pk.Count -eq 0) { foreach ($k in @($plan.SourcePk)) { $pk.Add($k) } }
		if ($pk.Count -eq 0) { throw "no primary key on the lane or the server table - cannot match rows (it would add duplicates)" }
		$missingKey = @($pk | Where-Object { $cols -notcontains $_ })
		if ($missingKey.Count -gt 0) { throw ("key column(s) not copied: " + ($missingKey -join ', ')) }

		$colList = ($cols | ForEach-Object { "[$_]" }) -join ', '
		$sCols = ($cols | ForEach-Object { "s.[$_]" }) -join ', '
		$on = ($pk | ForEach-Object { "d.[$_] = s.[$_]" }) -join ' AND '

		# Staging table with the lane's column types (the UNION drops the IDENTITY property)
		$null = (& $run "IF OBJECT_ID('tempdb..#pump') IS NOT NULL DROP TABLE #pump; SELECT $colList INTO #pump FROM [dbo].[$table] WHERE 1 = 0 UNION ALL SELECT $colList FROM [dbo].[$table] WHERE 1 = 0;").ExecuteNonQuery()

		$srcCmd = $src.CreateCommand()
		$srcCmd.CommandTimeout = 0
		$srcCmd.CommandText = "SELECT $colList FROM [dbo].[$table]"
		$srcReader = $srcCmd.ExecuteReader()
		try
		{
			$bulk = New-Object System.Data.SqlClient.SqlBulkCopy($dst, [System.Data.SqlClient.SqlBulkCopyOptions]::Default, $tx)
			$bulk.DestinationTableName = '#pump'
			$bulk.BatchSize = 5000
			$bulk.BulkCopyTimeout = 0
			foreach ($c in $cols) { $null = $bulk.ColumnMappings.Add($c, $c) }
			$bulk.WriteToServer($srcReader)
			$bulk.Close()
		}
		finally { $srcReader.Close() }

		# UPDATE rows that exist and differ (EXCEPT compares NULLs correctly)
		$updated = 0
		$nonKey = @($cols | Where-Object { $pk -notcontains $_ -and -not $destCols[$_].Identity })
		if ($nonKey.Count -gt 0)
		{
			$set = ($nonKey | ForEach-Object { "d.[$_] = s.[$_]" }) -join ', '
			$diff = ''
			if (@($nonKey | Where-Object { $noCompare -contains $destCols[$_].Type }).Count -eq 0)
			{
				$diff = " WHERE EXISTS (SELECT " + (($nonKey | ForEach-Object { "s.[$_]" }) -join ', ') + " EXCEPT SELECT " + (($nonKey | ForEach-Object { "d.[$_]" }) -join ', ') + ")"
			}
			$updated = (& $run "UPDATE d SET $set FROM [dbo].[$table] d JOIN #pump s ON $on$diff;" 0).ExecuteNonQuery()
		}

		# INSERT rows the lane does not have
		$insert = "INSERT INTO [dbo].[$table] ($colList) SELECT $sCols FROM #pump s WHERE NOT EXISTS (SELECT 1 FROM [dbo].[$table] d WHERE $on);"
		if (@($cols | Where-Object { $destCols[$_].Identity }).Count -gt 0) { $insert = "SET IDENTITY_INSERT [dbo].[$table] ON; $insert SET IDENTITY_INSERT [dbo].[$table] OFF;" }
		$inserted = (& $run $insert 0).ExecuteNonQuery()

		$null = (& $run "DROP TABLE #pump;").ExecuteNonQuery()
		$tx.Commit()
		$tx = $null
		& $say ("Lane {0} ({1}): {2} - {3} updated, {4} added, {5} server rows checked (lane-only rows kept)" -f $Lane, $Machine, $table, $updated, $inserted, $plan.Rows) 'green'
	}
	catch
	{
		if ($tx) { try { $tx.Rollback() } catch { } }
		& $say ("Lane {0} ({1}): {2} NOT pumped over SQL, lane table unchanged - {3}" -f $Lane, $Machine, $table, $_.Exception.Message) 'red'
		$failed.Add($table)
	}
}
try { $src.Dispose() } catch { }
try { $dst.Dispose() } catch { }
return [pscustomobject]@{ Lane = $Lane; Connected = $true; Failed = $failed.ToArray() }
'@

		$logQueue = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
		$poolSize = [Math]::Min(6, [Math]::Max(1, $sqlWork.Count))
		$pool = [RunspaceFactory]::CreateRunspacePool(1, $poolSize)
		$pool.Open()
		$plansArray = $tablePlans.ToArray()
		$tasks = New-Object System.Collections.Generic.List[object]
		Write_Log "`r`nPumping $($plansArray.Count) table(s) to $($sqlWork.Count) lane(s) over SQL, up to $poolSize at a time..." "blue"
		foreach ($w in $sqlWork)
		{
			$ps = [PowerShell]::Create()
			$ps.RunspacePool = $pool
			$null = $ps.AddScript($laneUpsertWorker).AddArgument($ConnectionString).AddArgument($w.ConnStr).AddArgument($w.Lane).AddArgument($w.Machine).AddArgument($w.Protocol).AddArgument($plansArray).AddArgument($logQueue)
			$tasks.Add([pscustomobject]@{ Work = $w; PS = $ps; Handle = $ps.BeginInvoke(); Done = $false })
		}

		# Wait without freezing the window: show worker messages as they come
		$drain = {
			$m = $null
			while ($logQueue.TryDequeue([ref]$m)) { Write_Log $m.Text $m.Color }
		}
		while (@($tasks | Where-Object { -not $_.Done }).Count -gt 0)
		{
			foreach ($tk in $tasks)
			{
				if ($tk.Done -or -not $tk.Handle.IsCompleted) { continue }
				$tk.Done = $true
				$res = $null
				try { $res = @($tk.PS.EndInvoke($tk.Handle)) | Select-Object -Last 1 }
				catch { Write_Log "Lane $($tk.Work.Lane): pump worker failed - $($_.Exception.Message)" "red" }
				finally { $tk.PS.Dispose() }
				& $drain
				if (-not $res) { $res = [pscustomobject]@{ Lane = $tk.Work.Lane; Connected = $false; Failed = @($tablePlans | ForEach-Object { $_.Table }) } }
				$failedTables = @($res.Failed)
				if ($res.Connected -and $failedTables.Count -lt $tablePlans.Count) { $protocolLanes += $tk.Work.Lane; $ProcessedLanes += $tk.Work.Lane }
				if ($failedTables.Count -gt 0)
				{
					$fileWork.Add([pscustomobject]@{ Lane = $tk.Work.Lane; Machine = $tk.Work.Machine; Tables = @($filteredAliasEntries | Where-Object { $failedTables -contains [string]$_.Table }) })
				}
			}
			& $drain
			$script:LastActivity = Get-Date
			try { [System.Windows.Forms.Application]::DoEvents() } catch { }
			Start-Sleep -Milliseconds 100
		}
		& $drain
		try { $pool.Close(); $pool.Dispose() } catch { }
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 8: File drop (lanes without SQL access, and tables whose SQL pump failed)
	#   SMS loads the file with @UPDATE_BATCH(JOB=ADDRPL ...) - also update-or-add by key.
	#   RULE: File-only lanes must ONLY pump tables <= 1000 rows
	# ------------------------------------------------------------------------------------------------
	$MaxFileRows = 1000
	$generatedSqlFiles = @{ }
	foreach ($fw in $fileWork)
	{
		$laneNum = $fw.Lane
		$machineName = $fw.Machine
		$fileEntries = @($fw.Tables)
		$LaneLocalPath = Join-Path $OfficePath "XF${StoreNumber}${laneNum}"
		if ($fileEntries.Count -gt 0)
		{
			if (Test-Path $LaneLocalPath)
			{
				Write_Log "`r`nCopying via FILE fallback for Lane #$laneNum ($machineName)..." "blue"

				foreach ($aliasEntry in $fileEntries)
				{
					$table = [string]$aliasEntry.Table
					if ([string]::IsNullOrWhiteSpace($table)) { continue }

					$rowCount = $null
					if ($tableRowCounts.ContainsKey($table)) { $rowCount = $tableRowCounts[$table] }
					if ($null -eq $rowCount)
					{
						Write_Log "Skipping '$table' for FILE lane ${laneNum}: rowcount unknown." "yellow"
						continue
					}
					if ($rowCount -gt $MaxFileRows)
					{
						Write_Log "Skipping '$table' for FILE lane ${laneNum}: $rowCount rows (> $MaxFileRows rule)." "yellow"
						continue
					}

					$baseTable = $table -replace '_TAB$', ''
					$sqlFileName = "${baseTable}_Load.sql"

					if (-not $generatedSqlFiles.ContainsKey($table))
					{
						$localTempPath = Join-Path $env:TEMP ("${StoreNumber}_${sqlFileName}")

						try
						{
							$ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)
							$streamWriter = New-Object System.IO.StreamWriter($localTempPath, $false, $ansiPcEncoding)
							$streamWriter.NewLine = "`r`n"

							$columnDataTypesQuery = @"
SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '$table'
ORDER BY ORDINAL_POSITION
"@
							$cmdColumnTypes = $srcSqlConnection.CreateCommand()
							$cmdColumnTypes.CommandText = $columnDataTypesQuery
							$readerColumnTypes = $cmdColumnTypes.ExecuteReader()
							$columnDataTypes = [ordered]@{ }
							while ($readerColumnTypes.Read())
							{
								$colName = [string]$readerColumnTypes["COLUMN_NAME"]
								$dataType = [string]$readerColumnTypes["DATA_TYPE"]
								$columnDataTypes[$colName] = $dataType
							}
							$readerColumnTypes.Close()

							$filterCmd = $srcSqlConnection.CreateCommand()
							$filterCmd.CommandTimeout = 60
							$filterCmd.CommandText = @"
SELECT c.name AS ColumnName, t.name AS DataType, c.is_computed AS IsComputed
FROM sys.columns c
JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.object_id = OBJECT_ID(@obj)
"@
							$null = $filterCmd.Parameters.Clear()
							$p2 = $filterCmd.Parameters.Add("@obj", [System.Data.SqlDbType]::NVarChar, 256)
							$p2.Value = "dbo.$table"
							$filterReader = $filterCmd.ExecuteReader()
							$skipCols = @{ }
							while ($filterReader.Read())
							{
								$cn = [string]$filterReader['ColumnName']
								$dt = [string]$filterReader['DataType']
								$ic = [bool]$filterReader['IsComputed']
								if ($ic -or $dt -in @('timestamp', 'rowversion')) { $skipCols[$cn] = $true }
							}
							$filterReader.Close()

							foreach ($k in @($skipCols.Keys))
							{
								if ($columnDataTypes.Contains($k)) { $columnDataTypes.Remove($k) }
							}

							if ($columnDataTypes.Count -eq 0)
							{
								throw "No eligible columns to pump after filtering computed/rowversion."
							}

							$pkQuery = @"
SELECT c.COLUMN_NAME
FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE c
	ON c.CONSTRAINT_NAME = tc.CONSTRAINT_NAME
	AND c.TABLE_NAME = tc.TABLE_NAME
WHERE tc.TABLE_NAME = '$table' AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
ORDER BY c.ORDINAL_POSITION
"@
							$cmdPK = $srcSqlConnection.CreateCommand()
							$cmdPK.CommandText = $pkQuery
							$readerPK = $cmdPK.ExecuteReader()
							$pkColumns = @()
							while ($readerPK.Read()) { $pkColumns += [string]$readerPK["COLUMN_NAME"] }
							$readerPK.Close()

							if ($pkColumns.Count -eq 0)
							{
								throw "File fallback requires a real PRIMARY KEY on '$table'."
							}
							else
							{
								$primaryKeyColumns = $pkColumns
							}

							$keyString = ($primaryKeyColumns | ForEach-Object { "$_=:$_" }) -join " AND "
							$viewName = $baseTable.Substring(0, 1).ToUpper() + $baseTable.Substring(1).ToLower() + '_Load'
							$columnList = ($columnDataTypes.Keys) -join ','

							$header = @"
@WIZRPL(DBASE_TIMEOUT=E);

CREATE VIEW $viewName AS SELECT $columnList FROM $table;

INSERT INTO $viewName VALUES
"@
							$header = $header -replace "(`r`n|`n|`r)", "`r`n"
							$streamWriter.WriteLine($header.TrimEnd())

							$dataQuery = "SELECT $columnList FROM [dbo].[$table]"
							$cmdData = $srcSqlConnection.CreateCommand()
							$cmdData.CommandText = $dataQuery
							$readerData = $cmdData.ExecuteReader()
							$firstRow = $true
							while ($readerData.Read())
							{
								if ($firstRow) { $firstRow = $false }
								else { $streamWriter.WriteLine(",") }
								$values = @()
								foreach ($col in $columnDataTypes.Keys)
								{
									$val = $readerData[$col]
									$dataType = $columnDataTypes[$col]
									if ($val -eq $null -or $val -is [System.DBNull])
									{
										$values += ""
										continue
									}
									switch -Wildcard ($dataType)
									{
										{ $_ -in @('char', 'nchar', 'varchar', 'nvarchar', 'text', 'ntext') } {
											$escapedVal = $val.ToString().Replace("'", "''")
											$escapedVal = $escapedVal -replace "(`r`n|`n|`r)", " "
											$values += "'$escapedVal'"
											break
										}
										{ $_ -in @('datetime', 'smalldatetime', 'date', 'datetime2') } {
											$dayOfYear = $val.DayOfYear.ToString("D3")
											$formattedDate = "'{0}{1} {2}'" -f $val.Year, $dayOfYear, $val.ToString("HH:mm:ss")
											$values += $formattedDate
											break
										}
										{ $_ -eq 'bit' } {
											$bitVal = if ($val) { "1" }
											else { "0" }
											$values += $bitVal
											break
										}
										{ $_ -in @('decimal', 'numeric', 'float', 'real', 'money', 'smallmoney') } {
											$values += ([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0}", $val))
											break
										}
										{ $_ -in @('tinyint', 'smallint', 'int', 'bigint') } {
											$values += $val.ToString()
											break
										}
										default {
											$escapedVal = $val.ToString().Replace("'", "''")
											$escapedVal = $escapedVal -replace "(`r`n|`n|`r)", " "
											$values += "'$escapedVal'"
											break
										}
									}
								}
								$insertStatement = "(" + ($values -join ",") + ")"
								$insertStatement = $insertStatement -replace "(`r`n|`n|`r)", " "
								$streamWriter.Write($insertStatement)
							}
							$readerData.Close()
							$streamWriter.WriteLine(";")
							$streamWriter.WriteLine()

							$footer = @"
@UPDATE_BATCH(JOB=ADDRPL,TAR=$table,
KEY=$keyString,
SRC=SELECT * FROM $viewName);

DROP VIEW $viewName;

@EXEC(INI=HOST_STORE[ACTIVATE_ACCEPT_ALL]);

@WIZCLR(DBASE_TIMEOUT);
"@
							$footer = $footer -replace "(`r`n|`n|`r)", "`r`n"
							$streamWriter.WriteLine($footer.TrimEnd())
							$streamWriter.WriteLine()
							$streamWriter.Flush()
							$streamWriter.Close()
							$streamWriter.Dispose()

							$generatedSqlFiles[$table] = $localTempPath
							Write_Log "Generated $sqlFileName (FILE mode) from SOURCE ($rowCount rows)." "green"
						}
						catch
						{
							try { if ($streamWriter) { $streamWriter.Close(); $streamWriter.Dispose() } }
							catch { }
							Write_Log "Error generating SQL for table '$table' (file fallback): $_" "red"
							continue
						}
					}

					try
					{
						$localTempPath = $generatedSqlFiles[$table]
						if (-not (Test-Path $localTempPath))
						{
							Write_Log "Temp SQL file missing for '$table' ($localTempPath). Skipping copy." "yellow"
							continue
						}

						$destinationPath = Join-Path $LaneLocalPath $sqlFileName
						Copy-Item -Path $localTempPath -Destination $destinationPath -Force -ErrorAction Stop

						try
						{
							$fileItem = Get-Item $destinationPath -ErrorAction Stop
							if ($fileItem.Attributes -band [System.IO.FileAttributes]::Archive)
							{
								$fileItem.Attributes -= [System.IO.FileAttributes]::Archive
							}
						}
						catch { }

						Write_Log "Copied $sqlFileName to Lane #$laneNum ($machineName) (file fallback)." "green"
						$fileCopyLanes += $laneNum
						$ProcessedLanes += $laneNum
					}
					catch
					{
						Write_Log "Error copying $sqlFileName to Lane #[$laneNum] ($machineName): $_" "red"
					}
				}
			}
			else
			{
				Write_Log "Lane #$laneNum ($machineName) not found at path: $LaneLocalPath (file fallback failed)" "yellow"
			}
		}
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 9: Clean up and final logging
	# ------------------------------------------------------------------------------------------------
	try { if ($srcSqlConnection) { $srcSqlConnection.Close() } }
	catch { }
	try { if ($srcSqlConnection) { $srcSqlConnection.Dispose() } }
	catch { }

	$uniqueProcessedLanes = @($ProcessedLanes | Select-Object -Unique)
	Write_Log "`r`nTotal Lanes processed: $($uniqueProcessedLanes.Count)" "green"
	if ($protocolLanes.Count -gt 0)
	{
		Write_Log "Lanes processed via SQL protocol: $((($protocolLanes | Select-Object -Unique) -join ', '))" "green"
	}
	if ($fileCopyLanes.Count -gt 0)
	{
		Write_Log "Lanes processed via FILE fallback: $((($fileCopyLanes | Select-Object -Unique) -join ', '))" "yellow"
	}
	Write_Log "`r`n==================== Pump_Tables Function Completed ====================" "blue"
}
