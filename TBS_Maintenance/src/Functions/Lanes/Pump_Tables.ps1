# ===================================================================================================
#                              FUNCTION: Pump_Tables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Extracts selected tables (from Get_Table_Aliases) from the store SQL Server and copies
#   the generated .sql files to the selected lanes. Uses the cached per-lane protocol
#   ($script:LaneProtocols) to push lanes over SQL where possible and falls back to a file
#   copy for the rest. Each selected table is generated once, then copied to many lanes.
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
	$Lanes = $Lanes | Where-Object { $LaneNumToMachineName.ContainsKey($_) } | Select-Object -Unique

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
	# FIX: Pre-calc row counts ONCE (used for file-lane 1000-row rule + better logging)
	# ------------------------------------------------------------------------------------------------
	$tableRowCounts = @{ }
	foreach ($aliasEntry in $filteredAliasEntries)
	{
		$tableName = [string]$aliasEntry.Table
		if ([string]::IsNullOrWhiteSpace($tableName)) { continue }
		if ($tableRowCounts.ContainsKey($tableName)) { continue }

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
			$cntCmd.CommandText = "SELECT COUNT(1) FROM [dbo].[$tableName]"
			$tableRowCounts[$tableName] = [int]$cntCmd.ExecuteScalar()
		}
		catch
		{
			Write_Log "Could not read rowcount for table '$tableName' from SOURCE. Skipping table. Error: $_" "yellow"
			$tableRowCounts[$tableName] = $null
		}
	}

	# ------------------------------------------------------------------------------------------------
	# STEP 5: Process each lane: Try protocol (Named Pipes or TCP), else fall back to file copy
	# RULE: File-only lanes must ONLY pump tables <= 1000 rows
	# ------------------------------------------------------------------------------------------------
	$MaxFileRows = 1000

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

		# Retrieve DB connection info for the lane
		$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNum
		if (-not $laneInfo)
		{
			Write_Log "Could not get DB info for lane $laneNum. Skipping." "yellow"
			continue
		}

		# FIX: Convert PSCustomObject -> hashtable for consistent indexing (matches your Process_Lanes fix)
		if ($laneInfo -and ($laneInfo -isnot [System.Collections.IDictionary]))
		{
			$tmp = @{ }
			foreach ($p in $laneInfo.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
			$laneInfo = $tmp
		}

		$namedPipesConnStr = [string]$laneInfo['NamedPipesConnStr']
		$tcpConnStr = [string]$laneInfo['TcpConnStr']

		# Always use numeric lane number for folder path
		$LaneLocalPath = Join-Path $OfficePath "XF${StoreNumber}${laneNum}"

		# FIX: Robust protocol lookup (Start_Lane_Protocol_Jobs stores several keys)
		$protocolType = $null
		if ($script:LaneProtocols)
		{
			if ($script:LaneProtocols.ContainsKey($laneNum)) { $protocolType = $script:LaneProtocols[$laneNum] }
			elseif ($script:LaneProtocols.ContainsKey($laneNum.PadLeft(3, '0'))) { $protocolType = $script:LaneProtocols[$laneNum.PadLeft(3, '0')] }
			elseif ($script:LaneProtocols.ContainsKey($machineName)) { $protocolType = $script:LaneProtocols[$machineName] }
			elseif ($script:LaneProtocols.ContainsKey($machineName.ToLower())) { $protocolType = $script:LaneProtocols[$machineName.ToLower()] }
		}
		if ([string]::IsNullOrWhiteSpace($protocolType)) { $protocolType = 'File' }

		$laneSqlConn = $null
		$protocolWorked = $false

		# ----------------------------------------------------------------------------------------
		# Try direct SQL protocol copy (Named Pipes/TCP) using FAST SqlBulkCopy
		# ----------------------------------------------------------------------------------------
		if ($protocolType -eq "Named Pipes" -or $protocolType -eq "TCP")
		{
			$workingConnStr = $null
			if ($protocolType -eq 'Named Pipes') { $workingConnStr = $namedPipesConnStr }
			elseif ($protocolType -eq 'TCP') { $workingConnStr = $tcpConnStr }

			$workingConnStrIsValid =
			(-not [string]::IsNullOrWhiteSpace($workingConnStr)) -and
			($workingConnStr -match '^\s*Server\s*=') -and
			($workingConnStr -match 'Database\s*=')

			if ($workingConnStrIsValid)
			{
				try
				{
					$laneSqlConn = New-Object System.Data.SqlClient.SqlConnection $workingConnStr
					$laneSqlConn.Open()
					if ($laneSqlConn.State -eq 'Open')
					{
						Write_Log "`r`nCopying data to Lane $laneNum ($machineName) via SQL protocol [$protocolType]..." "blue"

						foreach ($aliasEntry in $filteredAliasEntries)
						{
							$table = [string]$aliasEntry.Table
							if ([string]::IsNullOrWhiteSpace($table)) { continue }

							# Skip tables that we could not count / validate
							if (-not $tableRowCounts.ContainsKey($table) -or $null -eq $tableRowCounts[$table])
							{
								Write_Log "Skipping table '$table' (rowcount/validation failed on SOURCE)." "yellow"
								continue
							}

							Write_Log "Pumping table '$table' to lane $laneNum ($machineName) via SQL (SqlBulkCopy)..." "blue"

							try
							{
								# -----------------------------
								# 1) Get SOURCE column metadata (sys.columns) so we can:
								#    - create table if missing
								#    - identify identity/computed/rowversion columns
								# -----------------------------
								$metaCmd = $srcSqlConnection.CreateCommand()
								$metaCmd.CommandTimeout = 60
								$metaCmd.CommandText = @"
SELECT
	c.name AS ColumnName,
	t.name AS DataType,
	c.max_length AS MaxLength,
	c.precision AS [Precision],
	c.scale AS [Scale],
	c.is_nullable AS IsNullable,
	c.is_identity AS IsIdentity,
	c.is_computed AS IsComputed,
	ic.seed_value AS IdentitySeed,
	ic.increment_value AS IdentityIncrement
FROM sys.columns c
JOIN sys.types t
	ON c.user_type_id = t.user_type_id
LEFT JOIN sys.identity_columns ic
	ON c.object_id = ic.object_id
	AND c.column_id = ic.column_id
WHERE c.object_id = OBJECT_ID(@obj)
ORDER BY c.column_id;
"@
								$null = $metaCmd.Parameters.Clear()
								$param = $metaCmd.Parameters.Add("@obj", [System.Data.SqlDbType]::NVarChar, 256)
								$param.Value = "dbo.$table"
								$metaRows = @()
								$metaReader = $metaCmd.ExecuteReader()
								while ($metaReader.Read())
								{
									$metaRows += [pscustomobject]@{
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
									}
								}
								$metaReader.Close()

								if (-not $metaRows -or $metaRows.Count -eq 0)
								{
									throw "Source table metadata not found (OBJECT_ID failed)."
								}

								# Identify special columns
								$computedCols = @($metaRows | Where-Object { $_.IsComputed })
								$identityCols = @($metaRows | Where-Object { $_.IsIdentity })

								# Copy columns: exclude computed + rowversion/timestamp (SQL Server won't allow explicit insert)
								$copyCols = @($metaRows | Where-Object { -not $_.IsComputed -and $_.DataType -notin @('timestamp', 'rowversion') })

								# -----------------------------
								# 2) Ensure destination table exists and is empty
								#    - If exists: TRUNCATE (fallback to DELETE)
								#    - If missing: CREATE from source columns (computed columns cannot be created here -> fail)
								# -----------------------------
								$existsCmd = $laneSqlConn.CreateCommand()
								$existsCmd.CommandTimeout = 30
								$existsCmd.CommandText = "SELECT OBJECT_ID(N'[dbo].[$table]', N'U')"
								$destObjId = $existsCmd.ExecuteScalar()
								$destExists = ($destObjId -ne $null -and $destObjId -ne [DBNull]::Value)

								if ($destExists)
								{
									# Empty the table
									try
									{
										$emptyCmd = $laneSqlConn.CreateCommand()
										$emptyCmd.CommandTimeout = 60
										$emptyCmd.CommandText = "TRUNCATE TABLE [dbo].[$table]"
										$emptyCmd.ExecuteNonQuery() | Out-Null
									}
									catch
									{
										# TRUNCATE can fail (FKs). Fallback to DELETE.
										$emptyCmd2 = $laneSqlConn.CreateCommand()
										$emptyCmd2.CommandTimeout = 120
										$emptyCmd2.CommandText = "DELETE FROM [dbo].[$table]"
										$emptyCmd2.ExecuteNonQuery() | Out-Null
									}
								}
								else
								{
									# Cannot recreate computed columns without expressions; fail fast if source has them
									if ($computedCols.Count -gt 0)
									{
										throw "Destination table missing and source has computed columns. Cannot safely CREATE."
									}

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

									if (-not $colDefs -or $colDefs.Count -eq 0)
									{
										throw "No creatable columns discovered for '$table'."
									}

									$createSql = "CREATE TABLE [dbo].[$table] (" + ($colDefs -join ', ') + ")"
									$createCmd = $laneSqlConn.CreateCommand()
									$createCmd.CommandTimeout = 120
									$createCmd.CommandText = "IF OBJECT_ID(N'[dbo].[$table]', N'U') IS NOT NULL DROP TABLE [dbo].[$table]; $createSql;"
									$createCmd.ExecuteNonQuery() | Out-Null
									Write_Log "Created table '$table' on $machineName" "green"
								}

								# -----------------------------
								# 3) Bulk copy data
								# -----------------------------
								$identityInsertEnabled = $false
								if ($identityCols.Count -gt 0)
								{
									try
									{
										$identOn = $laneSqlConn.CreateCommand()
										$identOn.CommandTimeout = 30
										$identOn.CommandText = "SET IDENTITY_INSERT [dbo].[$table] ON"
										$identOn.ExecuteNonQuery() | Out-Null
										$identityInsertEnabled = $true
									}
									catch
									{
										Write_Log "Could not enable IDENTITY_INSERT for '$table' on $machineName. Will copy without identity columns. Error: $_" "yellow"
										$copyCols = @($copyCols | Where-Object { -not $_.IsIdentity })
									}
								}

								if (-not $copyCols -or $copyCols.Count -eq 0)
								{
									throw "No columns available to copy (computed/rowversion filtered everything)."
								}

								$colListSql = ($copyCols | ForEach-Object { "[" + $_.ColumnName + "]" }) -join ','
								$srcCmd = $srcSqlConnection.CreateCommand()
								$srcCmd.CommandTimeout = 0
								$srcCmd.CommandText = "SELECT $colListSql FROM [dbo].[$table]"
								$srcReader = $srcCmd.ExecuteReader()

								try
								{
									$bulkOptions = [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock
									if ($identityInsertEnabled) { $bulkOptions = $bulkOptions -bor [System.Data.SqlClient.SqlBulkCopyOptions]::KeepIdentity }

									$bulk = New-Object System.Data.SqlClient.SqlBulkCopy($laneSqlConn, $bulkOptions, $null)
									$bulk.DestinationTableName = "[dbo].[$table]"
									$bulk.BatchSize = 5000
									$bulk.BulkCopyTimeout = 0

									foreach ($c in $copyCols)
									{
										$bulk.ColumnMappings.Add($c.ColumnName, $c.ColumnName) | Out-Null
									}

									$bulk.WriteToServer($srcReader)
								}
								finally
								{
									try { $srcReader.Close() }
									catch { }
									try { $bulk.Close() }
									catch { }
								}

								if ($identityInsertEnabled)
								{
									try
									{
										$identOff = $laneSqlConn.CreateCommand()
										$identOff.CommandTimeout = 30
										$identOff.CommandText = "SET IDENTITY_INSERT [dbo].[$table] OFF"
										$identOff.ExecuteNonQuery() | Out-Null
									}
									catch { }
								}

								$rowCountCopied = $tableRowCounts[$table]
								Write_Log "Copied $rowCountCopied rows to '$table' on lane $laneNum ($machineName) (SQL protocol)." "green"
							}
							catch
							{
								Write_Log "Failed to copy table '$table' to lane $laneNum ($machineName) via SQL: $_" "red"
							}
						}

						$protocolWorked = $true
						$protocolLanes += $laneNum
						$ProcessedLanes += $laneNum
					}
				}
				catch
				{
					Write_Log "SQL protocol copy failed for lane [$laneNum] ($machineName) ($protocolType): $_" "yellow"
					$protocolWorked = $false
				}
				finally
				{
					try { if ($laneSqlConn) { $laneSqlConn.Close() } }
					catch { }
					try { if ($laneSqlConn) { $laneSqlConn.Dispose() } }
					catch { }
				}
			}
			else
			{
				Write_Log "Lane $laneNum ($machineName) protocol cached as '$protocolType' but connection string is invalid. Using file fallback." "yellow"
			}
		}

		# --------------------------------------------------------------------------------------------
		# File fallback (if protocol copy not possible)
		# RULE: File-only lanes must ONLY pump tables <= 1000 rows
		# --------------------------------------------------------------------------------------------
		if (-not $protocolWorked)
		{
			if (Test-Path $LaneLocalPath)
			{
				Write_Log "`r`nCopying via FILE fallback for Lane #$laneNum ($machineName)..." "blue"

				foreach ($aliasEntry in $filteredAliasEntries)
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
							$header = $header -replace "(
|
|
)", "`r`n"
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
											$escapedVal = $escapedVal -replace "(
|
|
)", " "
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
											$escapedVal = $escapedVal -replace "(
|
|
)", " "
											$values += "'$escapedVal'"
											break
										}
									}
								}
								$insertStatement = "(" + ($values -join ",") + ")"
								$insertStatement = $insertStatement -replace "(
|
|
)", " "
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
							$footer = $footer -replace "(
|
|
)", "`r`n"
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
	# STEP 6: Clean up and final logging
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
