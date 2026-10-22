# ===================================================================================================
#                              FUNCTION: Update_Lane_Config
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates the SQL load files for the selected lanes of a store: for each XF lane folder it
#   copies run_load.sql and generates lnk_load.sql, sto_load.sql and ter_load.sql, strictly
#   in ANSI (Windows-1252) encoding with CRLF line endings and no BOM.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $LoadPath: Office\Load directory with the source SQL files
#   - $script:FunctionResults['LaneMachineNames']: all-lane list when "All" is selected
#   - $script:ProcessedLanes: lanes processed so far
#
# Writes:
#   - run_load.sql / lnk_load.sql / sto_load.sql / ter_load.sql in each selected XF lane folder
# ===================================================================================================

function Update_Lane_Config
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Update_Lane_Config Function ====================" "blue"
	
	if (-not (Test-Path $LoadPath))
	{
		Write_Log "`r`nLoad Base Path not found: $LoadPath" "yellow"
		return
	}
	
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($selection -eq $null)
	{
		Write_Log "`r`nLane processing canceled by user." "yellow"
		return
	}
	
	$Type = $selection.Type
	$Lanes = $selection.Lanes
	$processAllLanes = $false
	if ($Type -eq "All") { $processAllLanes = $true }
	
	if ($processAllLanes)
	{
		try
		{
			$LaneMachineNames = $script:FunctionResults['LaneMachineNames']
			if ($LaneMachineNames -and $LaneMachineNames.Count -gt 0)
			{
				$Lanes = $LaneMachineNames
			}
			else
			{
				throw "LaneMachineNames is empty or not available."
			}
		}
		catch
		{
			Write_Log "Failed to retrieve LaneMachineNames: $_. Falling back to user-selected lanes." "yellow"
			$processAllLanes = $false
		}
	}
	
	$ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)
	$runLoadFilename = "run_load.sql"
	$lnkLoadFilename = "lnk_load.sql"
	$stoLoadFilename = "sto_load.sql"
	$terLoadFilename = "ter_load.sql"
	
	# Original SQL data (unchanged for file copy)
	$runLoadScript = @"
@CREATE(RUN_TAB,RUN);
CREATE VIEW Run_Load AS SELECT F1102,F1000,F1103,F1104,F1105,F1106,F1107,F1108,F1109,F1110,F1111,F1112,F1113,F1114,F1115,F1116,F1117 FROM RUN_TAB;

INSERT INTO Run_Load VALUES
(60,'@TER','sql=BACKUP_ALL','@TER','@DSSF',,'@DSS+001',1,'Backup programs',,0,,'',1,,,),
(62,'@TER','sql=BACKUP_DEVICE','@TER',,,,1,'Backup programs on device',,0,,'',1,,,),
(68,'@TER','sqi=BACKUP_RESELLER','@TER',,,,1,'Backup for reseller',,0,,'',1,,,),
(70,'@TER','sql=DBASE_MAINTENANCE','@TER',,,,1,'Database maintenance',,0,,'',1,,,),
(90,'@TER','sql=PURGE_ALL','@TER','@DSW-008',,'@DSW-001',0,'Purge all report data',,0,,'',7,,,),
(95,'@TER','sqi=PURGE_CLIENT','@TER',,,,0,'Purge customer sensitive data',,0,,'',1,,,),
(120,'@TER','sql=UPDATE_DOWNLOAD','@TER',,,,1,'Automatic update download',,0,,'',1,,,),
(125,'@TER','sql=UPDATE_AUTOMATIC','@TER',,,,1,'Automatic update execution',,0,,'',1,,,),
(250,'SMS','sqi=TRS_POS_BANK_EOS_PUP OUTPUT=RECEIPT','901',,,,1,'Automatic bank close',,,,,1,,,);

@UPDATE_BATCH(JOB=ADDRPL,TAR=RUN_TAB,
KEY=F1102=:F1102 AND F1000=:F1000,
SRC=SELECT * FROM Run_Load);

DROP TABLE Run_Load;
"@
	$lnkLoadHeader = @"
@CREATE(LNK_TAB,LNK);
CREATE VIEW Lnk_Load AS SELECT F1000,F1056,F1057 FROM LNK_TAB;

INSERT INTO Lnk_Load VALUES
"@
	$lnkLoadFooter = @"

@UPDATE_BATCH(JOB=ADDRPL,TAR=LNK_TAB,
KEY=F1000=:F1000 AND F1056=:F1056 AND F1057=:F1057,
SRC=SELECT * FROM Lnk_Load);

DROP TABLE Lnk_Load;
"@
	$stoLoadHeader = @"
@CREATE(STO_TAB,STO);
CREATE VIEW Sto_Load AS SELECT F1000,F1018,F1180,F1181,F1182,F1937,F1965,F1966,F2691 FROM STO_TAB;

INSERT INTO Sto_Load VALUES
"@
	$stoLoadFooter = @"

@UPDATE_BATCH(JOB=ADDRPL,TAR=STO_TAB,
KEY=F1000=:F1000,
SRC=SELECT * FROM Sto_Load);

DROP TABLE Sto_Load;
"@
	$terLoadHeader = @"
@CREATE(TER_TAB,TER); 
CREATE VIEW Ter_Load AS SELECT F1056,F1057,F1058,F1125,F1169 FROM TER_TAB;

INSERT INTO Ter_Load VALUES
"@
	$terLoadFooter = @"

@UPDATE_BATCH(JOB=ADDRPL,TAR=TER_TAB,
KEY=F1056=:F1056 AND F1057=:F1057,
SRC=SELECT * FROM Ter_Load);

DROP TABLE Ter_Load;
"@
	
	foreach ($laneNumber in $Lanes)
	{
		$protocolWorked = @{ }
		$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNumber
		$MachineName = $null
		$namedPipesConnStr = $null
		if ($laneInfo)
		{
			$MachineName = $laneInfo['MachineName']
			$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
			Write_Log "Lane #${laneNumber}: Connection info found. Machine: $MachineName" "green"
		}
		else
		{
			$MachineName = $script:FunctionResults['LaneNumToMachineName'][$laneNumber]
			if (-not $MachineName) { $MachineName = "POS${laneNumber}" }
			Write_Log "Lane #${laneNumber}: Fallback machine name '$MachineName'" "yellow"
		}
		
		$laneFolderName = "XF${StoreNumber}${laneNumber}"
		$laneFolderPath = Join-Path -Path $OfficePath -ChildPath $laneFolderName
		if (-not (Test-Path $laneFolderPath))
		{
			Write_Log "`r`nLane #$laneNumber not found at path: $laneFolderPath" "yellow"
			continue
		}
		$laneFolder = Get-Item -Path $laneFolderPath
		Write_Log "`r`nProcessing Lane #$laneNumber" "blue"
		$actionSummaries = @()
		$fileCopyNeeded = @()
		
		# --- FIX: Parse run_load rows ahead of time ---
		$runRows =
		$runLoadScript -split "INSERT INTO Run_Load VALUES", 2 |
		Select-Object -Last 1 |
		ForEach-Object { ($_ -replace "(?ms);.*", "") -split "`r?`n" } |
		Where-Object { $_ -match "^\s*\(" }
		
		if ($namedPipesConnStr -and (Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue))
		{
			foreach ($tableJob in @(
					@{
						Table    = 'RUN_TAB'
						Filename = $runLoadFilename
						Rows	 = $runRows
					}
					@{
						Table																																																										     = 'LNK_TAB'
						Filename																																																										 = $lnkLoadFilename
						Rows																																																											 = @(
							"('${laneNumber}','${StoreNumber}','${laneNumber}')",
							"('DSM','${StoreNumber}','${laneNumber}')",
							"('PAL','${StoreNumber}','${laneNumber}')",
							"('RAL','${StoreNumber}','${laneNumber}')",
							"('XAL','${StoreNumber}','${laneNumber}')"
						)
					}
					@{
						Table																																																		   = 'STO_TAB'
						Filename																																																	   = $stoLoadFilename
						Rows																																																		   = @(
							"('${laneNumber}','Terminal ${laneNumber}',1,1,1,,,,)",
							"('DSM','Deploy SMS',1,1,1,,,,)",
							"('PAL','Program all',0,0,1,1,,,)",
							"('RAL','Report all',1,0,0,,,,)",
							"('XAL','Exchange all',0,1,0,,,,)"
						)
					}
					@{
						Table																																																				     = 'TER_TAB'
						Filename																																																				 = $terLoadFilename
						Rows																																																					 = @(
							"('${StoreNumber}','${laneNumber}','Terminal ${laneNumber}','\\${MachineName}\storeman\office\XF${StoreNumber}${laneNumber}\','\\${MachineName}\storeman\office\XF${StoreNumber}901\')",
							"('${StoreNumber}','901','Server','','')"
						)
					}
				))
			{
				$table = $tableJob.Table
				$rows = $tableJob.Rows
				
				try
				{
					$srcConn = $script:FunctionResults['ConnectionString']
					$srcSqlConn = New-Object System.Data.SqlClient.SqlConnection($srcConn)
					$srcSqlConn.Open()
					$cmd = $srcSqlConn.CreateCommand()
					$cmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '$table'
ORDER BY ORDINAL_POSITION
"@
					$rdr = $cmd.ExecuteReader()
					$colDefs = @()
					while ($rdr.Read())
					{
						$col = $rdr["COLUMN_NAME"]
						$type = $rdr["DATA_TYPE"]
						$nullText = if ($rdr["IS_NULLABLE"] -eq "NO") { "NOT NULL" }
						else { "NULL" }
						switch ($type)
						{
							'nvarchar' { $typeText = "nvarchar($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'varchar'  { $typeText = "varchar($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'char'     { $typeText = "char($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'nchar'    { $typeText = "nchar($($rdr["CHARACTER_MAXIMUM_LENGTH"]))" }
							'decimal'  { $typeText = "decimal($($rdr["NUMERIC_PRECISION"]),$($rdr["NUMERIC_SCALE"]))" }
							'numeric'  { $typeText = "numeric($($rdr["NUMERIC_PRECISION"]),$($rdr["NUMERIC_SCALE"]))" }
							default    { $typeText = $type }
						}
						$colDefs += "[$col] $typeText $nullText"
					}
					$rdr.Close()
					$srcSqlConn.Close()
					$colDefsText = $colDefs -join ", "
					$dropCreate = "IF OBJECT_ID(N'[$table]', N'U') IS NOT NULL DROP TABLE [$table]; CREATE TABLE [$table] ($colDefsText);"
					Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $dropCreate -QueryTimeout 30 -ErrorAction Stop
					Write_Log "Dropped and recreated $table on lane $laneNumber via protocol." "green"
					
					# Get column names for the insert
					$schema = Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '$table' ORDER BY ORDINAL_POSITION" -QueryTimeout 30
					$colNames = $schema | ForEach-Object { $_.COLUMN_NAME }
					
					$successfulRows = 0
					foreach ($row in $rows)
					{
						$cleanRow = $row.Trim().TrimEnd(',', ';')
						if ($cleanRow.StartsWith("(") -and $cleanRow.EndsWith(")"))
						{
							$cleanRow = $cleanRow.Substring(1, $cleanRow.Length - 2)
						}
						$splitRegex = ',(?=(?:[^'']*''[^'']*'')*[^'']*$)'
						$values = [regex]::Split($cleanRow, $splitRegex) | ForEach-Object {
							$v = $_.Trim()
							if ($v -eq "" -or $v -eq "' '") { "NULL" }
							else
							{
								if ($v.StartsWith("'") -and $v.EndsWith("'"))
								{
									$unquoted = $v.Trim("'")
									$res = $unquoted
									$res = $res -replace "@STORE", $StoreNumber.PadLeft(3, "0")
									$res = $res -replace "@TER", $laneNumber.PadLeft(3, "0")
									$res = $res -replace "@USER", $env:USERNAME
									$res = $res -replace "@USERNAME", $env:USERNAME
									$res = $res -replace "@RUN", $script:BasePath
									$res = $res -replace "@OFFICE", $script:OfficePath
									$now = Get-Date
									$res = $res -replace "@TIME", $now.ToString("HHmm")
									$res = $res -replace "@NOW", $now.ToString("HHmmss")
									$res = $res -replace "@DSSF", $now.ToString("yyyyMMdd")
									$res = $res -replace "@DSW-001", $now.AddDays(-1).ToString("yyyyMMdd")
									$res = $res -replace "@DSW-008", $now.AddDays(-8).ToString("yyyyMMdd")
									$res = $res -replace "@DSS\+001", $now.AddDays(1).ToString("yyyyMMdd")
									$res = $res -replace "@DSS", $now.ToString("yyyyMMdd")
									"'" + $res + "'"
								}
								else { $v }
							}
						}
						# Fill with NULLs or trim
						if ($values.Count -lt $colNames.Count)
						{
							$values += @("NULL") * ($colNames.Count - $values.Count)
						}
						elseif ($values.Count -gt $colNames.Count)
						{
							$values = $values[0 .. ($colNames.Count - 1)]
						}
						$sql = "INSERT INTO $table ([{0}]) VALUES ({1})" -f ($colNames -join "],["), ($values -join ", ")
						try
						{
							Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $sql -QueryTimeout 30 -ErrorAction Stop
							$successfulRows++
						}
						catch
						{
							Write_Log "$table row insert failed: $sql; Error: $_" "red"
						}
					}
					if ($successfulRows -gt 0)
					{
						Write_Log "Inserted $successfulRows rows into $table on lane $laneNumber via Named Pipes protocol." "green"
						$protocolWorked[$table] = $true
						$actionSummaries += "Protocol loaded $($tableJob.Filename)"
					}
					else
					{
						Write_Log "No data inserted to $table on lane $laneNumber." "yellow"
						$fileCopyNeeded += $tableJob.Filename
						$protocolWorked[$table] = $false
					}
				}
				catch
				{
					Write_Log "Protocol copy failed for $table on lane [$laneNumber]: $_" "red"
					$fileCopyNeeded += $tableJob.Filename
					$protocolWorked[$table] = $false
				}
			}
		}
		else
		{
			# No protocol possible at all, mark all as needing fallback
			$fileCopyNeeded += $runLoadFilename, $lnkLoadFilename, $stoLoadFilename, $terLoadFilename
		}
		
		# --- Fallback file copy for only failed tables (preserving original format) ---
		$tableFileData = @{
			$runLoadFilename = $runLoadScript
			$lnkLoadFilename = $lnkLoadHeader + "`r`n" +
			"('${laneNumber}','${StoreNumber}','${laneNumber}'),`r`n" +
			"('DSM','${StoreNumber}','${laneNumber}'),`r`n" +
			"('PAL','${StoreNumber}','${laneNumber}'),`r`n" +
			"('RAL','${StoreNumber}','${laneNumber}'),`r`n" +
			"('XAL','${StoreNumber}','${laneNumber}');`r`n" +
			"`r`n" + $lnkLoadFooter.TrimStart() + "`r`n"
			$stoLoadFilename = $stoLoadHeader + "`r`n" +
			"('${laneNumber}','Terminal ${laneNumber}',1,1,1,,,,),`r`n" +
			"('DSM','Deploy SMS',1,1,1,,,,),`r`n" +
			"('PAL','Program all',0,0,1,1,,,),`r`n" +
			"('RAL','Report all',1,0,0,,,,),`r`n" +
			"('XAL','Exchange all',0,1,0,,,,);`r`n" +
			"`r`n" + $stoLoadFooter.TrimStart() + "`r`n"
			$terLoadFilename = $terLoadHeader + "`r`n" +
			"('${StoreNumber}','${laneNumber}','Terminal ${laneNumber}','\\${MachineName}\storeman\office\XF${StoreNumber}${laneNumber}\','\\${MachineName}\storeman\office\XF${StoreNumber}901\'),`r`n" +
			"('${StoreNumber}','901','Server','','');`r`n" +
			"`r`n" + $terLoadFooter.TrimStart() + "`r`n"
		}
		foreach ($file in $fileCopyNeeded | Select-Object -Unique)
		{
			try
			{
				$dest = Join-Path -Path $laneFolder.FullName -ChildPath $file
				[System.IO.File]::WriteAllText($dest, $tableFileData[$file], $ansiPcEncoding)
				Set-ItemProperty -Path $dest -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
				$actionSummaries += "Copied $file"
			}
			catch
			{
				$actionSummaries += "Failed to copy $file"
			}
		}
		
		$summaryMessage = "Lane ${laneNumber} (Machine: ${MachineName}): " + ($actionSummaries -join "; ")
		Write_Log $summaryMessage "green"
		if (-not ($script:ProcessedLanes -contains $laneNumber))
		{
			$script:ProcessedLanes += $laneNumber
		}
	}
	
	Write_Log "`r`n==================== Update_Lane_Config Function Completed ====================" "blue"
}
