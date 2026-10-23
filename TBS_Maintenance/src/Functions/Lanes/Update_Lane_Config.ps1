# ===================================================================================================
#                              FUNCTION: Update_Lane_Config
# ---------------------------------------------------------------------------------------------------
# Description:
#   Loads the standard lane configuration (RUN_TAB, LNK_TAB, STO_TAB, TER_TAB rows) on the
#   selected lanes. Each table is written straight into the lane database over its SQL
#   protocol (keyed delete + insert in one transaction, the same add/replace the load file
#   does); tables that cannot go over SQL are dropped as run_load.sql / lnk_load.sql /
#   sto_load.sql / ter_load.sql in the lane's XF folder, ANSI (Windows-1252), CRLF, no BOM.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $LoadPath: Office\Load directory (must exist)
#   - $script:LaneProtocols: per-lane protocol (TCP / Named Pipes / File / Offline)
#   - $script:ProcessedLanes: lanes processed so far
#
# Writes:
#   - lane RUN_TAB / LNK_TAB / STO_TAB / TER_TAB, or the matching *_load.sql in each XF lane folder
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
	if (-not $selection -or -not $selection.Lanes)
	{
		Write_Log "`r`nLane processing canceled by user." "yellow"
		return
	}

	$ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)

	# Rows per table, with the lane values filled in. One definition feeds both the SQL load
	# and the load file, so the two paths always write the same data.
	$getTableJobs = {
		param ([string]$Lane, [string]$Machine)
		@(
			@{
				Table    = 'RUN_TAB'; Alias = 'RUN'; View = 'Run_Load'; Filename = 'run_load.sql'
				Columns  = @('F1102', 'F1000', 'F1103', 'F1104', 'F1105', 'F1106', 'F1107', 'F1108', 'F1109', 'F1110', 'F1111', 'F1112', 'F1113', 'F1114', 'F1115', 'F1116', 'F1117')
				Keys     = @('F1102', 'F1000')
				Rows     = @(
					"(60,'@TER','sql=BACKUP_ALL','@TER','@DSSF',,'@DSS+001',1,'Backup programs',,0,,'',1,,,)",
					"(62,'@TER','sql=BACKUP_DEVICE','@TER',,,,1,'Backup programs on device',,0,,'',1,,,)",
					"(68,'@TER','sqi=BACKUP_RESELLER','@TER',,,,1,'Backup for reseller',,0,,'',1,,,)",
					"(70,'@TER','sql=DBASE_MAINTENANCE','@TER',,,,1,'Database maintenance',,0,,'',1,,,)",
					"(90,'@TER','sql=PURGE_ALL','@TER','@DSW-008',,'@DSW-001',0,'Purge all report data',,0,,'',7,,,)",
					"(95,'@TER','sqi=PURGE_CLIENT','@TER',,,,0,'Purge customer sensitive data',,0,,'',1,,,)",
					"(120,'@TER','sql=UPDATE_DOWNLOAD','@TER',,,,1,'Automatic update download',,0,,'',1,,,)",
					"(125,'@TER','sql=UPDATE_AUTOMATIC','@TER',,,,1,'Automatic update execution',,0,,'',1,,,)",
					"(250,'SMS','sqi=TRS_POS_BANK_EOS_PUP OUTPUT=RECEIPT','901',,,,1,'Automatic bank close',,,,,1,,,)"
				)
			}
			@{
				Table    = 'LNK_TAB'; Alias = 'LNK'; View = 'Lnk_Load'; Filename = 'lnk_load.sql'
				Columns  = @('F1000', 'F1056', 'F1057')
				Keys     = @('F1000', 'F1056', 'F1057')
				Rows     = @(
					"('$Lane','$StoreNumber','$Lane')",
					"('DSM','$StoreNumber','$Lane')",
					"('PAL','$StoreNumber','$Lane')",
					"('RAL','$StoreNumber','$Lane')",
					"('XAL','$StoreNumber','$Lane')"
				)
			}
			@{
				Table    = 'STO_TAB'; Alias = 'STO'; View = 'Sto_Load'; Filename = 'sto_load.sql'
				Columns  = @('F1000', 'F1018', 'F1180', 'F1181', 'F1182', 'F1937', 'F1965', 'F1966', 'F2691')
				Keys     = @('F1000')
				Rows     = @(
					"('$Lane','Terminal $Lane',1,1,1,,,,)",
					"('DSM','Deploy SMS',1,1,1,,,,)",
					"('PAL','Program all',0,0,1,1,,,)",
					"('RAL','Report all',1,0,0,,,,)",
					"('XAL','Exchange all',0,1,0,,,,)"
				)
			}
			@{
				Table    = 'TER_TAB'; Alias = 'TER'; View = 'Ter_Load'; Filename = 'ter_load.sql'
				Columns  = @('F1056', 'F1057', 'F1058', 'F1125', 'F1169')
				Keys     = @('F1056', 'F1057')
				Rows     = @(
					"('$StoreNumber','$Lane','Terminal $Lane','\\$Machine\storeman\office\XF$StoreNumber$Lane\','\\$Machine\storeman\office\XF${StoreNumber}901\')",
					"('$StoreNumber','901','Server','','')"
				)
			}
		)
	}

	# Load file in the SMS format (@CREATE / view / INSERT / @UPDATE_BATCH ADDRPL / DROP)
	$buildLoadFile = {
		param ($Job)
		$keyText = ($Job.Keys | ForEach-Object { "$_=:$_" }) -join ' AND '
		$lines = @(
			"@CREATE($($Job.Table),$($Job.Alias));"
			"CREATE VIEW $($Job.View) AS SELECT $($Job.Columns -join ',') FROM $($Job.Table);"
			""
			"INSERT INTO $($Job.View) VALUES"
			(($Job.Rows -join ",`r`n") + ";")
			""
			"@UPDATE_BATCH(JOB=ADDRPL,TAR=$($Job.Table),"
			"KEY=$keyText,"
			"SRC=SELECT * FROM $($Job.View));"
			""
			"DROP TABLE $($Job.View);"
		)
		($lines -join "`r`n") + "`r`n"
	}

	# One load row -> SQL literals, with the SMS macros the lane would expand resolved here
	$splitRegex = ',(?=(?:[^'']*''[^'']*'')*[^'']*$)'
	$rowToSqlValues = {
		param ([string]$Row, [string]$Lane)
		$inner = $Row.Trim()
		if ($inner.StartsWith('(') -and $inner.EndsWith(')')) { $inner = $inner.Substring(1, $inner.Length - 2) }
		$now = Get-Date
		foreach ($v in [regex]::Split($inner, $splitRegex))
		{
			$v = $v.Trim()
			if ($v -eq '') { 'NULL'; continue }
			if (-not ($v.StartsWith("'") -and $v.EndsWith("'"))) { $v; continue }
			$res = $v.Substring(1, $v.Length - 2)
			$res = $res -replace '@STORE', $StoreNumber.PadLeft(3, '0')
			$res = $res -replace '@TER', $Lane
			$res = $res -replace '@USERNAME', $env:USERNAME
			$res = $res -replace '@USER', $env:USERNAME
			$res = $res -replace '@RUN', $script:BasePath
			$res = $res -replace '@OFFICE', $script:OfficePath
			$res = $res -replace '@TIME', $now.ToString('HHmm')
			$res = $res -replace '@NOW', $now.ToString('HHmmss')
			$res = $res -replace '@DSSF', $now.ToString('yyyyMMdd')
			$res = $res -replace '@DSW-001', $now.AddDays(-1).ToString('yyyyMMdd')
			$res = $res -replace '@DSW-008', $now.AddDays(-8).ToString('yyyyMMdd')
			$res = $res -replace '@DSS\+001', $now.AddDays(1).ToString('yyyyMMdd')
			$res = $res -replace '@DSS', $now.ToString('yyyyMMdd')
			"'" + $res.Replace("'", "''") + "'"
		}
	}

	# Add/replace a table's rows by key in one transaction; returns the row count
	$loadTableViaSql = {
		param ([string]$ConnStr, $Job, [string]$Lane)
		$sb = New-Object System.Text.StringBuilder
		[void]$sb.AppendLine('SET XACT_ABORT ON; BEGIN TRAN;')
		$colList = '[' + ($Job.Columns -join '],[') + ']'
		foreach ($row in $Job.Rows)
		{
			$values = @(& $rowToSqlValues $row $Lane)
			if ($values.Count -ne $Job.Columns.Count) { throw "$($Job.Table): row has $($values.Count) values for $($Job.Columns.Count) columns: $row" }
			$where = ($Job.Keys | ForEach-Object { "[$_] = " + $values[[array]::IndexOf($Job.Columns, $_)] }) -join ' AND '
			[void]$sb.AppendLine("DELETE FROM [$($Job.Table)] WHERE $where;")
			[void]$sb.AppendLine("INSERT INTO [$($Job.Table)] ($colList) VALUES ($($values -join ', '));")
		}
		[void]$sb.AppendLine('COMMIT;')

		$conn = New-Object System.Data.SqlClient.SqlConnection($ConnStr)
		try
		{
			$conn.Open()
			$cmd = $conn.CreateCommand()
			$cmd.CommandText = $sb.ToString()
			$cmd.CommandTimeout = 30
			[void]$cmd.ExecuteNonQuery()
		}
		finally { $conn.Dispose() }
		$Job.Rows.Count
	}

	foreach ($laneNumber in @($selection.Lanes))
	{
		$laneNumber = ([string]$laneNumber).PadLeft(3, '0')
		$laneFolderPath = Join-Path -Path $OfficePath -ChildPath "XF${StoreNumber}${laneNumber}"
		if (-not (Test-Path $laneFolderPath))
		{
			Write_Log "`r`nLane #$laneNumber not found at path: $laneFolderPath" "yellow"
			continue
		}
		Write_Log "`r`nProcessing Lane #$laneNumber" "blue"

		$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNumber
		$MachineName = if ($laneInfo) { $laneInfo['MachineName'] } else { $script:FunctionResults['LaneNumToMachineName'][$laneNumber] }
		if (-not $MachineName) { $MachineName = "POS${laneNumber}" }

		# Lane connection for its protocol; unknown protocol tries Named Pipes (as before)
		$protocol = if ($script:LaneProtocols) { $script:LaneProtocols[$laneNumber] }
		$connStr = $null
		if ($laneInfo -and $protocol -notin @('File', 'Offline'))
		{
			$raw = if ($protocol -eq 'TCP') { $laneInfo['TcpConnStr'] } else { $laneInfo['NamedPipesConnStr'] }
			$connStr = Normalize_Sql_ConnectionString -RawConnectionString $raw
		}

		$actionSummaries = @()
		foreach ($job in (& $getTableJobs $laneNumber $MachineName))
		{
			if ($connStr)
			{
				try
				{
					$rowCount = & $loadTableViaSql $connStr $job $laneNumber
					Write_Log "Loaded $rowCount rows into $($job.Table) on lane $laneNumber over SQL." "green"
					$actionSummaries += "SQL loaded $($job.Table)"
					continue
				}
				catch
				{
					Write_Log "SQL load of $($job.Table) failed on lane $laneNumber, dropping $($job.Filename) instead: $($_.Exception.Message)" "yellow"
				}
			}

			try
			{
				$dest = Join-Path -Path $laneFolderPath -ChildPath $job.Filename
				Write_Sms_File -Path $dest -Content (& $buildLoadFile $job)
				$actionSummaries += "Copied $($job.Filename)"
			}
			catch
			{
				$actionSummaries += "Failed to copy $($job.Filename)"
			}
		}

		Write_Log ("Lane ${laneNumber} (Machine: ${MachineName}): " + ($actionSummaries -join "; ")) "green"
		if ($script:ProcessedLanes -notcontains $laneNumber) { $script:ProcessedLanes += $laneNumber }
	}

	Write_Log "`r`n==================== Update_Lane_Config Function Completed ====================" "blue"
}
