# ===================================================================================================
#                              FUNCTION: Repair_LOC_Databases_On_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Interactive database repair for lane (LOC) databases: pick lane(s) and a repair level
#   (Audit / Quick / Deep) in a dialog, then run DBCC repair with safety gates. Uses
#   Protocol_Results.txt as the protocol source of truth (File-protocol lanes are skipped),
#   refuses system DB names and restores MULTI_USER state in finally blocks.
#
# Parameters:
#   - StoreNumber    [string]: store number (required)
#   - CommandTimeout [int]: per-statement SQL timeout in seconds (default 900)
#
# Variables:
#   - $script:FunctionResults['LaneNumToMachineName']: required input (lane map)
#   - $script:FunctionResults['SqlModuleName']: required input (SQL module)
#
# Notes:
#   - Falls back to .NET SqlClient when Invoke-Sqlcmd lacks -ConnectionString.
#   - If tempdb is not ONLINE it offers to restart the lane SQL service, then re-checks.
# ===================================================================================================

function Repair_LOC_Databases_On_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[int]$CommandTimeout = 900 # per-statement timeout (seconds)
	)

	# ----------------------------------------
	# Banner: Start
	# ----------------------------------------
	Write_Log "`r`n==================== Starting Repair_LOC_Databases_On_Lanes Function ====================`r`n" "blue"

	# ----------------------------------------
	# Normalize store (3 digits) for any labels/logs
	# ----------------------------------------
	$StoreNumber = "$StoreNumber".PadLeft(3, '0')

	# ----------------------------------------
	# Validate required caches
	# ----------------------------------------
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }

	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "No lanes available. Please retrieve nodes first (LaneNumToMachineName missing/empty)." "red"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}

	# ----------------------------------------
	# Import SQL module and capture Invoke-Sqlcmd
	# ----------------------------------------
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	if (-not $SqlModuleName -or $SqlModuleName -eq "None")
	{
		Write_Log "No valid SQL module available for SQL operations!" "red"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}

	$invokeSqlcmdCmd = $null
	try
	{
		Import-Module $SqlModuleName -ErrorAction Stop
		$invokeSqlcmdCmd = Get-Command Invoke-Sqlcmd -ErrorAction Stop
	}
	catch
	{
		Write_Log "Failed to import SQL module or find Invoke-Sqlcmd: $_" "red"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}

	# Detect whether Invoke-Sqlcmd supports -ConnectionString
	$InvokeSupportsConnString = $false
	try
	{
		if ($invokeSqlcmdCmd -and $invokeSqlcmdCmd.Parameters -and $invokeSqlcmdCmd.Parameters.ContainsKey('ConnectionString'))
		{
			$InvokeSupportsConnString = $true
		}
	}
	catch { $InvokeSupportsConnString = $false }

	# Ensure SqlClient types are available for fallback mode
	try { Add-Type -AssemblyName System.Data 2>$null }
	catch { }

	$invokeSqlTable = {
		param (
			[string]$ConnectionString,
			[string]$Query,
			[int]$Timeout = 30
		)

		if ($InvokeSupportsConnString)
		{
			$rows = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $Query -QueryTimeout $Timeout -ErrorAction Stop
			if ($null -eq $rows) { return @() }
			return @($rows)
		}

		$dt = New-Object System.Data.DataTable
		$cn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
		try
		{
			$cn.Open()
			$cmd = $cn.CreateCommand()
			$cmd.CommandText = $Query
			$cmd.CommandTimeout = $Timeout
			$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
			[void]$da.Fill($dt)
		}
		finally
		{
			try { $cn.Close() }
			catch { }
			try { $cn.Dispose() }
			catch { }
		}

		return $dt
	}

	$invokeSqlNonQuery = {
		param (
			[string]$ConnectionString,
			[string]$Query,
			[int]$Timeout = 30
		)

		if ($InvokeSupportsConnString)
		{
			Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $Query -QueryTimeout $Timeout -ErrorAction Stop | Out-Null
			return
		}

		$cn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
		try
		{
			$cn.Open()
			$cmd = $cn.CreateCommand()
			$cmd.CommandText = $Query
			$cmd.CommandTimeout = $Timeout
			[void]$cmd.ExecuteNonQuery()
		}
		finally
		{
			try { $cn.Close() }
			catch { }
			try { $cn.Dispose() }
			catch { }
		}
	}

	$getRowCount = {
		param ($Rows)

		if ($Rows -is [System.Data.DataTable]) { return $Rows.Rows.Count }
		if ($null -eq $Rows) { return 0 }
		return @($Rows).Count
	}

	$getDatabaseState = {
		param (
			[string]$ConnectionString,
			[string]$DatabaseName
		)

		if ([string]::IsNullOrWhiteSpace($DatabaseName)) { return $null }

		$dbLiteral = $DatabaseName.Replace("'", "''")
		$rows = & $invokeSqlTable $ConnectionString "SELECT state_desc FROM sys.databases WHERE name = N'$dbLiteral';" 10

		if ($rows -is [System.Data.DataTable])
		{
			if ($rows.Rows.Count -gt 0) { return [string]$rows.Rows[0]['state_desc'] }
		}
		elseif ((& $getRowCount $rows) -gt 0)
		{
			return [string]($rows | Select-Object -First 1 -ExpandProperty state_desc)
		}

		return $null
	}

	$getSystemDbProblems = {
		param ([string]$ConnectionString)

		$rows = & $invokeSqlTable $ConnectionString "SELECT name, state_desc FROM sys.databases WHERE name IN ('master','model','msdb','tempdb');" 10
		$bad = New-Object System.Collections.Generic.List[pscustomobject]

		if ($rows -is [System.Data.DataTable])
		{
			foreach ($row in $rows.Rows)
			{
				$n = [string]$row['name']
				$s = [string]$row['state_desc']
				if ($s -and $s.ToUpper() -ne 'ONLINE')
				{
					$bad.Add([pscustomobject]@{ Name = $n; State = $s })
				}
			}
		}
		else
		{
			foreach ($r in @($rows))
			{
				if ($r -and $r.state_desc -and ($r.state_desc.ToString().ToUpper() -ne 'ONLINE'))
				{
					$bad.Add([pscustomobject]@{ Name = [string]$r.name; State = [string]$r.state_desc })
				}
			}
		}

		return @($bad)
	}

	$testSqlConnectivity = {
		param ([string]$ConnectionString)

		try
		{
			[void](& $invokeSqlTable $ConnectionString "SELECT 1 AS ok;" 5)
			return $true
		}
		catch
		{
			return $false
		}
	}

	$getCheckDbIssueCount = {
		param (
			[string]$ConnectionString,
			[string]$QuotedDbName,
			[int]$Timeout = 900
		)

		$rows = & $invokeSqlTable $ConnectionString "DBCC CHECKDB ($QuotedDbName) WITH NO_INFOMSGS, ALL_ERRORMSGS, TABLERESULTS;" $Timeout
		return (& $getRowCount $rows)
	}

	$detectLaneProtocol = {
		param ([string]$MachineName)

		if ([string]::IsNullOrWhiteSpace($MachineName)) { return $null }

		$candidates = @(
			@{ Label = 'TCP'; ConnectionString = "Data Source=tcp:$MachineName;Initial Catalog=master;Integrated Security=True;Connect Timeout=2;TrustServerCertificate=True;" },
			@{ Label = 'Named Pipes'; ConnectionString = "Data Source=np:$MachineName;Initial Catalog=master;Integrated Security=True;Connect Timeout=2;TrustServerCertificate=True;" }
		)

		foreach ($candidate in $candidates)
		{
			if (& $testSqlConnectivity $candidate.ConnectionString)
			{
				return $candidate.Label
			}
		}

		return $null
	}

	# ----------------------------------------
	# Load protocol results file (SOURCE OF TRUTH)
	# ----------------------------------------
	$protoFile = $script:ProtocolResultsFile

	$LaneProtocolsFromFile = @{ }
	if (Test-Path $protoFile)
	{
		try
		{
			$protoLines = Get-Content -LiteralPath $protoFile -ErrorAction SilentlyContinue
			foreach ($line in $protoLines)
			{
				# Format: LANE,PROTOCOL  (you wrote lane,protocol)
				if ($line -match '^\s*([^,]+)\s*,\s*([^,]+)\s*$')
				{
					$laneRaw = $matches[1].Trim()
					$proto = $matches[2].Trim()

					$laneDigits = ($laneRaw -replace '[^\d]', '')
					if (-not [string]::IsNullOrWhiteSpace($laneDigits))
					{
						$laneNum = $laneDigits.PadLeft(3, '0')
						$LaneProtocolsFromFile[$laneNum] = $proto
					}
				}
			}
			Write_Log "Loaded protocol results from: $protoFile" "gray"
		}
		catch
		{
			Write_Log "WARN: Failed reading ProtocolResultsFile '$protoFile'. Error: $_" "yellow"
		}
	}
	else
	{
		Write_Log "ProtocolResultsFile not found: $protoFile. Will fall back to in-memory/live protocol detection for this run." "yellow"
	}

	# ----------------------------------------
	# Lane selection (same UX pattern)
	# ----------------------------------------
	$selection = $null
	try
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Lanes for DB Repair"
	}
	catch
	{
		Write_Log "Lane selection failed: $_" "red"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}

	if ($selection -eq $null)
	{
		Write_Log "Lane DB repair canceled by user." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}
	if (-not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}

	# Normalize lane list from selection (handles PSCustomObject or strings/ints)
	$Lanes = @()
	if ($selection.Lanes[0] -is [PSCustomObject] -and $selection.Lanes[0].PSObject.Properties.Name -contains 'LaneNumber')
	{
		foreach ($item in $selection.Lanes)
		{
			$ln = "$($item.LaneNumber)" -replace '[^\d]', ''
			if ($ln) { $Lanes += $ln.PadLeft(3, '0') }
		}
	}
	else
	{
		foreach ($item in $selection.Lanes)
		{
			$ln = "$item" -replace '[^\d]', ''
			if ($ln) { $Lanes += $ln.PadLeft(3, '0') }
		}
	}
	$Lanes = @($Lanes | Where-Object { $_ } | Select-Object -Unique)

	# ----------------------------------------
	# In-function dialog to pick repair level
	# ----------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Database Repair Level"
	$form.Size = New-Object System.Drawing.Size(520, 240)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	$lbl = New-Object System.Windows.Forms.Label
	$lbl.Text = "Choose how aggressive the repair should be:"
	$lbl.AutoSize = $true
	$lbl.Location = New-Object System.Drawing.Point(12, 12)
	$form.Controls.Add($lbl)

	$grp = New-Object System.Windows.Forms.GroupBox
	$grp.Text = "Repair Level"
	$grp.Location = New-Object System.Drawing.Point(12, 36)
	$grp.Size = New-Object System.Drawing.Size(490, 110)
	$form.Controls.Add($grp)

	$rbAudit = New-Object System.Windows.Forms.RadioButton
	$rbAudit.Text = "Audit only (DBCC CHECKDB, no changes)"
	$rbAudit.AutoSize = $true
	$rbAudit.Location = New-Object System.Drawing.Point(16, 24)
	$rbAudit.Checked = $true
	$grp.Controls.Add($rbAudit)

	$rbQuick = New-Object System.Windows.Forms.RadioButton
	$rbQuick.Text = "Quick repair (REPAIR_REBUILD)"
	$rbQuick.AutoSize = $true
	$rbQuick.Location = New-Object System.Drawing.Point(16, 48)
	$grp.Controls.Add($rbQuick)

	$rbDeep = New-Object System.Windows.Forms.RadioButton
	$rbDeep.Text = "Deep repair (REPAIR_ALLOW_DATA_LOSS)"
	$rbDeep.AutoSize = $true
	$rbDeep.Location = New-Object System.Drawing.Point(16, 72)
	$grp.Controls.Add($rbDeep)

	$chkConfirm = New-Object System.Windows.Forms.CheckBox
	$chkConfirm.Text = "I understand deep repair can cause data loss."
	$chkConfirm.AutoSize = $true
	$chkConfirm.Location = New-Object System.Drawing.Point(28, 152)
	$chkConfirm.Enabled = $false
	$form.Controls.Add($chkConfirm)

	[void]$rbDeep.Add_CheckedChanged({
			if ($rbDeep.Checked) { $chkConfirm.Enabled = $true }
			else { $chkConfirm.Enabled = $false; $chkConfirm.Checked = $false }
		})
	[void]$rbAudit.Add_CheckedChanged({ if ($rbAudit.Checked) { $chkConfirm.Enabled = $false; $chkConfirm.Checked = $false } })
	[void]$rbQuick.Add_CheckedChanged({ if ($rbQuick.Checked) { $chkConfirm.Enabled = $false; $chkConfirm.Checked = $false } })

	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(314, 172)
	$btnOK.Size = New-Object System.Drawing.Size(85, 28)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(417, 172)
	$btnCancel.Size = New-Object System.Drawing.Size(85, 28)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)

	$dlg = $form.ShowDialog()
	if ($dlg -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Repair level selection canceled." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}

	$Level = 'Audit'
	if ($rbQuick.Checked) { $Level = 'Quick' }
	if ($rbDeep.Checked) { $Level = 'Deep' }

	if ($Level -eq 'Deep' -and -not $chkConfirm.Checked)
	{
		Write_Log "Deep repair selected but confirmation not checked. Aborting." "yellow"
		Write_Log "`r`n==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
		return
	}

	Write_Log ("Selected repair level: {0}" -f $Level) "gray"

	# ----------------------------------------
	# Summary container
	# ----------------------------------------
	$laneSummary = New-Object System.Collections.Generic.List[pscustomobject]

	# ----------------------------------------
	# Lane processing loop
	# ----------------------------------------
	foreach ($LaneNumber in ($Lanes | Sort-Object))
	{
		$laneKeyP = ("$LaneNumber" -replace '[^\d]', '').PadLeft(3, '0')

		# Resolve machine name
		$machineName = $null
		if ($LaneNumToMachineName.ContainsKey($laneKeyP)) { $machineName = $LaneNumToMachineName[$laneKeyP] }
		elseif ($LaneNumToMachineName.ContainsKey($LaneNumber)) { $machineName = $LaneNumToMachineName[$LaneNumber] }

		if (-not $machineName)
		{
			Write_Log ("Lane {0}: No machine mapping found. Skipping." -f $laneKeyP) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = ''; Protocol = ''; DBServer = ''; DBName = ''; Level = $Level; Result = 'NoMachine' })
			continue
		}

		# Protocol from file, then in-memory cache, then live detection
		$proto = $null
		if ($LaneProtocolsFromFile.ContainsKey($laneKeyP)) { $proto = [string]$LaneProtocolsFromFile[$laneKeyP] }

		if ([string]::IsNullOrWhiteSpace($proto) -and $script:LaneProtocols)
		{
			$tryKeys = @($laneKeyP, $LaneNumber)
			if ($machineName) { $tryKeys += @($machineName, $machineName.ToLower()) }

			foreach ($k in $tryKeys)
			{
				if (-not [string]::IsNullOrWhiteSpace("$k") -and $script:LaneProtocols.ContainsKey($k))
				{
					$proto = [string]$script:LaneProtocols[$k]
					break
				}
			}
		}

		if ([string]::IsNullOrWhiteSpace($proto))
		{
			$proto = & $detectLaneProtocol $machineName
			if (-not [string]::IsNullOrWhiteSpace($proto))
			{
				Write_Log ("Lane {0} ({1}): Protocol detected live as '{2}'." -f $laneKeyP, $machineName, $proto) "gray"
				$LaneProtocolsFromFile[$laneKeyP] = $proto
				if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
				$script:LaneProtocols[$laneKeyP] = $proto
				$script:LaneProtocols[$LaneNumber] = $proto
				$script:LaneProtocols[$machineName] = $proto
				$script:LaneProtocols[$machineName.ToLower()] = $proto
			}
		}

		if ([string]::IsNullOrWhiteSpace($proto))
		{
			Write_Log ("Lane {0} ({1}): No usable protocol found in file, memory, or live detection. Skipping." -f $laneKeyP, $machineName) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = ''; DBServer = ''; DBName = ''; Level = $Level; Result = 'NoProtocolDetected' })
			continue
		}

		if ($proto -ne 'TCP' -and $proto -ne 'Named Pipes')
		{
			# File / Unknown -> do not attempt DB operations
			Write_Log ("Lane {0} ({1}): Protocol is '{2}'. Skipping DB repair (not reachable)." -f $laneKeyP, $machineName, $proto) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = ''; DBName = ''; Level = $Level; Result = 'ProtocolNotReachable' })
			continue
		}

		# Resolve lane DB info (Startup.ini parsing cached by your helper)
		$dbInfo = $null
		try { $dbInfo = Get_All_Lanes_Database_Info -LaneNumber $laneKeyP }
		catch { $dbInfo = $null }

		if ($dbInfo -and ($dbInfo -isnot [hashtable]))
		{
			$tmpDbInfo = @{ }
			foreach ($p in $dbInfo.PSObject.Properties) { $tmpDbInfo[$p.Name] = $p.Value }
			$dbInfo = $tmpDbInfo
		}

		if (-not $dbInfo)
		{
			Write_Log ("Lane {0} ({1}): Could not get DB info. Skipping." -f $laneKeyP, $machineName) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = ''; DBName = ''; Level = $Level; Result = 'NoDbInfo' })
			continue
		}

		$dbName = $dbInfo['DBName']
		$dbServer = $dbInfo['DBServer']
		$instanceName = $dbInfo['InstanceName']
		$csNamedPipes = $dbInfo['NamedPipesConnStr']
		$csTcp = $dbInfo['TcpConnStr']

		Write_Log ("`r`n--- Lane {0} ({1}) | Protocol={2} | DB: {3} on {4} ---" -f $laneKeyP, $machineName, $proto, $dbName, $dbServer) "blue"

		if ([string]::IsNullOrWhiteSpace($dbName) -or [string]::IsNullOrWhiteSpace($dbServer))
		{
			Write_Log "Incomplete DB info (DBName or DBServer missing). Skipping lane." "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'IncompleteDbInfo' })
			continue
		}

		# Guardrail: refuse if the application DB name is a system DB (protects from bad Startup.ini)
		$sysDbNames = @('master', 'model', 'msdb', 'tempdb')
		if ($sysDbNames -contains $dbName.ToLower())
		{
			Write_Log ("REFUSING: Startup.ini DBName resolved to system DB '{0}'. Skipping lane." -f $dbName) "red"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'Refused_SystemDbName' })
			continue
		}

		# Choose the correct raw connection string based on protocol from file
		$rawCs = $null
		if ($proto -eq 'TCP') { $rawCs = $csTcp }
		if ($proto -eq 'Named Pipes') { $rawCs = $csNamedPipes }

		if ([string]::IsNullOrWhiteSpace($rawCs))
		{
			Write_Log ("Lane {0} ({1}): Missing {2} connection string in dbInfo. Skipping." -f $laneKeyP, $machineName, $proto) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'NoConnStrForProtocol' })
			continue
		}

		# Build MASTER connection string for system checks + repairs (no fallbacks: we obey protocol file)
		$connStr = $rawCs
		if ($connStr -match '(?i)\bDatabase\s*=\s*[^;]+;')
		{
			$connStr = [regex]::Replace($connStr, '(?i)\bDatabase\s*=\s*[^;]+;', 'Database=master;')
		}
		else
		{
			if ($connStr.Trim().EndsWith(';')) { $connStr = $connStr + 'Database=master;' }
			else { $connStr = $connStr + ';Database=master;' }
		}

		# Add safe extras if not already present (do NOT override auth)
		if ($connStr -notmatch '(?i)\bTrustServerCertificate\s*=') { $connStr += 'TrustServerCertificate=True;' }
		if ($connStr -notmatch '(?i)\bApplication\s*Name\s*=') { $connStr += 'Application Name=TBS_DBRepair;' }
		if ($connStr -notmatch '(?i)\bIntegrated\s*Security\s*=' -and $connStr -notmatch '(?i)\bUser\s*ID\s*=') { $connStr += 'Integrated Security=True;' }

		# Quote db identifier safely for T-SQL
		$qDb = "[" + ($dbName -replace "]", "]]") + "]"
		$dbNameQuoted = $dbName.Replace("'", "''")

		# ----------------------------------------
		# Unified query execution (Invoke-Sqlcmd if possible; else SqlClient)
		# ----------------------------------------
		$ExecError = $null

		# Probe: connection + DB existence/state
		$probeRows = $null
		$probeTsql = @"
SELECT name, state_desc
FROM sys.databases
WHERE name = N'$dbNameQuoted';
"@

		try
		{
			if ($InvokeSupportsConnString)
			{
				$probeRows = Invoke-Sqlcmd -ConnectionString $connStr -Query $probeTsql -QueryTimeout 8 -ErrorAction Stop
			}
			else
			{
				# SqlClient fallback
				$dt = New-Object System.Data.DataTable
				$cn = New-Object System.Data.SqlClient.SqlConnection $connStr
				try
				{
					$cn.Open()
					$cmd = $cn.CreateCommand()
					$cmd.CommandText = $probeTsql
					$cmd.CommandTimeout = 8
					$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
					[void]$da.Fill($dt)
				}
				finally
				{
					try { $cn.Close() }
					catch { }
					try { $cn.Dispose() }
					catch { }
				}
				$probeRows = $dt
			}
		}
		catch
		{
			$ExecError = $_
		}

		if ($ExecError)
		{
			Write_Log ("Lane {0} ({1}): Probe failed using {2}. Skipping. Error: {3}" -f $laneKeyP, $machineName, $proto, $ExecError.Exception.Message) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'ConnProbeFailed' })
			continue
		}

		# Normalize probe result handling for both object array and DataTable
		$targetDbFound = $false
		$targetDbState = $null

		if ($probeRows -is [System.Data.DataTable])
		{
			if ($probeRows.Rows.Count -gt 0)
			{
				$targetDbFound = $true
				$targetDbState = [string]$probeRows.Rows[0]['state_desc']
			}
		}
		else
		{
			if ($probeRows -and ($probeRows | Measure-Object).Count -gt 0)
			{
				$targetDbFound = $true
				$targetDbState = [string]($probeRows | Select-Object -First 1 -ExpandProperty state_desc)
			}
		}

		if (-not $targetDbFound)
		{
			Write_Log ("Lane {0} ({1}): Target DB '{2}' not found in sys.databases. Skipping." -f $laneKeyP, $machineName, $dbName) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'DbMissing' })
			continue
		}

		if ([string]::IsNullOrWhiteSpace($targetDbState)) { $targetDbState = 'Unknown' }
		Write_Log ("Target DB state (pre): {0}" -f $targetDbState) "gray"

		# ----------------------------------------
		# System DB health gate (tempdb/model/msdb/master states)
		# ----------------------------------------
		try
		{
			$badSystemDbs = @(& $getSystemDbProblems $connStr)
		}
		catch
		{
			Write_Log ("Lane {0} ({1}): WARN: could not query system DB states: {2}" -f $laneKeyP, $machineName, $_.Exception.Message) "yellow"
			$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'SystemDbStateUnknown' })
			continue
		}

		if ($badSystemDbs.Count -gt 0)
		{
			Write_Log ("Lane {0} ({1}): System DBs not ONLINE detected." -f $laneKeyP, $machineName) "yellow"
			foreach ($b in $badSystemDbs) { Write_Log ("  - {0}: {1}" -f $b.Name, $b.State) "yellow" }

			# Determine SQL service name based on instance
			$svcName = "MSSQLSERVER"
			if ($instanceName -and ($instanceName.ToUpper() -ne "MSSQLSERVER"))
			{
				$svcName = "MSSQL`$$instanceName"
			}

			# tempdb handling: offer restart SQL service
			$hasTempdbIssue = $false
			foreach ($b in $badSystemDbs) { if ($b.Name -eq 'tempdb') { $hasTempdbIssue = $true } }

			if ($hasTempdbIssue)
			{
				$msg = "Lane $laneKeyP ($machineName): tempdb is not ONLINE.`r`n`r`n" +
				"Would you like to restart the SQL Server service ($svcName) on $machineName now?`r`n" +
				"This will drop connections and recreate tempdb."

				$resp = [System.Windows.Forms.MessageBox]::Show(
					$msg,
					"Restart SQL Service for tempdb",
					[System.Windows.Forms.MessageBoxButtons]::YesNo,
					[System.Windows.Forms.MessageBoxIcon]::Warning
				)

				if ($resp -eq [System.Windows.Forms.DialogResult]::Yes)
				{
					Write_Log ("Attempting to restart service {0} on {1}..." -f $svcName, $machineName) "gray"

					# Stop service
					try { & sc.exe "\\$machineName" stop $svcName | Out-Null }
					catch { }

					# Wait for stopped up to ~60s
					$stopped = $false
					for ($i = 0; $i -lt 30; $i++)
					{
						Start-Sleep -Milliseconds 2000
						try
						{
							$chk = sc.exe "\\$machineName" query $svcName 2>$null
							$txt = if ($chk) { $chk | Out-String }
							else { "" }
							if ($txt -match 'STATE\s*:\s*\d+\s*STOPPED') { $stopped = $true; break }
						}
						catch { }
					}
					if ($stopped) { Write_Log "Service stopped." "gray" }
					else { Write_Log "WARN: Service did not report STOPPED in time." "yellow" }

					# Start service
					try { & sc.exe "\\$machineName" start $svcName | Out-Null }
					catch { }

					# Wait for running up to ~90s
					$running = $false
					for ($i = 0; $i -lt 45; $i++)
					{
						Start-Sleep -Milliseconds 2000
						try
						{
							$chk = sc.exe "\\$machineName" query $svcName 2>$null
							$txt = if ($chk) { $chk | Out-String }
							else { "" }
							if ($txt -match 'STATE\s*:\s*\d+\s*RUNNING') { $running = $true; break }
						}
						catch { }
					}
					if ($running) { Write_Log "Service started." "green" }
					else { Write_Log "ERROR: Service did not report RUNNING in time." "red" }

					# Re-probe system DBs after restart (simple loop)
					$reprobeOK = $false
					for ($t = 1; $t -le 60; $t++)
					{
						try
						{
							if ($InvokeSupportsConnString)
							{
								Invoke-Sqlcmd -ConnectionString $connStr -Query "SELECT 1 AS ok;" -QueryTimeout 5 -ErrorAction Stop | Out-Null
							}
							else
							{
								$cnx = New-Object System.Data.SqlClient.SqlConnection $connStr
								try { $cnx.Open(); $cmdx = $cnx.CreateCommand(); $cmdx.CommandText = "SELECT 1"; $cmdx.CommandTimeout = 5; [void]$cmdx.ExecuteScalar() }
								finally
								{
									try { $cnx.Close() }
									catch { }
									try { $cnx.Dispose() }
									catch { }
								}
							}
							$reprobeOK = $true
							break
						}
						catch
						{
							Start-Sleep -Milliseconds 2000
						}
					}

					if (-not $reprobeOK)
					{
						Write_Log "ERROR: Could not reconnect after service restart; skipping lane." "red"
						$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'ServiceRestartFailed' })
						continue
					}

					try
					{
						$badSystemDbs = @(& $getSystemDbProblems $connStr)
					}
					catch
					{
						Write_Log ("ERROR: Could not re-check system DB states after restart: {0}" -f $_.Exception.Message) "red"
						$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'SystemDbReprobeFailed' })
						continue
					}

					$tempdbStillBad = $false
					foreach ($b in $badSystemDbs)
					{
						if ($b.Name -eq 'tempdb') { $tempdbStillBad = $true; break }
					}

					if ($tempdbStillBad)
					{
						Write_Log "ERROR: tempdb is still not ONLINE after SQL service restart. Skipping lane." "red"
						$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'TempdbStillNotOnline' })
						continue
					}
				}
				else
				{
					Write_Log "User declined SQL service restart for tempdb. Skipping lane with guidance: restart SQL service to recreate tempdb." "yellow"
					$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'TempdbNotOnline_Skipped' })
					continue
				}
			}

			# If master is bad -> do not automate
			$hasMaster = $false
			$hasModel = $false
			$hasMsdb = $false
			foreach ($b in $badSystemDbs)
			{
				if ($b.Name -eq 'master') { $hasMaster = $true }
				if ($b.Name -eq 'model') { $hasModel = $true }
				if ($b.Name -eq 'msdb') { $hasMsdb = $true }
			}

			if ($hasMaster)
			{
				Write_Log "CRITICAL: master is not ONLINE. Automated repair is not supported here." "red"
				Write_Log "Action: Restore master from known-good backup OR rebuild system DBs (setup.exe /ACTION=REBUILDDATABASE), then restore." "gray"
				$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'MasterNotOnline' })
				continue
			}

			# model/msdb -> optional emergency deep repair prompt
			$targets = @()
			if ($hasModel) { $targets += 'model' }
			if ($hasMsdb) { $targets += 'msdb' }

			if ($targets.Count -gt 0)
			{
				$list = ($targets -join ', ')
				$warn = "Lane $laneKeyP ($machineName): $list is not ONLINE.`r`n`r`n" +
				"Attempt DEEP REPAIR (EMERGENCY + REPAIR_ALLOW_DATA_LOSS) for these system DB(s)?`r`n" +
				"This can cause data loss. Recommended alternative is RESTORE from backup."

				$resp2 = [System.Windows.Forms.MessageBox]::Show(
					$warn,
					"Deep Repair system DB(s)?",
					[System.Windows.Forms.MessageBoxButtons]::YesNo,
					[System.Windows.Forms.MessageBoxIcon]::Warning
				)

				if ($resp2 -ne [System.Windows.Forms.DialogResult]::Yes)
				{
					Write_Log ("User declined deep repair for system DB(s): {0}. Skipping lane with guidance to RESTORE." -f $list) "yellow"
					$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'SystemDbNotOnline_Skipped' })
					continue
				}

				foreach ($sysDb in $targets)
				{
					Write_Log ("Attempting deep repair on system DB: {0}" -f $sysDb) "yellow"
					$qSys = "[" + ($sysDb -replace "]", "]]") + "]"

					# Best-effort sequence
					try
					{
						& $invokeSqlNonQuery $connStr "ALTER DATABASE $qSys SET EMERGENCY;" 30
					}
					catch { Write_Log ("WARN: EMERGENCY failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "yellow" }

					try
					{
						& $invokeSqlNonQuery $connStr "ALTER DATABASE $qSys SET SINGLE_USER WITH ROLLBACK IMMEDIATE;" 30
					}
					catch { Write_Log ("WARN: SINGLE_USER failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "yellow" }

					try
					{
						$ts = "DBCC CHECKDB ($qSys, REPAIR_ALLOW_DATA_LOSS) WITH ALL_ERRORMSGS;"
						[void](& $invokeSqlTable $connStr $ts $CommandTimeout)
						Write_Log ("Deep repair completed on {0}." -f $sysDb) "green"
					}
					catch { Write_Log ("Deep repair failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "red" }

					try
					{
						& $invokeSqlNonQuery $connStr "ALTER DATABASE $qSys SET ONLINE;" 30
					}
					catch { Write_Log ("WARN: ONLINE restore failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "yellow" }

					try
					{
						& $invokeSqlNonQuery $connStr "ALTER DATABASE $qSys SET MULTI_USER;" 30
					}
					catch { Write_Log ("WARN: MULTI_USER restore failed on {0}: {1}" -f $sysDb, $_.Exception.Message) "yellow" }
				}

				try
				{
					$badSystemDbs = @(& $getSystemDbProblems $connStr)
				}
				catch
				{
					Write_Log ("ERROR: Could not re-check system DB states after system DB repair: {0}" -f $_.Exception.Message) "red"
					$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'SystemDbReprobeFailed' })
					continue
				}

				if ($badSystemDbs.Count -gt 0)
				{
					Write_Log "System DBs still not ONLINE after repair attempt. Skipping lane." "red"
					foreach ($b in $badSystemDbs) { Write_Log ("  - {0}: {1}" -f $b.Name, $b.State) "red" }
					$laneSummary.Add([pscustomobject]@{ Lane = $laneKeyP; Machine = $machineName; Protocol = $proto; DBServer = $dbServer; DBName = $dbName; Level = $Level; Result = 'SystemDbStillNotOnline' })
					continue
				}
			}
		}

		# ----------------------------------------
		# Snapshot DB state BEFORE repair (sys.databases)
		# ----------------------------------------
		$stateBefore = 'Unknown'
		try
		{
			$stateBefore = & $getDatabaseState $connStr $dbName
		}
		catch { }

		if ([string]::IsNullOrWhiteSpace($stateBefore)) { $stateBefore = 'Unknown' }
		Write_Log ("DB state before: {0}" -f $stateBefore) "gray"

		# ----------------------------------------
		# Repair execution (always try to restore DB to MULTI_USER in finally)
		# ----------------------------------------
		$ok = $false
		$usedQuick = $false
		$usedDeep = $false
		$auditIssueCount = $null

		# In all repair modes we may toggle SINGLE_USER / EMERGENCY
		$didSingleUser = $false
		$didEmergency = $false

		try
		{
			if ($Level -eq 'Audit')
			{
				Write_Log ("{0}: Running DBCC CHECKDB (Audit)..." -f $dbName) "gray"
				$auditIssueCount = & $getCheckDbIssueCount $connStr $qDb $CommandTimeout
				if ($auditIssueCount -gt 0)
				{
					$ok = $false
					Write_Log ("Audit completed. DBCC reported {0} row(s)." -f $auditIssueCount) "yellow"
				}
				else
				{
					$ok = $true
					Write_Log "Audit completed. DBCC reported no issues." "green"
				}
			}
			elseif ($Level -eq 'Quick')
			{
				Write_Log ("{0}: Quick repair (REPAIR_REBUILD)..." -f $dbName) "gray"

				# SINGLE_USER
				try
				{
					$tsSU = "ALTER DATABASE $qDb SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"
					if ($InvokeSupportsConnString) { Invoke-Sqlcmd -ConnectionString $connStr -Query $tsSU -QueryTimeout 60 -ErrorAction Stop | Out-Null }
					else
					{
						$cnQ = New-Object System.Data.SqlClient.SqlConnection $connStr
						try { $cnQ.Open(); $cmdQ = $cnQ.CreateCommand(); $cmdQ.CommandText = $tsSU; $cmdQ.CommandTimeout = 60; [void]$cmdQ.ExecuteNonQuery() }
						finally
						{
							try { $cnQ.Close() }
							catch { }
							try { $cnQ.Dispose() }
							catch { }
						}
					}
					$didSingleUser = $true
				}
				catch
				{
					Write_Log ("WARN: SINGLE_USER failed: {0}" -f $_.Exception.Message) "yellow"
				}

				# DBCC repair
				$ts = "DBCC CHECKDB ($qDb, REPAIR_REBUILD) WITH ALL_ERRORMSGS, TABLERESULTS;"
				if ($InvokeSupportsConnString)
				{
					Invoke-Sqlcmd -ConnectionString $connStr -Query $ts -QueryTimeout $CommandTimeout -ErrorAction Stop | Out-Null
				}
				else
				{
					$cnQ2 = New-Object System.Data.SqlClient.SqlConnection $connStr
					try { $cnQ2.Open(); $cmdQ2 = $cnQ2.CreateCommand(); $cmdQ2.CommandText = $ts; $cmdQ2.CommandTimeout = $CommandTimeout; [void]$cmdQ2.ExecuteNonQuery() }
					finally
					{
						try { $cnQ2.Close() }
						catch { }
						try { $cnQ2.Dispose() }
						catch { }
					}
				}

				$ok = $true
				$usedQuick = $true
				Write_Log "Quick repair completed." "green"
			}
			elseif ($Level -eq 'Deep')
			{
				Write_Log ("{0}: Deep repair (EMERGENCY + REPAIR_ALLOW_DATA_LOSS)..." -f $dbName) "gray"

				# EMERGENCY
				try
				{
					$tsE = "ALTER DATABASE $qDb SET EMERGENCY;"
					if ($InvokeSupportsConnString) { Invoke-Sqlcmd -ConnectionString $connStr -Query $tsE -QueryTimeout 60 -ErrorAction Stop | Out-Null }
					else
					{
						$cnD = New-Object System.Data.SqlClient.SqlConnection $connStr
						try { $cnD.Open(); $cmdD = $cnD.CreateCommand(); $cmdD.CommandText = $tsE; $cmdD.CommandTimeout = 60; [void]$cmdD.ExecuteNonQuery() }
						finally
						{
							try { $cnD.Close() }
							catch { }
							try { $cnD.Dispose() }
							catch { }
						}
					}
					$didEmergency = $true
				}
				catch
				{
					Write_Log ("WARN: EMERGENCY failed: {0}" -f $_.Exception.Message) "yellow"
				}

				# SINGLE_USER
				try
				{
					$tsSU = "ALTER DATABASE $qDb SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"
					if ($InvokeSupportsConnString) { Invoke-Sqlcmd -ConnectionString $connStr -Query $tsSU -QueryTimeout 60 -ErrorAction Stop | Out-Null }
					else
					{
						$cnD2 = New-Object System.Data.SqlClient.SqlConnection $connStr
						try { $cnD2.Open(); $cmdD2 = $cnD2.CreateCommand(); $cmdD2.CommandText = $tsSU; $cmdD2.CommandTimeout = 60; [void]$cmdD2.ExecuteNonQuery() }
						finally
						{
							try { $cnD2.Close() }
							catch { }
							try { $cnD2.Dispose() }
							catch { }
						}
					}
					$didSingleUser = $true
				}
				catch
				{
					Write_Log ("WARN: SINGLE_USER failed: {0}" -f $_.Exception.Message) "yellow"
				}

				# DBCC deep repair
				$ts = "DBCC CHECKDB ($qDb, REPAIR_ALLOW_DATA_LOSS) WITH ALL_ERRORMSGS, TABLERESULTS;"
				if ($InvokeSupportsConnString)
				{
					Invoke-Sqlcmd -ConnectionString $connStr -Query $ts -QueryTimeout $CommandTimeout -ErrorAction Stop | Out-Null
				}
				else
				{
					$cnD3 = New-Object System.Data.SqlClient.SqlConnection $connStr
					try { $cnD3.Open(); $cmdD3 = $cnD3.CreateCommand(); $cmdD3.CommandText = $ts; $cmdD3.CommandTimeout = $CommandTimeout; [void]$cmdD3.ExecuteNonQuery() }
					finally
					{
						try { $cnD3.Close() }
						catch { }
						try { $cnD3.Dispose() }
						catch { }
					}
				}

				$ok = $true
				$usedDeep = $true
				Write_Log "Deep repair completed." "green"
			}
		}
		catch
		{
			$ok = $false
			if ($Level -eq 'Audit')
			{
				Write_Log ("Audit reported errors or failed: {0}" -f $_.Exception.Message) "yellow"
			}
			else
			{
				Write_Log ("Repair failed: {0}" -f $_.Exception.Message) "red"
			}
		}
		finally
		{
			# ALWAYS attempt to restore MULTI_USER if we toggled SINGLE_USER
			if ($didSingleUser)
			{
				try
				{
					$tsMU = "ALTER DATABASE $qDb SET MULTI_USER;"
					if ($InvokeSupportsConnString) { Invoke-Sqlcmd -ConnectionString $connStr -Query $tsMU -QueryTimeout 60 -ErrorAction Stop | Out-Null }
					else
					{
						$cnF = New-Object System.Data.SqlClient.SqlConnection $connStr
						try { $cnF.Open(); $cmdF = $cnF.CreateCommand(); $cmdF.CommandText = $tsMU; $cmdF.CommandTimeout = 60; [void]$cmdF.ExecuteNonQuery() }
						finally
						{
							try { $cnF.Close() }
							catch { }
							try { $cnF.Dispose() }
							catch { }
						}
					}
				}
				catch
				{
					Write_Log ("WARN: MULTI_USER restore failed: {0}" -f $_.Exception.Message) "yellow"
				}
			}

			# Best-effort: if we set EMERGENCY, try to clear it by setting ONLINE
			if ($didEmergency)
			{
				try
				{
					$tsOn = "ALTER DATABASE $qDb SET ONLINE;"
					if ($InvokeSupportsConnString) { Invoke-Sqlcmd -ConnectionString $connStr -Query $tsOn -QueryTimeout 60 -ErrorAction Stop | Out-Null }
					else
					{
						$cnF2 = New-Object System.Data.SqlClient.SqlConnection $connStr
						try { $cnF2.Open(); $cmdF2 = $cnF2.CreateCommand(); $cmdF2.CommandText = $tsOn; $cmdF2.CommandTimeout = 60; [void]$cmdF2.ExecuteNonQuery() }
						finally
						{
							try { $cnF2.Close() }
							catch { }
							try { $cnF2.Dispose() }
							catch { }
						}
					}
				}
				catch
				{
					# Not fatal; just log
					Write_Log ("WARN: ONLINE restore after EMERGENCY failed: {0}" -f $_.Exception.Message) "yellow"
				}
			}
		}

		# ----------------------------------------
		# Snapshot state AFTER repair
		# ----------------------------------------
		$stateAfter = 'Unknown'
		try
		{
			$stateAfter = & $getDatabaseState $connStr $dbName
		}
		catch { }

		if ([string]::IsNullOrWhiteSpace($stateAfter)) { $stateAfter = 'Unknown' }
		Write_Log ("DB state after:  {0}" -f $stateAfter) "gray"

		# ----------------------------------------
		# Summarize lane result
		# ----------------------------------------
		$result = 'OK'
		if (-not $ok -and $Level -eq 'Audit')
		{
			$result = 'AuditFoundErrorsOrFailed'
		}
		elseif (-not $ok)
		{
			$result = 'Failed'
		}
		elseif ([string]::IsNullOrWhiteSpace($stateAfter) -or $stateAfter.ToUpper() -ne 'ONLINE')
		{
			$ok = $false
			if ($Level -eq 'Audit') { $result = 'AuditFinishedButStateNotOnline' }
			else { $result = 'PostRepairStateNotOnline' }
		}

		$laneSummary.Add([pscustomobject]@{
				Lane	    = $laneKeyP
				Machine	    = $machineName
				Protocol    = $proto
				DBServer    = $dbServer
				DBName	    = $dbName
				Level	    = $Level
				QuickRepair = $usedQuick
				DeepRepair  = $usedDeep
				Result	    = $result
			})

		Write_Log ("Lane {0} | DB {1} -> Level={2}, Protocol={3}, Quick={4}, Deep={5}, Result={6}" -f $laneKeyP, $dbName, $Level, $proto, $usedQuick, $usedDeep, $result) "gray"
	}

	# ----------------------------------------
	# Summary table + finish banner
	# ----------------------------------------
	Write_Log "`r`n================ Lane DB Repair Summary ================" "blue"
	if ($laneSummary.Count -gt 0)
	{
		Write_Log ((
				$laneSummary |
				Sort-Object Lane |
				Format-Table -AutoSize Lane, Machine, Protocol, DBServer, DBName, Level, QuickRepair, DeepRepair, Result |
				Out-String
			)) "gray"
	}
	else
	{
		Write_Log "No lanes processed." "yellow"
	}

	Write_Log "==================== Repair_LOC_Databases_On_Lanes Function Completed ====================" "blue"
}
