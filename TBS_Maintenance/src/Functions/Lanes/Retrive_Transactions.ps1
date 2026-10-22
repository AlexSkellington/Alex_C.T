# ===================================================================================================
#                              FUNCTION: Retrive_Transactions
# ---------------------------------------------------------------------------------------------------
# Description:
#   PowerShell port of trs_clt_reprocess.sqi: rebuilds the per-lane transaction list table
#   (SAL_HDR_SUS@TER) from SAL_HDR for a chosen date range and transaction range. Connects
#   per the Start_Lane_Protocol_Jobs preference (TCP / Named Pipes / File); lanes on File
#   protocol or any SQL failure fall back to writing trs_clt_reprocess.sqi to the lane XF folder.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#   - ForceSQI     [switch]: skip SQL and write the .sqi file for every lane
#
# Variables:
#   - $script:OfficePath: Storeman office root (resolved and back-filled)
#   - $script:FunctionResults['SqlModuleName']: SQL module to import when Invoke-Sqlcmd is missing
#
# Writes:
#   - XF<Store><Lane>\trs_clt_reprocess.sqi on File-protocol / failed lanes
#
# Returns:
#   PSCustomObject with a per-lane Summary (method, deleted/inserted counts).
# ===================================================================================================

function Retrive_Transactions
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[switch]$ForceSQI
	)

	# -----------------------------
	# Logging fallback (PS 5.1 safe)
	# -----------------------------
	$hasWriteLog = [bool](Get-Command Write_Log -ErrorAction SilentlyContinue)
	$log = {
		param ([string]$Msg,
			[string]$Color = "white")
		if ($hasWriteLog) { Write_Log $Msg $Color }
		else { Write-Host $Msg }
	}

	& $log "`r`n==================== Starting Retrive_Transactions ====================`r`n" "blue"

	# -----------------------------
	# Resolve OfficePath
	# -----------------------------
	$office = $null
	if ($script:OfficePath -and (Test-Path $script:OfficePath)) { $office = $script:OfficePath }
	elseif ($global:OfficePath -and (Test-Path $global:OfficePath)) { $office = $global:OfficePath }
	elseif ($OfficePath -and (Test-Path $OfficePath)) { $office = $OfficePath }
	else
	{
		$office = $script:StoremanOfficeRoots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
		if (-not $office) { $office = $script:PrimaryStoremanOfficePath }
	}

	if (-not $script:OfficePath -or -not (Test-Path $script:OfficePath)) { $script:OfficePath = $office }
	$OfficePath = $script:OfficePath

	$st = $StoreNumber.Trim()
	if ($st -match '^\d+$') { $st = $st.PadLeft(3, '0') }

	# -----------------------------
	# Ensure Invoke-Sqlcmd exists
	#   If unavailable, continue in SQI-only mode
	# -----------------------------
	$canUseSql = $true
	if (-not (Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue))
	{
		$modName = $null
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('SqlModuleName'))
		{
			$modName = $script:FunctionResults['SqlModuleName']
		}
		foreach ($m in @($modName, "SqlServer", "SQLPS") | Where-Object { $_ })
		{
			try { Import-Module $m -ErrorAction Stop; break }
			catch { }
		}
	}
	$invokeSql = Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue
	if (-not $invokeSql)
	{
		$canUseSql = $false
		& $log "Invoke-Sqlcmd not available. SQI mode will be used for all selected lanes." "yellow"
	}
	# SqlServer 22+ encrypts by default; lanes use self-signed certificates
	$sqlExtra = @{ }
	if ($invokeSql -and $invokeSql.Parameters.ContainsKey('TrustServerCertificate')) { $sqlExtra['TrustServerCertificate'] = $true }

	if (-not (Get-Command Get_All_Lanes_Database_Info -ErrorAction SilentlyContinue))
	{
		& $log "Get_All_Lanes_Database_Info not found. Cannot guarantee correct DB connections." "red"
		return
	}

	# -----------------------------
	# UI: Date range + transaction range (FIL1/FIL2)
	# -----------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "TRS CLT Reprocess - Select Range"
	$form.Size = New-Object System.Drawing.Size(420, 240)
	$form.StartPosition = "CenterScreen"
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.FormBorderStyle = 'FixedDialog'

	$lblStart = New-Object System.Windows.Forms.Label
	$lblStart.Text = "Start Date:"
	$lblStart.Location = New-Object System.Drawing.Point(12, 20)
	$lblStart.AutoSize = $true
	$form.Controls.Add($lblStart)

	$dtStart = New-Object System.Windows.Forms.DateTimePicker
	$dtStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$dtStart.Location = New-Object System.Drawing.Point(120, 16)
	$dtStart.Width = 120
	$form.Controls.Add($dtStart)

	$lblStop = New-Object System.Windows.Forms.Label
	$lblStop.Text = "Stop Date:"
	$lblStop.Location = New-Object System.Drawing.Point(12, 55)
	$lblStop.AutoSize = $true
	$form.Controls.Add($lblStop)

	$dtStop = New-Object System.Windows.Forms.DateTimePicker
	$dtStop.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$dtStop.Location = New-Object System.Drawing.Point(120, 51)
	$dtStop.Width = 120
	$form.Controls.Add($dtStop)

	$lblFil1 = New-Object System.Windows.Forms.Label
	$lblFil1.Text = "Tran Start (FIL1):"
	$lblFil1.Location = New-Object System.Drawing.Point(12, 92)
	$lblFil1.AutoSize = $true
	$form.Controls.Add($lblFil1)

	$numFil1 = New-Object System.Windows.Forms.NumericUpDown
	$numFil1.Location = New-Object System.Drawing.Point(160, 88)
	$numFil1.Width = 140
	$numFil1.Minimum = 0
	$numFil1.Maximum = 99999999
	$numFil1.Value = 0
	$form.Controls.Add($numFil1)

	$lblFil2 = New-Object System.Windows.Forms.Label
	$lblFil2.Text = "Tran Stop (FIL2):"
	$lblFil2.Location = New-Object System.Drawing.Point(12, 127)
	$lblFil2.AutoSize = $true
	$form.Controls.Add($lblFil2)

	$numFil2 = New-Object System.Windows.Forms.NumericUpDown
	$numFil2.Location = New-Object System.Drawing.Point(160, 123)
	$numFil2.Width = 140
	$numFil2.Minimum = 0
	$numFil2.Maximum = 99999999
	$numFil2.Value = 99999999
	$form.Controls.Add($numFil2)

	$btnOk = New-Object System.Windows.Forms.Button
	$btnOk.Text = "OK"
	$btnOk.Location = New-Object System.Drawing.Point(95, 165)
	$btnOk.Add_Click({
			if ([int]$numFil1.Value -gt [int]$numFil2.Value)
			{
				[System.Windows.Forms.MessageBox]::Show(
					"FIL1 must be <= FIL2.",
					"Invalid Range",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Warning
				) | Out-Null
				return
			}

			$form.Tag = @{
				Start = $dtStart.Value
				Stop  = $dtStop.Value
				FIL1  = [int]$numFil1.Value
				FIL2  = [int]$numFil2.Value
			}
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})
	$form.Controls.Add($btnOk)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(210, 165)
	$btnCancel.Add_Click({
			$form.Tag = $null
			$form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$form.Close()
		})
	$form.Controls.Add($btnCancel)

	$form.AcceptButton = $btnOk
	$form.CancelButton = $btnCancel

	$dlg = $form.ShowDialog()
	if ($dlg -ne [System.Windows.Forms.DialogResult]::OK -or -not $form.Tag)
	{
		& $log "User cancelled range selection." "yellow"
		& $log "==================== Retrive_Transactions Completed ====================" "blue"
		return
	}

	$startDate = $form.Tag.Start.Date
	$stopDate = $form.Tag.Stop.Date
	$startDT = $startDate
	$stopDT = $stopDate.AddDays(1).AddSeconds(-1)

	$fil1 = [int]$form.Tag.FIL1
	$fil2 = [int]$form.Tag.FIL2

	# ALWAYS clear (no duplicates allowed)
	$doClear = $true

	& $log ("Range: {0} 00:00:00  ->  {1} 23:59:59 | FIL1={2} FIL2={3} | ClearExisting=True (forced)" -f `
		$startDate.ToString("MM/dd/yyyy"), $stopDate.ToString("MM/dd/yyyy"), $fil1, $fil2) "green"

	# -----------------------------
	# Lane selection
	# -----------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Lanes to Build Transaction List"
	if (-not $selection -or -not $selection.ContainsKey('Lanes') -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		& $log "Lane selection cancelled / empty." "yellow"
		return
	}

	$lanes = @($selection.Lanes)
	& $log ("Target lanes: {0}" -f ($lanes -join ', ')) "yellow"

	# -----------------------------
	# Processing
	# -----------------------------
	$summary = New-Object System.Collections.Generic.List[object]

	foreach ($laneRaw in $lanes)
	{
		$l = $laneRaw.ToString().Trim()
		if ($l -match '^\d+$') { $l = ("{0:D3}" -f [int]$l) }
		$lane = $l

		$laneInfo = $null
		try { $laneInfo = Get_All_Lanes_Database_Info -LaneNumber $lane }
		catch { $laneInfo = $null }

		if ($laneInfo -and ($laneInfo -isnot [hashtable]))
		{
			$tmp = @{ }
			foreach ($p in $laneInfo.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
			$laneInfo = $tmp
		}

		$dbName = $null
		if ($laneInfo -and $laneInfo.ContainsKey('DBName')) { $dbName = [string]$laneInfo['DBName'] }

		$machine = $null
		if ($laneInfo -and $laneInfo.ContainsKey('MachineName')) { $machine = [string]$laneInfo['MachineName'] }
		if ([string]::IsNullOrWhiteSpace($machine) -and $script:FunctionResults -and $script:FunctionResults.ContainsKey('LaneNumToMachineName'))
		{
			try { $machine = [string]$script:FunctionResults['LaneNumToMachineName'][$lane] }
			catch { }
		}

		$preferredProtocolRaw = $null
		if ($script:LaneProtocols)
		{
			$tryKeys = @($lane, $laneRaw)
			if ($machine) { $tryKeys += @($machine, $machine.ToLower()) }
			foreach ($k in $tryKeys)
			{
				if ($script:LaneProtocols.ContainsKey($k)) { $preferredProtocolRaw = [string]$script:LaneProtocols[$k]; break }
			}
		}
		$protoNorm = if ($preferredProtocolRaw) { $preferredProtocolRaw.Trim().ToLower() }
		else { $null }

		$useSql = $canUseSql
		if ($ForceSQI) { $useSql = $false }
		elseif ($protoNorm -in @('file', 'offline')) { $useSql = $false }
		elseif ([string]::IsNullOrWhiteSpace($dbName)) { $useSql = $false }

		$transListTable = "SAL_HDR_SUS$lane"
		$targetTbl = "dbo.$transListTable"

		$startIso = $startDT.ToString("yyyy-MM-ddTHH:mm:ss")
		$stopIso = $stopDT.ToString("yyyy-MM-ddTHH:mm:ss")

		$sql = @"
SET NOCOUNT ON;

DECLARE @StartDate datetime = CONVERT(datetime, '$startIso', 126);
DECLARE @StopDate  datetime = CONVERT(datetime, '$stopIso', 126);
DECLARE @Deleted int = 0;
DECLARE @Inserted int = 0;

IF OBJECT_ID('$targetTbl','U') IS NULL
BEGIN
    SELECT TOP 0 * INTO $targetTbl FROM dbo.SAL_HDR;
END

DELETE FROM $targetTbl;
SET @Deleted = @@ROWCOUNT;

INSERT INTO $targetTbl
SELECT *
FROM dbo.SAL_HDR
WHERE F1067 = 'CLOSE'
  AND F254 >= @StartDate
  AND F254 <= @StopDate
  AND F1032 >= $fil1
  AND F1032 <= $fil2;

SET @Inserted = @@ROWCOUNT;

SELECT @Deleted AS DeletedRows, @Inserted AS InsertedRows;
"@

		$method = $null
		$usedProtoDisplay = $preferredProtocolRaw
		$inserted = $null
		$deleted = $null
		$sqiPath = $null
		$errText = $null

		if ($useSql)
		{
			$npServer = if ($laneInfo -and $laneInfo.ContainsKey('NamedPipes')) { [string]$laneInfo['NamedPipes'] }
			else { $null }
			$tcpServer = if ($laneInfo -and $laneInfo.ContainsKey('TcpServer')) { [string]$laneInfo['TcpServer'] }
			else { $null }

			$attempts = @()
			if ($protoNorm -eq 'tcp')
			{
				$attempts += [pscustomobject]@{ Proto = 'TCP'; Server = $tcpServer }
				$attempts += [pscustomobject]@{ Proto = 'Named Pipes'; Server = $npServer }
			}
			elseif ($protoNorm -eq 'named pipes' -or $protoNorm -match 'pipe')
			{
				$attempts += [pscustomobject]@{ Proto = 'Named Pipes'; Server = $npServer }
				$attempts += [pscustomobject]@{ Proto = 'TCP'; Server = $tcpServer }
			}
			else
			{
				$attempts += [pscustomobject]@{ Proto = 'Named Pipes'; Server = $npServer }
				$attempts += [pscustomobject]@{ Proto = 'TCP'; Server = $tcpServer }
			}

			$connected = $false
			foreach ($a in $attempts)
			{
				if ([string]::IsNullOrWhiteSpace($a.Server)) { continue }

				try
				{
					$res = Invoke-Sqlcmd -ServerInstance $a.Server -Database $dbName -Query $sql -QueryTimeout 120 -ErrorAction Stop @sqlExtra

					# Robust parse
					$row = $null
					if ($res -is [System.Data.DataSet])
					{
						if ($res.Tables.Count -gt 0 -and $res.Tables[0].Rows.Count -gt 0) { $row = $res.Tables[0].Rows[0] }
					}
					elseif ($res -is [System.Data.DataTable])
					{
						if ($res.Rows.Count -gt 0) { $row = $res.Rows[0] }
					}
					elseif ($res -is [System.Array])
					{
						if ($res.Count -gt 0) { $row = $res[0] }
					}
					else { $row = $res }

					if ($row -is [System.Data.DataRow])
					{
						if ($row.Table.Columns.Contains('InsertedRows')) { $inserted = [int]$row['InsertedRows'] }
						if ($row.Table.Columns.Contains('DeletedRows')) { $deleted = [int]$row['DeletedRows'] }
					}
					elseif ($row)
					{
						try { if ($row.PSObject.Properties.Match('InsertedRows').Count -gt 0) { $inserted = [int]$row.InsertedRows } }
						catch { }
						try { if ($row.PSObject.Properties.Match('DeletedRows').Count -gt 0) { $deleted = [int]$row.DeletedRows } }
						catch { }
					}

					$connected = $true
					$method = "SQL"
					$usedProtoDisplay = $a.Proto
					break
				}
				catch
				{
					$errText = $_.Exception.Message
					$connected = $false
				}
			}

			if (-not $connected)
			{
				& $log ("SQL failed on lane {0} ({1}). Falling back to SQI. Error: {2}" -f $lane, $machine, $errText) "yellow"
				$useSql = $false
			}
		}

		if (-not $useSql)
		{
			$method = "SQI"

			$xfDir = Join-Path $office ("XF{0}{1}" -f $st, $lane)
			if (-not (Test-Path $xfDir))
			{
				try { New-Item -Path $xfDir -ItemType Directory -Force | Out-Null }
				catch { }
			}

			$sqiPath = Join-Path $xfDir "trs_clt_reprocess.sqi"

			$startFmt = $startDate.ToString("MM/dd/yyyy")
			$stopFmt = $stopDate.ToString("MM/dd/yyyy")

			$sqi = New-Object System.Text.StringBuilder
			[void]$sqi.AppendLine("@WIZSET(DETAIL=D);")
			[void]$sqi.AppendLine("@WIZINIT;")
			[void]$sqi.AppendLine("@WIZDATES(START=$startFmt,STOP=$stopFmt);")
			[void]$sqi.AppendLine("@WIZSET(FIL1=$fil1);")
			[void]$sqi.AppendLine("@WIZSET(FIL2=$fil2);")
			[void]$sqi.AppendLine("")
			[void]$sqi.AppendLine("@WIZRPL(TRANS_LIST=SAL_HDR_SUS@TER);")
			[void]$sqi.AppendLine("@CREATE(@WIZGET(TRANS_LIST),HDRSAL);")
			[void]$sqi.AppendLine("")
			# Always clear (forced)
			[void]$sqi.AppendLine("DELETE FROM @WIZGET(TRANS_LIST);")
			[void]$sqi.AppendLine("")
			[void]$sqi.AppendLine("INSERT INTO @WIZGET(TRANS_LIST) SELECT @DBFLD(@WIZGET(TRANS_LIST)) FROM SAL_HDR@WIZGET(TRANS_LOCAL)")
			[void]$sqi.AppendLine("WHERE F1067='CLOSE' and F254>='@WIZGET(START)' and F254<='@WIZGET(STOP)' AND")
			[void]$sqi.AppendLine("F1032>=@WIZGET(FIL1) and F1032<=@WIZGET(FIL2);")

			try
			{
				$sqi.ToString() | Out-File -FilePath $sqiPath -Encoding ASCII -Force
				try { [System.IO.File]::SetAttributes($sqiPath, [System.IO.FileAttributes]::Normal) }
				catch { }
				$usedProtoDisplay = "File"
				# We can't know the rowcount from SQI here
				$deleted = "N/A"
				& $log ("SQI deployed for lane {0} -> {1}" -f $lane, $sqiPath) "green"
			}
			catch
			{
				$errText = $_.Exception.Message
				& $log ("Failed to write SQI for lane {0} ({1}): {2}" -f $lane, $machine, $errText) "red"
			}
		}

		$protoOut = $usedProtoDisplay
		if ([string]::IsNullOrWhiteSpace($protoOut)) { $protoOut = $preferredProtocolRaw }
		if ([string]::IsNullOrWhiteSpace($protoOut)) { $protoOut = "Unknown" }

		$summary.Add([pscustomobject]@{
				Lane	 = $lane
				Machine  = $machine
				Protocol = $protoOut
				Method   = $method
				Deleted  = $deleted
				Inserted = $inserted
				SQIPath  = $sqiPath
				Error    = $errText
			})
	}

	& $log "`r`n----- Build Summary -----" "yellow"
	try { $summary | Sort-Object Lane | Format-Table -AutoSize | Out-String | ForEach-Object { & $log $_ "blue" } }
	catch { & $log ("Could not format summary table: {0}" -f $_.Exception.Message) "red" }

	$totInserted = ($summary | Where-Object { $_.Method -eq 'SQL' -and $_.Inserted -ne $null } | Measure-Object Inserted -Sum).Sum
	$totDeleted = ($summary | Where-Object { $_.Method -eq 'SQL' -and $_.Deleted -ne $null -and $_.Deleted -is [int] } | Measure-Object Deleted -Sum).Sum
	& $log ("Totals (SQL lanes): Deleted={0}  Inserted={1}" -f ($totDeleted -as [int]), ($totInserted -as [int])) "green"

	& $log "`r`n==================== Retrive_Transactions Completed ====================`r`n" "blue"

	return [pscustomobject]@{ Summary = $summary }
}
