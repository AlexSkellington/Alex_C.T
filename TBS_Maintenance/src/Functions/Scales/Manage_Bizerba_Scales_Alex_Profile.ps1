# ===================================================================================================
#                              FUNCTION: Manage_Bizerba_Scales_Alex_Profile
# ---------------------------------------------------------------------------------------------------
# Description:
#   All-in-one Bizerba scale manager that modifies ONLY the dynamic config (BMS.xml stays
#   untouched): BMS gets ONE EventProcess per event (SEND/DEL+SEND/STATUS/SALES) with one
#   InputPath (shared deployment folder) and Terminals = only the selected scales.
#
# Parameters:
#   - StoreNumber                         [string] (required): store to process
#   - BmsRootPath                         [string]: BMS root (default $script:BmsRootPath)
#   - BmsConfigName / BmsExeName          [string]: defaults "BMS.xml" / "BMS.exe"
#   - AlexDynamicConfigName               [string]: default "dynamic_config_Alex_C.T.xml"
#   - ScaleManagementAppPath              [string]: default $script:ScaleManagementAppExePath
#   - ScaleManagementAppUpdateSpecialsPath [string]: default $script:ScaleManagementAppUpdateSpecialsPath
#
# Variables:
#   - $script:FunctionResults['ScaleCodeToIPInfo']: scale IPs (from Retrieve_Nodes)
#
# Notes:
#   - Captures ATST/PLST exports from GLOBAL toBizerba into memory and deletes them immediately.
#   - Deploys SCL files once to Scales_Deployment\toBizerba; creates extensionless triggers
#     (BatchSendFull, BatchScl, DBatch, DBatchScl, Status, Sales) in the original toBizerba.
#   - Runs: BMS.exe -config .\BMS.xml -dynamic .\dynamic_config_Alex_C.T.xml
# ===================================================================================================

function Manage_Bizerba_Scales_Alex_Profile
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$BmsRootPath = $script:BmsRootPath,
		[Parameter(Mandatory = $false)]
		[string]$BmsConfigName = "BMS.xml",
		[Parameter(Mandatory = $false)]
		[string]$AlexDynamicConfigName = "dynamic_config_Alex_C.T.xml",
		[Parameter(Mandatory = $false)]
		[string]$BmsExeName = "BMS.exe",
		[Parameter(Mandatory = $false)]
		[string]$ScaleManagementAppPath = $script:ScaleManagementAppExePath,
		[Parameter(Mandatory = $false)]
		[string]$ScaleManagementAppUpdateSpecialsPath = $script:ScaleManagementAppUpdateSpecialsPath
	)

	# ------------------------------------------------------------------------------------------------
	# 0) Banner
	# ------------------------------------------------------------------------------------------------
	Write_Log "`r`n==================== Manage Bizerba Scales (Dynamic Only) ====================`r`n" "blue"
	Write_Log "Store Number : $StoreNumber" "blue"
	Write_Log "BMS Root Path: $BmsRootPath" "blue"

	# Normalize base BMS root (original working directory)
	$BmsRootNorm = $BmsRootPath.TrimEnd('\') + "\"

	# Global (normal service) locations - KEEP and reuse
	$BmsExePath = Join-Path $BmsRootNorm $BmsExeName
	$BmsConfigPath = Join-Path $BmsRootNorm $BmsConfigName
	$GlobalToBizerba = Join-Path $BmsRootNorm "toBizerba"
	$GlobalBackup = Join-Path $BmsRootNorm "Backup"
	$GlobalLogPath = Join-Path $BmsRootNorm "log"
	$GlobalFromBiz = Join-Path $BmsRootNorm "fromBizerba"
	$AlexDynamicPath = Join-Path $BmsRootNorm $AlexDynamicConfigName
	$GlobalToBizerbaDir = $GlobalToBizerba.TrimEnd('\') + "\"

	# Shared deployment folder (InputPath for BMS SEND)
	$ScaleDeployRoot = Join-Path $BmsRootNorm "Scales_Deployment"
	$DeployInputPath = Join-Path $ScaleDeployRoot "toBizerba"

	# ------------------------------------------------------------------------------------------------
	# 1) Validate core paths/files
	# ------------------------------------------------------------------------------------------------
	if (-not (Test-Path -LiteralPath $BmsExePath))
	{
		Write_Log "ERROR: BMS.exe not found at: $BmsExePath" "red"
		return
	}
	if (-not (Test-Path -LiteralPath $BmsConfigPath))
	{
		Write_Log "ERROR: BMS config not found at: $BmsConfigPath" "red"
		return
	}
	if (-not (Test-Path -LiteralPath $GlobalToBizerba))
	{
		Write_Log "ERROR: Global toBizerba folder not found at: $GlobalToBizerba" "red"
		return
	}

	foreach ($p in @($ScaleDeployRoot, $DeployInputPath))
	{
		if (-not (Test-Path -LiteralPath $p))
		{
			try
			{
				New-Item -ItemType Directory -Path $p -Force | Out-Null
				Write_Log "Created deployment folder: $p" "green"
			}
			catch
			{
				Write_Log "ERROR: Failed to create deployment folder '$p': $_" "red"
				return
			}
		}
	}

	# ------------------------------------------------------------------------------------------------
	# 2) Operation Selection (GUI) - ONLY 3 OPTIONS + optional "Delete first then Deploy"
	# ------------------------------------------------------------------------------------------------
	try
	{
		Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
		Add-Type -AssemblyName System.Drawing -ErrorAction Stop
	}
	catch
	{
		Write_Log "ERROR: Failed to load WinForms assemblies: $_" "red"
		return
	}

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Bizerba Scales (Dynamic Only) - Operation Selection"
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.Size = New-Object System.Drawing.Size(560, 320)

	$grpOp = New-Object System.Windows.Forms.GroupBox
	$grpOp.Text = "Operation"
	$grpOp.Location = New-Object System.Drawing.Point(10, 10)
	$grpOp.Size = New-Object System.Drawing.Size(530, 140)
	$form.Controls.Add($grpOp)

	$rdoSendAll = New-Object System.Windows.Forms.RadioButton
	$rdoSendAll.Text = "Send ALL items (Full Deploy)  [Trigger: BatchSendFull]"
	$rdoSendAll.Location = New-Object System.Drawing.Point(15, 25)
	$rdoSendAll.AutoSize = $true
	$rdoSendAll.Checked = $true
	$grpOp.Controls.Add($rdoSendAll)

	$rdoRecent = New-Object System.Windows.Forms.RadioButton
	$rdoRecent.Text = "Send RECENT price changes  [Trigger: BatchScl / UpdateSpecials capture]"
	$rdoRecent.Location = New-Object System.Drawing.Point(15, 55)
	$rdoRecent.AutoSize = $true
	$grpOp.Controls.Add($rdoRecent)

	$rdoDeleteOnly = New-Object System.Windows.Forms.RadioButton
	$rdoDeleteOnly.Text = "DELETE ALL items from selected scales (Delete-only)  [Trigger: DBatch]"
	$rdoDeleteOnly.Location = New-Object System.Drawing.Point(15, 85)
	$rdoDeleteOnly.AutoSize = $true
	$grpOp.Controls.Add($rdoDeleteOnly)

	$grpSeq = New-Object System.Windows.Forms.GroupBox
	$grpSeq.Text = "Optional Sequence"
	$grpSeq.Location = New-Object System.Drawing.Point(10, 155)
	$grpSeq.Size = New-Object System.Drawing.Size(530, 60)
	$form.Controls.Add($grpSeq)

	$chkDeleteThenDeploy = New-Object System.Windows.Forms.CheckBox
	$chkDeleteThenDeploy.Text = "Run DELETE ALL first, then run the selected deploy (BMS runs twice: DBatch -> Batch*)"
	$chkDeleteThenDeploy.Location = New-Object System.Drawing.Point(15, 25)
	$chkDeleteThenDeploy.AutoSize = $true
	$grpSeq.Controls.Add($chkDeleteThenDeploy)

	# Enable/disable sequence checkbox based on which op is selected
	$handlerUpdateSeq = {
		if ($rdoDeleteOnly.Checked)
		{
			$chkDeleteThenDeploy.Checked = $false
			$chkDeleteThenDeploy.Enabled = $false
		}
		else
		{
			$chkDeleteThenDeploy.Enabled = $true
		}
	}
	$rdoSendAll.Add_CheckedChanged($handlerUpdateSeq)
	$rdoRecent.Add_CheckedChanged($handlerUpdateSeq)
	$rdoDeleteOnly.Add_CheckedChanged($handlerUpdateSeq)

	$lblInfo = New-Object System.Windows.Forms.Label
	$lblInfo.AutoSize = $true
	$lblInfo.Location = New-Object System.Drawing.Point(10, 225)
	$lblInfo.Text = "Note: Delete-only does NOT export. Deploy options build/capture ATST/PLST and then trigger Batch*."
	$form.Controls.Add($lblInfo)

	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(175, 245)
	$btnOK.Size = New-Object System.Drawing.Size(100, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(290, 245)
	$btnCancel.Size = New-Object System.Drawing.Size(100, 32)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)

	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Cancelled by user at operation selection." "yellow"
		return
	}

	# Translate selection into modes
	$mode = ""
	if ($rdoSendAll.Checked) { $mode = "FULL" }
	elseif ($rdoRecent.Checked) { $mode = "RECENT" }
	elseif ($rdoDeleteOnly.Checked) { $mode = "DELETE" }

	$deleteFirstThenDeploy = $false
	if ($mode -ne "DELETE" -and $chkDeleteThenDeploy.Checked) { $deleteFirstThenDeploy = $true }

	# Derive actions
	$sendFull = ($mode -eq "FULL")
	$sendRecent = ($mode -eq "RECENT")
	$deleteOnly = ($mode -eq "DELETE")

	$opSummary = ""
	if ($deleteOnly) { $opSummary = "DeleteAllItems" }
	elseif ($sendFull -and $deleteFirstThenDeploy) { $opSummary = "DeleteAllItems THEN SendAllItems" }
	elseif ($sendRecent -and $deleteFirstThenDeploy) { $opSummary = "DeleteAllItems THEN SendRecentPriceChanges" }
	elseif ($sendFull) { $opSummary = "SendAllItems" }
	elseif ($sendRecent) { $opSummary = "SendRecentPriceChanges" }

	Write_Log "Selected: $opSummary" "blue"

	# ------------------------------------------------------------------------------------------------
	# 3) Select Bizerba scales via existing node selector
	# ------------------------------------------------------------------------------------------------
	Write_Log "Opening scale selection..." "blue"

	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Scale" -OnlyBizerbaScales -Title "Select Bizerba Scales (Dynamic Only)"
	if (-not $selection -or -not $selection.Scales -or $selection.Scales.Count -eq 0)
	{
		Write_Log "No scales selected. Aborting." "yellow"
		return
	}

	$selectedScaleCodes = @($selection.Scales)

	# ------------------------------------------------------------------------------------------------
	# 4) Resolve scale metadata (ScaleCodeToIPInfo)
	# ------------------------------------------------------------------------------------------------
	if (-not $script:FunctionResults -or -not $script:FunctionResults.ContainsKey('ScaleCodeToIPInfo'))
	{
		Write_Log "ERROR: ScaleCodeToIPInfo not found in FunctionResults. Run Retrieve_Nodes first." "red"
		return
	}

	$scaleCodeToIpInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$selectedScaleMeta = @()

	foreach ($code in $selectedScaleCodes)
	{
		if (-not $scaleCodeToIpInfo.ContainsKey($code))
		{
			Write_Log "WARNING: No metadata found for scale code '$code'. Skipping." "yellow"
			continue
		}

		$info = $scaleCodeToIpInfo[$code]

		# Resolve IP
		$ip = $null
		if ($info.FullIP) { $ip = $info.FullIP }
		elseif ($info.IPNetwork -and $info.IPDevice) { $ip = (([string]$info.IPNetwork).TrimEnd('.') + '.' + [string]$info.IPDevice) }

		if (-not $ip)
		{
			Write_Log "WARNING: No IP resolved for scale code '$code'. Skipping." "yellow"
			continue
		}

		# Resolve Name
		$scaleName = $info.ScaleName
		if (-not $scaleName -or [string]::IsNullOrWhiteSpace($scaleName)) { $scaleName = "Scale $code" }
		$scaleName = $scaleName.Trim()

		$selectedScaleMeta += [PSCustomObject]@{
			ScaleCode = "$code"
			ScaleName = $scaleName
			IPAddress = $ip
		}
	}

	if ($selectedScaleMeta.Count -eq 0)
	{
		Write_Log "ERROR: None of the selected scales could be mapped. Aborting." "red"
		return
	}

	$targetLine = ($selectedScaleMeta | ForEach-Object { "{0} ({1})" -f $_.ScaleName, $_.IPAddress }) -join "; "
	Write_Log "Targeting: $targetLine" "green"

	# ------------------------------------------------------------------------------------------------
	# 5) Build export data (FULL from SQL, RECENT via UpdateSpecials capture)
	#     - For DELETE-only: nothing is built/captured.
	# ------------------------------------------------------------------------------------------------
	$atstLines = $null
	$plstLines = $null
	$atstOrigName = $null
	$plstOrigName = $null

	# ------------------------- FULL (Send All Items) -------------------------
	if ($sendFull)
	{
		Write_Log "Export: FULL (Build ATST/PLST from SQL for ALL eligible items) [matches BatchSendFull]" "blue"

		if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }

		$needConn =
		(-not $script:FunctionResults.ContainsKey('ConnectionString')) -or
		[string]::IsNullOrWhiteSpace($script:FunctionResults['ConnectionString']) -or
		($script:FunctionResults['ConnectionString'] -eq 'N/A')

		if ($needConn)
		{
			Write_Log "Store/DB info not loaded yet. Running Get_Store_And_Database_Info..." "yellow"
			Get_Store_And_Database_Info -WinIniPath $WinIniPath -SmsStartIniPath $SmsStartIniPath -StartupIniPath $StartupIniPath -SystemIniPath $SystemIniPath
		}

		$connString = $script:FunctionResults['ConnectionString']
		if ([string]::IsNullOrWhiteSpace($connString) -or $connString -eq 'N/A')
		{
			Write_Log "ERROR: ConnectionString is empty. Cannot query SQL." "red"
			Write_Log "DBSERVER=$($script:FunctionResults['DBSERVER'])  DBNAME=$($script:FunctionResults['DBNAME'])" "yellow"
			return
		}

		Write_Log "Using SQL: $($script:FunctionResults['DBSERVER']) / $($script:FunctionResults['DBNAME'])" "green"

		$sql = @"
;WITH Eligible AS (
    SELECT o.F01
    FROM OBJ_TAB o
    WHERE SUBSTRING(o.F01,1,3) = '$StoreNumber'
      AND SUBSTRING(o.F01,9,5) = '00000'
      AND SUBSTRING(o.F01,4,5) NOT LIKE '%[^0-9]%'   -- digits-only PLU5
),
LatestPrice AS (
    SELECT p1.F01, p1.F30
    FROM PRICE_TAB p1
    INNER JOIN (
        SELECT F01, MAX(F253) AS MaxF253
        FROM PRICE_TAB
        GROUP BY F01
    ) px ON px.F01 = p1.F01 AND px.MaxF253 = p1.F253
),
PosPlu AS (
    SELECT
        p.F01,
        p.F82,
        CASE
            WHEN p.F01 IS NOT NULL
             AND LTRIM(RTRIM(p.F01)) <> ''
             AND LTRIM(RTRIM(p.F01)) NOT LIKE '%[^0-9]%'   -- digits-only
             AND LEN(LTRIM(RTRIM(p.F01))) <= 18            -- safe for BIGINT
            THEN CAST(LTRIM(RTRIM(p.F01)) AS bigint)
            ELSE NULL
        END AS F01_NUM
    FROM POS_TAB p
)
SELECT
    e.F01,
    SUBSTRING(e.F01,4,5) AS PLU5,
    CAST(SUBSTRING(e.F01,4,5) AS int) AS PLU_INT,

    COALESCE(posFull.F82, posPluJoin.F82) AS POS_F82_ProductType,

    pr.F30    AS RetailPriceMoney,
    obj.F29   AS OBJ_F29,
    scl.F1952 AS SCL_F1952,
    scl.F2581 AS SCL_F2581,
    scl.F2582 AS SCL_F2582,
    scl.F263  AS SCL_F263_Field20,
    scl.F258  AS SCL_F258_Field26,
    scl.F272  AS SCL_F272_Field42,
    txt.F297  AS IngredientsText
FROM Eligible e

LEFT JOIN POS_TAB posFull
    ON posFull.F01 = e.F01

LEFT JOIN PosPlu posPluJoin
    ON posPluJoin.F01_NUM = CAST(SUBSTRING(e.F01,4,5) AS bigint)

LEFT JOIN OBJ_TAB     obj ON obj.F01 = e.F01
LEFT JOIN SCL_TAB     scl ON scl.F01 = e.F01
LEFT JOIN SCL_TXT_TAB txt ON txt.F267 = CAST(SUBSTRING(e.F01,4,5) AS int)
LEFT JOIN LatestPrice pr  ON pr.F01  = e.F01
ORDER BY e.F01;
"@

		$dt = New-Object System.Data.DataTable
		try
		{
			$cn = New-Object System.Data.SqlClient.SqlConnection $connString
			$cn.Open()
			$cmd = $cn.CreateCommand()
			$cmd.CommandText = $sql
			$cmd.CommandTimeout = 240
			$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
			[void]$da.Fill($dt)
			$cn.Close()
		}
		catch
		{
			try { if ($cn -and $cn.State -eq 'Open') { $cn.Close() } }
			catch { }
			Write_Log "ERROR: SQL query failed (BatchSendFull build): $($_.Exception.Message)" "red"
			return
		}

		if ($dt.Rows.Count -eq 0)
		{
			Write_Log "ERROR: SQL returned 0 eligible items for Store $StoreNumber." "red"
			return
		}

		$plstOut = New-Object System.Collections.Generic.List[string]
		$atstOut = New-Object System.Collections.Generic.List[string]

		foreach ($r in $dt.Rows)
		{
			$plu5 = [string]$r["PLU5"]
			if ([string]::IsNullOrWhiteSpace($plu5)) { continue }

			$pluInt = 0
			try { $pluInt = [int]$r["PLU_INT"] }
			catch { continue }

			# Field 2 (Product type)
			$field2 = ""
			if ($r["POS_F82_ProductType"] -ne [DBNull]::Value -and $r["POS_F82_ProductType"] -ne $null)
			{
				try { $field2 = [string]([int]$r["POS_F82_ProductType"]) }
				catch { $field2 = "" }
			}

			# Field 10 (price in cents)
			$field10 = ""
			if ($r["RetailPriceMoney"] -ne [DBNull]::Value -and $r["RetailPriceMoney"] -ne $null)
			{
				try
				{
					$priceDec = [decimal]$r["RetailPriceMoney"]
					$field10 = [string]([int][Math]::Round(($priceDec * 100), 0))
				}
				catch { $field10 = "" }
			}

			# Field 20 (blank if NULL)
			$field20 = ""
			if ($r["SCL_F263_Field20"] -ne [DBNull]::Value -and $r["SCL_F263_Field20"] -ne $null)
			{
				try { $field20 = [string]([int]$r["SCL_F263_Field20"]) }
				catch { $field20 = "" }
			}

			# Field 26 (blank if NULL)
			$field26 = ""
			if ($r["SCL_F258_Field26"] -ne [DBNull]::Value -and $r["SCL_F258_Field26"] -ne $null)
			{
				try { $field26 = [string]([int]$r["SCL_F258_Field26"]) }
				catch { $field26 = "" }
			}

			# Field 42 (default 0)
			$field42 = "0"
			if ($r["SCL_F272_Field42"] -ne [DBNull]::Value -and $r["SCL_F272_Field42"] -ne $null)
			{
				try { $field42 = [string]([int][Math]::Round(([double]$r["SCL_F272_Field42"]), 0)) }
				catch { $field42 = "0" }
			}

			# Descriptions joined with @0A
			$descParts = New-Object System.Collections.Generic.List[string]
			$main = ""; if ($r["OBJ_F29"] -ne [DBNull]::Value -and $r["OBJ_F29"]) { $main = ([string]$r["OBJ_F29"]).Trim() }
			$d2 = ""; if ($r["SCL_F1952"] -ne [DBNull]::Value -and $r["SCL_F1952"]) { $d2 = ([string]$r["SCL_F1952"]).Trim() }
			$d3 = ""; if ($r["SCL_F2581"] -ne [DBNull]::Value -and $r["SCL_F2581"]) { $d3 = ([string]$r["SCL_F2581"]).Trim() }
			$d4 = ""; if ($r["SCL_F2582"] -ne [DBNull]::Value -and $r["SCL_F2582"]) { $d4 = ([string]$r["SCL_F2582"]).Trim() }
			if ($main) { [void]$descParts.Add($main) }
			if ($d2) { [void]$descParts.Add($d2) }
			if ($d3) { [void]$descParts.Add($d3) }
			if ($d4) { [void]$descParts.Add($d4) }
			$descBlock = ($descParts -join "@0A")

			# Ingredients
			$ing = ""
			if ($r["IngredientsText"] -ne [DBNull]::Value -and $r["IngredientsText"] -ne $null) { $ing = [string]$r["IngredientsText"] }
			$hasIngredients = (-not [string]::IsNullOrWhiteSpace($ing))

			# Build 61-field PLST with defaults = "0"
			$plst = New-Object string[] 61
			for ($i = 0; $i -lt 61; $i++) { $plst[$i] = "0" }

			# Real BatchSendFull "blank fields"
			$plst[14] = "" # field 15
			$plst[30] = "" # field 31
			$plst[60] = "" # field 61 blank unless ingredients exist

			# Populate fields
			$plst[0] = "0"
			$plst[1] = $field2
			$plst[2] = $plu5
			$plst[7] = "2"
			$plst[8] = "4"
			$plst[9] = $field10

			$plst[19] = $field20
			$plst[25] = $field26

			$plst[41] = $field42
			$plst[42] = $descBlock
			$plst[43] = [string]$pluInt

			if ($hasIngredients) { $plst[60] = $ing }

			[void]$plstOut.Add(($plst -join "|"))

			if ($hasIngredients)
			{
				$atstBody = "^4;^Al;^b+^u+Ingredients:^u-^b- @0A^3;^Al;" + $ing
				[void]$atstOut.Add(("0|{0}|{1}" -f $pluInt, $atstBody))
			}
		}

		$plstLines = @($plstOut)
		$atstLines = @($atstOut)

		if ($plstLines.Count -eq 0)
		{
			Write_Log "ERROR: Built 0 PLST lines. Nothing to deploy." "red"
			return
		}

		$stamp = Get-Date -Format "MMdd.HHmmssfff"
		$plstOrigName = "PLST_SCL$StoreNumber.$stamp`_1.BatchSendFull"
		$atstOrigName = "ATST_SCL$StoreNumber.$stamp`_1.BatchSendFull"

		Write_Log "Built FULL deploy files from SQL (PLST: $($plstLines.Count) line(s), ATST: $($atstLines.Count) line(s))." "green"
	}

	# ------------------------- RECENT (Send Price Changes / UpdateSpecials capture) -------------------------
	if ($sendRecent)
	{
		$scaleExeToRun = $ScaleManagementAppUpdateSpecialsPath
		Write_Log "Export: RECENT price changes (ScaleManagementAppUpdateSpecials) [capture + delete from GLOBAL]" "blue"

		if (-not (Test-Path -LiteralPath $scaleExeToRun))
		{
			Write_Log "ERROR: UpdateSpecials EXE not found: $scaleExeToRun" "red"
			return
		}

		$atstPattern = "ATST_SCL*.BatchScl"
		$plstPattern = "PLST_SCL*.BatchScl"

		Write_Log "Watching GLOBAL toBizerba for export files..." "blue"

		$scaleExeWorkingDir = Split-Path -Parent $scaleExeToRun
		try
		{
			$scaleProc = Start-Process -FilePath $scaleExeToRun -WorkingDirectory $scaleExeWorkingDir -PassThru -ErrorAction Stop
			Write_Log "ScaleManagementAppUpdateSpecials PID: $($scaleProc.Id)" "green"
		}
		catch
		{
			Write_Log "ERROR: Failed to start UpdateSpecials export EXE: $_" "red"
			return
		}

		$atstCaptured = $false
		$plstCaptured = $false
		$maxWaitSeconds = 20
		$pollIntervalMs = 50
		$elapsedMs = 0

		while ($elapsedMs -lt ($maxWaitSeconds * 1000))
		{
			try { $scaleProc.Refresh() }
			catch { }

			if (-not $atstCaptured)
			{
				$atstCandidate = Get-ChildItem -LiteralPath $GlobalToBizerba -Filter $atstPattern -ErrorAction SilentlyContinue |
				Sort-Object LastWriteTime -Descending | Select-Object -First 1

				if ($atstCandidate)
				{
					$atstOrigName = $atstCandidate.Name
					try
					{
						$atstLines = Get-Content -LiteralPath $atstCandidate.FullName -Encoding Default -ErrorAction Stop
						$atstCaptured = $true
						Remove-Item -LiteralPath $atstCandidate.FullName -Force -ErrorAction SilentlyContinue
					}
					catch { Write_Log "ERROR: Failed to capture ATST export: $_" "red" }
				}
			}

			if (-not $plstCaptured)
			{
				$plstCandidate = Get-ChildItem -LiteralPath $GlobalToBizerba -Filter $plstPattern -ErrorAction SilentlyContinue |
				Sort-Object LastWriteTime -Descending | Select-Object -First 1

				if ($plstCandidate)
				{
					$plstOrigName = $plstCandidate.Name
					try
					{
						$plstLines = Get-Content -LiteralPath $plstCandidate.FullName -Encoding Default -ErrorAction Stop
						$plstCaptured = $true
						Remove-Item -LiteralPath $plstCandidate.FullName -Force -ErrorAction SilentlyContinue
					}
					catch { Write_Log "ERROR: Failed to capture PLST export: $_" "red" }
				}
			}

			$scaleExited = $false
			try { $scaleExited = $scaleProc.HasExited }
			catch { $scaleExited = $true }

			if ($atstCaptured -and $plstCaptured) { break }
			if ($scaleExited -and ($atstCaptured -or $plstCaptured)) { break }

			Start-Sleep -Milliseconds $pollIntervalMs
			$elapsedMs += $pollIntervalMs
		}

		$atstCount = if ($atstLines) { $atstLines.Count }
		else { 0 }
		$plstCount = if ($plstLines) { $plstLines.Count }
		else { 0 }

		if ($atstCount -eq 0 -and $plstCount -eq 0)
		{
			Write_Log "ERROR: No export data captured; cannot send recent price changes." "red"
			return
		}

		Write_Log "Captured RECENT export data (ATST: $atstCount line(s), PLST: $plstCount line(s))." "green"
	}

	# ------------------------------------------------------------------------------------------------
	# 6) Build dynamic_config_Alex_C.T.xml
	#     IMPORTANT CHANGE:
	#       - Only TWO events exist now:
	#           1) SEND_DATA (Batch*) -> EP_BMS_ITEM (uses InputPath)
	#           2) DELETE_ALL_ONLY (DBatch*) -> EP_BMS_DELETEALL ONLY (no send)
	# ------------------------------------------------------------------------------------------------
	Write_Log "Building dynamic config for selected scales..." "blue"

	$globalWorkDir = $BmsRootNorm
	$globalBackup = $GlobalBackup.TrimEnd('\') + "\"
	$globalLogDir = $GlobalLogPath.TrimEnd('\') + "\"
	$inputPathXml = $DeployInputPath.TrimEnd('\') + "\"

	try
	{
		$dynDoc = New-Object System.Xml.XmlDocument
		$decl = $dynDoc.CreateXmlDeclaration("1.0", "utf-8", $null)
		[void]$dynDoc.AppendChild($decl)

		$root = $dynDoc.CreateElement("Configuration")
		[void]$dynDoc.AppendChild($root)

		$globalElem = $dynDoc.CreateElement("Global")
		[void]$root.AppendChild($globalElem)

		$logLevelElem = $dynDoc.CreateElement("LogLevel"); $logLevelElem.InnerText = "0"; [void]$globalElem.AppendChild($logLevelElem)
		[void]$globalElem.AppendChild($dynDoc.CreateComment(" 0 for error only; 2 for complete info "))

		$wdElem = $dynDoc.CreateElement("WorkingDirectory"); $wdElem.InnerText = $globalWorkDir; [void]$globalElem.AppendChild($wdElem)
		$backupElem = $dynDoc.CreateElement("BackupPath"); $backupElem.InnerText = $globalBackup; [void]$globalElem.AppendChild($backupElem)
		$backupTimeElem = $dynDoc.CreateElement("BackupTime"); $backupTimeElem.InnerText = "30"; [void]$globalElem.AppendChild($backupTimeElem)
		$logPathElem = $dynDoc.CreateElement("LogPath"); $logPathElem.InnerText = $globalLogDir; [void]$globalElem.AppendChild($logPathElem)

		$termsRoot = $dynDoc.CreateElement("Terminals")
		$attrExisting = $dynDoc.CreateAttribute("existingTerminals"); $attrExisting.Value = "replace"
		[void]$termsRoot.Attributes.Append($attrExisting)
		[void]$root.AppendChild($termsRoot)

		foreach ($m in $selectedScaleMeta)
		{
			$term = $dynDoc.CreateElement("Terminal"); [void]$termsRoot.AppendChild($term)
			$nameElem = $dynDoc.CreateElement("Name"); $nameElem.InnerText = $m.ScaleName; [void]$term.AppendChild($nameElem)
			$ipElem = $dynDoc.CreateElement("TCPIPAddress"); $ipElem.InnerText = $m.IPAddress; [void]$term.AppendChild($ipElem)

			$e = $dynDoc.CreateElement("Active"); $e.InnerText = "Yes"; [void]$term.AppendChild($e)
			$e = $dynDoc.CreateElement("SystemLocalTime"); $e.InnerText = "Yes"; [void]$term.AppendChild($e)
			$e = $dynDoc.CreateElement("AutoStart"); $e.InnerText = "Yes"; [void]$term.AppendChild($e)
			$e = $dynDoc.CreateElement("AutoTransmit"); $e.InnerText = "Yes"; [void]$term.AppendChild($e)
			$e = $dynDoc.CreateElement("BufferTime"); $e.InnerText = "5"; [void]$term.AppendChild($e)
		}

		$eventsElem = $dynDoc.CreateElement("Events")
		$attrTermMap = $dynDoc.CreateAttribute("terminalMappings"); $attrTermMap.Value = "replace"
		[void]$eventsElem.Attributes.Append($attrTermMap)
		[void]$root.AppendChild($eventsElem)

		# ------------------------- SEND_DATA (Batch*) -------------------------
		[void]$eventsElem.AppendChild($dynDoc.CreateComment(" Send Data only (Batch*) "))

		$eventSend = $dynDoc.CreateElement("Event")
		$a = $dynDoc.CreateAttribute("name"); $a.Value = "SEND_DATA"; [void]$eventSend.Attributes.Append($a)
		$a = $dynDoc.CreateAttribute("path"); $a.Value = $GlobalToBizerbaDir; [void]$eventSend.Attributes.Append($a)
		$a = $dynDoc.CreateAttribute("filename"); $a.Value = "Batch*"; [void]$eventSend.Attributes.Append($a)
		[void]$eventsElem.AppendChild($eventSend)

		$eps = $dynDoc.CreateElement("EventProcesses"); [void]$eventSend.AppendChild($eps)
		$epItem = $dynDoc.CreateElement("EventProcess")
		$a = $dynDoc.CreateAttribute("name"); $a.Value = "EP_BMS_ITEM"; [void]$epItem.Attributes.Append($a)
		$a = $dynDoc.CreateAttribute("no"); $a.Value = "1"; [void]$epItem.Attributes.Append($a)
		[void]$eps.AppendChild($epItem)

		$vars = $dynDoc.CreateElement("Variables"); [void]$epItem.AppendChild($vars)
		$v = $dynDoc.CreateElement("Variable"); $a = $dynDoc.CreateAttribute("name"); $a.Value = "InputPath"; [void]$v.Attributes.Append($a); $v.InnerText = $inputPathXml; [void]$vars.AppendChild($v)
		$v = $dynDoc.CreateElement("Variable"); $a = $dynDoc.CreateAttribute("name"); $a.Value = "NutriTemplateNumber"; [void]$v.Attributes.Append($a); $v.InnerText = "0"; [void]$vars.AppendChild($v)
		$v = $dynDoc.CreateElement("Variable"); $a = $dynDoc.CreateAttribute("name"); $a.Value = "ResponseFile"; [void]$v.Attributes.Append($a); $v.InnerText = ($globalLogDir + "Response.log"); [void]$vars.AppendChild($v)

		$terms = $dynDoc.CreateElement("Terminals"); [void]$epItem.AppendChild($terms)
		foreach ($mm in $selectedScaleMeta)
		{
			$n = $dynDoc.CreateElement("Name"); $n.InnerText = $mm.ScaleName; [void]$terms.AppendChild($n)
		}

		# ------------------------- DELETE ALL ONLY (DBatch*) -------------------------
		# Key change: this is DELETE-ONLY. No second EventProcess for EP_BMS_ITEM.
		[void]$eventsElem.AppendChild($dynDoc.CreateComment(" Delete entire PLU/TEXT/NUTRITION table ONLY (DBatch*) "))

		$eventDelOnly = $dynDoc.CreateElement("Event")
		$a = $dynDoc.CreateAttribute("name"); $a.Value = "DELETE_ALL_ONLY"; [void]$eventDelOnly.Attributes.Append($a)
		$a = $dynDoc.CreateAttribute("path"); $a.Value = $GlobalToBizerbaDir; [void]$eventDelOnly.Attributes.Append($a)
		$a = $dynDoc.CreateAttribute("filename"); $a.Value = "DBatch*"; [void]$eventDelOnly.Attributes.Append($a)
		[void]$eventsElem.AppendChild($eventDelOnly)

		$epsDel = $dynDoc.CreateElement("EventProcesses"); [void]$eventDelOnly.AppendChild($epsDel)
		$epDel = $dynDoc.CreateElement("EventProcess")
		$a = $dynDoc.CreateAttribute("name"); $a.Value = "EP_BMS_DELETEALL"; [void]$epDel.Attributes.Append($a)
		$a = $dynDoc.CreateAttribute("no"); $a.Value = "1"; [void]$epDel.Attributes.Append($a)
		[void]$epsDel.AppendChild($epDel)

		$varsDel = $dynDoc.CreateElement("Variables"); [void]$epDel.AppendChild($varsDel)
		$v = $dynDoc.CreateElement("Variable"); $a = $dynDoc.CreateAttribute("name"); $a.Value = "ResponseFile"; [void]$v.Attributes.Append($a); $v.InnerText = ($globalLogDir + "Response.log"); [void]$varsDel.AppendChild($v)

		$termsDel = $dynDoc.CreateElement("Terminals"); [void]$epDel.AppendChild($termsDel)
		foreach ($mm in $selectedScaleMeta)
		{
			$n = $dynDoc.CreateElement("Name"); $n.InnerText = $mm.ScaleName; [void]$termsDel.AppendChild($n)
		}

		$dynDoc.Save($AlexDynamicPath)
		Write_Log "Dynamic config saved: $AlexDynamicPath" "green"
	}
	catch
	{
		Write_Log "ERROR: Failed to build dynamic config: $_" "red"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# 7) Helper: clean deployment folder + remove old triggers (prevents accidental mixed runs)
	# ------------------------------------------------------------------------------------------------
	$timestampSlug = Get-Date -Format "MMdd.HHmmssfff"

	# Clean deploy input every time (safe)
	try { Get-ChildItem -LiteralPath $DeployInputPath -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue }
	catch { }

	# Remove any old trigger files that can interfere with this run
	foreach ($t in @("BatchSendFull", "BatchScl", "DBatch", "DBatchScl", "Status", "Sales"))
	{
		$tp = Join-Path $GlobalToBizerba $t
		if (Test-Path -LiteralPath $tp)
		{
			try { Remove-Item -LiteralPath $tp -Force -ErrorAction SilentlyContinue }
			catch { }
		}
	}

	# Determine trigger names for deploy (Batch*) and delete (DBatch*)
	$deployTriggerName = $null
	if ($sendFull) { $deployTriggerName = "BatchSendFull" }
	if ($sendRecent) { $deployTriggerName = "BatchScl" }

	$deleteTriggerName = "DBatch" # Always DBatch for "Delete all items"

	# ------------------------------------------------------------------------------------------------
	# 8) Run sequence:
	#     - DELETE ONLY: create DBatch trigger and run BMS once
	#     - DEPLOY ONLY: write files + create Batch* trigger and run BMS once
	#     - DELETE THEN DEPLOY: run BMS twice (DBatch first, then Batch*)
	# ------------------------------------------------------------------------------------------------
	$responseLogPath = Join-Path $GlobalLogPath "Response.log"
	$relConfig = ".\{0}" -f $BmsConfigName
	$relDynamic = ".\{0}" -f $AlexDynamicConfigName

	# Collect per-stage summaries (so delete stage doesn't get overwritten by deploy stage)
	$summaryDeleteStage = New-Object System.Collections.Generic.List[string]
	$summaryDeployStage = New-Object System.Collections.Generic.List[string]

	# ------------------------- Stage runner -------------------------
	# Runs BMS.exe once and summarizes Response.log per selected scale. Only lines BMS appended
	# during THIS run count: the log keeps history, and "latest line per IP" used to report an
	# older run's result when this run wrote nothing for a scale.
	$runBmsStage = {
		param ([string]$StageName, [string]$TriggerLike, $Summary)

		$linesBefore = 0
		if (Test-Path -LiteralPath $responseLogPath) { $linesBefore = @(Get-Content -LiteralPath $responseLogPath -ErrorAction SilentlyContinue).Count }

		Write_Log "Running BMS.exe ($StageName stage)..." "blue"
		$exitCode = $null
		try
		{
			Push-Location $BmsRootNorm
			& $BmsExePath -config $relConfig -dynamic $relDynamic
			$exitCode = $LASTEXITCODE
		}
		catch { Write_Log "ERROR: Failed to run BMS.exe ($StageName stage): $_" "red" }
		finally { Pop-Location }

		if ($null -ne $exitCode) { Write_Log "BMS.exe exit code ($StageName stage): $exitCode" "green" }
		else { Write_Log "WARNING: BMS.exe exit code unknown ($StageName stage)." "yellow" }

		if (-not (Test-Path -LiteralPath $responseLogPath))
		{
			foreach ($m in $selectedScaleMeta) { [void]$Summary.Add(("{0} ({1}): Response.log not found" -f $m.ScaleName, $m.IPAddress)) }
			return
		}

		$allLines = @(Get-Content -LiteralPath $responseLogPath -ErrorAction SilentlyContinue)
		# A shorter log means BMS rotated it: then every line is from this run
		$newLines = if ($allLines.Count -ge $linesBefore) { @($allLines | Select-Object -Skip $linesBefore) } else { $allLines }

		# Response.log: ts|name|ip|rc|?|trigger|items|errors|type
		$latestByIP = @{ }
		foreach ($l in $newLines)
		{
			if ([string]::IsNullOrWhiteSpace($l)) { continue }
			$parts = $l -split '\|', -1
			if ($parts.Count -lt 9) { continue }
			$ip = $parts[2].Trim()
			$trigger = $parts[5].Trim()
			if ($trigger -notlike $TriggerLike -or [string]::IsNullOrWhiteSpace($ip)) { continue }

			$tsNum = 0L; [void][int64]::TryParse(($parts[0].Trim() -replace '_', ''), [ref]$tsNum)
			$rec = [pscustomobject]@{
				TimestampText = $parts[0].Trim()
				TimestampNum  = $tsNum
				Trigger	      = $trigger
				RC		      = $parts[3].Trim()
				Items		  = $parts[6].Trim()
				Errors	      = $parts[7].Trim()
				Type		  = $parts[8].Trim()
				ScaleNameLog  = $parts[1].Trim()
			}
			if (-not $latestByIP.ContainsKey($ip) -or $rec.TimestampNum -ge $latestByIP[$ip].TimestampNum) { $latestByIP[$ip] = $rec }
		}

		foreach ($m in $selectedScaleMeta)
		{
			$ipKey = $m.IPAddress.Trim()
			if (-not $latestByIP.ContainsKey($ipKey))
			{
				[void]$Summary.Add(("{0} ({1}): no {2} response line detected" -f $m.ScaleName, $m.IPAddress, $StageName))
				continue
			}
			$r = $latestByIP[$ipKey]
			$orUnknown = { param ($v) if ([string]::IsNullOrWhiteSpace($v)) { '?' } else { $v } }
			$typeText = if ($r.Type) { ", type=$($r.Type)" } else { "" }
			[void]$Summary.Add(("{0} ({1}): rc={2}, trigger={3}, items={4}, errs={5}{6}, ts={7}, logName={8}" -f
					$m.ScaleName, $m.IPAddress, (& $orUnknown $r.RC), $r.Trigger, (& $orUnknown $r.Items), (& $orUnknown $r.Errors), $typeText, $r.TimestampText, $r.ScaleNameLog))
		}
	}
	# ========== STAGE 1: DELETE (if deleteOnly OR deleteFirstThenDeploy) ==========
	if ($deleteOnly -or $deleteFirstThenDeploy)
	{
		Write_Log "Preparing DELETE ALL trigger only..." "blue"

		# Ensure no deploy trigger exists before delete run
		if ($deployTriggerName)
		{
			$tp = Join-Path $GlobalToBizerba $deployTriggerName
			if (Test-Path -LiteralPath $tp)
			{
				try { Remove-Item -LiteralPath $tp -Force -ErrorAction SilentlyContinue }
				catch { }
			}
		}

		# Create DBatch trigger
		$dbatchPath = Join-Path $GlobalToBizerba $deleteTriggerName
		try
		{
			"DBatch trigger created at $timestampSlug" | Set-Content -LiteralPath $dbatchPath -Encoding ASCII -ErrorAction Stop
			Write_Log "Prepared delete trigger: $deleteTriggerName" "green"
		}
		catch
		{
			Write_Log "ERROR: Failed to create delete trigger '$dbatchPath': $_" "red"
			return
		}

		& $runBmsStage 'DELETE' 'DBatch*' $summaryDeleteStage
		# If DELETE ONLY, we stop here (no deploy stage)
		if ($deleteOnly)
		{
			Write_Log "`r`nManage_Bizerba_Scales_Alex_Profile completed." "green"
			Write_Log "  Actions : $opSummary" "green"
			Write_Log "  Scales  : $($selectedScaleMeta.ScaleName -join ', ')" "green"

			Write_Log "  DELETE Stage Responses:" "cyan"
			foreach ($s in $summaryDeleteStage) { Write_Log "    $s" "magenta" }

			Write_Log "`r`n==================== Completed Manage Bizerba Scales (Dynamic Only) ====================`r`n" "blue"
			return
		}

		# For delete-then-deploy flow: remove DBatch trigger so it won't re-run during deploy stage
		$dbatchPathCleanup = Join-Path $GlobalToBizerba $deleteTriggerName
		if (Test-Path -LiteralPath $dbatchPathCleanup)
		{
			try { Remove-Item -LiteralPath $dbatchPathCleanup -Force -ErrorAction SilentlyContinue }
			catch { }
		}

		Start-Sleep -Milliseconds 250
	}

	# ========== STAGE 2: DEPLOY (if FULL or RECENT) ==========
	if ($sendFull -or $sendRecent)
	{
		Write_Log "Preparing DEPLOY stage..." "blue"

		# Clean deploy input before writing deploy files
		try { Get-ChildItem -LiteralPath $DeployInputPath -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue }
		catch { }

		# Write ATST/PLST into DeployInputPath
		if ($atstLines -and $atstLines.Count -gt 0)
		{
			$atstOutName = if ($atstOrigName) { $atstOrigName }
			else { "ATST_SCL.$timestampSlug.txt" }
			$atstOut = Join-Path $DeployInputPath $atstOutName
			try { $atstLines | Set-Content -LiteralPath $atstOut -Encoding Default -ErrorAction Stop }
			catch { Write_Log "ERROR: Failed to write ATST: $_" "red"; return }
		}

		if ($plstLines -and $plstLines.Count -gt 0)
		{
			$plstOutName = if ($plstOrigName) { $plstOrigName }
			else { "PLST_SCL.$timestampSlug.txt" }
			$plstOut = Join-Path $DeployInputPath $plstOutName
			try { $plstLines | Set-Content -LiteralPath $plstOut -Encoding Default -ErrorAction Stop }
			catch { Write_Log "ERROR: Failed to write PLST: $_" "red"; return }
		}

		if (-not $deployTriggerName)
		{
			Write_Log "ERROR: Deploy trigger name could not be resolved." "red"
			return
		}

		# Ensure delete trigger is NOT present for deploy stage
		$dbatchPathCleanup2 = Join-Path $GlobalToBizerba $deleteTriggerName
		if (Test-Path -LiteralPath $dbatchPathCleanup2)
		{
			try { Remove-Item -LiteralPath $dbatchPathCleanup2 -Force -ErrorAction SilentlyContinue }
			catch { }
		}

		# Create deploy trigger (Batch*)
		$batchTriggerPath = Join-Path $GlobalToBizerba $deployTriggerName
		try
		{
			"Batch trigger created at $timestampSlug" | Set-Content -LiteralPath $batchTriggerPath -Encoding ASCII -ErrorAction Stop
			Write_Log "Prepared deploy trigger: $deployTriggerName" "green"
		}
		catch
		{
			Write_Log "ERROR: Failed to create deploy trigger '$batchTriggerPath': $_" "red"
			return
		}

		& $runBmsStage 'DEPLOY' 'Batch*' $summaryDeployStage
	}
	# ------------------------------------------------------------------------------------------------
	# 9) Completion Summary (shows delete stage separately when used)
	# ------------------------------------------------------------------------------------------------
	Write_Log "`r`nManage_Bizerba_Scales_Alex_Profile completed." "green"
	Write_Log "  Actions : $opSummary" "green"
	Write_Log "  Scales  : $($selectedScaleMeta.ScaleName -join ', ')" "green"

	if ($deleteFirstThenDeploy)
	{
		Write_Log "  DELETE Stage Responses:" "cyan"
		foreach ($s in $summaryDeleteStage) { Write_Log "    $s" "magenta" }

		Write_Log "  DEPLOY Stage Responses:" "cyan"
		foreach ($s in $summaryDeployStage) { Write_Log "    $s" "magenta" }
	}
	else
	{
		Write_Log "  Responses:" "cyan"
		# For single-stage deploy, show deploy summaries; for single-stage delete-only, we already returned earlier.
		foreach ($s in $summaryDeployStage) { Write_Log "    $s" "magenta" }
	}

	Write_Log "`r`n==================== Completed Manage Bizerba Scales (Dynamic Only) ====================`r`n" "blue"
}
