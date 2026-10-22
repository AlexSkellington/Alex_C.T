# ===================================================================================================
#                              FUNCTION: Configure_Scale_Sub-Departments
# ---------------------------------------------------------------------------------------------------
# Description:
#   WinForms dialog to configure scale sub-department tables: auto-picks the first EMPTY table
#   in order FAM_TAB > RPC_TAB > CAT_TAB (shows which was used), individual choices prompt to
#   Back Out or Overwrite (exact 7 rows), and SdpDefault is updated to OBJ.F16 / OBJ.F18 /
#   OBJ.F17 or restored to "1" (no DB change).
#
# Parameters:
#   - ServerInstance  [string]: SQL server (fallback: FunctionResults, else localhost\SQLEXPRESS)
#   - Database        [string]: database (fallback: FunctionResults, else STORESQL)
#   - ConfigFolder    [string]: ScaleComm config folder (default $script:PrimaryScaleCommRoot)
#
# Notes:
#   - Schema-aware, transactional SQL; PS 5.1 safe (no ternary, no nested functions).
#   - Start and completion banners are written at all exits.
# ===================================================================================================

function Configure_Scale_Sub-Departments
{
	[CmdletBinding()]
	param (
		[string]$ServerInstance,
		[string]$Database,
		[string]$ConfigFolder = $script:PrimaryScaleCommRoot
	)
	
	# ---- REQUIRED START BANNER (keep exactly as requested earlier) ----
	Write_Log "`r`n==================== Starting Update_ScaleConfig_And_DB Function ====================`r`n" "blue"
	
	# -------------------- Resolve Server/DB (params > cache > defaults) --------------------
	try
	{
		if (-not $ServerInstance -or $ServerInstance.Trim() -eq '')
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBSERVER') -and $script:FunctionResults['DBSERVER'])
			{
				$ServerInstance = $script:FunctionResults['DBSERVER']
			}
			else
			{
				$ServerInstance = 'localhost\SQLEXPRESS'
			}
		}
		if (-not $Database -or $Database.Trim() -eq '')
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBNAME') -and $script:FunctionResults['DBNAME'])
			{
				$Database = $script:FunctionResults['DBNAME']
			}
			else
			{
				$Database = 'STORESQL'
			}
		}
		Write_Log ("Target SQL: " + $ServerInstance + " | DB: " + $Database) "cyan"
	}
	catch
	{
		Write_Log ("Failed resolving server/database: " + $_.Exception.Message) "red"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# -------------------- Ensure Invoke-Sqlcmd is available --------------------
	try
	{
		$moduleLoaded = $false
		if (Get-Module -ListAvailable -Name SqlServer | Select-Object -First 1)
		{
			Import-Module SqlServer -ErrorAction Stop
			$moduleLoaded = $true
		}
		elseif (Get-Module -ListAvailable -Name SQLPS | Select-Object -First 1)
		{
			Import-Module SQLPS -DisableNameChecking -ErrorAction Stop
			$moduleLoaded = $true
		}
		else
		{
			Import-Module SqlServer -ErrorAction SilentlyContinue | Out-Null
			Import-Module SQLPS -DisableNameChecking -ErrorAction SilentlyContinue | Out-Null
		}
		if (-not $moduleLoaded)
		{
			Write_Log "Warning: SqlServer/SQLPS module not found; Invoke-Sqlcmd may be unavailable." "yellow"
		}
	}
	catch
	{
		Write_Log ("Module import error: " + $_.Exception.Message) "red"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# ==============================
	# SQL templates (literal strings)
	# ==============================
	$Tsql_Check_For = @"
SET NOCOUNT ON;
DECLARE @Schema sysname = NULL, @Cnt bigint = NULL, @Tbl sysname = N'%TABLE%';
SELECT TOP (1) @Schema = s.name
FROM sys.tables AS t
JOIN sys.schemas AS s ON s.schema_id = t.schema_id
WHERE t.name = @Tbl;
IF @Schema IS NOT NULL
BEGIN
    DECLARE @sql nvarchar(max) =
        N'SELECT @c = COUNT_BIG(1) FROM ' + QUOTENAME(@Schema) + N'.' + QUOTENAME(@Tbl) + N';';
    EXEC sp_executesql @sql, N'@c bigint OUTPUT', @c = @Cnt OUTPUT;
END
SELECT @Schema AS SchemaName, @Cnt AS RowCnt;
"@
	
	$Tsql_Merge_FAM_Keep = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.FAM_TAB AS T
USING src S ON T.F16 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1040=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F16,F49,F1040,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL);
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_FAM_Overwrite = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.FAM_TAB AS T
USING src S ON T.F16 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1040=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F16,F49,F1040,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL)
WHEN NOT MATCHED BY SOURCE THEN
 DELETE;
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_RPC_Keep = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.RPC_TAB AS T
USING src S ON T.F18 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1024=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F18,F49,F1024,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL);
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_RPC_Overwrite = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (SELECT vID, vName FROM (VALUES
 (1,N'Produce'),(2,N'Deli'),(3,N'Meat'),(4,N'Seafood'),(5,N'Cafeteria'),(6,N'Hot Foods'),(7,N'Bakery')
) AS x(vID,vName))
MERGE /*+ HOLDLOCK */ %SCHEMA%.RPC_TAB AS T
USING src S ON T.F18 = S.vID
WHEN MATCHED THEN UPDATE SET
 T.F49=NULL, T.F1024=S.vName, T.F1123=NULL, T.F1124=NULL, T.F1168=NULL, T.F1965=NULL, T.F1966=NULL, T.F1967=NULL
WHEN NOT MATCHED BY TARGET THEN
 INSERT (F18,F49,F1024,F1123,F1124,F1168,F1965,F1966,F1967)
 VALUES (S.vID,NULL,S.vName,NULL,NULL,NULL,NULL,NULL,NULL)
WHEN NOT MATCHED BY SOURCE THEN
 DELETE;
COMMIT;
END TRY
BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
"@
	
	$Tsql_Merge_CAT_Keep = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (
    SELECT vID, vName FROM (VALUES
        (1,N'Produce'),
        (2,N'Deli'),
        (3,N'Meat'),
        (4,N'Seafood'),
        (5,N'Cafeteria'),
        (6,N'Hot Foods'),
        (7,N'Bakery')
    ) AS x(vID,vName)
)
MERGE /*+ HOLDLOCK */ %SCHEMA%.CAT_TAB AS T
USING src AS S
    ON T.F17 = S.vID

WHEN MATCHED THEN
    UPDATE SET
        T.F49   = NULL,
        T.F1023 = S.vName,
        T.F1123 = NULL,
        T.F1124 = NULL,
        T.F1168 = NULL,
        T.F1943 = NULL,
        T.F1944 = NULL,
        T.F1945 = NULL,
        T.F1946 = NULL,
        T.F1947 = NULL,
        T.F1965 = NULL,
        T.F1966 = NULL,
        T.F1967 = NULL,
        T.F2653 = NULL,
        T.F2654 = NULL

WHEN NOT MATCHED BY TARGET THEN
    INSERT (
        F17,   F49,  F1023,  F1123, F1124, F1168,
        F1943, F1944, F1945, F1946, F1947,
        F1965, F1966, F1967,
        F2653, F2654
    )
    VALUES (
        S.vID,        -- F17
        NULL,         -- F49
        S.vName,      -- F1023
        NULL,         -- F1123
        NULL,         -- F1124
        NULL,         -- F1168
        NULL,         -- F1943
        NULL,         -- F1944
        NULL,         -- F1945
        NULL,         -- F1946
        NULL,         -- F1947
        NULL,         -- F1965
        NULL,         -- F1966
        NULL,         -- F1967
        NULL,         -- F2653
        NULL          -- F2654
    );

COMMIT;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK;
    THROW;
END CATCH;
"@
	$Tsql_Merge_CAT_Overwrite = @"
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRY
BEGIN TRAN;
WITH src AS (
    SELECT vID, vName FROM (VALUES
        (1,N'Produce'),
        (2,N'Deli'),
        (3,N'Meat'),
        (4,N'Seafood'),
        (5,N'Cafeteria'),
        (6,N'Hot Foods'),
        (7,N'Bakery')
    ) AS x(vID,vName)
)
MERGE /*+ HOLDLOCK */ %SCHEMA%.CAT_TAB AS T
USING src AS S
    ON T.F17 = S.vID

WHEN MATCHED THEN
    UPDATE SET
        T.F49   = NULL,
        T.F1023 = S.vName,
        T.F1123 = NULL,
        T.F1124 = NULL,
        T.F1168 = NULL,
        T.F1943 = NULL,
        T.F1944 = NULL,
        T.F1945 = NULL,
        T.F1946 = NULL,
        T.F1947 = NULL,
        T.F1965 = NULL,
        T.F1966 = NULL,
        T.F1967 = NULL,
        T.F2653 = NULL,
        T.F2654 = NULL

WHEN NOT MATCHED BY TARGET THEN
    INSERT (
        F17,   F49,  F1023,  F1123, F1124, F1168,
        F1943, F1944, F1945, F1946, F1947,
        F1965, F1966, F1967,
        F2653, F2654
    )
    VALUES (
        S.vID,        -- F17
        NULL,         -- F49
        S.vName,      -- F1023
        NULL,         -- F1123
        NULL,         -- F1124
        NULL,         -- F1168
        NULL,         -- F1943
        NULL,         -- F1944
        NULL,         -- F1945
        NULL,         -- F1946
        NULL,         -- F1947
        NULL,         -- F1965
        NULL,         -- F1966
        NULL,         -- F1967
        NULL,         -- F2653
        NULL          -- F2654
    )

WHEN NOT MATCHED BY SOURCE THEN
    DELETE;

COMMIT;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK;
    THROW;
END CATCH;
"@
	# ===================
	# WinForms UI (no PreferredSize math; fixed DPI-aware size to avoid clipping)
	# ===================
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	[void][System.Windows.Forms.Application]::EnableVisualStyles()
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Scale Sub-Departments Configuration"
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.Topmost = $true
	$form.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	
	# ---- FIX: no PreferredSize/ClientSize math. Set roomy, DPI-aware size and lock minimum. ----
	$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
	$form.AutoSize = $false
	$form.Size = New-Object System.Drawing.Size -ArgumentList 700, 400 # width, height
	$form.MinimumSize = $form.Size
	
	# Root layout
	$root = New-Object System.Windows.Forms.TableLayoutPanel
	$root.Dock = 'Fill'
	$root.AutoSize = $true
	$root.AutoSizeMode = 'GrowAndShrink'
	$root.ColumnCount = 1
	$root.RowCount = 5
	$root.Padding = New-Object System.Windows.Forms.Padding(12, 12, 12, 12)
	$root.BackColor = $form.BackColor
	$form.Controls.Add($root)
	
	# Title
	$title = New-Object System.Windows.Forms.Label
	$title.Text = "Choose how to configure Sub-Deparments for the scales"
	$title.Font = New-Object System.Drawing.Font("Segoe UI", [float]11, [System.Drawing.FontStyle]::Bold)
	$title.AutoSize = $true
	$title.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 8)
	$root.Controls.Add($title)
	
	# Group box
	$grp = New-Object System.Windows.Forms.GroupBox
	$grp.Text = "Mode"
	$grp.Dock = 'Top'
	$grp.AutoSize = $true
	$grp.AutoSizeMode = 'GrowAndShrink'
	$grp.Padding = New-Object System.Windows.Forms.Padding(12, 12, 12, 12)
	$root.Controls.Add($grp)
	
	# Grid inside group
	$grid = New-Object System.Windows.Forms.TableLayoutPanel
	$grid.Dock = 'Top'
	$grid.AutoSize = $true
	$grid.AutoSizeMode = 'GrowAndShrink'
	$grid.ColumnCount = 1
	$grid.RowCount = 8
	$grp.Controls.Add($grid)
	
	# Auto radio + desc
	$rbAuto = New-Object System.Windows.Forms.RadioButton
	$rbAuto.Text = "Auto (pick empty table: FAM > RPC > CAT)"
	$rbAuto.AutoSize = $true
	$rbAuto.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$rbAuto.Checked = $true
	$grid.Controls.Add($rbAuto)
	
	$descAuto = New-Object System.Windows.Forms.Label
	$descAuto.Text = "   Recommended. Populates the first EMPTY table and sets SdpDefault accordingly."
	$descAuto.AutoSize = $true
	$descAuto.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descAuto.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descAuto)
	
	# FAM radio + desc
	$rbFAM = New-Object System.Windows.Forms.RadioButton
	$rbFAM.Text = "Use FAM_TAB (Family) (OBJ.F16)"
	$rbFAM.AutoSize = $true
	$rbFAM.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$grid.Controls.Add($rbFAM)
	
	$descFAM = New-Object System.Windows.Forms.Label
	$descFAM.Text = "   Populates FAM_TAB with 7 rows. If it already has rows, you can back out or overwrite."
	$descFAM.AutoSize = $true
	$descFAM.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descFAM.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descFAM)
	
	# RPC radio + desc
	$rbRPC = New-Object System.Windows.Forms.RadioButton
	$rbRPC.Text = "Use RPC_TAB (Report) (OBJ.F18)"
	$rbRPC.AutoSize = $true
	$rbRPC.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$grid.Controls.Add($rbRPC)
	
	$descRPC = New-Object System.Windows.Forms.Label
	$descRPC.Text = "   Populates RPC_TAB with 7 rows. If it already has rows, you can back out or overwrite."
	$descRPC.AutoSize = $true
	$descRPC.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descRPC.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descRPC)
	
	# CAT radio + desc
	$rbCAT = New-Object System.Windows.Forms.RadioButton
	$rbCAT.Text = "Use CAT_TAB (Category) (OBJ.F17)"
	$rbCAT.AutoSize = $true
	$rbCAT.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 0)
	$grid.Controls.Add($rbCAT)
	
	$descCAT = New-Object System.Windows.Forms.Label
	$descCAT.Text = "   Populates CAT_TAB with 7 rows. If it already has rows, you can back out or overwrite."
	$descCAT.AutoSize = $true
	$descCAT.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$descCAT.Margin = New-Object System.Windows.Forms.Padding(22, 0, 4, 6)
	$grid.Controls.Add($descCAT)
	
	# Restore default
	$chkRestore = New-Object System.Windows.Forms.CheckBox
	$chkRestore.Text = "Restore default SdpDefault=1 (no DB changes)"
	$chkRestore.AutoSize = $true
	$chkRestore.Margin = New-Object System.Windows.Forms.Padding(8, 8, 4, 4)
	$root.Controls.Add($chkRestore)
	
	# Footer: server|db
	$status = New-Object System.Windows.Forms.Label
	$status.Text = "Resolved target: " + $ServerInstance + " | " + $Database
	$status.AutoSize = $true
	$status.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 118)
	$status.Margin = New-Object System.Windows.Forms.Padding(8, 8, 4, 8)
	$root.Controls.Add($status)
	
	# Buttons row
	$btnPanel = New-Object System.Windows.Forms.TableLayoutPanel
	$btnPanel.AutoSize = $true
	$btnPanel.AutoSizeMode = 'GrowAndShrink'
	$btnPanel.ColumnCount = 3
	$btnPanel.RowCount = 1
	$btnPanel.Dock = 'Top'
	$btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	$btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$root.Controls.Add($btnPanel)
	
	$spacer = New-Object System.Windows.Forms.Label
	$spacer.AutoSize = $true
	$btnPanel.Controls.Add($spacer, 0, 0)
	
	$btnRun = New-Object System.Windows.Forms.Button
	$btnRun.Text = "Run"
	$btnRun.AutoSize = $true
	$btnRun.Margin = New-Object System.Windows.Forms.Padding(4, 0, 8, 0)
	$btnPanel.Controls.Add($btnRun, 1, 0)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.AutoSize = $true
	$btnCancel.Margin = New-Object System.Windows.Forms.Padding(4, 0, 8, 0)
	$btnPanel.Controls.Add($btnCancel, 2, 0)
	
	# Default buttons
	$form.AcceptButton = $btnRun
	$form.CancelButton = $btnCancel
	
	# Tooltips
	$tip = New-Object System.Windows.Forms.ToolTip
	$tip.SetToolTip($rbAuto, "Automatically picks the first EMPTY table: FAM > RPC > CAT.")
	$tip.SetToolTip($rbFAM, "Use FAM_TAB. If it already has rows, you can overwrite or back out.")
	$tip.SetToolTip($rbRPC, "Use RPC_TAB. If it already has rows, you can overwrite or back out.")
	$tip.SetToolTip($rbCAT, "Use CAT_TAB. If it already has rows, you can overwrite or back out.")
	$tip.SetToolTip($chkRestore, "Only updates configs to SdpDefault=1. No database changes.")
	
	# Decide action (no nested funcs)
	$script:_Sdp_Action = $null # "AUTO" | "FAM" | "RPC" | "CAT" | "RESTORE"
	$btnRun.Add_Click({
			if ($chkRestore.Checked)
			{
				$script:_Sdp_Action = "RESTORE"
			}
			else
			{
				if ($rbAuto.Checked) { $script:_Sdp_Action = "AUTO" }
				elseif ($rbFAM.Checked) { $script:_Sdp_Action = "FAM" }
				elseif ($rbRPC.Checked) { $script:_Sdp_Action = "RPC" }
				elseif ($rbCAT.Checked) { $script:_Sdp_Action = "CAT" }
			}
			if (-not $script:_Sdp_Action)
			{
				[System.Windows.Forms.MessageBox]::Show("Pick a mode or check 'Restore default'.", "No Action Selected",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})
	$btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
	
	$dr = $form.ShowDialog()
	if ($dr -ne [System.Windows.Forms.DialogResult]::OK -or -not $script:_Sdp_Action)
	{
		Write_Log "User canceled subdepartment configuration." "yellow"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# =========================
	# Behavior: Restore default (no DB)
	# =========================
	if ($script:_Sdp_Action -eq 'RESTORE')
	{
		Write_Log "Restoring SdpDefault=1 in configs (no DB changes)..." "blue"
		$files = @('ScaleManagementApp.exe.config', 'ScaleManagementAppUpdateSpecials.exe.config', 'ScaleManagementAppSetup.exe.config', 'ScaleManagementApp_FastDEPLOY.exe.config')
		$key = 'SdpDefault'
		$value = '1'
		foreach ($file in $files)
		{
			try
			{
				$full = Join-Path $ConfigFolder $file
				if (-not (Test-Path $full)) { Write_Log ("[" + $file + "] Not found in '" + $ConfigFolder + "' - skipped.") "yellow"; continue }
				[xml]$xml = Get-Content -Path $full -Raw
				if (-not $xml.configuration) { $xml.AppendChild($xml.CreateElement('configuration')) | Out-Null }
				if (-not $xml.configuration.appSettings) { $xml.configuration.AppendChild($xml.CreateElement('appSettings')) | Out-Null }
				$settings = $xml.configuration.appSettings
				$existing = $settings.add | Where-Object { $_.key -eq $key }
				if ($existing) { $existing.value = $value; Write_Log ("[" + $file + "] Updated '" + $key + "' to '1'.") "green" }
				else { $add = $xml.CreateElement('add'); $add.SetAttribute('key', $key); $add.SetAttribute('value', $value); $settings.AppendChild($add) | Out-Null; Write_Log ("[" + $file + "] Inserted '" + $key + "' = '1'.") "green" }
				$xml.Save($full)
			}
			catch { Write_Log ("[" + $file + "] XML update error: " + $_.Exception.Message) "red" }
		}
		Write_Log "SdpDefault restored to '1'." "green"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	
	# =========================
	# Table checks (schema + rowcount)
	# =========================
	$famInfo = $null; $rpcInfo = $null; $catInfo = $null
	try
	{
		$famInfo = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query ($Tsql_Check_For.Replace('%TABLE%', 'FAM_TAB')) -ErrorAction Stop | Select-Object -First 1
		$rpcInfo = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query ($Tsql_Check_For.Replace('%TABLE%', 'RPC_TAB')) -ErrorAction Stop | Select-Object -First 1
		$catInfo = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query ($Tsql_Check_For.Replace('%TABLE%', 'CAT_TAB')) -ErrorAction Stop | Select-Object -First 1
	}
	catch
	{
		Write_Log ("Error reading table info: " + $_.Exception.Message) "red"
		Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
		return
	}
	$famCount = 0; if ($famInfo -and $famInfo.RowCnt -ne $null) { $famCount = [int64]$famInfo.RowCnt }
	$rpcCount = 0; if ($rpcInfo -and $rpcInfo.RowCnt -ne $null) { $rpcCount = [int64]$rpcInfo.RowCnt }
	$catCount = 0; if ($catInfo -and $catInfo.RowCnt -ne $null) { $catCount = [int64]$catInfo.RowCnt }
	
	# =========================
	# Execute selected action (parse-safe .Replace and clean Query build)
	# =========================
	$chosenShort = $null
	$chosenSchema = $null
	$sdpValue = $null
	$overwrite = $false
	
	if ($script:_Sdp_Action -eq 'AUTO')
	{
		if ($famInfo -and $famInfo.SchemaName -and $famCount -eq 0)
		{
			$chosenShort = 'FAM_TAB'; $chosenSchema = $famInfo.SchemaName; $sdpValue = 'OBJ.F16'
			$q = $Tsql_Merge_FAM_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') # no space before .Replace
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
		}
		elseif ($rpcInfo -and $rpcInfo.SchemaName -and $rpcCount -eq 0)
		{
			$chosenShort = 'RPC_TAB'; $chosenSchema = $rpcInfo.SchemaName; $sdpValue = 'OBJ.F18'
			$q = $Tsql_Merge_RPC_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']')
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
		}
		elseif ($catInfo -and $catInfo.SchemaName -and $catCount -eq 0)
		{
			$chosenShort = 'CAT_TAB'; $chosenSchema = $catInfo.SchemaName; $sdpValue = 'OBJ.F17'
			$q = $Tsql_Merge_CAT_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']')
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
		}
		else
		{
			$why = "No empty candidate found. FAM rows=" + ($famInfo.RowCnt) + ", RPC rows=" + ($rpcInfo.RowCnt) + ", CAT rows=" + ($catInfo.RowCnt) + "."
			[void][System.Windows.Forms.MessageBox]::Show($why, "Auto: No Empty Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			Write_Log ("Auto aborted: " + $why) "yellow"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		Write_Log ("Auto used: [" + $chosenSchema + "]." + $chosenShort) "green"
		[void][System.Windows.Forms.MessageBox]::Show(("Auto used: [" + $chosenSchema + "]." + $chosenShort), "Auto Selection",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
	}
	elseif ($script:_Sdp_Action -eq 'FAM')
	{
		if (-not ($famInfo -and $famInfo.SchemaName))
		{
			[void][System.Windows.Forms.MessageBox]::Show(("FAM_TAB does not exist in " + $Database + "."), "Missing Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			Write_Log "FAM_TAB missing." "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		$chosenShort = 'FAM_TAB'; $chosenSchema = $famInfo.SchemaName; $sdpValue = 'OBJ.F16'
		if ($famCount -gt 0)
		{
			$msg = "FAM_TAB already has " + $famCount + " row(s)." + [Environment]::NewLine + "Overwrite table to the standard 7 rows?"
			$res = [System.Windows.Forms.MessageBox]::Show($msg, "FAM_TAB Occupied",
				[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if ($res -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User backed out of FAM overwrite." "yellow"
				Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
				return
			}
			$overwrite = $true
		}
		try
		{
			if ($overwrite) { $q = $Tsql_Merge_FAM_Overwrite.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			else { $q = $Tsql_Merge_FAM_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
			Write_Log ("FAM_TAB configured. Overwrite=" + $overwrite) "green"
		}
		catch
		{
			Write_Log ("FAM SQL error: " + $_.Exception.Message) "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
	}
	elseif ($script:_Sdp_Action -eq 'RPC')
	{
		if (-not ($rpcInfo -and $rpcInfo.SchemaName))
		{
			[void][System.Windows.Forms.MessageBox]::Show(("RPC_TAB does not exist in " + $Database + "."), "Missing Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			Write_Log "RPC_TAB missing." "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		$chosenShort = 'RPC_TAB'; $chosenSchema = $rpcInfo.SchemaName; $sdpValue = 'OBJ.F18'
		if ($rpcCount -gt 0)
		{
			$msg = "RPC_TAB already has " + $rpcCount + " row(s)." + [Environment]::NewLine + "Overwrite table to the standard 7 rows?"
			$res = [System.Windows.Forms.MessageBox]::Show($msg, "RPC_TAB Occupied",
				[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if ($res -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User backed out of RPC overwrite." "yellow"
				Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
				return
			}
			$overwrite = $true
		}
		try
		{
			if ($overwrite) { $q = $Tsql_Merge_RPC_Overwrite.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			else { $q = $Tsql_Merge_RPC_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
			Write_Log ("RPC_TAB configured. Overwrite=" + $overwrite) "green"
		}
		catch
		{
			Write_Log ("RPC SQL error: " + $_.Exception.Message) "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
	}
	elseif ($script:_Sdp_Action -eq 'CAT')
	{
		if (-not ($catInfo -and $catInfo.SchemaName))
		{
			[void][System.Windows.Forms.MessageBox]::Show(("CAT_TAB does not exist in " + $Database + "."), "Missing Table",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			Write_Log "CAT_TAB missing." "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
		$chosenShort = 'CAT_TAB'; $chosenSchema = $catInfo.SchemaName; $sdpValue = 'OBJ.F17'
		if ($catCount -gt 0)
		{
			$msg = "CAT_TAB already has " + $catCount + " row(s)." + [Environment]::NewLine + "Overwrite table to the standard 7 rows?"
			$res = [System.Windows.Forms.MessageBox]::Show($msg, "CAT_TAB Occupied",
				[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if ($res -ne [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User backed out of CAT overwrite." "yellow"
				Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
				return
			}
			$overwrite = $true
		}
		try
		{
			if ($overwrite) { $q = $Tsql_Merge_CAT_Overwrite.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			else { $q = $Tsql_Merge_CAT_Keep.Replace('%SCHEMA%', '[' + $chosenSchema + ']') }
			Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $q -QueryTimeout 120 -ErrorAction Stop | Out-Null
			Write_Log ("CAT_TAB configured. Overwrite=" + $overwrite) "green"
		}
		catch
		{
			Write_Log ("CAT SQL error: " + $_.Exception.Message) "red"
			Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
			return
		}
	}
	
	# =========================
	# Update SdpDefault in config files
	# =========================
	if ($sdpValue)
	{
		Write_Log ("Updating SdpDefault to '" + $sdpValue + "' in configs...") "blue"
		$files = @('ScaleManagementApp.exe.config', 'ScaleManagementAppUpdateSpecials.exe.config', 'ScaleManagementAppSetup.exe.config', 'ScaleManagementApp_FastDEPLOY.exe.config')
		$key = 'SdpDefault'
		foreach ($file in $files)
		{
			try
			{
				$full = Join-Path $ConfigFolder $file
				if (-not (Test-Path $full)) { Write_Log ("[" + $file + "] Not found in '" + $ConfigFolder + "' - skipped.") "yellow"; continue }
				[xml]$xml = Get-Content -Path $full -Raw
				if (-not $xml.configuration) { $xml.AppendChild($xml.CreateElement('configuration')) | Out-Null }
				if (-not $xml.configuration.appSettings) { $xml.configuration.AppendChild($xml.CreateElement('appSettings')) | Out-Null }
				$settings = $xml.configuration.appSettings
				$existing = $settings.add | Where-Object { $_.key -eq $key }
				if ($existing)
				{
					if ($existing.value -ne $sdpValue) { $existing.value = $sdpValue; Write_Log ("[" + $file + "] Updated '" + $key + "' to '" + $sdpValue + "'.") "green" }
					else { Write_Log ("[" + $file + "] '" + $key + "' already '" + $sdpValue + "' - no change.") "gray" }
				}
				else
				{
					$add = $xml.CreateElement('add'); $add.SetAttribute('key', $key); $add.SetAttribute('value', $sdpValue)
					$settings.AppendChild($add) | Out-Null
					Write_Log ("[" + $file + "] Inserted '" + $key + "' = '" + $sdpValue + "'.") "green"
				}
				$xml.Save($full)
			}
			catch { Write_Log ("[" + $file + "] XML update error: " + $_.Exception.Message) "red" }
		}
		Write_Log ("SdpDefault successfully set to '" + $sdpValue + "'.") "green"
	}
	
	# ---- REQUESTED COMPLETION BANNER (exact string) ----
	Write_Log "`r`n==================== Configure_Scale_Sub-Departments Completed ====================`r`n" "blue"
}
