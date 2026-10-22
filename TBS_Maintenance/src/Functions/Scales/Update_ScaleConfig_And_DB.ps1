# ===================================================================================================
#                              FUNCTION: Update_ScaleConfig_And_DB
# ---------------------------------------------------------------------------------------------------
# Description:
#   Weighted-item marking wizard: choose SCL mode or Default/POS mode, updates the
#   ProdType_F272 setting in all ScaleCommApp XML config files (insert or update), then runs
#   the SQL update/merge over SCL_TAB and POS_TAB keyed on the item key (default F01).
#   Reverting to POS clears F272 only, and only when the column allows NULL (a NOT NULL F272
#   is left alone). The SQL connection is checked before any config file is changed.
#
# Parameters:
#   (none - defaults set inline at the top of the function)
#
# Variables:
#   - $script:ScaleCommRoots / $script:PrimaryScaleCommRoot: config XML location
#   - $script:FunctionResults: ConnectionString / DBSERVER / DBNAME (SQL goes through SqlClient)
#
# Notes:
#   - Range/field defaults (MinItem, MaxItem, POS_Field/Value, SCL_Field/Value, ItemKey) are inline.
# ===================================================================================================

function Update_ScaleConfig_And_DB
{
	Write_Log "`r`n==================== Starting Update_ScaleConfig_And_DB Function ====================`r`n" "blue"

	# ---- Set all defaults up front (edit as needed) ----
	$Folder = $script:ScaleCommRoots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	if (-not $Folder) { $Folder = $script:PrimaryScaleCommRoot }
	$MinItem = '0020000100000'
	$MaxItem = '0029999900000'
	$POS_Field = 'F82'
	$POS_Value = 0
	$SCL_Field = 'F272'
	$SCL_Value = 3
	$ItemKey = 'F01'

	# ---------------------------------------------------------------------------
	# Prompt the user for which mode to use (GUI)
	# ---------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Weighted Item Marking Method"
	$form.Size = New-Object System.Drawing.Size(420, 170)
	$form.StartPosition = "CenterScreen"
	$form.Topmost = $true
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Which marking method do you want to implement for weighted items?"
	$label.Size = New-Object System.Drawing.Size(390, 40)
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$form.Controls.Add($label)
	$radioSCL = New-Object System.Windows.Forms.RadioButton
	$radioSCL.Text = "New Way (SCL.F272 = 3, config to SCL)"
	$radioSCL.Location = New-Object System.Drawing.Point(30, 50)
	$radioSCL.Size = New-Object System.Drawing.Size(350, 20)
	$radioSCL.Checked = $true
	$form.Controls.Add($radioSCL)
	$radioPOS = New-Object System.Windows.Forms.RadioButton
	$radioPOS.Text = "Default / Old Way (POS.F82, config to POS, F272 cleared)"
	$radioPOS.Location = New-Object System.Drawing.Point(30, 75)
	$radioPOS.Size = New-Object System.Drawing.Size(350, 20)
	$form.Controls.Add($radioPOS)
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(220, 110)
	$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
	$form.Controls.Add($okButton)
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Location = New-Object System.Drawing.Point(310, 110)
	$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
	$form.Controls.Add($cancelButton)
	$form.AcceptButton = $okButton
	$form.CancelButton = $cancelButton
	$form.Tag = $null
	$form.ShowDialog() | Out-Null
	if ($form.Tag -ne "OK")
	{
		Write_Log "User canceled Update_ScaleConfig_And_DB." "yellow"
		Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
		return
	}
	$Mode = if ($radioSCL.Checked) { "SCL" }
	else { "POS" }

	# ---------------------------------------------------------------------------
	# SQL connection first: configs and DB must change together, so nothing is
	# written when the DB cannot be reached
	# ---------------------------------------------------------------------------
	$ConnectionString = Normalize_Sql_ConnectionString -RawConnectionString $script:FunctionResults['ConnectionString'] -Server $script:FunctionResults['DBSERVER'] -Database $script:FunctionResults['DBNAME']
	if (-not $ConnectionString)
	{
		Write_Log "DB server, DB name, or connection string not found. Nothing was changed." "red"
		Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
		return
	}
	$sqlConn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
	try { $sqlConn.Open() }
	catch
	{
		$sqlConn.Dispose()
		Write_Log "Cannot connect to the store database: $($_.Exception.Message). Nothing was changed." "red"
		Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
		return
	}

	# ---------------------------------------------------------------------------
	# XML Config Update (robust, always update or insert key)
	# ---------------------------------------------------------------------------
	$Files = @(
		'ScaleManagementApp.exe.config',
		'ScaleManagementAppUpdateSpecials.exe.config',
		'ScaleManagementAppSetup.exe.config',
		'ScaleManagementApp_FastDEPLOY.exe.config'
	)
	$ProdTypeKey = "ProdType_F272"
	$ProdTypeValue = if ($Mode -eq "SCL") { "SCL.F272" }
	else { "POS.F82" }
	foreach ($file in $Files)
	{
		$FullPath = Join-Path $Folder $file
		if (-not (Test-Path $FullPath))
		{
			Write_Log "[$file] Not found, skipped" "yellow"
			continue
		}
		try
		{
			[xml]$xml = Get-Content $FullPath -Raw
			$settings = $xml.configuration.appSettings
			$existing = $settings.add | Where-Object { $_.key -eq $ProdTypeKey }
			if ($existing)
			{
				if ($existing.value -ne $ProdTypeValue)
				{
					$existing.value = $ProdTypeValue
					Write_Log "[$file] Updated key '$ProdTypeKey' to '$ProdTypeValue'" "green"
				}
				else
				{
					Write_Log "[$file] '$ProdTypeKey' already set to '$ProdTypeValue', no change needed" "gray"
				}
			}
			else
			{
				$addElem = $xml.CreateElement("add")
				$addElem.SetAttribute("key", $ProdTypeKey)
				$addElem.SetAttribute("value", $ProdTypeValue)
				$settings.AppendChild($addElem) | Out-Null
				Write_Log "[$file] Inserted key '$ProdTypeKey' = '$ProdTypeValue'" "green"
			}
			$xml.Save($FullPath)
		}
		catch
		{
			Write_Log "[$file] XML update error: $_" "red"
		}
	}

	if ($Mode -eq "SCL")
	{
		$SqlQuery = @"
MERGE INTO SCL_TAB AS Target
USING (
    SELECT
        $ItemKey,
        F1000,
		TRY_CAST(FLOOR(CAST(SUBSTRING($ItemKey, 4, LEN($ItemKey) - 8) AS FLOAT)) AS INT) AS F267,
        $POS_Field -- e.g., F82
    FROM POS_TAB
    WHERE $ItemKey BETWEEN '$MinItem' AND '$MaxItem'
) AS Source
ON Target.$ItemKey = Source.$ItemKey
WHEN MATCHED THEN
    UPDATE SET
        $SCL_Field = CASE WHEN Source.$POS_Field = 1 THEN 0 ELSE $SCL_Value END,
        F1000 = Source.F1000,
        F267 = Source.F267
WHEN NOT MATCHED BY TARGET THEN
    INSERT ($ItemKey, F1000, $SCL_Field, F267)
    VALUES (
        Source.$ItemKey,
        Source.F1000,
        CASE WHEN Source.$POS_Field = 1 THEN 0 ELSE $SCL_Value END,
        Source.F267
    );
SELECT @@ROWCOUNT AS RowsAffected;
"@
	}
	else
	{
		# -1 = column is NOT NULL: leave it alone
		$SqlQuery = @"
IF COLUMNPROPERTY(OBJECT_ID('dbo.SCL_TAB'), '$SCL_Field', 'AllowsNull') = 1
BEGIN
    UPDATE SCL_TAB
    SET $SCL_Field = NULL
    WHERE $ItemKey BETWEEN '$MinItem' AND '$MaxItem'
    AND $SCL_Field IS NOT NULL;
    SELECT @@ROWCOUNT AS RowsAffected;
END
ELSE
    SELECT -1 AS RowsAffected;
"@
	}
	try
	{
		$cmd = $sqlConn.CreateCommand()
		$cmd.CommandText = $SqlQuery
		$cmd.CommandTimeout = 0
		$rowsAffected = [int]$cmd.ExecuteScalar()
		if ($rowsAffected -lt 0) { Write_Log "SCL_TAB.$SCL_Field does not allow NULL - left unchanged." "yellow" }
		else { Write_Log "Database updated for $Mode mode. ($rowsAffected) items changed." "green" }
	}
	catch
	{
		Write_Log "Error executing SQL update for $Mode mode: $_" "red"
	}
	finally { $sqlConn.Dispose() }
	Write_Log "`r`n==================== Update_ScaleConfig_And_DB Function Completed ====================" "blue"
}
