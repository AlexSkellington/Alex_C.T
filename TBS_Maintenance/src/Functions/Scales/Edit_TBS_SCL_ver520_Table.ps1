# ===================================================================================================
#                              FUNCTION: Edit_TBS_SCL_ver520_Table
# ---------------------------------------------------------------------------------------------------
# Description:
#   Interactive grid editor for the TBS_SCL_ver520 scale table (no row dragging). Auto Sort
#   orders BIZERBA first (by numeric IP) then ISHIDA, renumbers them, applies the naming and
#   BufferTime templates and saves. Save validates and writes exactly what is visible in the
#   grid; Refresh reloads from the DB. IPNetwork must end with a trailing period.
#
# Variables:
#   - $script:FunctionResults: DBSERVER, DBNAME, ConnectionString, SQL_UserID/SQL_UserPwd
#
# Writes:
#   - store database table TBS_SCL_ver520
#
# Notes:
#   - SQL goes through System.Data.SqlClient (no SQL module needed).
#   - Rows are matched on their ORIGINAL brand / model / IPNetwork / IPDevice (NULLs match '').
# ===================================================================================================

function Edit_TBS_SCL_ver520_Table
{
	[CmdletBinding()]
	param ()

	Write_Log "`r`n==================== Starting Edit_TBS_SCL_ver520_Table (No Drag / No Checkbox) ====================`r`n" "blue"

	# ----------------------------------------------------------------------------------------------
	# 1) Resolve DB info and build connection string (SQL Auth preferred, else Integrated)
	# ----------------------------------------------------------------------------------------------
	$server = $script:FunctionResults['DBSERVER']
	$database = $script:FunctionResults['DBNAME']
	$connectionString = Normalize_Sql_ConnectionString -RawConnectionString $script:FunctionResults['ConnectionString'] -Server $server -Database $database -SqlUser $script:FunctionResults['SQL_UserID'] -SqlPwd $script:FunctionResults['SQL_UserPwd']

	if (-not $server -or -not $database)
	{
		Write_Log "DB server or DB name missing - cannot proceed." "red"
		return
	}

	if (-not $connectionString)
	{
		Write_Log "Unable to build a valid SQL connection string from cached values." "red"
		return
	}
	$script:FunctionResults['ConnectionString'] = $connectionString

	# ----------------------------------------------------------------------------------------------
	# 2) SQL helpers (one connection per call)
	# ----------------------------------------------------------------------------------------------
	# Rows of a query (PowerShell unrolls the DataTable: DataRow(s) or nothing)
	$runQuery = {
		param ([string]$Sql)
		$conn = New-Object System.Data.SqlClient.SqlConnection $connectionString
		try
		{
			$conn.Open()
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 0
			$td = New-Object System.Data.DataTable
			[void](New-Object System.Data.SqlClient.SqlDataAdapter $cmd).Fill($td)
			$td
		}
		finally { $conn.Dispose() }
	}

	$runScalar = {
		param ([string]$Sql)
		$conn = New-Object System.Data.SqlClient.SqlConnection $connectionString
		try
		{
			$conn.Open()
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 0
			$cmd.ExecuteScalar()
		}
		finally { $conn.Dispose() }
	}

	# ----------------------------------------------------------------------------------------------
	# 3) Discover the /24 prefix in use; prefer NIC that can reach the SQL server; else NIC w/ gateway; else first RFC1918
	#     Result: $allowedPrefix like '192.168.5' (used to validate/normalize IPNetwork).
	# ----------------------------------------------------------------------------------------------
	$allowedPrefix = $null
	try
	{
		# Resolve SQL server IPv4s (best effort; instance names may not resolve)
		$serverIPv4s = @()
		try
		{
			$serverIPs = [System.Net.Dns]::GetHostAddresses((($server -replace '^(?i)tcp:', '') -split '[\\,]')[0])
			foreach ($ip in $serverIPs)
			{
				if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork)
				{
					$serverIPv4s += $ip
				}
			}
		}
		catch { }

		$allNics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object {
			$_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up
		}

		# CHANGE: Inline subnet comparison (no helper function). Compare (addr & mask) for equality.
		$pickedIP = $null

		# Pass 1: NIC whose IPv4 shares subnet with SQL server
		if ($serverIPv4s.Count -gt 0)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					if ($ua.IPv4Mask -eq $null) { continue }

					$aBytes = $ua.Address.GetAddressBytes()
					$mBytes = $ua.IPv4Mask.GetAddressBytes()

					$matchFound = $false
					foreach ($srv in $serverIPv4s)
					{
						$bBytes = $srv.GetAddressBytes()
						$same = $true
						for ($i = 0; $i -lt 4; $i++)
						{
							if (($aBytes[$i] -band $mBytes[$i]) -ne ($bBytes[$i] -band $mBytes[$i]))
							{
								$same = $false
								break
							}
						}
						if ($same) { $matchFound = $true; break }
					}
					if ($matchFound) { $pickedIP = $ua.Address.ToString(); break }
				}
				if ($pickedIP) { break }
			}
		}

		# Pass 2: NIC with default gateway
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				$hasGw = $false
				foreach ($gw in $ipProps.GatewayAddresses)
				{
					if ($gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
				}
				if (-not $hasGw) { continue }

				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$pickedIP = $ua.Address.ToString()
					break
				}
				if ($pickedIP) { break }
			}
		}

		# Pass 3: first RFC1918 IPv4
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$ip = $ua.Address.ToString()
					if ($ip.StartsWith('10.') -or $ip.StartsWith('192.168.') -or (
							$ip.StartsWith('172.') -and (([int]($ip.Split('.')[1])) -ge 16) -and (([int]($ip.Split('.')[1])) -le 31)
						))
					{
						$pickedIP = $ip
						break
					}
				}
				if ($pickedIP) { break }
			}
		}

		if ($pickedIP)
		{
			$parts = $pickedIP.Split('.')
			if ($parts.Length -ge 3)
			{
				$allowedPrefix = "$($parts[0]).$($parts[1]).$($parts[2])"
			}
		}
	}
	catch { }

	if ($allowedPrefix)
	{
		Write_Log "IPNetwork allowed prefix detected (active NIC): $allowedPrefix" "cyan"
	}
	else
	{
		Write_Log "WARNING: Could not determine active NIC / RFC1918 IPv4. IPNetwork validation will be skipped (but trailing '.' will still be enforced if parsable)." "yellow" # CHANGE: note trailing '.' enforcement still happens
	}

	# ----------------------------------------------------------------------------------------------
	# 4) Load table -> DataTable (include Orig* columns to anchor updates)
	# ----------------------------------------------------------------------------------------------
	$selectQuery = @"
SELECT
    ScaleCode, ScaleName, ScaleBrand, ScaleModel, IPNetwork, IPDevice, BufferTime, Active
FROM [TBS_SCL_ver520]
ORDER BY ScaleCode ASC;
"@

	$dt = New-Object System.Data.DataTable
	foreach ($c in 'ScaleCode', 'ScaleName', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice', 'BufferTime', 'Active',
		'OrigScaleBrand', 'OrigScaleModel', 'OrigIPNetwork', 'OrigIPDevice')
	{
		[void]$dt.Columns.Add($c)
	}

	# DB rows -> grid rows, with the original key columns snapshotted for the UPDATE match
	$fillGrid = {
		param ($Rows)
		$dt.Rows.Clear()
		foreach ($r in @($Rows))
		{
			if (-not $r) { continue }
			$nr = $dt.NewRow()
			foreach ($col in 'ScaleCode', 'ScaleName', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice', 'BufferTime', 'Active') { $nr[$col] = [string]$r.$col }
			$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
			$nr['OrigScaleModel'] = [string]$r.ScaleModel
			$nr['OrigIPNetwork'] = [string]$r.IPNetwork
			$nr['OrigIPDevice'] = [string]$r.IPDevice
			[void]$dt.Rows.Add($nr)
		}
	}

	# Reload the grid from the DB; $true when it worked
	$reloadFromDb = {
		param ([string]$What = 'refresh')
		try
		{
			& $fillGrid (& $runQuery $selectQuery)
			if ($grid) { $grid.Refresh() }
			return $true
		}
		catch
		{
			Write_Log "SQL error ($What): $($_.Exception.Message)" "red"
			return $false
		}
	}

	Write_Log "Loading table for edit..." "blue"
	if (-not (& $reloadFromDb 'load')) { return }
	if ($dt.Rows.Count -eq 0)
	{
		Write_Log "TBS_SCL_ver520 has no rows - nothing to edit." "yellow"
		return
	}

	# ----------------------------------------------------------------------------------------------
	# 5) WinForms UI - simple layout (no drag UI, no checkbox)
	# ----------------------------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Edit Placings - TBS_SCL_ver520"
	$form.StartPosition = 'CenterScreen'
	$form.Size = New-Object System.Drawing.Size(1100, 720)
	$form.MinimumSize = New-Object System.Drawing.Size(960, 600)
	$form.AutoScaleMode = 'Dpi'
	$form.TopMost = $true

	$grid = New-Object System.Windows.Forms.DataGridView
	$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
	$grid.AutoGenerateColumns = $false
	$grid.AllowUserToAddRows = $false
	$grid.AllowUserToDeleteRows = $false
	$grid.ReadOnly = $false
	$grid.SelectionMode = 'FullRowSelect'
	$grid.MultiSelect = $false
	$grid.AutoSizeColumnsMode = 'Fill'
	$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
	$grid.AllowUserToOrderColumns = $true
	$grid.RowHeadersWidth = 30

	# CHANGE: Inline column creation (removed $AddCol helper).
	$col_ScaleCode = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleCode.HeaderText = 'ScaleCode'
	$col_ScaleCode.DataPropertyName = 'ScaleCode'
	$col_ScaleCode.ReadOnly = $false
	$col_ScaleCode.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleCode) | Out-Null

	$col_ScaleName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleName.HeaderText = 'ScaleName'
	$col_ScaleName.DataPropertyName = 'ScaleName'
	$col_ScaleName.ReadOnly = $false
	$col_ScaleName.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleName) | Out-Null

	$col_ScaleBrand = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleBrand.HeaderText = 'ScaleBrand'
	$col_ScaleBrand.DataPropertyName = 'ScaleBrand'
	$col_ScaleBrand.ReadOnly = $false
	$col_ScaleBrand.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleBrand) | Out-Null

	$col_ScaleModel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleModel.HeaderText = 'ScaleModel'
	$col_ScaleModel.DataPropertyName = 'ScaleModel'
	$col_ScaleModel.ReadOnly = $false
	$col_ScaleModel.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleModel) | Out-Null

	$col_IPNetwork = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_IPNetwork.HeaderText = 'IPNetwork'
	$col_IPNetwork.DataPropertyName = 'IPNetwork'
	$col_IPNetwork.ReadOnly = $false
	$col_IPNetwork.SortMode = 'NotSortable'
	$grid.Columns.Add($col_IPNetwork) | Out-Null

	$col_IPDevice = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_IPDevice.HeaderText = 'IPDevice'
	$col_IPDevice.DataPropertyName = 'IPDevice'
	$col_IPDevice.ReadOnly = $false
	$col_IPDevice.SortMode = 'NotSortable'
	$grid.Columns.Add($col_IPDevice) | Out-Null

	$col_BufferTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_BufferTime.HeaderText = 'BufferTime'
	$col_BufferTime.DataPropertyName = 'BufferTime'
	$col_BufferTime.ReadOnly = $false
	$col_BufferTime.SortMode = 'NotSortable'
	$grid.Columns.Add($col_BufferTime) | Out-Null

	$col_Active = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_Active.HeaderText = 'Active'
	$col_Active.DataPropertyName = 'Active'
	$col_Active.ReadOnly = $false
	$col_Active.SortMode = 'NotSortable'
	$grid.Columns.Add($col_Active) | Out-Null

	# Hidden originals
	$col_OrigScaleBrand = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigScaleBrand.HeaderText = 'OrigScaleBrand'
	$col_OrigScaleBrand.DataPropertyName = 'OrigScaleBrand'
	$col_OrigScaleBrand.ReadOnly = $true
	$col_OrigScaleBrand.Visible = $false
	$col_OrigScaleBrand.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigScaleBrand) | Out-Null

	$col_OrigScaleModel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigScaleModel.HeaderText = 'OrigScaleModel'
	$col_OrigScaleModel.DataPropertyName = 'OrigScaleModel'
	$col_OrigScaleModel.ReadOnly = $true
	$col_OrigScaleModel.Visible = $false
	$col_OrigScaleModel.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigScaleModel) | Out-Null

	$col_OrigIPNetwork = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigIPNetwork.HeaderText = 'OrigIPNetwork'
	$col_OrigIPNetwork.DataPropertyName = 'OrigIPNetwork'
	$col_OrigIPNetwork.ReadOnly = $true
	$col_OrigIPNetwork.Visible = $false
	$col_OrigIPNetwork.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigIPNetwork) | Out-Null

	$col_OrigIPDevice = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigIPDevice.HeaderText = 'OrigIPDevice'
	$col_OrigIPDevice.DataPropertyName = 'OrigIPDevice'
	$col_OrigIPDevice.ReadOnly = $true
	$col_OrigIPDevice.Visible = $false
	$col_OrigIPDevice.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigIPDevice) | Out-Null

	$grid.DataSource = $dt
	$form.Controls.Add($grid)

	# Bottom bar
	$bottom = New-Object System.Windows.Forms.Panel
	$bottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
	$bottom.Height = 86
	$form.Controls.Add($bottom)

	$lbl = New-Object System.Windows.Forms.Label
	$lbl.AutoSize = $true
	$lbl.Location = New-Object System.Drawing.Point(10, 10)
	if ($allowedPrefix)
	{
		# CHANGE: Reflect the enforced trailing period examples only.
		$lbl.Text = "Edit cells directly  |  Auto Sort renumbers and saves  |  Active: Y/N  |  IPNetwork: '$allowedPrefix.' (accepts $allowedPrefix, $allowedPrefix.0, $allowedPrefix.*, $allowedPrefix.x)"
	}
	else
	{
		# CHANGE: Note that a trailing period is still enforced if we can parse A.B.C.
		$lbl.Text = "Edit cells directly. Auto Sort applies standard order + numbering and saves immediately.  Active must be Y/N. IPNetwork will be normalized to 'A.B.C.' (trailing dot) if parsable."
	}
	$bottom.Controls.Add($lbl)

	$btnAuto = New-Object System.Windows.Forms.Button
	$btnAuto.Text = "Auto Sort (BIZERBA first)"
	$btnAuto.Size = New-Object System.Drawing.Size(220, 28)
	$btnAuto.Location = New-Object System.Drawing.Point(10, 44)
	$bottom.Controls.Add($btnAuto)

	$btnRefresh = New-Object System.Windows.Forms.Button
	$btnRefresh.Text = "Refresh"
	$btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
	$btnRefresh.Location = New-Object System.Drawing.Point(240, 44)
	$bottom.Controls.Add($btnRefresh)

	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = "Close"
	$btnClose.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnClose)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnCancel)

	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Text = "Save"
	$btnSave.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnSave)

	# CHANGE: Inline initial placement (removed $placeButtons helper)
	$bw = [int]$bottom.ClientSize.Width
	$btnSave.Location = New-Object System.Drawing.Point(($bw - 110), 44)
	$btnCancel.Location = New-Object System.Drawing.Point(($bw - 220), 44)
	$btnClose.Location = New-Object System.Drawing.Point(($bw - 330), 44)

	# CHANGE: Inline resize handler computes placements directly
	$form.add_Resize({
			$bw2 = [int]$bottom.ClientSize.Width
			$btnSave.Location = New-Object System.Drawing.Point(($bw2 - 110), 44)
			$btnCancel.Location = New-Object System.Drawing.Point(($bw2 - 220), 44)
			$btnClose.Location = New-Object System.Drawing.Point(($bw2 - 330), 44)
		})

	# ----------------------------------------------------------------------------------------------
	# Save: validate + write exactly what is in the grid (shared by Save and Auto Sort)
	# ----------------------------------------------------------------------------------------------
	# ISNULL on the target keys: a scale with a NULL brand / model / IP part was never matched
	# (and silently not saved); the row count tells when a row still did not match.
	$sqlSaveTemplate = @'
DECLARE @Updated int = 0;
BEGIN TRY
  BEGIN TRAN;

  DECLARE @Placings TABLE(
      OrigScaleBrand nvarchar(100) NOT NULL,
      OrigScaleModel nvarchar(100) NOT NULL,
      OrigIPNetwork  nvarchar(100) NOT NULL,
      OrigIPDevice   nvarchar(100) NOT NULL,
      NewScaleBrand  nvarchar(100) NOT NULL,
      NewScaleModel  nvarchar(100) NOT NULL,
      NewIPNetwork   nvarchar(100) NOT NULL,
      NewIPDevice    nvarchar(100) NOT NULL,
      NewScaleCode   int           NOT NULL,
      NewBufferTime  nvarchar(20)  NOT NULL,
      NewScaleName   nvarchar(200) NOT NULL,
      NewActive      nvarchar(20)  NOT NULL
  );

  INSERT INTO @Placings (
      OrigScaleBrand, OrigScaleModel, OrigIPNetwork, OrigIPDevice,
      NewScaleBrand, NewScaleModel, NewIPNetwork, NewIPDevice,
      NewScaleCode, NewBufferTime, NewScaleName, NewActive
  )
  VALUES
  $(VALUES_BLOCK);

  UPDATE tgt
     SET tgt.ScaleBrand = src.NewScaleBrand,
         tgt.ScaleModel = src.NewScaleModel,
         tgt.IPNetwork  = src.NewIPNetwork,
         tgt.IPDevice   = src.NewIPDevice,
         tgt.ScaleCode  = src.NewScaleCode,
         tgt.BufferTime = src.NewBufferTime,
         tgt.ScaleName  = src.NewScaleName,
         tgt.Active     = src.NewActive
  FROM [TBS_SCL_ver520] AS tgt
  JOIN @Placings AS src
    ON  ISNULL(CAST(tgt.ScaleBrand AS nvarchar(100)), N'') = src.OrigScaleBrand
    AND ISNULL(CAST(tgt.ScaleModel AS nvarchar(100)), N'') = src.OrigScaleModel
    AND ISNULL(CAST(tgt.IPNetwork AS nvarchar(100)), N'')  = src.OrigIPNetwork
    AND ISNULL(CAST(tgt.IPDevice AS nvarchar(100)), N'')   = src.OrigIPDevice;
  SET @Updated = @@ROWCOUNT;

  COMMIT TRAN;
END TRY
BEGIN CATCH
  IF @@TRANCOUNT > 0 ROLLBACK TRAN;
  DECLARE @msg nvarchar(4000) = ERROR_MESSAGE();
  RAISERROR(@msg, 16, 1);
END CATCH;
SELECT @Updated;
'@

	$ipnPattern = '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$'

	# Validate every grid row (shows the first problem); normalizes Active and IPNetwork in place
	$validateGrid = {
		$seenCodes = @{ }
		$rowIdx = 0
		foreach ($row in $dt.Rows)
		{
			$rowIdx++
			$raw = $row['ScaleCode']
			$val = $null
			if (-not [int]::TryParse([string]$raw, [ref]$val))
			{
				[void][System.Windows.Forms.MessageBox]::Show("Invalid ScaleCode at row #${rowIdx}: '$raw' is not a number.", "Invalid Code", 'OK', 'Error')
				return $false
			}
			if ($seenCodes.ContainsKey($val))
			{
				[void][System.Windows.Forms.MessageBox]::Show("Duplicate ScaleCode '$val' detected (row #$rowIdx).", "Duplicate Code", 'OK', 'Error')
				return $false
			}
			$seenCodes[$val] = $true

			$actRaw = [string]$row['Active']
			$actUp = $actRaw.Trim().ToUpper()
			if ($actUp -ne 'Y' -and $actUp -ne 'N')
			{
				[void][System.Windows.Forms.MessageBox]::Show("Invalid Active value at row #${rowIdx}: '$actRaw'. Use Y or N.", "Invalid Active", 'OK', 'Error')
				return $false
			}
			$row['Active'] = $actUp

			# IPNetwork: A.B.C(.anything) normalized to "A.B.C." (always a trailing period);
			# with a detected NIC prefix it must be that /24
			$ipnRaw = ([string]$row['IPNetwork']).Trim()
			$m = [regex]::Match($ipnRaw, $ipnPattern)
			if (-not $m.Success)
			{
				$expect = if ($allowedPrefix) { "something within '$allowedPrefix.*'" } else { "a pattern like A.B.C or A.B.C.X" }
				[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #${rowIdx}: '$ipnRaw'. Expected $expect.", "Invalid IPNetwork", 'OK', 'Error')
				return $false
			}
			$prefix = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)"
			if ($allowedPrefix -and $prefix -ne $allowedPrefix)
			{
				[void][System.Windows.Forms.MessageBox]::Show("IPNetwork at row #$rowIdx must be within '$allowedPrefix.*'. Got '$ipnRaw'.", "IPNetwork Out of Range", 'OK', 'Error')
				return $false
			}
			$row['IPNetwork'] = "$prefix."
		}
		return $true
	}

	# Validate, save, reload; $true when saved. $What names the action in the log.
	$saveGrid = {
		param ([string]$What)

		# Commit any in-progress grid edit
		if ($grid.IsCurrentCellInEditMode) { [void]$grid.EndEdit() }
		else { [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }

		if (-not (& $validateGrid)) { return $false }

		$sqlText = { param ($v) (([string]$v) -replace "[`r`n]", '') -replace "'", "''" }
		$values = New-Object System.Collections.Generic.List[string]
		foreach ($row in $dt.Rows)
		{
			$parts = @(
				'OrigScaleBrand', 'OrigScaleModel', 'OrigIPNetwork', 'OrigIPDevice', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice' |
				ForEach-Object { "'" + (& $sqlText $row[$_]) + "'" }
			)
			$parts += [string][int]$row['ScaleCode']
			$parts += @('BufferTime', 'ScaleName', 'Active' | ForEach-Object { "'" + (& $sqlText $row[$_]) + "'" })
			$values.Add('(' + ($parts -join ',') + ')')
		}
		if ($values.Count -eq 0)
		{
			Write_Log "Nothing to save ($What)." "yellow"
			return $false
		}

		Write_Log "Saving to database ($What)..." "blue"
		try
		{
			$updated = [int](& $runScalar ($sqlSaveTemplate.Replace('$(VALUES_BLOCK)', [string]::Join(",`r`n  ", $values))))
		}
		catch
		{
			Write_Log "Save ($What) failed: $($_.Exception.Message)" "red"
			[void][System.Windows.Forms.MessageBox]::Show("Save failed. Check logs.", "Error", 'OK', 'Error')
			return $false
		}
		if ($updated -ne $values.Count) { Write_Log "Saved $updated of $($values.Count) row(s) - the others no longer match the table (changed elsewhere?). Refreshing." "yellow" }

		Write_Log "Refreshing view..." "blue"
		[void](& $reloadFromDb "refresh after $What")
		return $true
	}

	# ----------------------------------------------------------------------------------------------
	# 6) Button: Save (validate + write exactly what's in grid; no renumber)
	# ----------------------------------------------------------------------------------------------
	$btnSave.add_Click({
			if (& $saveGrid 'Save')
			{
				[void][System.Windows.Forms.MessageBox]::Show("Saved and refreshed.", "Success", 'OK', 'Information')
				Write_Log "Saved and refreshed." "green"
			}
		})

	# ----------------------------------------------------------------------------------------------
	# 7) Auto Sort: BIZERBA by numeric IP then ISHIDA, renumber, template names, immediate SAVE
	# ----------------------------------------------------------------------------------------------
	$btnAuto.add_Click({
			# Deep snapshots split by brand to avoid DataRow detach issues
			$snapB = @()
			$snapI = @()
			foreach ($row in $dt.Rows)
			{
				$o = [ordered]@{ }
				foreach ($col in $dt.Columns) { $o[$col.ColumnName] = $row[$col.ColumnName] }
				if ($o.ScaleBrand -eq 'BIZERBA') { $snapB += $o }
				elseif ($o.ScaleBrand -eq 'ISHIDA') { $snapI += $o }
			}

			# Sort by numeric IPDevice (non-numeric last), then by raw IPDevice
			$byDevice = @(
				@{ Expression = { $n = $null; if ([int]::TryParse([string]$_.IPDevice, [ref]$n)) { $n } else { [int]::MaxValue } } },
				@{ Expression = { $_.IPDevice } }
			)
			$snapB = @($snapB | Sort-Object $byDevice)
			$snapI = @($snapI | Sort-Object $byDevice)

			$dt.BeginLoadData()
			try
			{
				$dt.Rows.Clear()
				foreach ($s in @($snapB) + @($snapI)) { $nr = $dt.NewRow(); foreach ($k in $s.Keys) { $nr[$k] = [string]$s[$k] }; [void]$dt.Rows.Add($nr) }
			}
			finally { $dt.EndLoadData() }

			# Template naming + BufferTime for BIZERBA; Ishida WMAI name rule
			$firstB = $true
			foreach ($r in $dt.Rows)
			{
				if ($r['ScaleBrand'] -eq 'BIZERBA')
				{
					$r['ScaleName'] = "Scale $($r['IPDevice'])"
					if ($firstB) { $r['BufferTime'] = '1'; $firstB = $false }
					else { $r['BufferTime'] = '5' }
				}
			}
			$ishWmai = @($dt.Rows | Where-Object { $_['ScaleBrand'] -eq 'ISHIDA' -and $_['ScaleModel'] -eq 'WMAI' })
			if ($ishWmai.Count -gt 1) { foreach ($r in $ishWmai) { $r['ScaleName'] = "Ishida Wrapper $($r['IPDevice'])" } }
			elseif ($ishWmai.Count -eq 1) { $ishWmai[0]['ScaleName'] = 'Ishida Wrapper' }

			# Renumber: BIZERBA start at 10; ISHIDA start at (last BIZERBA)+10 or 20 if none
			$codeB = 10
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'BIZERBA') { $r['ScaleCode'] = [string]$codeB; $codeB = $codeB + 1 } }
			$ishidaStart = 20
			if ($codeB -gt 10) { $ishidaStart = ($codeB - 1) + 10 }
			$codeI = $ishidaStart
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'ISHIDA') { $r['ScaleCode'] = [string]$codeI; $codeI = $codeI + 1 } }

			$grid.Refresh()

			# Immediately SAVE (same validation rules as Save button)
			if (& $saveGrid 'Auto Sort')
			{
				[void][System.Windows.Forms.MessageBox]::Show("Auto Sort applied, saved, and refreshed.", "Success", 'OK', 'Information')
				Write_Log "Auto Sort applied, saved, and refreshed." "green"
			}
		})

	# ----------------------------------------------------------------------------------------------
	# 8) Refresh / Close / Cancel wiring
	# ----------------------------------------------------------------------------------------------
	$btnRefresh.add_Click({
			Write_Log "Refreshing view from database..." "blue"
			[void](& $reloadFromDb 'refresh')
		})

	$btnClose.add_Click({ $form.Close() })
	$btnCancel.add_Click({
			Write_Log "Cancel: reloading from database..." "yellow"
			$btnRefresh.PerformClick()
		})

	# ----------------------------------------------------------------------------------------------
	# 9) Show dialog
	# ----------------------------------------------------------------------------------------------
	[void]$form.ShowDialog()
	Write_Log "==================== Edit_TBS_SCL_ver520_Table Closed ====================" "blue"
}
