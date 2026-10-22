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
#   - $script:FunctionResults: DBSERVER, DBNAME, ConnectionString, SQL_UserID/SQL_UserPwd, SqlModuleName
#
# Writes:
#   - store database table TBS_SCL_ver520
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
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
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
	# 2) Pick SQL runner (Invoke-Sqlcmd if available, else .NET SqlClient)
	# ----------------------------------------------------------------------------------------------
	$InvokeSqlCmd = $null
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch
		{
			Write_Log "SQL module '$SqlModuleName' not loaded. Using .NET SqlClient fallback." "yellow"
		}
	}
	else
	{
		Write_Log "No SQL module name provided. Using .NET SqlClient fallback." "yellow"
	}
	$supportsConnectionString = $InvokeSqlCmd -and ($InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString')

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
			$serverIPs = [System.Net.Dns]::GetHostAddresses($server)
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

	Write_Log "Loading table for edit..." "blue"

	# CHANGE: Inline query execution (no $runQuery helper). Return rows.
	$rows = $null
	try
	{
		if ($InvokeSqlCmd)
		{
			if ($supportsConnectionString)
			{
				$rows = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
			}
			else
			{
				$rows = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
			}
		}
		else
		{
			$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
			$csb.ConnectionString = $connectionString
			$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
			$conn.Open()
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $selectQuery; $cmd.CommandTimeout = 0
			$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
			$td = New-Object System.Data.DataTable
			[void]$da.Fill($td)
			$conn.Close()
			$rows = $td
		}
	}
	catch
	{
		Write_Log "SQL error: $($_.Exception.Message)" "red"
		return
	}
	if (-not $rows) { return }

	$dt = New-Object System.Data.DataTable
	foreach ($c in 'ScaleCode', 'ScaleName', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice', 'BufferTime', 'Active',
		'OrigScaleBrand', 'OrigScaleModel', 'OrigIPNetwork', 'OrigIPDevice')
	{
		[void]$dt.Columns.Add($c)
	}

	foreach ($r in $rows)
	{
		$nr = $dt.NewRow()
		$nr['ScaleCode'] = [string]$r.ScaleCode
		$nr['ScaleName'] = [string]$r.ScaleName
		$nr['ScaleBrand'] = [string]$r.ScaleBrand
		$nr['ScaleModel'] = [string]$r.ScaleModel
		$nr['IPNetwork'] = [string]$r.IPNetwork
		$nr['IPDevice'] = [string]$r.IPDevice
		$nr['BufferTime'] = [string]$r.BufferTime
		$nr['Active'] = [string]$r.Active
		# ORIGINAL KEYS SNAPSHOT
		$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
		$nr['OrigScaleModel'] = [string]$r.ScaleModel
		$nr['OrigIPNetwork'] = [string]$r.IPNetwork
		$nr['OrigIPDevice'] = [string]$r.IPDevice
		[void]$dt.Rows.Add($nr)
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
		$lbl.Text = "Edit directly [?][?][?] Auto-Sort renumbers & saves [?][?][?] Active: Y/N [?][?][?] IPNetwork [?][?][?] '$allowedPrefix.' (accepts $allowedPrefix, $allowedPrefix.0/*/x/.)"
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
	# Helper-less inline SAVE logic (duplicated in Save and Auto Sort to keep "all inline" constraint)
	# ----------------------------------------------------------------------------------------------
	$sqlSaveTemplate = @'
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
    ON  tgt.ScaleBrand = src.OrigScaleBrand
    AND tgt.ScaleModel = src.OrigScaleModel
    AND tgt.IPNetwork  = src.OrigIPNetwork
    AND tgt.IPDevice   = src.OrigIPDevice;

  COMMIT TRAN;
END TRY
BEGIN CATCH
  IF @@TRANCOUNT > 0 ROLLBACK TRAN;
  DECLARE @msg nvarchar(4000) = ERROR_MESSAGE();
  RAISERROR(@msg, 16, 1);
END CATCH;
'@

	# ----------------------------------------------------------------------------------------------
	# 6) Button: Save (validate + write exactly what's in grid; no renumber)
	# ----------------------------------------------------------------------------------------------
	$btnSave.add_Click({
			# Commit any in-progress grid edits (inline, no helper)
			if ($grid -and $grid.IsCurrentCellInEditMode) { [void]$grid.EndEdit() }
			else { if ($grid) { [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } }

			# Validate ScaleCode unique & numeric; Active Y/N; IPNetwork matches allowed /24 if available
			$seenCodes = @{ }
			$rowIdx = 0
			foreach ($row in $dt.Rows)
			{
				$raw = $row['ScaleCode']
				$val = $null
				if (-not [int]::TryParse([string]$raw, [ref]$val))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid ScaleCode at row #$($rowIdx + 1): '$raw' is not a number.", "Invalid Code", 'OK', 'Error')
					return
				}
				if ($seenCodes.ContainsKey($val))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Duplicate ScaleCode '$val' detected (row #$($rowIdx + 1)).", "Duplicate Code", 'OK', 'Error')
					return
				}
				$seenCodes[$val] = $true

				$actRaw = [string]$row['Active']; if ($null -eq $actRaw) { $actRaw = '' }
				$actUp = $actRaw.Trim().ToUpper()
				if ($actUp -ne 'Y' -and $actUp -ne 'N')
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid Active value at row #$($rowIdx + 1): '$actRaw'. Use Y or N.", "Invalid Active", 'OK', 'Error')
					return
				}
				$row['Active'] = $actUp

				# --- IPNetwork normalization/enforcement (ALWAYS ends with a period) -----------------
				$ipnRaw = [string]$row['IPNetwork']; if ($null -eq $ipnRaw) { $ipnRaw = '' }
				$ipnRaw = $ipnRaw.Trim()

				if ($allowedPrefix)
				{
					# Validate input looks like allowedPrefix.(something optional), then normalize strictly to "$allowedPrefix."
					$m = [regex]::Match($ipnRaw, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if (-not $m.Success)
					{
						[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #$($rowIdx + 1): '$ipnRaw'. Expected something within '$allowedPrefix.*'.", "Invalid IPNetwork", 'OK', 'Error')
						return
					}
					$prefix = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)"
					if ($prefix -ne $allowedPrefix)
					{
						[void][System.Windows.Forms.MessageBox]::Show("IPNetwork at row #$($rowIdx + 1) must be within '$allowedPrefix.*'. Got '$ipnRaw'.", "IPNetwork Out of Range", 'OK', 'Error')
						return
					}

					# CHANGE: enforce trailing period
					$row['IPNetwork'] = "$allowedPrefix." # <-- always ends with a '.'
				}
				else
				{
					# NEW: Even without allowedPrefix, if parsable A.B.C(.anything), normalize to A.B.C.
					$mGen = [regex]::Match($ipnRaw, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if ($mGen.Success)
					{
						$prefixGen = "$($mGen.Groups[1].Value).$($mGen.Groups[2].Value).$($mGen.Groups[3].Value)"
						$row['IPNetwork'] = "$prefixGen." # <-- enforce trailing '.'
					}
					else
					{
						# If totally unparsable, keep original but warn & block save
						[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #$($rowIdx + 1): '$ipnRaw'. Expected a pattern like A.B.C or A.B.C.X", "Invalid IPNetwork", 'OK', 'Error')
						return
					}
				}
				# --------------------------------------------------------------------------------------

				$rowIdx = $rowIdx + 1
			}

			# Build VALUES block with proper SQL escaping
			$values = New-Object System.Collections.Generic.List[string]
			foreach ($row in $dt.Rows)
			{
				$nBrand = ([string]$row['ScaleBrand']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nModel = ([string]$row['ScaleModel']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nIPNet = ([string]$row['IPNetwork']).Replace("`r", "").Replace("`n", "") -replace "'", "''" # already ends with '.'
				$nIPDev = ([string]$row['IPDevice']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nCode = [int]$row['ScaleCode']
				$nBuf = ([string]$row['BufferTime']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nName = ([string]$row['ScaleName']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nAct = ([string]$row['Active']).Replace("`r", "").Replace("`n", "") -replace "'", "''"

				$oBrand = (([string]$row['OrigScaleBrand'])) -replace "'", "''"
				$oModel = (([string]$row['OrigScaleModel'])) -replace "'", "''"
				$oIPNet = (([string]$row['OrigIPNetwork'])) -replace "'", "''"
				$oIPDev = (([string]$row['OrigIPDevice'])) -replace "'", "''"

				$values.Add("('$oBrand','$oModel','$oIPNet','$oIPDev','$nBrand','$nModel','$nIPNet','$nIPDev',$nCode,'$nBuf','$nName','$nAct')")
			}
			if ($values.Count -eq 0)
			{
				Write_Log "Nothing to save." "yellow"
				return
			}

			$sqlSave = $sqlSaveTemplate.Replace('$(VALUES_BLOCK)', ([string]::Join(",`r`n  ", $values)))

			Write_Log "Saving to database..." "blue"

			# Execute non-query inline
			$ok = $true
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						& $InvokeSqlCmd -ConnectionString $connectionString -Query $sqlSave -ErrorAction Stop | Out-Null
					}
					else
					{
						& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $sqlSave -ErrorAction Stop | Out-Null
					}
				}
				else
				{
					$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb.ConnectionString = $connectionString
					$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
					$conn.Open()
					$cmd = $conn.CreateCommand(); $cmd.CommandText = $sqlSave; $cmd.CommandTimeout = 0
					[void]$cmd.ExecuteNonQuery()
					$conn.Close()
				}
			}
			catch
			{
				$ok = $false
				Write_Log "Save failed: $($_.Exception.Message)" "red"
				[void][System.Windows.Forms.MessageBox]::Show("Save failed. Check logs.", "Error", 'OK', 'Error')
			}
			if (-not $ok) { return }

			# Refresh view from DB
			Write_Log "Refreshing view..." "blue"
			$fresh = $null
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						$fresh = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
					}
					else
					{
						$fresh = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					}
				}
				else
				{
					$csb2 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb2.ConnectionString = $connectionString
					$conn2 = New-Object System.Data.SqlClient.SqlConnection $csb2.ConnectionString
					$conn2.Open()
					$cmd2 = $conn2.CreateCommand(); $cmd2.CommandText = $selectQuery; $cmd2.CommandTimeout = 0
					$da2 = New-Object System.Data.SqlClient.SqlDataAdapter $cmd2
					$td2 = New-Object System.Data.DataTable
					[void]$da2.Fill($td2)
					$conn2.Close()
					$fresh = $td2
				}
			}
			catch
			{
				Write_Log "SQL error (refresh): $($_.Exception.Message)" "red"
				return
			}

			if ($fresh)
			{
				$dt.Rows.Clear() | Out-Null
				foreach ($r in $fresh)
				{
					$nr = $dt.NewRow()
					$nr['ScaleCode'] = [string]$r.ScaleCode
					$nr['ScaleName'] = [string]$r.ScaleName
					$nr['ScaleBrand'] = [string]$r.ScaleBrand
					$nr['ScaleModel'] = [string]$r.ScaleModel
					$nr['IPNetwork'] = [string]$r.IPNetwork
					$nr['IPDevice'] = [string]$r.IPDevice
					$nr['BufferTime'] = [string]$r.BufferTime
					$nr['Active'] = [string]$r.Active
					$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
					$nr['OrigScaleModel'] = [string]$r.ScaleModel
					$nr['OrigIPNetwork'] = [string]$r.IPNetwork
					$nr['OrigIPDevice'] = [string]$r.IPDevice
					[void]$dt.Rows.Add($nr)
				}
				$grid.Refresh()
			}

			[void][System.Windows.Forms.MessageBox]::Show("Saved and refreshed.", "Success", 'OK', 'Information')
			Write_Log "Saved and refreshed." "green"
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
				$o = [ordered]@{
					ScaleCode	   = $row['ScaleCode']; ScaleName = $row['ScaleName']; ScaleBrand = $row['ScaleBrand']; ScaleModel = $row['ScaleModel']
					IPNetwork	   = $row['IPNetwork']; IPDevice = $row['IPDevice']; BufferTime = $row['BufferTime']; Active = $row['Active']
					OrigScaleBrand = $row['OrigScaleBrand']; OrigScaleModel = $row['OrigScaleModel']; OrigIPNetwork = $row['OrigIPNetwork']; OrigIPDevice = $row['OrigIPDevice']
				}
				if ($o.ScaleBrand -eq 'BIZERBA') { $snapB += $o }
				elseif ($o.ScaleBrand -eq 'ISHIDA') { $snapI += $o }
			}

			# Sort by numeric IPDevice (non-numeric last), then by raw IPDevice
			$snapB = $snapB | Sort-Object @{
				Expression = {
					$n = $null; if ([int]::TryParse([string]$_.IPDevice, [ref]$n)) { $n }
					else { [int]::MaxValue }
				}
			},
										  @{ Expression = { $_.IPDevice } }
			$snapI = $snapI | Sort-Object @{
				Expression = {
					$n = $null; if ([int]::TryParse([string]$_.IPDevice, [ref]$n)) { $n }
					else { [int]::MaxValue }
				}
			},
										  @{ Expression = { $_.IPDevice } }

			$dt.BeginLoadData()
			try
			{
				$dt.Rows.Clear()
				foreach ($s in $snapB) { $nr = $dt.NewRow(); foreach ($k in $s.Keys) { $nr[$k] = [string]$s[$k] }; [void]$dt.Rows.Add($nr) }
				foreach ($s in $snapI) { $nr = $dt.NewRow(); foreach ($k in $s.Keys) { $nr[$k] = [string]$s[$k] }; [void]$dt.Rows.Add($nr) }
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
			$ishWmai = @()
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'ISHIDA' -and $r['ScaleModel'] -eq 'WMAI') { $ishWmai += $r } }
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
			if ($grid -and $grid.IsCurrentCellInEditMode) { [void]$grid.EndEdit() }
			else { if ($grid) { [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } }

			$seenCodes2 = @{ }
			$rowIdx2 = 0
			foreach ($row2 in $dt.Rows)
			{
				$raw2 = $row2['ScaleCode']
				$val2 = $null
				if (-not [int]::TryParse([string]$raw2, [ref]$val2))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid ScaleCode at row #$($rowIdx2 + 1): '$raw2' is not a number.", "Invalid Code", 'OK', 'Error')
					return
				}
				if ($seenCodes2.ContainsKey($val2))
				{
					[void][System.Windows.Forms.MessageBox]::Show("Duplicate ScaleCode '$val2' detected (row #$($rowIdx2 + 1)).", "Duplicate Code", 'OK', 'Error')
					return
				}
				$seenCodes2[$val2] = $true

				$actRaw2 = [string]$row2['Active']; if ($null -eq $actRaw2) { $actRaw2 = '' }
				$actUp2 = $actRaw2.Trim().ToUpper()
				if ($actUp2 -ne 'Y' -and $actUp2 -ne 'N')
				{
					[void][System.Windows.Forms.MessageBox]::Show("Invalid Active value at row #$($rowIdx2 + 1): '$actRaw2'. Use Y or N.", "Invalid Active", 'OK', 'Error')
					return
				}
				$row2['Active'] = $actUp2

				# --- IPNetwork normalization/enforcement (ALWAYS ends with a period) -----------------
				$ipnRaw2 = [string]$row2['IPNetwork']; if ($null -eq $ipnRaw2) { $ipnRaw2 = '' }
				$ipnRaw2 = $ipnRaw2.Trim()

				if ($allowedPrefix)
				{
					$m2 = [regex]::Match($ipnRaw2, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if (-not $m2.Success)
					{
						[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #$($rowIdx2 + 1): '$ipnRaw2'. Expected something within '$allowedPrefix.*'.", "Invalid IPNetwork", 'OK', 'Error')
						return
					}
					$prefix2 = "$($m2.Groups[1].Value).$($m2.Groups[2].Value).$($m2.Groups[3].Value)"
					if ($prefix2 -ne $allowedPrefix)
					{
						[void][System.Windows.Forms.MessageBox]::Show("IPNetwork at row #$($rowIdx2 + 1) must be within '$allowedPrefix.*'. Got '$ipnRaw2'.", "IPNetwork Out of Range", 'OK', 'Error')
						return
					}

					# CHANGE: enforce trailing period
					$row2['IPNetwork'] = "$allowedPrefix." # <-- always ends with '.'
				}
				else
				{
					# NEW: Even without allowedPrefix, normalize any parsable A.B.C(.anything) to A.B.C.
					$mGen2 = [regex]::Match($ipnRaw2, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if ($mGen2.Success)
					{
						$prefixGen2 = "$($mGen2.Groups[1].Value).$($mGen2.Groups[2].Value).$($mGen2.Groups[3].Value)"
						$row2['IPNetwork'] = "$prefixGen2." # <-- enforce trailing '.'
					}
					else
					{
						[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #$($rowIdx2 + 1): '$ipnRaw2'. Expected a pattern like A.B.C or A.B.C.X", "Invalid IPNetwork", 'OK', 'Error')
						return
					}
				}
				# --------------------------------------------------------------------------------------

				$rowIdx2 = $rowIdx2 + 1
			}

			$values2 = New-Object System.Collections.Generic.List[string]
			foreach ($row3 in $dt.Rows)
			{
				$nBrand2 = ([string]$row3['ScaleBrand']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nModel2 = ([string]$row3['ScaleModel']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nIPNet2 = ([string]$row3['IPNetwork']).Replace("`r", "").Replace("`n", "") -replace "'", "''" # already ends with '.'
				$nIPDev2 = ([string]$row3['IPDevice']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nCode2 = [int]$row3['ScaleCode']
				$nBuf2 = ([string]$row3['BufferTime']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nName2 = ([string]$row3['ScaleName']).Replace("`r", "").Replace("`n", "") -replace "'", "''"
				$nAct2 = ([string]$row3['Active']).Replace("`r", "").Replace("`n", "") -replace "'", "''"

				$oBrand2 = (([string]$row3['OrigScaleBrand'])) -replace "'", "''"
				$oModel2 = (([string]$row3['OrigScaleModel'])) -replace "'", "''"
				$oIPNet2 = (([string]$row3['OrigIPNetwork'])) -replace "'", "''"
				$oIPDev2 = (([string]$row3['OrigIPDevice'])) -replace "'", "''"

				$values2.Add("('$oBrand2','$oModel2','$oIPNet2','$oIPDev2','$nBrand2','$nModel2','$nIPNet2','$nIPDev2',$nCode2,'$nBuf2','$nName2','$nAct2')")
			}
			if ($values2.Count -eq 0)
			{
				Write_Log "Nothing to save after Auto Sort." "yellow"
				return
			}

			$sqlSave2 = $sqlSaveTemplate.Replace('$(VALUES_BLOCK)', ([string]::Join(",`r`n  ", $values2)))

			Write_Log "Saving to database after Auto Sort..." "blue"

			$ok2 = $true
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						& $InvokeSqlCmd -ConnectionString $connectionString -Query $sqlSave2 -ErrorAction Stop | Out-Null
					}
					else
					{
						& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $sqlSave2 -ErrorAction Stop | Out-Null
					}
				}
				else
				{
					$csb3 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb3.ConnectionString = $connectionString
					$conn3 = New-Object System.Data.SqlClient.SqlConnection $csb3.ConnectionString
					$conn3.Open()
					$cmd3 = $conn3.CreateCommand(); $cmd3.CommandText = $sqlSave2; $cmd3.CommandTimeout = 0
					[void]$cmd3.ExecuteNonQuery()
					$conn3.Close()
				}
			}
			catch
			{
				$ok2 = $false
				Write_Log "Save (Auto Sort) failed: $($_.Exception.Message)" "red"
				[void][System.Windows.Forms.MessageBox]::Show("Save after Auto Sort failed. Check logs.", "Error", 'OK', 'Error')
			}
			if (-not $ok2) { return }

			# Refresh view from DB
			Write_Log "Refreshing view after Auto Sort..." "blue"
			$fresh2 = $null
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						$fresh2 = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
					}
					else
					{
						$fresh2 = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					}
				}
				else
				{
					$csb4 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb4.ConnectionString = $connectionString
					$conn4 = New-Object System.Data.SqlClient.SqlConnection $csb4.ConnectionString
					$conn4.Open()
					$cmd4 = $conn4.CreateCommand(); $cmd4.CommandText = $selectQuery; $cmd4.CommandTimeout = 0
					$da4 = New-Object System.Data.SqlClient.SqlDataAdapter $cmd4
					$td4 = New-Object System.Data.DataTable
					[void]$da4.Fill($td4)
					$conn4.Close()
					$fresh2 = $td4
				}
			}
			catch
			{
				Write_Log "SQL error (refresh after Auto Sort): $($_.Exception.Message)" "red"
				return
			}

			if ($fresh2)
			{
				$dt.Rows.Clear() | Out-Null
				foreach ($r in $fresh2)
				{
					$nr = $dt.NewRow()
					$nr['ScaleCode'] = [string]$r.ScaleCode
					$nr['ScaleName'] = [string]$r.ScaleName
					$nr['ScaleBrand'] = [string]$r.ScaleBrand
					$nr['ScaleModel'] = [string]$r.ScaleModel
					$nr['IPNetwork'] = [string]$r.IPNetwork
					$nr['IPDevice'] = [string]$r.IPDevice
					$nr['BufferTime'] = [string]$r.BufferTime
					$nr['Active'] = [string]$r.Active
					$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
					$nr['OrigScaleModel'] = [string]$r.ScaleModel
					$nr['OrigIPNetwork'] = [string]$r.IPNetwork
					$nr['OrigIPDevice'] = [string]$r.IPDevice
					[void]$dt.Rows.Add($nr)
				}
				$grid.Refresh()
			}

			[void][System.Windows.Forms.MessageBox]::Show("Auto Sort applied, saved, and refreshed.", "Success", 'OK', 'Information')
			Write_Log "Auto Sort applied, saved, and refreshed." "green"
		})

	# ----------------------------------------------------------------------------------------------
	# 8) Refresh / Close / Cancel wiring (inline)
	# ----------------------------------------------------------------------------------------------
	$btnRefresh.add_Click({
			Write_Log "Refreshing view from database..." "blue"
			$fresh = $null
			try
			{
				if ($InvokeSqlCmd)
				{
					if ($supportsConnectionString)
					{
						$fresh = & $InvokeSqlCmd -ConnectionString $connectionString -Query $selectQuery -ErrorAction Stop
					}
					else
					{
						$fresh = & $InvokeSqlCmd -ServerInstance $server -Database $database -Query $selectQuery -ErrorAction Stop
					}
				}
				else
				{
					$csb5 = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
					$csb5.ConnectionString = $connectionString
					$conn5 = New-Object System.Data.SqlClient.SqlConnection $csb5.ConnectionString
					$conn5.Open()
					$cmd5 = $conn5.CreateCommand(); $cmd5.CommandText = $selectQuery; $cmd5.CommandTimeout = 0
					$da5 = New-Object System.Data.SqlClient.SqlDataAdapter $cmd5
					$td5 = New-Object System.Data.DataTable
					[void]$da5.Fill($td5)
					$conn5.Close()
					$fresh = $td5
				}
			}
			catch
			{
				Write_Log "SQL error (refresh): $($_.Exception.Message)" "red"
				return
			}

			if ($fresh)
			{
				$dt.Rows.Clear() | Out-Null
				foreach ($r in $fresh)
				{
					$nr = $dt.NewRow()
					$nr['ScaleCode'] = [string]$r.ScaleCode
					$nr['ScaleName'] = [string]$r.ScaleName
					$nr['ScaleBrand'] = [string]$r.ScaleBrand
					$nr['ScaleModel'] = [string]$r.ScaleModel
					$nr['IPNetwork'] = [string]$r.IPNetwork
					$nr['IPDevice'] = [string]$r.IPDevice
					$nr['BufferTime'] = [string]$r.BufferTime
					$nr['Active'] = [string]$r.Active
					$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
					$nr['OrigScaleModel'] = [string]$r.ScaleModel
					$nr['OrigIPNetwork'] = [string]$r.IPNetwork
					$nr['OrigIPDevice'] = [string]$r.IPDevice
					[void]$dt.Rows.Add($nr)
				}
				$grid.Refresh()
			}
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
