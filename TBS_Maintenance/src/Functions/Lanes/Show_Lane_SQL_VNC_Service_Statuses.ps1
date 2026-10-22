# ===================================================================================================
#                              FUNCTION: Show_Lane_SQL_VNC_Service_Statuses
# ---------------------------------------------------------------------------------------------------
# Description:
#   Opens a styled popup that checks the SQL and VNC services on all lanes of the store in
#   the background and fills a grid as results arrive. Shows a message box when no usable
#   lanes are found.
#
# Variables:
#   - $script:FunctionResults['StoreNumber']: store header text
#   - $script:FunctionResults['LaneNumToMachineName']: lane map (input)
# ===================================================================================================

function Show_Lane_SQL_VNC_Service_Statuses
{
	$storeNumber = $script:FunctionResults['StoreNumber']
	$laneMap = $script:FunctionResults['LaneNumToMachineName']

	if (-not $laneMap -or $laneMap.Count -eq 0)
	{
		[System.Windows.Forms.MessageBox]::Show(
			"No lanes were found for this store.",
			"Lane Services",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Information
		) | Out-Null
		return
	}

	$selectedLanes = @($laneMap.Keys | Where-Object { $_ -and $laneMap.ContainsKey([string]$_) } | Sort-Object { [int]$_ })
	if (-not $selectedLanes -or $selectedLanes.Count -eq 0)
	{
		[System.Windows.Forms.MessageBox]::Show(
			"No usable lanes were found.",
			"Lane Services",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Information
		) | Out-Null
		return
	}

	$formSvc = New-Object System.Windows.Forms.Form
	$formSvc.Text = "Lane SQL / VNC Services"
	$formSvc.StartPosition = 'CenterParent'
	$formSvc.Size = New-Object System.Drawing.Size(980, 520)
	$formSvc.MinimumSize = New-Object System.Drawing.Size(860, 420)
	$formSvc.MaximizeBox = $true
	$formSvc.MinimizeBox = $false

	$layoutSvc = New-Object System.Windows.Forms.TableLayoutPanel
	$layoutSvc.Dock = 'Fill'
	$layoutSvc.ColumnCount = 1
	$layoutSvc.RowCount = 3
	[void]$layoutSvc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutSvc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutSvc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$formSvc.Controls.Add($layoutSvc)

	$lblSvcHeader = New-Object System.Windows.Forms.Label
	$lblSvcHeader.Text = ("Store {0} - SQL and VNC service status for all {1} lane(s)" -f $storeNumber, $selectedLanes.Count)
	$lblSvcHeader.AutoSize = $true
	$lblSvcHeader.Margin = New-Object System.Windows.Forms.Padding(12, 12, 12, 6)
	$layoutSvc.Controls.Add($lblSvcHeader, 0, 0)

	$gridSvc = New-Object System.Windows.Forms.DataGridView
	$gridSvc.Dock = 'Fill'
	$gridSvc.AllowUserToAddRows = $false
	$gridSvc.AllowUserToDeleteRows = $false
	$gridSvc.AllowUserToResizeRows = $false
	$gridSvc.ReadOnly = $true
	$gridSvc.MultiSelect = $false
	$gridSvc.RowHeadersVisible = $false
	$gridSvc.SelectionMode = 'FullRowSelect'
	$gridSvc.AutoSizeColumnsMode = 'Fill'
	$gridSvc.Margin = New-Object System.Windows.Forms.Padding(12, 0, 12, 8)
	[void]$gridSvc.Columns.Add("Lane", "Lane")
	[void]$gridSvc.Columns.Add("Machine", "Machine")
	[void]$gridSvc.Columns.Add("SqlService", "SQL Service")
	[void]$gridSvc.Columns.Add("SqlStatus", "SQL Status")
	[void]$gridSvc.Columns.Add("VncService", "VNC Service")
	[void]$gridSvc.Columns.Add("VncStatus", "VNC Status")
	$gridSvc.Columns["Lane"].FillWeight = 55
	$gridSvc.Columns["Machine"].FillWeight = 95
	$gridSvc.Columns["SqlService"].FillWeight = 120
	$gridSvc.Columns["SqlStatus"].FillWeight = 85
	$gridSvc.Columns["VncService"].FillWeight = 120
	$gridSvc.Columns["VncStatus"].FillWeight = 85
	$layoutSvc.Controls.Add($gridSvc, 0, 1)

	$pnlSvcBottom = New-Object System.Windows.Forms.TableLayoutPanel
	$pnlSvcBottom.Dock = 'Fill'
	$pnlSvcBottom.ColumnCount = 1
	$pnlSvcBottom.RowCount = 2
	$pnlSvcBottom.Margin = New-Object System.Windows.Forms.Padding(12, 0, 12, 12)
	[void]$pnlSvcBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$pnlSvcBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$pnlSvcBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$layoutSvc.Controls.Add($pnlSvcBottom, 0, 2)

	$lblSvcStatus = New-Object System.Windows.Forms.Label
	$lblSvcStatus.Text = "Checking all lanes..."
	$lblSvcStatus.AutoSize = $true
	$lblSvcStatus.Anchor = 'Left'
	$lblSvcStatus.Margin = New-Object System.Windows.Forms.Padding(0, 8, 12, 0)
	$pnlSvcBottom.Controls.Add($lblSvcStatus, 0, 0)

	$pnlSvcActions = New-Object System.Windows.Forms.FlowLayoutPanel
	$pnlSvcActions.AutoSize = $true
	$pnlSvcActions.WrapContents = $false
	$pnlSvcActions.AutoScroll = $true
	$pnlSvcActions.FlowDirection = 'LeftToRight'
	$pnlSvcActions.Dock = 'Fill'
	$pnlSvcActions.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
	$pnlSvcBottom.Controls.Add($pnlSvcActions, 0, 1)

	$btnSvcRefresh = New-Object System.Windows.Forms.Button
	$btnSvcRefresh.Text = "Refresh All"
	$btnSvcRefresh.Size = New-Object System.Drawing.Size(110, 32)
	$pnlSvcActions.Controls.Add($btnSvcRefresh)

	$btnSvcClose = New-Object System.Windows.Forms.Button
	$btnSvcClose.Text = "Close"
	$btnSvcClose.Size = New-Object System.Drawing.Size(110, 32)
	$btnSvcClose.Add_Click({ $formSvc.Close() })
	$pnlSvcActions.Controls.Add($btnSvcClose)
	$formSvc.AcceptButton = $btnSvcClose
	$formSvc.CancelButton = $btnSvcClose

	if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $formSvc }

	$svcJobs = @{ }
	$svcActionJobs = @{ }
	$rowMap = @{ }
	$totalRows = $selectedLanes.Count

	$startActionJobForSelectedRow = {
		param (
			[string]$ServiceKind,
			[string]$RequestedAction
		)

		if (-not $gridSvc.SelectedRows -or $gridSvc.SelectedRows.Count -eq 0)
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Select a lane first.",
				"Lane Services",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			) | Out-Null
			return
		}

		$row = $gridSvc.SelectedRows[0]
		if (-not $row) { return }

		$laneKey = [string]$row.Cells["Lane"].Value
		$machineName = [string]$row.Cells["Machine"].Value
		$expectedSqlService = [string]$row.Cells["SqlService"].Value
		$serviceName = if ($ServiceKind -eq 'SQL') { [string]$row.Cells["SqlService"].Value }
		else { [string]$row.Cells["VncService"].Value }
		$statusColumn = if ($ServiceKind -eq 'SQL') { 'SqlStatus' }
		else { 'VncStatus' }

		if ([string]::IsNullOrWhiteSpace($laneKey) -or [string]::IsNullOrWhiteSpace($machineName)) { return }
		if ($svcActionJobs.ContainsKey($laneKey))
		{
			[System.Windows.Forms.MessageBox]::Show(
				"A service action is already running for lane $laneKey.",
				"Lane Services",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			) | Out-Null
			return
		}

		if ([string]::IsNullOrWhiteSpace($serviceName) -or $serviceName -in @('Checking...', 'Not Found', 'Unavailable'))
		{
			[System.Windows.Forms.MessageBox]::Show(
				("$ServiceKind service was not found for lane $laneKey."),
				"Lane Services",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			) | Out-Null
			return
		}

		$row.Cells[$statusColumn].Value = "$RequestedAction..."
		$lblSvcStatus.Text = ("{0} {1} on lane {2}..." -f $RequestedAction, $ServiceKind, $laneKey)

		$job = Start-Job -Name ("LaneServiceAction_{0}_{1}_{2}" -f $machineName, $ServiceKind, $RequestedAction) -ArgumentList $laneKey, $machineName, $expectedSqlService, $ServiceKind, $serviceName, $RequestedAction -ScriptBlock {
			param (
				[string]$Lane,
				[string]$Machine,
				[string]$ExpectedSqlService,
				[string]$ServiceKind,
				[string]$ServiceName,
				[string]$RequestedAction
			)

			$getServiceSnapshot = {
				param (
					[string]$SnapshotLane,
					[string]$SnapshotMachine,
					[string]$SnapshotExpectedSqlService
				)

				$result = [ordered]@{
					Lane	   = $SnapshotLane
					Machine    = $SnapshotMachine
					SqlService = if ([string]::IsNullOrWhiteSpace($SnapshotExpectedSqlService)) { 'MSSQLSERVER' } else { $SnapshotExpectedSqlService }
					SqlStatus  = 'Not Found'
					VncService = 'Not Found'
					VncStatus  = 'Not Found'
					ActionSummary = $null
				}

				$services = $null
				$serviceFetchError = $null
				try
				{
					$services = Get-Service -ComputerName $SnapshotMachine -ErrorAction Stop | Select-Object Name, DisplayName, Status
				}
				catch
				{
					$serviceFetchError = $_.Exception.Message
				}

				if (-not $services)
				{
					try
					{
						$services = Get-CimInstance -ClassName Win32_Service -ComputerName $SnapshotMachine -ErrorAction Stop | Select-Object Name, DisplayName, @{ Name = 'Status'; Expression = { $_.State } }
						$serviceFetchError = $null
					}
					catch
					{
						$serviceFetchError = $_.Exception.Message
					}
				}

				if (-not $services)
				{
					$result.SqlStatus = if ([string]::IsNullOrWhiteSpace($serviceFetchError)) { 'Error' }
					else { "Error: $serviceFetchError" }
					$result.VncStatus = if ([string]::IsNullOrWhiteSpace($serviceFetchError)) { 'Error' }
					else { "Error: $serviceFetchError" }
					$result.VncService = 'Unavailable'
					return [pscustomobject]$result
				}

				$sqlService = $services | Where-Object { $_.Name -ieq $SnapshotExpectedSqlService } | Select-Object -First 1
				if (-not $sqlService)
				{
					$sqlService = $services | Where-Object { $_.Name -like 'MSSQL*' -or $_.DisplayName -match '(?i)^SQL Server' } | Sort-Object Name | Select-Object -First 1
				}
				if ($sqlService)
				{
					$result.SqlService = [string]$sqlService.Name
					$result.SqlStatus = [string]$sqlService.Status
				}

				$vncService = $services | Where-Object {
					$svcName = [string]$_.Name
					$svcDisplay = [string]$_.DisplayName
					($svcName -match '(?i)^(uvnc_service|winvnc.*|tvnserver|vncserver|tigervnc.*|tightvnc.*)$') -or
					($svcDisplay -match '(?i)ultravnc|tightvnc|tigervnc|vnc')
				} | Sort-Object Name | Select-Object -First 1
				if ($vncService)
				{
					$result.VncService = [string]$vncService.Name
					$result.VncStatus = [string]$vncService.Status
				}

				return [pscustomobject]$result
			}

			$actionSummary = "$RequestedAction request sent."
			$serviceEscaped = $ServiceName.Replace("'", "''")
			try
			{
				$svcObject = Get-CimInstance -ClassName Win32_Service -ComputerName $Machine -Filter "Name='$serviceEscaped'" -ErrorAction Stop
				switch ($RequestedAction)
				{
					'Start'   { $null = Invoke-CimMethod -InputObject $svcObject -MethodName StartService -ErrorAction Stop }
					'Stop'    { $null = Invoke-CimMethod -InputObject $svcObject -MethodName StopService -ErrorAction Stop }
					'Restart' { $null = Invoke-CimMethod -InputObject $svcObject -MethodName StopService -ErrorAction Stop; Start-Sleep -Seconds 2; $svcObject = Get-CimInstance -ClassName Win32_Service -ComputerName $Machine -Filter "Name='$serviceEscaped'" -ErrorAction Stop; $null = Invoke-CimMethod -InputObject $svcObject -MethodName StartService -ErrorAction Stop }
				}
				$actionSummary = "$RequestedAction $ServiceKind on lane $Lane."
			}
			catch
			{
				try
				{
					switch ($RequestedAction)
					{
						'Start'   { $null = & sc.exe "\\$Machine" start $ServiceName 2>$null }
						'Stop'    { $null = & sc.exe "\\$Machine" stop $ServiceName 2>$null }
						'Restart' { $null = & sc.exe "\\$Machine" stop $ServiceName 2>$null; Start-Sleep -Seconds 2; $null = & sc.exe "\\$Machine" start $ServiceName 2>$null }
					}
					$actionSummary = "$RequestedAction $ServiceKind on lane $Lane."
				}
				catch
				{
					$actionSummary = "Failed to $RequestedAction $ServiceKind on lane ${Lane}: $($_.Exception.Message)"
				}
			}

			$targetStatus = if ($RequestedAction -eq 'Stop') { 'Stopped' }
			else { 'Running' }
			for ($i = 0; $i -lt 18; $i++)
			{
				Start-Sleep -Milliseconds 750
				$snapshot = & $getServiceSnapshot $Lane $Machine $ExpectedSqlService
				$currentStatus = if ($ServiceKind -eq 'SQL') { [string]$snapshot.SqlStatus }
				else { [string]$snapshot.VncStatus }
				if ($currentStatus -eq $targetStatus -or $currentStatus -like 'Error:*' -or $currentStatus -eq 'Not Found')
				{
					$snapshot.ActionSummary = $actionSummary
					return $snapshot
				}
			}

			$snapshot = & $getServiceSnapshot $Lane $Machine $ExpectedSqlService
			$snapshot.ActionSummary = $actionSummary
			return $snapshot
		}

		$svcActionJobs[$laneKey] = @{
			Job		    = $job
			Row		    = $row.Index
			Action	    = $RequestedAction
			ServiceKind = $ServiceKind
		}
	}

	$refreshSvcLanes = {
		param (
			[string[]]$LaneKeysToRefresh
		)

		foreach ($laneKey in @($LaneKeysToRefresh))
		{
			if (-not $laneMap.ContainsKey($laneKey)) { continue }

			if ($svcJobs.ContainsKey($laneKey))
			{
				try { Stop-Job -Job $svcJobs[$laneKey].Job -ErrorAction SilentlyContinue }
				catch { }
				try { Remove-Job -Job $svcJobs[$laneKey].Job -Force -ErrorAction SilentlyContinue }
				catch { }
				[void]$svcJobs.Remove($laneKey)
			}

			$machineName = [string]$laneMap[$laneKey]
			$expectedSqlService = 'MSSQLSERVER'
			try
			{
				$dbInfo = Get_All_Lanes_Database_Info -LaneNumber $laneKey
				if ($dbInfo -and ($dbInfo -isnot [hashtable]))
				{
					$tmpDbInfo = @{ }
					foreach ($p in $dbInfo.PSObject.Properties) { $tmpDbInfo[$p.Name] = $p.Value }
					$dbInfo = $tmpDbInfo
				}
				if ($dbInfo -and $dbInfo.ContainsKey('InstanceName'))
				{
					$instanceName = [string]$dbInfo['InstanceName']
					if (-not [string]::IsNullOrWhiteSpace($instanceName) -and $instanceName.ToUpper() -ne 'MSSQLSERVER')
					{
						$expectedSqlService = "MSSQL`$$instanceName"
					}
				}
			}
			catch { }

			$rowIndex = [int]$rowMap[$laneKey]
			if ($rowIndex -lt 0 -or $rowIndex -ge $gridSvc.Rows.Count) { continue }
			$gridSvc.Rows[$rowIndex].Cells["SqlService"].Value = $expectedSqlService
			$gridSvc.Rows[$rowIndex].Cells["SqlStatus"].Value = 'Checking...'
			$gridSvc.Rows[$rowIndex].Cells["VncService"].Value = 'Checking...'
			$gridSvc.Rows[$rowIndex].Cells["VncStatus"].Value = 'Checking...'

			$job = Start-Job -Name ("LaneServiceStatus_{0}" -f $machineName) -ArgumentList $laneKey, $machineName, $expectedSqlService -ScriptBlock {
				param (
					[string]$Lane,
					[string]$Machine,
					[string]$ExpectedSqlService
				)

				$result = [ordered]@{
					Lane	   = $Lane
					Machine    = $Machine
					SqlService = if ([string]::IsNullOrWhiteSpace($ExpectedSqlService)) { 'MSSQLSERVER' } else { $ExpectedSqlService }
					SqlStatus  = 'Not Found'
					VncService = 'Not Found'
					VncStatus  = 'Not Found'
				}

				$reachable = $false
				try
				{
					$ping = New-Object System.Net.NetworkInformation.Ping
					$reply = $ping.Send($Machine, 700)
					if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) { $reachable = $true }
				}
				catch { }

				if (-not $reachable)
				{
					$tcp = $null
					$async = $null
					try
					{
						$tcp = New-Object System.Net.Sockets.TcpClient
						$async = $tcp.BeginConnect($Machine, 445, $null, $null)
						if ($async.AsyncWaitHandle.WaitOne(700, $false) -and $tcp.Connected)
						{
							try { $tcp.EndConnect($async) }
							catch { }
							$reachable = $true
						}
					}
					catch { }
					finally
					{
						try { if ($async) { $async.AsyncWaitHandle.Close() } }
						catch { }
						try { if ($tcp) { $tcp.Close() } }
						catch { }
					}
				}

				if (-not $reachable)
				{
					$result.SqlStatus = 'Offline'
					$result.VncStatus = 'Offline'
					$result.VncService = 'Offline'
					return [pscustomobject]$result
				}

				$services = $null
				$serviceFetchError = $null
				try
				{
					$services = Get-Service -ComputerName $Machine -ErrorAction Stop | Select-Object Name, DisplayName, Status
				}
				catch
				{
					$serviceFetchError = $_.Exception.Message
				}

				if (-not $services)
				{
					try
					{
						$services = Get-CimInstance -ClassName Win32_Service -ComputerName $Machine -ErrorAction Stop | Select-Object Name, DisplayName, @{ Name = 'Status'; Expression = { $_.State } }
						$serviceFetchError = $null
					}
					catch
					{
						$serviceFetchError = $_.Exception.Message
					}
				}

				if (-not $services)
				{
					$result.SqlStatus = if ([string]::IsNullOrWhiteSpace($serviceFetchError)) { 'Error' }
					else { "Error: $serviceFetchError" }
					$result.VncStatus = if ([string]::IsNullOrWhiteSpace($serviceFetchError)) { 'Error' }
					else { "Error: $serviceFetchError" }
					$result.VncService = 'Unavailable'
					return [pscustomobject]$result
				}

				$sqlService = $services | Where-Object { $_.Name -ieq $ExpectedSqlService } | Select-Object -First 1
				if (-not $sqlService)
				{
					$sqlService = $services | Where-Object { $_.Name -like 'MSSQL*' -or $_.DisplayName -match '(?i)^SQL Server' } | Sort-Object Name | Select-Object -First 1
				}
				if ($sqlService)
				{
					$result.SqlService = [string]$sqlService.Name
					$result.SqlStatus = [string]$sqlService.Status
				}

				$vncService = $services | Where-Object {
					$svcName = [string]$_.Name
					$svcDisplay = [string]$_.DisplayName
					($svcName -match '(?i)^(uvnc_service|winvnc.*|tvnserver|vncserver|tigervnc.*|tightvnc.*)$') -or
					($svcDisplay -match '(?i)ultravnc|tightvnc|tigervnc|vnc')
				} | Sort-Object Name | Select-Object -First 1
				if ($vncService)
				{
					$result.VncService = [string]$vncService.Name
					$result.VncStatus = [string]$vncService.Status
				}

				return [pscustomobject]$result
			}

			$svcJobs[$laneKey] = @{
				Job	    = $job
				Row	    = $rowIndex
				Machine = $machineName
			}
		}
	}

	$svcRowMenu = New-Object System.Windows.Forms.ContextMenuStrip
	if ($Apply_Modern_Menu_Style) { & $Apply_Modern_Menu_Style $svcRowMenu }
	$miSvcRefreshLane = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh This Lane")
	$miSvcRefreshLane.Add_Click({
			$script:LastActivity = Get-Date
			if ($gridSvc.SelectedRows -and $gridSvc.SelectedRows.Count -gt 0)
			{
				$laneKey = [string]$gridSvc.SelectedRows[0].Cells["Lane"].Value
				if (-not [string]::IsNullOrWhiteSpace($laneKey))
				{
					& $refreshSvcLanes @($laneKey)
					$lblSvcStatus.Text = ("Refreshing lane {0}..." -f $laneKey)
				}
			}
		})
	[void]$svcRowMenu.Items.Add($miSvcRefreshLane)
	[void]$svcRowMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
	$miSvcStartSql = New-Object System.Windows.Forms.ToolStripMenuItem("Start SQL")
	$miSvcStartSql.Add_Click({ $script:LastActivity = Get-Date; & $startActionJobForSelectedRow 'SQL' 'Start' })
	[void]$svcRowMenu.Items.Add($miSvcStartSql)
	$miSvcStopSql = New-Object System.Windows.Forms.ToolStripMenuItem("Stop SQL")
	$miSvcStopSql.Add_Click({ $script:LastActivity = Get-Date; & $startActionJobForSelectedRow 'SQL' 'Stop' })
	[void]$svcRowMenu.Items.Add($miSvcStopSql)
	$miSvcRestartSql = New-Object System.Windows.Forms.ToolStripMenuItem("Restart SQL")
	$miSvcRestartSql.Add_Click({ $script:LastActivity = Get-Date; & $startActionJobForSelectedRow 'SQL' 'Restart' })
	[void]$svcRowMenu.Items.Add($miSvcRestartSql)
	[void]$svcRowMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
	$miSvcStartVnc = New-Object System.Windows.Forms.ToolStripMenuItem("Start VNC")
	$miSvcStartVnc.Add_Click({ $script:LastActivity = Get-Date; & $startActionJobForSelectedRow 'VNC' 'Start' })
	[void]$svcRowMenu.Items.Add($miSvcStartVnc)
	$miSvcStopVnc = New-Object System.Windows.Forms.ToolStripMenuItem("Stop VNC")
	$miSvcStopVnc.Add_Click({ $script:LastActivity = Get-Date; & $startActionJobForSelectedRow 'VNC' 'Stop' })
	[void]$svcRowMenu.Items.Add($miSvcStopVnc)
	$miSvcRestartVnc = New-Object System.Windows.Forms.ToolStripMenuItem("Restart VNC")
	$miSvcRestartVnc.Add_Click({ $script:LastActivity = Get-Date; & $startActionJobForSelectedRow 'VNC' 'Restart' })
	[void]$svcRowMenu.Items.Add($miSvcRestartVnc)
	$gridSvc.ContextMenuStrip = $svcRowMenu
	$gridSvc.Add_MouseDown({
			param ($sender,
				$e)
			if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
			$hit = $gridSvc.HitTest($e.X, $e.Y)
			if ($hit.RowIndex -lt 0) { return }
			try
			{
				$gridSvc.ClearSelection()
				$gridSvc.Rows[$hit.RowIndex].Selected = $true
				$gridSvc.CurrentCell = $gridSvc.Rows[$hit.RowIndex].Cells[0]
			}
			catch { }
		})

	foreach ($lane in $selectedLanes)
	{
		$laneKey = [string]$lane
		$machineName = [string]$laneMap[$laneKey]
		$expectedSqlService = 'MSSQLSERVER'

		try
		{
			$dbInfo = Get_All_Lanes_Database_Info -LaneNumber $laneKey
			if ($dbInfo -and ($dbInfo -isnot [hashtable]))
			{
				$tmpDbInfo = @{ }
				foreach ($p in $dbInfo.PSObject.Properties) { $tmpDbInfo[$p.Name] = $p.Value }
				$dbInfo = $tmpDbInfo
			}

			if ($dbInfo -and $dbInfo.ContainsKey('InstanceName'))
			{
				$instanceName = [string]$dbInfo['InstanceName']
				if (-not [string]::IsNullOrWhiteSpace($instanceName) -and $instanceName.ToUpper() -ne 'MSSQLSERVER')
				{
					$expectedSqlService = "MSSQL`$$instanceName"
				}
			}
		}
		catch { }

		$rowIndex = $gridSvc.Rows.Add($laneKey, $machineName, $expectedSqlService, 'Checking...', 'Checking...', 'Checking...')
		$rowMap[$laneKey] = $rowIndex

		$job = Start-Job -Name ("LaneServiceStatus_{0}" -f $machineName) -ArgumentList $laneKey, $machineName, $expectedSqlService -ScriptBlock {
			param (
				[string]$Lane,
				[string]$Machine,
				[string]$ExpectedSqlService
			)

			$result = [ordered]@{
				Lane	   = $Lane
				Machine    = $Machine
				SqlService = if ([string]::IsNullOrWhiteSpace($ExpectedSqlService)) { 'MSSQLSERVER' } else { $ExpectedSqlService }
				SqlStatus  = 'Not Found'
				VncService = 'Not Found'
				VncStatus  = 'Not Found'
			}

			$reachable = $false
			try
			{
				$ping = New-Object System.Net.NetworkInformation.Ping
				$reply = $ping.Send($Machine, 700)
				if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) { $reachable = $true }
			}
			catch { }

			if (-not $reachable)
			{
				$tcp = $null
				$async = $null
				try
				{
					$tcp = New-Object System.Net.Sockets.TcpClient
					$async = $tcp.BeginConnect($Machine, 445, $null, $null)
					if ($async.AsyncWaitHandle.WaitOne(700, $false) -and $tcp.Connected)
					{
						try { $tcp.EndConnect($async) }
						catch { }
						$reachable = $true
					}
				}
				catch { }
				finally
				{
					try { if ($async) { $async.AsyncWaitHandle.Close() } }
					catch { }
					try { if ($tcp) { $tcp.Close() } }
					catch { }
				}
			}

			if (-not $reachable)
			{
				$result.SqlStatus = 'Offline'
				$result.VncStatus = 'Offline'
				$result.VncService = 'Offline'
				return [pscustomobject]$result
			}

			$services = $null
			$serviceFetchError = $null

			try
			{
				$services = Get-Service -ComputerName $Machine -ErrorAction Stop | Select-Object Name, DisplayName, Status
			}
			catch
			{
				$serviceFetchError = $_.Exception.Message
			}

			if (-not $services)
			{
				try
				{
					$services = Get-CimInstance -ClassName Win32_Service -ComputerName $Machine -ErrorAction Stop | Select-Object Name, DisplayName, @{ Name = 'Status'; Expression = { $_.State } }
					$serviceFetchError = $null
				}
				catch
				{
					$serviceFetchError = $_.Exception.Message
				}
			}

			if (-not $services)
			{
				$result.SqlStatus = 'Error'
				$result.VncStatus = 'Error'
				$result.SqlService = if ($ExpectedSqlService) { $ExpectedSqlService }
				else { 'Unavailable' }
				$result.VncService = 'Unavailable'
				if (-not [string]::IsNullOrWhiteSpace($serviceFetchError))
				{
					$result.SqlStatus = "Error: $serviceFetchError"
					$result.VncStatus = "Error: $serviceFetchError"
				}
				return [pscustomobject]$result
			}

			$sqlService = $services | Where-Object { $_.Name -ieq $ExpectedSqlService } | Select-Object -First 1
			if (-not $sqlService)
			{
				$sqlService = $services | Where-Object {
					$_.Name -like 'MSSQL*' -or
					$_.DisplayName -match '(?i)^SQL Server'
				} | Sort-Object Name | Select-Object -First 1
			}
			if ($sqlService)
			{
				$result.SqlService = [string]$sqlService.Name
				$result.SqlStatus = [string]$sqlService.Status
			}

			$vncService = $services | Where-Object {
				$svcName = [string]$_.Name
				$svcDisplay = [string]$_.DisplayName
				($svcName -match '(?i)^(uvnc_service|winvnc.*|tvnserver|vncserver|tigervnc.*|tightvnc.*)$') -or
				($svcDisplay -match '(?i)ultravnc|tightvnc|tigervnc|vnc')
			} | Sort-Object Name | Select-Object -First 1
			if ($vncService)
			{
				$result.VncService = [string]$vncService.Name
				$result.VncStatus = [string]$vncService.Status
			}

			return [pscustomobject]$result
		}

		$svcJobs[$laneKey] = @{
			Job	    = $job
			Row	    = $rowIndex
			Machine = $machineName
		}
	}

	$btnSvcRefresh.Add_Click({
			$script:LastActivity = Get-Date
			& $refreshSvcLanes @($selectedLanes | ForEach-Object { [string]$_ })
			$lblSvcStatus.Text = "Refreshing all lanes..."
		})

	$svcTimer = New-Object System.Windows.Forms.Timer
	$svcTimer.Interval = 700
	$svcTimer.Add_Tick({
			foreach ($laneKey in @($svcJobs.Keys))
			{
				$entry = $svcJobs[$laneKey]
				if (-not $entry) { continue }

				$job = $entry.Job
				if (-not $job)
				{
					$doneCount++
					continue
				}

				if ($job.State -in @('Completed', 'Failed', 'Stopped'))
				{
					$result = $null
					try { $result = Receive-Job -Job $job -ErrorAction SilentlyContinue }
					catch { }

					$rowIndex = [int]$entry.Row
					if ($rowIndex -ge 0 -and $rowIndex -lt $gridSvc.Rows.Count)
					{
						$row = $gridSvc.Rows[$rowIndex]
						if ($result)
						{
							$row.Cells["SqlService"].Value = [string]$result.SqlService
							$row.Cells["SqlStatus"].Value = [string]$result.SqlStatus
							$row.Cells["VncService"].Value = [string]$result.VncService
							$row.Cells["VncStatus"].Value = [string]$result.VncStatus
							$row.Tag = @{
								Lane = [string]$result.Lane
								Machine = [string]$result.Machine
								SqlService = [string]$result.SqlService
								SqlStatus = [string]$result.SqlStatus
								VncService = [string]$result.VncService
								VncStatus = [string]$result.VncStatus
							}
						}
						else
						{
							$row.Cells["SqlStatus"].Value = 'Error'
							$row.Cells["VncStatus"].Value = 'Error'
						}
					}

					try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
					catch { }
					[void]$svcJobs.Remove($laneKey)
				}
			}

			foreach ($laneKey in @($svcActionJobs.Keys))
			{
				$entry = $svcActionJobs[$laneKey]
				if (-not $entry -or -not $entry.Job) { continue }

				if ($entry.Job.State -in @('Completed', 'Failed', 'Stopped'))
				{
					$result = $null
					try { $result = Receive-Job -Job $entry.Job -ErrorAction SilentlyContinue }
					catch { }

					$rowIndex = [int]$entry.Row
					if ($rowIndex -ge 0 -and $rowIndex -lt $gridSvc.Rows.Count -and $result)
					{
						$row = $gridSvc.Rows[$rowIndex]
						$row.Cells["SqlService"].Value = [string]$result.SqlService
						$row.Cells["SqlStatus"].Value = [string]$result.SqlStatus
						$row.Cells["VncService"].Value = [string]$result.VncService
						$row.Cells["VncStatus"].Value = [string]$result.VncStatus
						$row.Tag = @{
							Lane = [string]$result.Lane
							Machine = [string]$result.Machine
							SqlService = [string]$result.SqlService
							SqlStatus = [string]$result.SqlStatus
							VncService = [string]$result.VncService
							VncStatus = [string]$result.VncStatus
						}
					}

					if ($result -and -not [string]::IsNullOrWhiteSpace([string]$result.ActionSummary))
					{
						$lblSvcStatus.Text = [string]$result.ActionSummary
					}

					try { Remove-Job -Job $entry.Job -Force -ErrorAction SilentlyContinue }
					catch { }
					[void]$svcActionJobs.Remove($laneKey)
				}
			}

			$pendingStatusCount = $svcJobs.Count
			$pendingActionCount = $svcActionJobs.Count
			if ($pendingActionCount -gt 0)
			{
				$lblSvcStatus.Text = ("Working... {0} action(s), {1} refresh check(s) pending" -f $pendingActionCount, $pendingStatusCount)
			}
			elseif ($pendingStatusCount -gt 0)
			{
				$doneCount = $totalRows - $pendingStatusCount
				$lblSvcStatus.Text = ("Checking all lanes... {0}/{1} complete" -f $doneCount, $totalRows)
			}
			elseif ($lblSvcStatus.Text -match '^(Checking|Refreshing|Working)')
			{
				$lblSvcStatus.Text = ("Ready. Checked {0} lane(s)." -f $totalRows)
			}
		})
	$svcTimer.Start()

	$formSvc.Add_FormClosing({
			try
			{
				if ($svcTimer)
				{
					$svcTimer.Stop()
					$svcTimer.Dispose()
				}
			}
			catch { }

			foreach ($laneKey in @($svcJobs.Keys))
			{
				$entry = $svcJobs[$laneKey]
				if (-not $entry -or -not $entry.Job) { continue }
				try { Stop-Job -Job $entry.Job -ErrorAction SilentlyContinue }
				catch { }
				try { Remove-Job -Job $entry.Job -Force -ErrorAction SilentlyContinue }
				catch { }
			}
			$svcJobs.Clear()

			foreach ($laneKey in @($svcActionJobs.Keys))
			{
				$entry = $svcActionJobs[$laneKey]
				if (-not $entry -or -not $entry.Job) { continue }
				try { Stop-Job -Job $entry.Job -ErrorAction SilentlyContinue }
				catch { }
				try { Remove-Job -Job $entry.Job -Force -ErrorAction SilentlyContinue }
				catch { }
			}
			$svcActionJobs.Clear()
		})

	[void]$formSvc.ShowDialog()
}
