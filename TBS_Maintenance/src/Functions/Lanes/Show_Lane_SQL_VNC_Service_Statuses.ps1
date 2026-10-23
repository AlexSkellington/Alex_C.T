# ===================================================================================================
#                              FUNCTION: Show_Lane_SQL_VNC_Service_Statuses
# ---------------------------------------------------------------------------------------------------
# Description:
#   Opens a styled popup that checks the SQL and VNC services on all lanes of the store in
#   the background (runspace pool, up to 8 lanes at a time) and fills a grid as results
#   arrive. Right-click a lane to refresh it or start / stop / restart its SQL or VNC
#   service. Shows a message box when no usable lanes are found.
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

	# ---- Background work: one runspace pool (Start-Job spawned a powershell.exe per lane) ----------
	$svcPool = [RunspaceFactory]::CreateRunspacePool(1, [Math]::Min(8, [Math]::Max(2, $selectedLanes.Count)))
	$svcPool.Open()
	$svcJobs = @{ }
	$svcActionJobs = @{ }
	$rowMap = @{ }
	$totalRows = $selectedLanes.Count

	# Lane status: reachability (ping, then SMB 445), then the SQL and VNC services.
	# Text, not a scriptblock: the action script rebuilds it to re-read the status.
	$statusCheckText = @'
param ([string]$Lane, [string]$Machine, [string]$ExpectedSqlService)

$result = [pscustomobject][ordered]@{
	Lane          = $Lane
	Machine       = $Machine
	SqlService    = if ([string]::IsNullOrWhiteSpace($ExpectedSqlService)) { 'MSSQLSERVER' } else { $ExpectedSqlService }
	SqlStatus     = 'Not Found'
	VncService    = 'Not Found'
	VncStatus     = 'Not Found'
	ActionSummary = $null
}

$reachable = $false
try { $reachable = ((New-Object System.Net.NetworkInformation.Ping).Send($Machine, 700).Status -eq 'Success') }
catch { }
if (-not $reachable)
{
	$tcp = New-Object System.Net.Sockets.TcpClient
	try
	{
		$async = $tcp.BeginConnect($Machine, 445, $null, $null)
		if ($async.AsyncWaitHandle.WaitOne(700, $false) -and $tcp.Connected) { $reachable = $true }
	}
	catch { }
	finally { $tcp.Close() }
}
if (-not $reachable)
{
	$result.SqlStatus = 'Offline'
	$result.VncService = 'Offline'
	$result.VncStatus = 'Offline'
	return $result
}

$services = $null
$serviceFetchError = $null
try { $services = Get-Service -ComputerName $Machine -ErrorAction Stop | Select-Object Name, DisplayName, Status }
catch { $serviceFetchError = $_.Exception.Message }
if (-not $services)
{
	try
	{
		$services = Get-CimInstance -ClassName Win32_Service -ComputerName $Machine -ErrorAction Stop |
		Select-Object Name, DisplayName, @{ Name = 'Status'; Expression = { $_.State } }
		$serviceFetchError = $null
	}
	catch { $serviceFetchError = $_.Exception.Message }
}
if (-not $services)
{
	$errorText = if ([string]::IsNullOrWhiteSpace($serviceFetchError)) { 'Error' } else { "Error: $serviceFetchError" }
	$result.SqlStatus = $errorText
	$result.VncStatus = $errorText
	$result.VncService = 'Unavailable'
	return $result
}

$sqlService = $services | Where-Object { $_.Name -ieq $result.SqlService } | Select-Object -First 1
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
	([string]$_.Name -match '(?i)^(uvnc_service|winvnc.*|tvnserver|vncserver|tigervnc.*|tightvnc.*)$') -or
	([string]$_.DisplayName -match '(?i)ultravnc|tightvnc|tigervnc|vnc')
} | Sort-Object Name | Select-Object -First 1
if ($vncService)
{
	$result.VncService = [string]$vncService.Name
	$result.VncStatus = [string]$vncService.Status
}

$result
'@

	# Start / Stop / Restart one service, then poll the status until it settles (~14 s max)
	$serviceActionText = @'
param ([string]$Lane, [string]$Machine, [string]$ExpectedSqlService, [string]$ServiceKind, [string]$ServiceName, [string]$RequestedAction, [string]$StatusCheckText)

$getSnapshot = [scriptblock]::Create($StatusCheckText)
$actionSummary = "$RequestedAction $ServiceKind on lane $Lane."
try
{
	$filter = "Name='{0}'" -f $ServiceName.Replace("'", "''")
	if ($RequestedAction -in @('Stop', 'Restart'))
	{
		$svcObject = Get-CimInstance -ClassName Win32_Service -ComputerName $Machine -Filter $filter -ErrorAction Stop
		$null = Invoke-CimMethod -InputObject $svcObject -MethodName StopService -ErrorAction Stop
	}
	if ($RequestedAction -eq 'Restart') { Start-Sleep -Seconds 2 }
	if ($RequestedAction -in @('Start', 'Restart'))
	{
		$svcObject = Get-CimInstance -ClassName Win32_Service -ComputerName $Machine -Filter $filter -ErrorAction Stop
		$null = Invoke-CimMethod -InputObject $svcObject -MethodName StartService -ErrorAction Stop
	}
}
catch
{
	# CIM blocked (firewall / WinRM): the Service Control Manager over SMB usually still works
	try
	{
		if ($RequestedAction -in @('Stop', 'Restart')) { $null = & sc.exe "\\$Machine" stop $ServiceName 2>$null }
		if ($RequestedAction -eq 'Restart') { Start-Sleep -Seconds 2 }
		if ($RequestedAction -in @('Start', 'Restart')) { $null = & sc.exe "\\$Machine" start $ServiceName 2>$null }
		# sc.exe never throws; 1056 = already running, 1062 = not started (already stopped)
		if ($LASTEXITCODE -notin @(0, 1056, 1062)) { $actionSummary = "Failed to $RequestedAction $ServiceKind on lane $Lane (sc.exe error $LASTEXITCODE)." }
	}
	catch { $actionSummary = "Failed to $RequestedAction $ServiceKind on lane ${Lane}: $($_.Exception.Message)" }
}

$targetStatus = if ($RequestedAction -eq 'Stop') { 'Stopped' } else { 'Running' }
for ($i = 0; $i -lt 18; $i++)
{
	Start-Sleep -Milliseconds 750
	$snapshot = & $getSnapshot $Lane $Machine $ExpectedSqlService
	$currentStatus = if ($ServiceKind -eq 'SQL') { [string]$snapshot.SqlStatus } else { [string]$snapshot.VncStatus }
	if ($currentStatus -eq $targetStatus -or $currentStatus -like 'Error*' -or $currentStatus -in @('Not Found', 'Offline')) { break }
}
$snapshot.ActionSummary = $actionSummary
$snapshot
'@

	$startPoolWork = {
		param ([string]$ScriptText, [object[]]$Arguments)
		$ps = [PowerShell]::Create()
		$ps.RunspacePool = $svcPool
		[void]$ps.AddScript($ScriptText)
		foreach ($a in $Arguments) { [void]$ps.AddArgument($a) }
		@{ PS = $ps; Handle = $ps.BeginInvoke() }
	}

	# Finished work item -> its result object (or $null), and the PowerShell instance disposed
	$collectPoolWork = {
		param ($Entry)
		$res = $null
		try { $res = @($Entry.PS.EndInvoke($Entry.Handle)) | Select-Object -Last 1 }
		catch { }
		try { $Entry.PS.Dispose() }
		catch { }
		$res
	}

	$stopPoolWork = {
		param ($Entry)
		try { [void]$Entry.PS.BeginStop($null, $null) }
		catch { }
	}

	# Service name the lane's SQL instance should have (MSSQLSERVER or MSSQL$<instance>)
	$getExpectedSqlService = {
		param ([string]$LaneKey)
		try
		{
			$dbInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneKey
			$instanceName = if ($dbInfo) { [string]$dbInfo['InstanceName'] }
			if (-not [string]::IsNullOrWhiteSpace($instanceName) -and $instanceName.ToUpper() -ne 'MSSQLSERVER') { return "MSSQL`$$instanceName" }
		}
		catch { }
		'MSSQLSERVER'
	}

	$applyResultToRow = {
		param ([int]$RowIndex, $Result)
		if ($RowIndex -lt 0 -or $RowIndex -ge $gridSvc.Rows.Count) { return }
		$row = $gridSvc.Rows[$RowIndex]
		if (-not $Result)
		{
			$row.Cells["SqlStatus"].Value = 'Error'
			$row.Cells["VncStatus"].Value = 'Error'
			return
		}
		$row.Cells["SqlService"].Value = [string]$Result.SqlService
		$row.Cells["SqlStatus"].Value = [string]$Result.SqlStatus
		$row.Cells["VncService"].Value = [string]$Result.VncService
		$row.Cells["VncStatus"].Value = [string]$Result.VncStatus
	}

	$startActionJobForSelectedRow = {
		param (
			[string]$ServiceKind,
			[string]$RequestedAction
		)

		if (-not $gridSvc.SelectedRows -or $gridSvc.SelectedRows.Count -eq 0)
		{
			[void][System.Windows.Forms.MessageBox]::Show("Select a lane first.", "Lane Services", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			return
		}

		$row = $gridSvc.SelectedRows[0]
		$laneKey = [string]$row.Cells["Lane"].Value
		$machineName = [string]$row.Cells["Machine"].Value
		$expectedSqlService = [string]$row.Cells["SqlService"].Value
		$serviceName = if ($ServiceKind -eq 'SQL') { $expectedSqlService } else { [string]$row.Cells["VncService"].Value }
		$statusColumn = if ($ServiceKind -eq 'SQL') { 'SqlStatus' } else { 'VncStatus' }

		if ([string]::IsNullOrWhiteSpace($laneKey) -or [string]::IsNullOrWhiteSpace($machineName)) { return }
		if ($svcActionJobs.ContainsKey($laneKey))
		{
			[void][System.Windows.Forms.MessageBox]::Show("A service action is already running for lane $laneKey.", "Lane Services", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			return
		}
		if ([string]::IsNullOrWhiteSpace($serviceName) -or $serviceName -in @('Checking...', 'Not Found', 'Unavailable', 'Offline'))
		{
			[void][System.Windows.Forms.MessageBox]::Show("$ServiceKind service was not found for lane $laneKey.", "Lane Services", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			return
		}

		$row.Cells[$statusColumn].Value = "$RequestedAction..."
		$lblSvcStatus.Text = ("{0} {1} on lane {2}..." -f $RequestedAction, $ServiceKind, $laneKey)

		$work = & $startPoolWork $serviceActionText @($laneKey, $machineName, $expectedSqlService, $ServiceKind, $serviceName, $RequestedAction, $statusCheckText)
		$work.Row = $row.Index
		$svcActionJobs[$laneKey] = $work
	}

	$refreshSvcLanes = {
		param (
			[string[]]$LaneKeysToRefresh
		)

		foreach ($laneKey in @($LaneKeysToRefresh))
		{
			if (-not $laneMap.ContainsKey($laneKey) -or -not $rowMap.ContainsKey($laneKey)) { continue }

			# A refresh replaces a check still in flight for the lane
			if ($svcJobs.ContainsKey($laneKey))
			{
				& $stopPoolWork $svcJobs[$laneKey]
				[void]$svcJobs.Remove($laneKey)
			}

			$machineName = [string]$laneMap[$laneKey]
			$expectedSqlService = & $getExpectedSqlService $laneKey
			$rowIndex = [int]$rowMap[$laneKey]
			$gridSvc.Rows[$rowIndex].Cells["SqlService"].Value = $expectedSqlService
			$gridSvc.Rows[$rowIndex].Cells["SqlStatus"].Value = 'Checking...'
			$gridSvc.Rows[$rowIndex].Cells["VncService"].Value = 'Checking...'
			$gridSvc.Rows[$rowIndex].Cells["VncStatus"].Value = 'Checking...'

			$work = & $startPoolWork $statusCheckText @($laneKey, $machineName, $expectedSqlService)
			$work.Row = $rowIndex
			$svcJobs[$laneKey] = $work
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

	# Start / Stop / Restart for SQL and VNC; the menu item's Tag carries kind + action
	foreach ($kind in @('SQL', 'VNC'))
	{
		[void]$svcRowMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
		foreach ($action in @('Start', 'Stop', 'Restart'))
		{
			$mi = New-Object System.Windows.Forms.ToolStripMenuItem("$action $kind")
			$mi.Tag = @($kind, $action)
			$mi.Add_Click({
					param ($sender, $e)
					$script:LastActivity = Get-Date
					& $startActionJobForSelectedRow $sender.Tag[0] $sender.Tag[1]
				})
			[void]$svcRowMenu.Items.Add($mi)
		}
	}
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

	# One row per lane, then the first check for all of them
	foreach ($lane in $selectedLanes)
	{
		$laneKey = [string]$lane
		$rowMap[$laneKey] = $gridSvc.Rows.Add($laneKey, [string]$laneMap[$laneKey], 'MSSQLSERVER', 'Checking...', 'Checking...', 'Checking...')
	}
	& $refreshSvcLanes @($selectedLanes | ForEach-Object { [string]$_ })

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
				if (-not $entry.Handle.IsCompleted) { continue }
				& $applyResultToRow ([int]$entry.Row) (& $collectPoolWork $entry)
				[void]$svcJobs.Remove($laneKey)
			}

			foreach ($laneKey in @($svcActionJobs.Keys))
			{
				$entry = $svcActionJobs[$laneKey]
				if (-not $entry.Handle.IsCompleted) { continue }
				$result = & $collectPoolWork $entry
				if ($result) { & $applyResultToRow ([int]$entry.Row) $result }
				if ($result -and -not [string]::IsNullOrWhiteSpace([string]$result.ActionSummary)) { $lblSvcStatus.Text = [string]$result.ActionSummary }
				[void]$svcActionJobs.Remove($laneKey)
			}

			$pendingStatusCount = $svcJobs.Count
			$pendingActionCount = $svcActionJobs.Count
			if ($pendingActionCount -gt 0)
			{
				$lblSvcStatus.Text = ("Working... {0} action(s), {1} refresh check(s) pending" -f $pendingActionCount, $pendingStatusCount)
			}
			elseif ($pendingStatusCount -gt 0)
			{
				$lblSvcStatus.Text = ("Checking all lanes... {0}/{1} complete" -f ($totalRows - $pendingStatusCount), $totalRows)
			}
			elseif ($lblSvcStatus.Text -match '^(Checking|Refreshing|Working)')
			{
				$lblSvcStatus.Text = ("Ready. Checked {0} lane(s)." -f $totalRows)
			}
		})
	$svcTimer.Start()

	$formSvc.Add_FormClosing({
			try { $svcTimer.Stop(); $svcTimer.Dispose() }
			catch { }
			# Ask running checks to stop and close the pool without waiting on slow remote calls
			foreach ($entry in @($svcJobs.Values) + @($svcActionJobs.Values)) { & $stopPoolWork $entry }
			$svcJobs.Clear()
			$svcActionJobs.Clear()
			try { [void]$svcPool.BeginClose($null, $null) }
			catch { }
		})

	[void]$formSvc.ShowDialog()
	$formSvc.Dispose()
}
