# ------------------------------------------------------------------------------------------------
# ========================= Protocol Table Popup (kept; slightly modern styling) ==================
# ------------------------------------------------------------------------------------------------
$rowHeight = 19
$rowCount = 25
$gridHeight = ($rowCount * $rowHeight) + 28

if (-not $script:LaneDbSizeMap) { $script:LaneDbSizeMap = @{ } }
if (-not $script:LaneDbSizeJobs) { $script:LaneDbSizeJobs = @{ } }
if (-not $script:LaneDbSizeRefreshSeconds -or $script:LaneDbSizeRefreshSeconds -lt 5) { $script:LaneDbSizeRefreshSeconds = 5 }
if (-not $script:LaneDbSizePool)
{
	$minThreads = 1
	$maxThreads = [Math]::Max(8, [Environment]::ProcessorCount * 2)
	$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
	$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool($minThreads, $maxThreads, $iss, $Host)
	try { $pool.ApartmentState = 'MTA' }
	catch { }
	$pool.Open()
	$script:LaneDbSizePool = $pool
}
if (-not $script:LaneDbSizeWorker)
{
	$script:LaneDbSizeWorker = @'
param(
	[string]$lane,
	[string]$protocol,
	[string]$connectionString,
	[string]$machineName,
	[string]$dbName,
	[string]$smsStartIniPath
)

Add-Type -AssemblyName System.Data 2>$null
Add-Type -AssemblyName System 2>$null

$result = [ordered]@{
	Lane      = $lane
	DataText  = 'N/A'
	LogText   = 'N/A'
	Status    = 'Unavailable'
	OpenPath  = $null
	UpdatedAt = [DateTime]::Now
}

$machineNameNorm = [string]$machineName
$isLocalMachine = $false
try
{
	if ([string]::IsNullOrWhiteSpace($machineNameNorm) -or $machineNameNorm -match '^(?i)(localhost|\.|\(local\))$' -or $machineNameNorm -ieq $env:COMPUTERNAME)
	{
		$isLocalMachine = $true
	}
}
catch { }

if (-not $isLocalMachine)
{
	$reachable = $false
	try
	{
		$ping = New-Object System.Net.NetworkInformation.Ping
		$reply = $ping.Send($machineNameNorm, 700)
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
			$async = $tcp.BeginConnect($machineNameNorm, 445, $null, $null)
			if ($async.AsyncWaitHandle.WaitOne(700, $false) -and $tcp.Connected)
			{
				try { $tcp.EndConnect($async) } catch { }
				$reachable = $true
			}
		}
		catch { }
		finally
		{
			try { if ($async) { $async.AsyncWaitHandle.Close() } } catch { }
			try { if ($tcp) { $tcp.Close() } } catch { }
		}
	}

	if (-not $reachable)
	{
		$result.DataText = 'Offline'
		$result.LogText = 'Offline'
		$result.Status = 'Offline'
		[pscustomobject]$result
		return
	}
}

if ($protocol -eq 'File')
{
	if (-not [string]::IsNullOrWhiteSpace($machineName) -and -not [string]::IsNullOrWhiteSpace($dbName))
	{
		$candidateDirs = New-Object System.Collections.ArrayList

		foreach ($driveShare in @('c$', 'd$', 'e$', 'f$'))
		{
			foreach ($sqlBaseRoot in @(
				("\\{0}\{1}\Program Files\Microsoft SQL Server" -f $machineName, $driveShare),
				("\\{0}\{1}\Program Files (x86)\Microsoft SQL Server" -f $machineName, $driveShare)
			))
			{
				if (-not (Test-Path -LiteralPath $sqlBaseRoot)) { continue }

				try
				{
					$instanceDirs = Get-ChildItem -LiteralPath $sqlBaseRoot -Directory -ErrorAction Stop
					foreach ($instanceDir in $instanceDirs)
					{
						foreach ($dataDir in @(
							([System.IO.Path]::Combine($instanceDir.FullName, 'MSSQL', 'DATA')),
							([System.IO.Path]::Combine($instanceDir.FullName, 'MSSQL', 'Data'))
						))
						{
							if (Test-Path -LiteralPath $dataDir)
							{
								[void]$candidateDirs.Add($dataDir)
							}
						}
					}
				}
				catch { }
			}
		}

		$candidateDirs = @($candidateDirs | Select-Object -Unique)

		$dataFile = $null
		$logFile = $null

		foreach ($dir in $candidateDirs)
		{
			if ([string]::IsNullOrWhiteSpace($dir)) { continue }
			if (-not (Test-Path -LiteralPath $dir)) { continue }

			$dataExactPath = [System.IO.Path]::Combine($dir, ($dbName + '.mdf'))
			if ((-not $dataFile) -and (Test-Path -LiteralPath $dataExactPath))
			{
				try { $dataFile = Get-Item -LiteralPath $dataExactPath -ErrorAction Stop }
				catch { $dataFile = $null }
			}

			if (-not $dataFile)
			{
				try
				{
					$dataFile = Get-ChildItem -LiteralPath $dir -Filter ($dbName + '*.mdf') -File -ErrorAction Stop |
						Sort-Object Length -Descending |
						Select-Object -First 1
				}
				catch { $dataFile = $null }
			}

			foreach ($logName in @(($dbName + '_Log.ldf'), ($dbName + '_log.ldf'), ($dbName + '.ldf')))
			{
				if ($logFile) { break }
				$logExactPath = [System.IO.Path]::Combine($dir, $logName)
				if (Test-Path -LiteralPath $logExactPath)
				{
					try { $logFile = Get-Item -LiteralPath $logExactPath -ErrorAction Stop }
					catch { $logFile = $null }
				}
			}

			if (-not $logFile)
			{
				try
				{
					$logFile = Get-ChildItem -LiteralPath $dir -Filter ($dbName + '*.ldf') -File -ErrorAction Stop |
						Sort-Object Length -Descending |
						Select-Object -First 1
				}
				catch { $logFile = $null }
			}

			if ($dataFile -or $logFile) { break }
		}

		if ($dataFile)
		{
			$dataMb = [double]($dataFile.Length / 1MB)
			if ($dataMb -ge 1024) { $result.DataText = ('{0:N2} GB' -f ($dataMb / 1024.0)) }
			else { $result.DataText = ('{0:N0} MB' -f $dataMb) }
		}

		if ($logFile)
		{
			$logMb = [double]($logFile.Length / 1MB)
			if ($logMb -ge 1024) { $result.LogText = ('{0:N2} GB' -f ($logMb / 1024.0)) }
			else { $result.LogText = ('{0:N0} MB' -f $logMb) }
		}

		if ($dataFile)
		{
			try { $result.OpenPath = [string]$dataFile.DirectoryName } catch { }
		}
		elseif ($logFile)
		{
			try { $result.OpenPath = [string]$logFile.DirectoryName } catch { }
		}

		if ($dataFile -or $logFile) { $result.Status = 'FileReady' }
		else { $result.Status = 'FileMissing' }
	}
	else
	{
		$result.Status = 'FileMissing'
	}
}

if (($result.Status -ne 'Ready') -and ($result.Status -ne 'FileReady') -and
	(-not [string]::IsNullOrWhiteSpace($smsStartIniPath)) -and (Test-Path -LiteralPath $smsStartIniPath))
{
	try
	{
		$iniLines = Get-Content -LiteralPath $smsStartIniPath -ErrorAction Stop
		$inDatabaseSection = $false
		$dbSizeMb = $null
		$dbLogSizeMb = $null

		foreach ($line in $iniLines)
		{
			$text = [string]$line
			if ($text -match '^\s*\[(.+?)\]\s*$')
			{
				$inDatabaseSection = ($matches[1] -eq 'Database')
				continue
			}
			if (-not $inDatabaseSection) { continue }
			if ($text -match '^\s*DBSIZE\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*$') { $dbSizeMb = [double]$matches[1]; continue }
			if ($text -match '^\s*DBLOGSIZE\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*$') { $dbLogSizeMb = [double]$matches[1]; continue }
		}

		if ($null -ne $dbSizeMb)
		{
			if ($dbSizeMb -ge 1024) { $result.DataText = ('{0:N2} GB' -f ($dbSizeMb / 1024.0)) }
			else { $result.DataText = ('{0:N0} MB' -f $dbSizeMb) }
		}

		if ($null -ne $dbLogSizeMb)
		{
			if ($dbLogSizeMb -ge 1024) { $result.LogText = ('{0:N2} GB' -f ($dbLogSizeMb / 1024.0)) }
			else { $result.LogText = ('{0:N0} MB' -f $dbLogSizeMb) }
		}

		if (($null -ne $dbSizeMb) -or ($null -ne $dbLogSizeMb))
		{
			$result.Status = 'SmsStartFallback'
		}
	}
	catch { }
}

if ($protocol -eq 'File')
{
	[pscustomobject]$result
	return
}

if ([string]::IsNullOrWhiteSpace($connectionString))
{
	[pscustomobject]$result
	return
}

$cn = $null
try
{
	$cn = New-Object System.Data.SqlClient.SqlConnection $connectionString
	$cn.Open()

	$cmd = $cn.CreateCommand()
	$cmd.CommandTimeout = 5
	$cmd.CommandText = @"
SET NOCOUNT ON;
SELECT
	type_desc,
	size,
	physical_name
FROM sys.database_files;
"@

	$rdr = $cmd.ExecuteReader()
	$dataMb = 0.0
	$logMb = 0.0
	$dataPhysicalPath = $null
	$logPhysicalPath = $null
	while ($rdr.Read())
	{
		$typeDesc = $null
		$sizePages = $null
		$physicalName = $null
		if (-not $rdr.IsDBNull(0)) { $typeDesc = [string]$rdr.GetValue(0) }
		if (-not $rdr.IsDBNull(1)) { $sizePages = [double]$rdr.GetValue(1) }
		if (-not $rdr.IsDBNull(2)) { $physicalName = [string]$rdr.GetValue(2) }

		if ($typeDesc -eq 'ROWS')
		{
			if ($null -ne $sizePages) { $dataMb += ($sizePages * 8.0 / 1024.0) }
			if (-not $dataPhysicalPath -and -not [string]::IsNullOrWhiteSpace($physicalName)) { $dataPhysicalPath = $physicalName }
		}
		elseif ($typeDesc -eq 'LOG')
		{
			if ($null -ne $sizePages) { $logMb += ($sizePages * 8.0 / 1024.0) }
			if (-not $logPhysicalPath -and -not [string]::IsNullOrWhiteSpace($physicalName)) { $logPhysicalPath = $physicalName }
		}
	}

	if ($dataMb -gt 0)
	{
		if ($dataMb -ge 1024) { $result.DataText = ('{0:N2} GB' -f ($dataMb / 1024.0)) }
		else { $result.DataText = ('{0:N0} MB' -f $dataMb) }
	}

	if ($logMb -gt 0)
	{
		if ($logMb -ge 1024) { $result.LogText = ('{0:N2} GB' -f ($logMb / 1024.0)) }
		else { $result.LogText = ('{0:N0} MB' -f $logMb) }
	}

	$rawOpenPath = $null
	if (-not [string]::IsNullOrWhiteSpace($dataPhysicalPath))
	{
		try { $rawOpenPath = [System.IO.Path]::GetDirectoryName($dataPhysicalPath) } catch { $rawOpenPath = $null }
	}
	elseif (-not [string]::IsNullOrWhiteSpace($logPhysicalPath))
	{
		try { $rawOpenPath = [System.IO.Path]::GetDirectoryName($logPhysicalPath) } catch { $rawOpenPath = $null }
	}

	if (-not [string]::IsNullOrWhiteSpace($rawOpenPath))
	{
		$machineNameNorm = [string]$machineName
		if ([string]::IsNullOrWhiteSpace($machineNameNorm) -or $machineNameNorm -match '^(?i)(localhost|\.|\(local\))$' -or $machineNameNorm -ieq $env:COMPUTERNAME)
		{
			$result.OpenPath = $rawOpenPath
		}
		elseif ($rawOpenPath -match '^(?<drive>[A-Za-z]):\\(?<rest>.*)$')
		{
			$driveShare = ($matches['drive'].ToLowerInvariant() + '$')
			$pathRest = $matches['rest']
			if ([string]::IsNullOrWhiteSpace($pathRest)) { $result.OpenPath = ('\\{0}\{1}' -f $machineNameNorm, $driveShare) }
			else { $result.OpenPath = ('\\{0}\{1}\{2}' -f $machineNameNorm, $driveShare, $pathRest) }
		}
		else
		{
			$result.OpenPath = $rawOpenPath
		}
	}

	if (($dataMb -gt 0) -or ($logMb -gt 0) -or (-not [string]::IsNullOrWhiteSpace([string]$result.OpenPath)))
	{
		$result.Status = 'Ready'
	}
	try { $rdr.Close() } catch { }
	try { $rdr.Dispose() } catch { }
}
catch
{
	$result.DataText = 'N/A'
	$result.LogText = 'N/A'
	$result.Status = 'Error'
}
finally
{
	try { if ($cn) { $cn.Close() } } catch { }
	try { if ($cn) { $cn.Dispose() } } catch { }
}

if (($result.Status -ne 'Ready') -and
	(-not [string]::IsNullOrWhiteSpace($smsStartIniPath)) -and (Test-Path -LiteralPath $smsStartIniPath))
{
	try
	{
		$iniLines = Get-Content -LiteralPath $smsStartIniPath -ErrorAction Stop
		$inDatabaseSection = $false
		$dbSizeMb = $null
		$dbLogSizeMb = $null

		foreach ($line in $iniLines)
		{
			$text = [string]$line
			if ($text -match '^\s*\[(.+?)\]\s*$')
			{
				$inDatabaseSection = ($matches[1] -eq 'Database')
				continue
			}
			if (-not $inDatabaseSection) { continue }
			if ($text -match '^\s*DBSIZE\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*$') { $dbSizeMb = [double]$matches[1]; continue }
			if ($text -match '^\s*DBLOGSIZE\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*$') { $dbLogSizeMb = [double]$matches[1]; continue }
		}

		if ($null -ne $dbSizeMb)
		{
			if ($dbSizeMb -ge 1024) { $result.DataText = ('{0:N2} GB' -f ($dbSizeMb / 1024.0)) }
			else { $result.DataText = ('{0:N0} MB' -f $dbSizeMb) }
		}

		if ($null -ne $dbLogSizeMb)
		{
			if ($dbLogSizeMb -ge 1024) { $result.LogText = ('{0:N2} GB' -f ($dbLogSizeMb / 1024.0)) }
			else { $result.LogText = ('{0:N0} MB' -f $dbLogSizeMb) }
		}

		if (($null -ne $dbSizeMb) -or ($null -ne $dbLogSizeMb))
		{
			$result.Status = 'SmsStartFallback'
		}
	}
	catch { }
}

[pscustomobject]$result
'@
}

if (-not $global:ProtocolForm)
{
	$global:ProtocolForm = New-Object System.Windows.Forms.Form
	$global:ProtocolForm.Text = "Lane Protocols / DB Sizes"
	$global:ProtocolForm.Size = New-Object System.Drawing.Size(500, 520)
	$global:ProtocolForm.StartPosition = "CenterScreen"
	$global:ProtocolForm.Topmost = $true
	$global:ProtocolForm.BackColor = $uiBg
	$global:ProtocolForm.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
	if ($script:TbsAppIcon) { try { $global:ProtocolForm.Icon = $script:TbsAppIcon } catch { } }

	# No minimize/maximize
	$global:ProtocolForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$global:ProtocolForm.MaximizeBox = $false
	$global:ProtocolForm.MinimizeBox = $false

	# Hide when focus leaves
	$global:ProtocolForm.add_Deactivate({ $global:ProtocolForm.Hide() })

	$global:ProtocolGrid = New-Object System.Windows.Forms.DataGridView
	$global:ProtocolGrid.Location = New-Object System.Drawing.Point(12, 12)
	$global:ProtocolGrid.Size = New-Object System.Drawing.Size(460, 396)
	$global:ProtocolGrid.RowTemplate.Height = 22
	$global:ProtocolGrid.ColumnCount = 4
	$global:ProtocolGrid.Columns[0].Name = "Lane"
	$global:ProtocolGrid.Columns[1].Name = "Protocol"
	$global:ProtocolGrid.Columns[2].Name = "DB"
	$global:ProtocolGrid.Columns[3].Name = "Log"
	$global:ProtocolGrid.ReadOnly = $true
	$global:ProtocolGrid.RowHeadersVisible = $false
	$global:ProtocolGrid.AllowUserToAddRows = $false
	$global:ProtocolGrid.AllowUserToDeleteRows = $false
	$global:ProtocolGrid.AllowUserToResizeRows = $false
	$global:ProtocolGrid.AllowUserToResizeColumns = $false
	$global:ProtocolGrid.SelectionMode = "FullRowSelect"
	$global:ProtocolGrid.Font = New-Object System.Drawing.Font("Consolas", 10)
	$global:ProtocolGrid.Add_CellDoubleClick({
			param ($sender,
				$e)
			if ($e.RowIndex -lt 0) { return }
			try
			{
				$row = $sender.Rows[$e.RowIndex]
				if (-not $row) { return }
				$targetPath = [string]$row.Tag
				if ([string]::IsNullOrWhiteSpace($targetPath)) { return }
				Start-Process -FilePath 'explorer.exe' -ArgumentList @($targetPath) | Out-Null
			}
			catch { }
		})

	# Modern-ish grid styling (safe/light)
	try
	{
		$global:ProtocolGrid.BackgroundColor = $uiCardBg
		$global:ProtocolGrid.BorderStyle = 'FixedSingle'
		$global:ProtocolGrid.CellBorderStyle = 'SingleHorizontal'
		$global:ProtocolGrid.EnableHeadersVisualStyles = $false
		$global:ProtocolGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 247)
		$global:ProtocolGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
		$global:ProtocolGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
		$global:ProtocolGrid.DefaultCellStyle.BackColor = $uiCardBg
		$global:ProtocolGrid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
		$global:ProtocolGrid.GridColor = [System.Drawing.Color]::FromArgb(235, 238, 242)
		$global:ProtocolGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 253)
		$global:ProtocolGrid.AlternatingRowsDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
		$global:ProtocolGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(216, 236, 255)
		$global:ProtocolGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
	}
	catch { }

	$global:ProtocolForm.Controls.Add($global:ProtocolGrid)

	# Hint line under the grid
	$protocolHint = New-Object System.Windows.Forms.Label
	$protocolHint.Text = "Double-click a row to open its database folder."
	$protocolHint.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
	$protocolHint.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 125)
	$protocolHint.AutoSize = $true
	$protocolHint.Location = New-Object System.Drawing.Point(14, 412)
	$global:ProtocolForm.Controls.Add($protocolHint)

	$closeBtn = New-Object System.Windows.Forms.Button
	$closeBtn.Text = "Hide"
	$closeBtn.Location = New-Object System.Drawing.Point(176, 438)
	$closeBtn.Size = New-Object System.Drawing.Size(120, 32)
	$closeBtn.FlatStyle = 'Flat'
	$closeBtn.UseVisualStyleBackColor = $false
	$closeBtn.BackColor = [System.Drawing.Color]::White
	$closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
	$closeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
	$closeBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 205, 212)
	$closeBtn.FlatAppearance.BorderSize = 1
	$closeBtn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
	$closeBtn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(235, 242, 252)
	$closeBtn.Add_Click({ $global:ProtocolForm.Hide() })
	$global:ProtocolForm.Controls.Add($closeBtn)

	# Prevent disposal on X; just hide
	$global:ProtocolForm.add_FormClosing({
			$_.Cancel = $true
			$global:ProtocolForm.Hide()
		})
}

if (-not $global:ProtocolFormTimer)
{
	$global:ProtocolFormTimer = New-Object System.Windows.Forms.Timer
	$global:ProtocolFormTimer.Interval = 500
	$global:ProtocolFormTimer.add_Tick({
			if ($script:LaneDbSizeJobs -and $script:LaneDbSizeJobs.Count -gt 0)
			{
				foreach ($laneKey in @($script:LaneDbSizeJobs.Keys))
				{
					$st = $script:LaneDbSizeJobs[$laneKey]
					if (-not $st) { continue }

					$handle = $st.Handle
					if ($handle -and $handle.IsCompleted)
					{
						$ps = $st.PS
						$resultList = $null
						try { $resultList = $ps.EndInvoke($handle) }
						catch { $resultList = $null }
						try { if ($ps) { $ps.Dispose() } }
						catch { }

						[void]$script:LaneDbSizeJobs.Remove($laneKey)

						$result = $null
						if ($resultList -and $resultList.Count -ge 1) { $result = $resultList[0] }

						if ($result)
						{
							$script:LaneDbSizeMap[$laneKey] = [pscustomobject]@{
								DataText = [string]$result.DataText
								LogText  = [string]$result.LogText
								Status   = [string]$result.Status
								OpenPath = [string]$result.OpenPath
								LastUpdate = Get-Date
							}
						}
						else
						{
							$script:LaneDbSizeMap[$laneKey] = [pscustomobject]@{
								DataText   = 'N/A'
								LogText    = 'N/A'
								Status	   = 'Error'
								OpenPath   = $null
								LastUpdate = Get-Date
							}
						}
					}
				}
			}

			if (-not $global:ProtocolForm.Visible) { return }

			# Save scroll & selection BEFORE clearing rows
			$prevRowCount = $global:ProtocolGrid.Rows.Count
			$scrollIndex = 0
			if ($prevRowCount -gt 0)
			{
				try { $scrollIndex = $global:ProtocolGrid.FirstDisplayedScrollingRowIndex }
				catch { $scrollIndex = 0 }
			}
			$selIndex = $null
			if ($global:ProtocolGrid.SelectedRows.Count -gt 0)
			{
				try { $selIndex = $global:ProtocolGrid.SelectedRows[0].Index }
				catch { $selIndex = $null }
			}

			# Refresh grid
			$global:ProtocolGrid.Rows.Clear()

			$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
			$storeNumberForPopup = [string]$script:FunctionResults['StoreNumber']
			if ($storeNumberForPopup) { $storeNumberForPopup = ($storeNumberForPopup -replace '[^\d]', '').PadLeft(3, '0') }
			$dbsRootForPopup = $null
			try { $dbsRootForPopup = [string]$script:DbsPath }
			catch { $dbsRootForPopup = $null }
			if ([string]::IsNullOrWhiteSpace($dbsRootForPopup))
			{
				try { $dbsRootForPopup = [string]$DbsPath }
				catch { $dbsRootForPopup = $null }
			}
			if ([string]::IsNullOrWhiteSpace($dbsRootForPopup))
			{
				try { $dbsRootForPopup = [System.IO.Path]::Combine($OfficePath, 'Dbs') }
				catch { $dbsRootForPopup = $null }
			}

			$popupMode = [string]$script:ProtocolPopupMode
			if ([string]::IsNullOrWhiteSpace($popupMode)) { $popupMode = 'Lane' }

			if ($popupMode -eq 'Server')
			{
				$global:ProtocolForm.Text = "Server DB / Log Sizes"
				$global:ProtocolGrid.Columns[0].Name = "Server"

				$now = Get-Date
				$serverKey = '__SERVER__'
				$serverDisplay = [string]$script:FunctionResults['DBSERVER']
				if ([string]::IsNullOrWhiteSpace($serverDisplay) -or $serverDisplay -eq 'N/A') { $serverDisplay = $env:COMPUTERNAME }

				$dbName = [string]$script:FunctionResults['DBNAME']
				if ($dbName -eq 'N/A') { $dbName = $null }

				$connStr = [string]$script:FunctionResults['ConnectionString']
				if ($connStr -eq 'N/A') { $connStr = $null }

				$smsStartFallbackPath = $null
				if (-not [string]::IsNullOrWhiteSpace($dbsRootForPopup) -and -not [string]::IsNullOrWhiteSpace($storeNumberForPopup))
				{
					try { $smsStartFallbackPath = [System.IO.Path]::Combine($dbsRootForPopup, ("INFO_{0}901_SMSStart.ini" -f $storeNumberForPopup)) }
					catch { $smsStartFallbackPath = $null }
				}

				$proto = 'Local'
				try
				{
					if ($connStr -match '(?i)(?:Data Source|Server)\s*=\s*np:') { $proto = 'Named Pipes' }
					elseif ($connStr -match '(?i)(?:Data Source|Server)\s*=\s*tcp:') { $proto = 'TCP' }
					elseif ($serverDisplay -match '^(?i)(localhost|\(local\)|\.)(\\|$)') { $proto = 'Local' }
					elseif ($serverDisplay -match '\\') { $proto = 'Instance' }
					else { $proto = 'Server' }
				}
				catch { }

				if (-not ($script:LaneDbSizeJobs.ContainsKey($serverKey)))
				{
					$sizeEntry = $script:LaneDbSizeMap[$serverKey]
					$lastUpdate = $null
					try { $lastUpdate = $sizeEntry.LastUpdate }
					catch { $lastUpdate = $null }

					$needsRefresh = $false
					if (-not $sizeEntry) { $needsRefresh = $true }
					elseif (-not ($lastUpdate -is [datetime])) { $needsRefresh = $true }
					elseif ((($now) - $lastUpdate).TotalSeconds -ge $script:LaneDbSizeRefreshSeconds) { $needsRefresh = $true }

					if ($needsRefresh -and $script:LaneDbSizePool)
					{
						$workerProtocol = if ([string]::IsNullOrWhiteSpace($connStr)) { 'File' }
						else { 'SQL' }
						$canRunWorker = $false
						if ($workerProtocol -eq 'File')
						{
							$canRunWorker = -not [string]::IsNullOrWhiteSpace($serverDisplay) -and -not [string]::IsNullOrWhiteSpace($dbName)
						}
						else
						{
							$canRunWorker = -not [string]::IsNullOrWhiteSpace($connStr)
						}

						if (-not $canRunWorker)
						{
							$script:LaneDbSizeMap[$serverKey] = [pscustomobject]@{
								DataText   = 'N/A'
								LogText    = 'N/A'
								Status	   = 'NoConn'
								OpenPath   = $null
								LastUpdate = $now
							}
						}
						else
						{
							$script:LaneDbSizeMap[$serverKey] = [pscustomobject]@{
								DataText = if ($workerProtocol -eq 'File') { 'Scan...' } else { '...' }
								LogText  = if ($workerProtocol -eq 'File') { 'Scan...' } else { '...' }
								Status   = 'Pending'
								OpenPath = $null
								LastUpdate = $lastUpdate
							}

							$ps = [System.Management.Automation.PowerShell]::Create()
							$ps.RunspacePool = $script:LaneDbSizePool
							$null = $ps.AddScript($script:LaneDbSizeWorker).
							AddArgument([string]$serverKey).
							AddArgument([string]$workerProtocol).
							AddArgument([string]$connStr).
							AddArgument([string]$env:COMPUTERNAME).
							AddArgument([string]$dbName).
							AddArgument([string]$smsStartFallbackPath)
							$handle = $ps.BeginInvoke()

							$script:LaneDbSizeJobs[$serverKey] = @{ PS = $ps; Handle = $handle }
						}
					}
				}

				$sizeEntry = $script:LaneDbSizeMap[$serverKey]
				$dataText = '...'
				$logText = '...'
				if ($sizeEntry)
				{
					try
					{
						if (-not [string]::IsNullOrWhiteSpace([string]$sizeEntry.DataText)) { $dataText = [string]$sizeEntry.DataText }
						if (-not [string]::IsNullOrWhiteSpace([string]$sizeEntry.LogText)) { $logText = [string]$sizeEntry.LogText }
					}
					catch { }
				}

				$r = $global:ProtocolGrid.Rows.Add()
				$global:ProtocolGrid.Rows[$r].Cells[0].Value = $serverDisplay
				$global:ProtocolGrid.Rows[$r].Cells[1].Value = $proto
				$global:ProtocolGrid.Rows[$r].Cells[2].Value = $dataText
				$global:ProtocolGrid.Rows[$r].Cells[3].Value = $logText
				try { $global:ProtocolGrid.Rows[$r].Tag = [string]$sizeEntry.OpenPath }
				catch { }
			}
			else
			{
				$global:ProtocolForm.Text = "Lane Protocols / DB Sizes"
				$global:ProtocolGrid.Columns[0].Name = "Lane"

				# Prefer the new map (best), else fall back to LaneProtocols
				$map = $script:ProtocolResultsMap
				if (-not $map) { $map = $script:LaneProtocols }

				if ($map -and $map.Count -gt 0)
				{
					# Only numeric 3-digit lanes (LaneProtocols can also contain machine-name keys)
					$laneKeys = @($map.Keys | Where-Object { "$_" -match '^\d{3}$' } | Sort-Object { [int]$_ })

					$laneDbInfoMap = $script:FunctionResults['LaneDatabaseInfo']
					$now = Get-Date

					foreach ($ln in $laneKeys)
					{
						$proto = [string]$map[$ln]
						if ([string]::IsNullOrWhiteSpace($proto)) { $proto = "File" }

						$machineName = $ln
						if ($LaneNumToMachineName)
						{
							try
							{
								$mn = $LaneNumToMachineName[$ln]
								if (-not [string]::IsNullOrWhiteSpace($mn)) { $machineName = $mn }
							}
							catch { }
						}

						$dbInfo = $null
						try
						{
							if ($laneDbInfoMap -and $laneDbInfoMap.ContainsKey($ln))
							{
								$dbInfo = $laneDbInfoMap[$ln]
							}
						}
						catch { }

						if (-not $dbInfo)
						{
							try { $dbInfo = Get_All_Lanes_Database_Info -LaneNumber $ln }
							catch { $dbInfo = $null }
						}

						$dbName = $null
						$connStr = $null
						$smsStartFallbackPath = $null
						if ($dbInfo)
						{
							try
							{
								if ($dbInfo -is [hashtable])
								{
									$dbName = [string]$dbInfo['DBName']
									if ($proto -eq 'Named Pipes') { $connStr = [string]$dbInfo['NamedPipesConnStr'] }
									else { $connStr = [string]$dbInfo['TcpConnStr'] }
								}
								else
								{
									$dbName = [string]$dbInfo.DBName
									if ($proto -eq 'Named Pipes') { $connStr = [string]$dbInfo.NamedPipesConnStr }
									else { $connStr = [string]$dbInfo.TcpConnStr }
								}
							}
							catch
							{
								$dbName = $null
								$connStr = $null
							}
						}
						if (-not [string]::IsNullOrWhiteSpace($machineName) -and -not [string]::IsNullOrWhiteSpace($storeNumberForPopup))
						{
							try { $smsStartFallbackPath = ("\\{0}\storeman\office\dbs\INFO_{1}{2}_SMSStart.ini" -f $machineName, $storeNumberForPopup, $ln) }
							catch { $smsStartFallbackPath = $null }
						}

						if (-not ($script:LaneDbSizeJobs.ContainsKey($ln)))
						{
							$sizeEntry = $script:LaneDbSizeMap[$ln]
							$lastUpdate = $null
							try { $lastUpdate = $sizeEntry.LastUpdate }
							catch { $lastUpdate = $null }

							$needsRefresh = $false
							if (-not $sizeEntry) { $needsRefresh = $true }
							elseif (-not ($lastUpdate -is [datetime])) { $needsRefresh = $true }
							elseif ((($now) - $lastUpdate).TotalSeconds -ge $script:LaneDbSizeRefreshSeconds) { $needsRefresh = $true }

							if ($proto -eq 'Offline')
							{
								$script:LaneDbSizeMap[$ln] = [pscustomobject]@{
									DataText   = 'Offline'
									LogText    = 'Offline'
									Status	   = 'Offline'
									OpenPath   = $null
									LastUpdate = $now
								}
							}
							elseif ($needsRefresh -and $script:LaneDbSizePool)
							{
								$canRunWorker = $false
								if ($proto -eq 'File')
								{
									$canRunWorker = -not [string]::IsNullOrWhiteSpace($machineName) -and -not [string]::IsNullOrWhiteSpace($dbName)
								}
								else
								{
									$canRunWorker = -not [string]::IsNullOrWhiteSpace($connStr)
								}

								if (-not $canRunWorker)
								{
									$script:LaneDbSizeMap[$ln] = [pscustomobject]@{
										DataText   = 'N/A'
										LogText    = 'N/A'
										Status	   = 'NoConn'
										OpenPath   = $null
										LastUpdate = $now
									}
								}
								else
								{
									$script:LaneDbSizeMap[$ln] = [pscustomobject]@{
										DataText = if ($proto -eq 'File') { 'Scan...' } else { '...' }
										LogText  = if ($proto -eq 'File') { 'Scan...' } else { '...' }
										Status   = 'Pending'
										OpenPath = $null
										LastUpdate = $lastUpdate
									}

									$ps = [System.Management.Automation.PowerShell]::Create()
									$ps.RunspacePool = $script:LaneDbSizePool
									$null = $ps.AddScript($script:LaneDbSizeWorker).
									AddArgument([string]$ln).
									AddArgument([string]$proto).
									AddArgument([string]$connStr).
									AddArgument([string]$machineName).
									AddArgument([string]$dbName).
									AddArgument([string]$smsStartFallbackPath)
									$handle = $ps.BeginInvoke()

									$script:LaneDbSizeJobs[$ln] = @{ PS = $ps; Handle = $handle }
								}
							}
						}

						$sizeEntry = $script:LaneDbSizeMap[$ln]
						$dataText = '...'
						$logText = '...'
						if ($sizeEntry)
						{
							try
							{
								if (-not [string]::IsNullOrWhiteSpace([string]$sizeEntry.DataText)) { $dataText = [string]$sizeEntry.DataText }
								if (-not [string]::IsNullOrWhiteSpace([string]$sizeEntry.LogText)) { $logText = [string]$sizeEntry.LogText }
							}
							catch { }
						}

						$r = $global:ProtocolGrid.Rows.Add()
						$global:ProtocolGrid.Rows[$r].Cells[0].Value = $machineName
						$global:ProtocolGrid.Rows[$r].Cells[1].Value = $proto
						$global:ProtocolGrid.Rows[$r].Cells[2].Value = $dataText
						$global:ProtocolGrid.Rows[$r].Cells[3].Value = $logText
						try { $global:ProtocolGrid.Rows[$r].Tag = [string]$sizeEntry.OpenPath }
						catch { }
					}

					# Optional: rebuild $script:ProtocolResults so any other old code keeps working
					$script:ProtocolResults = @(
						foreach ($ln in $laneKeys)
						{
							[pscustomobject]@{ Lane = $ln; Protocol = [string]$map[$ln] }
						}
					)
				}
			}

			# Column widths (Lane/Protocol fixed, DB/Log fill; account for scrollbar)
			$col0Width = if ($popupMode -eq 'Server') { 180 }
			else { 130 }
			$col1Width = 90
			$global:ProtocolGrid.Columns[0].Width = $col0Width
			$global:ProtocolGrid.Columns[1].Width = $col1Width
			$visibleRowCount = [math]::Floor($global:ProtocolGrid.DisplayRectangle.Height / $global:ProtocolGrid.RowTemplate.Height)
			$scrollBarVisible = $global:ProtocolGrid.Rows.Count -gt $visibleRowCount
			$remainingWidth = $global:ProtocolGrid.Width - $col0Width - $col1Width - 4
			if ($scrollBarVisible)
			{
				$remainingWidth -= [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
			}
			if ($remainingWidth -lt 180) { $remainingWidth = 180 }
			$dbWidth = [math]::Floor($remainingWidth / 2)
			$logWidth = $remainingWidth - $dbWidth
			if ($dbWidth -lt 90) { $dbWidth = 90 }
			if ($logWidth -lt 90) { $logWidth = 90 }
			$global:ProtocolGrid.Columns[2].Width = $dbWidth
			$global:ProtocolGrid.Columns[3].Width = $logWidth

			if ($global:ProtocolGrid.Columns[2].Width + $global:ProtocolGrid.Columns[3].Width -gt $remainingWidth)
			{
				$global:ProtocolGrid.Columns[3].Width = [math]::Max(90, $remainingWidth - $global:ProtocolGrid.Columns[2].Width)
			}

			# Restore scroll & selection safely
			$rowCount = $global:ProtocolGrid.Rows.Count
			if ($rowCount -gt 0)
			{
				if ($scrollIndex -lt 0) { $scrollIndex = 0 }
				if ($scrollIndex -ge $rowCount) { $scrollIndex = $rowCount - 1 }
				try { $global:ProtocolGrid.FirstDisplayedScrollingRowIndex = $scrollIndex }
				catch { }
			}
			if ($selIndex -ne $null -and $rowCount -gt $selIndex -and $selIndex -ge 0)
			{
				try { $global:ProtocolGrid.Rows[$selIndex].Selected = $true }
				catch { }
			}
		})
	$global:ProtocolFormTimer.Start()
}
