######################################################################################################################
# 
# Lane Tools Button
#
######################################################################################################################

############################################################################
# Lane Tools Anchor Button (Blue)
############################################################################

# Lane Context Menu
$ContextMenuLane = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuLane.ShowItemToolTips = $true
& $Apply_Modern_Menu_Style $ContextMenuLane

# Submenu containers (the assembled layout is at the bottom of this menu section)
$SubLaneMaintenance = New-Object System.Windows.Forms.ToolStripMenuItem("Maintenance")
$SubLaneMaintenance.ToolTipText = "Lane database and program maintenance tasks"
$SubLaneSchedule = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule")
$SubLaneSchedule.ToolTipText = "Scheduled maintenance and backup tasks"
$SubLaneFiles = New-Object System.Windows.Forms.ToolStripMenuItem("Files & Data")
$SubLaneFiles.ToolTipText = "Lane files, shares and data retrieval"
$SubLanePower = New-Object System.Windows.Forms.ToolStripMenuItem("Power & Services")
$SubLanePower.ToolTipText = "Reboots, restarts, time sync and service tasks"

$LaneToolsButton = New-Object System.Windows.Forms.Button
$LaneToolsButton.Name = "LaneToolsButton"
$LaneToolsButton.Text = "Lane Tools"
$LaneToolsButton.Location = New-Object System.Drawing.Point(275, 475)
$LaneToolsButton.Size = New-Object System.Drawing.Size(200, 52)
$LaneToolsButton.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Regular)
$LaneToolsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$LaneToolsButton.UseVisualStyleBackColor = $false
$LaneToolsButton.BackColor = [System.Drawing.Color]::White
$LaneToolsButton.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
$LaneToolsButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$LaneToolsButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$LaneToolsButton.FlatAppearance.BorderSize = 0
$LaneToolsButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(230, 243, 255)
$LaneToolsButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(215, 235, 250)
$LaneToolsButton.Tag = [pscustomobject]@{ CategoryColor = $uiCategoryLane }

# Bottom accent bar
$LaneToolsButton.add_Paint({
		param ($s, $e)
		try
		{
			$pen = New-Object System.Drawing.Pen($uiCategoryLane, 3)
			$e.Graphics.DrawLine($pen, 4, $s.Height - 4, $s.Width - 4, $s.Height - 4)
			$pen.Dispose()
		}
		catch { }
	})

# Hover effects
$LaneToolsButton.add_MouseEnter({
		param ($s, $e)
		try { $s.BackColor = [System.Drawing.Color]::FromArgb(230, 243, 255); $s.ForeColor = $uiCategoryLane }
		catch { }
	})
$LaneToolsButton.add_MouseLeave({
		param ($s, $e)
		try { $s.BackColor = [System.Drawing.Color]::White; $s.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32) }
		catch { }
	})
$LaneToolsButton.add_Click({
		$script:LastActivity = Get-Date
		$ContextMenuLane.Show($LaneToolsButton, 0, $LaneToolsButton.Height)
	})
$LaneToolsButton.add_MouseDown({
		param ($sender, $e)
		if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right)
		{
			$script:ProtocolPopupMode = 'Lane'
			$global:ProtocolForm.Show()
			$global:ProtocolForm.BringToFront()
		}
	})
$toolTip.SetToolTip($LaneToolsButton, "Click for lane-related tools. Right-click for lane protocol window.")
$form.Controls.Add($LaneToolsButton)

############################################################################
# 1) Lane DB Maintenance Button
############################################################################
$LaneDBMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Lane DB Maintenance")
$LaneDBMaintenanceItem.ToolTipText = "Runs maintenance at the lane(s) databases for the selected lane(s)."
$LaneDBMaintenanceItem.Add_Click({
		$script:LastActivity = Get-Date
		Process_Lanes -StoreNumber $StoreNumber
	})
[void]$ContextMenuLane.Items.Add($LaneDBMaintenanceItem)

############################################################################
# 2) Schedule the DB maintenance at the lanes
############################################################################
$LaneScheduleMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Lane DB Maintenance")
$LaneScheduleMaintenanceItem.ToolTipText = "Schedule a task to run maintenance at the lane/s database."
$LaneScheduleMaintenanceItem.Add_Click({
		$script:LastActivity = Get-Date
		Schedule_Lane_DB_Maintenance -StoreNumber "$StoreNumber"
	})
[void]$SubLaneSchedule.DropDownItems.Add($LaneScheduleMaintenanceItem)

############################################################################
# 3) Schedule DB backup of the lanes
############################################################################
$LaneScheduleBackupItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Lane DB Backups")
$LaneScheduleBackupItem.ToolTipText = "Schedule a task to run backups of the selected lanes' databases."
$LaneScheduleBackupItem.Add_Click({
		$script:LastActivity = Get-Date
		Schedule_LaneDB_Backup
	})
[void]$SubLaneSchedule.DropDownItems.Add($LaneScheduleBackupItem)

############################################################################
#  X) Install/Check LOC Options (Lanes) - lane picker + options picker  (kept disabled)
############################################################################
<# ############################################################################
$InstallCheckLOCOptionsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Install/Check LOC Options")
$InstallCheckLOCOptionsItem.ToolTipText = "Pick lanes, then pick LOC Options to audit/install/reinstall (with categories & search)."
$InstallCheckLOCOptionsItem.Add_Click({
		$script:LastActivity = Get-Date
		Install_And_Check_LOC_SMS_Options_On_Lanes -StoreNumber "$StoreNumber"
	})
[void]$ContextMenuLane.Items.Add($InstallCheckLOCOptionsItem)
#>	############################################################################

############################################################################
#  X) Audit/Repair Lane Databases - lane picker + in-function level picker
############################################################################
$RepairLaneDatabasesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Audit/Repair Lane Databases")
$RepairLaneDatabasesItem.ToolTipText = "Pick lanes, then choose Audit/Quick/Deep in the next dialog. Uses Startup.ini via Get_All_Lanes_Database_Info."
$RepairLaneDatabasesItem.Add_Click({
		$script:LastActivity = Get-Date
		Repair_LOC_Databases_On_Lanes -StoreNumber "$StoreNumber"
	})
[void]$SubLaneMaintenance.DropDownItems.Add($RepairLaneDatabasesItem)

############################################################################
# 4) Pump Table to Lane Menu Item
############################################################################
$PumpTableToLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Pump Table to Lane")
$PumpTableToLaneItem.ToolTipText = "Pump the selected tables to the lane/s databases."
$PumpTableToLaneItem.Add_Click({
		$script:LastActivity = Get-Date
		Pump_Tables -StoreNumber $StoreNumber
	})
[void]$SubLaneMaintenance.DropDownItems.Add($PumpTableToLaneItem)

############################################################################
# 5) Update Lane Configuration Menu Item
############################################################################
$UpdateLaneConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Lane Configuration")
$UpdateLaneConfigItem.ToolTipText = "Update the configuration files for the lanes. Fixes connectivity errors and mistakes made during lane ghosting."
$UpdateLaneConfigItem.Add_Click({
		$script:LastActivity = Get-Date
		Update_Lane_Config -StoreNumber $StoreNumber
	})
[void]$SubLaneMaintenance.DropDownItems.Add($UpdateLaneConfigItem)

############################################################################
# 6) Close Open Transactions Menu Item
############################################################################
$CloseOpenTransItem = New-Object System.Windows.Forms.ToolStripMenuItem("Close Open Transactions")
$CloseOpenTransItem.ToolTipText = "Close any open transactions at the lane/s."
$CloseOpenTransItem.Add_Click({
		$script:LastActivity = Get-Date
		Close_Open_Transactions -StoreNumber $StoreNumber
	})
[void]$SubLaneMaintenance.DropDownItems.Add($CloseOpenTransItem)

############################################################################
# 7) Open Lane C$ Share(s)
############################################################################
$OpenLaneCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Lane C$ Share(s)")
$OpenLaneCShareItem.ToolTipText = "Select lanes and open their administrative C$ shares in Explorer."
$OpenLaneCShareItem.Add_Click({
		$script:LastActivity = Get-Date
		Open_Selected_Node_C_Path -StoreNumber $StoreNumber -NodeTypes Lane
	})
[void]$SubLaneFiles.DropDownItems.Add($OpenLaneCShareItem)

############################################################################
# 8) Delete DBS Menu Item
############################################################################
$DeleteDBSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Delete DBS")
$DeleteDBSItem.ToolTipText = "Delete the DBS files (*.txt, *.dwr, if selected *.sus as well) at the lane."
$DeleteDBSItem.Add_Click({
		$script:LastActivity = Get-Date
		Delete_DBS -StoreNumber "$StoreNumber"
	})
[void]$SubLaneFiles.DropDownItems.Add($DeleteDBSItem)

############################################################################
# 9) Refresh PIN Pad Files Menu Item
############################################################################
$RefreshPinPadFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh PIN Pad Files")
$RefreshPinPadFilesItem.ToolTipText = "Refresh the PIN pad files for the lane/s."
$RefreshPinPadFilesItem.Add_Click({
		$script:LastActivity = Get-Date
		Refresh_PIN_Pad_Files -StoreNumber "$StoreNumber"
	})
[void]$SubLaneFiles.DropDownItems.Add($RefreshPinPadFilesItem)

############################################################################
# 10) Drawer Control Item
############################################################################
$DrawerControlItem = New-Object System.Windows.Forms.ToolStripMenuItem("Drawer Control")
$DrawerControlItem.ToolTipText = "Set the Drawer Control for a lane for testing"
$DrawerControlItem.Add_Click({
		$script:LastActivity = Get-Date
		Drawer_Control -StoreNumber $StoreNumber
	})
[void]$SubLanePower.DropDownItems.Add($DrawerControlItem)

############################################################################
# 11) Refresh Database
############################################################################
$RefreshDatabaseItem = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Database")
$RefreshDatabaseItem.ToolTipText = "Refresh the database at the lane/s"
$RefreshDatabaseItem.Add_Click({
		$script:LastActivity = Get-Date
		Refresh_Database -StoreNumber $StoreNumber
	})
[void]$ContextMenuLane.Items.Add($RefreshDatabaseItem)

############################################################################
# 12) Send Restart Command Menu Item
############################################################################
$SendRestartCommandItem = New-Object System.Windows.Forms.ToolStripMenuItem("Send Restart All Programs")
$SendRestartCommandItem.ToolTipText = "Send restart all programs to selected lane(s) for the store."
$SendRestartCommandItem.Add_Click({
		$script:LastActivity = Get-Date
		Send_Restart_All_Programs -StoreNumber "$StoreNumber"
	})
[void]$SubLanePower.DropDownItems.Add($SendRestartCommandItem)

############################################################################
# 13) Enable SQL Protocols Menu Item
############################################################################
$EnableSQLProtocolsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Enable SQL Protocols")
$EnableSQLProtocolsItem.ToolTipText = "Enable TCP/IP, Named Pipes, and set static port for SQL Server on selected lane(s)."
$EnableSQLProtocolsItem.Add_Click({
		$script:LastActivity = Get-Date
		Enable_SQL_Protocols_On_Selected_Lanes -StoreNumber $StoreNumber
	})
[void]$SubLaneMaintenance.DropDownItems.Add($EnableSQLProtocolsItem)

############################################################################
# 14) Check Lane SQL / VNC Services
############################################################################
$LaneServiceStatusItem = New-Object System.Windows.Forms.ToolStripMenuItem("Check Lane SQL / VNC Services")
$LaneServiceStatusItem.ToolTipText = "Open the all-lanes SQL/VNC service viewer. Right-click a lane in the grid for actions."
$LaneServiceStatusItem.Add_Click({
		$script:LastActivity = Get-Date
		Show_Lane_SQL_VNC_Service_Statuses
	})
[void]$ContextMenuLane.Items.Add($LaneServiceStatusItem)

############################################################################
# 15) Set the time on the lanes
############################################################################
$SetLaneTimeFromLocalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Set/Schedule Time on Lanes")
$SetLaneTimeFromLocalItem.ToolTipText = "Synchronize or schedule time sync for selected lanes."
$SetLaneTimeFromLocalItem.Add_Click({
		$script:LastActivity = Get-Date
		$modeResult = [System.Windows.Forms.MessageBox]::Show(
			"Do you want to schedule recurring sync? (Yes for schedule, No for one-time)",
			"Choose Mode",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Question
		)
		$isSchedule = ($modeResult -eq [System.Windows.Forms.DialogResult]::Yes)
		Send_SERVER_time_to_Lanes -StoreNumber $StoreNumber -Schedule:$isSchedule
	})
[void]$SubLanePower.DropDownItems.Add($SetLaneTimeFromLocalItem)

############################################################################
# 16) Sync Server Key (SYS_TAB.F1243) + Optional Scheduled Task
############################################################################
$SyncServerKeyItem = New-Object System.Windows.Forms.ToolStripMenuItem("Sync Server Key")
$SyncServerKeyItem.ToolTipText = "Reads SYS_TAB.F1243 from the store DB, deploys it to selected lane DBs (SQL -> SQI fallback), restarts programs, and can add/merge lanes into an auto-sync Scheduled Task."
$SyncServerKeyItem.Add_Click({
		$script:LastActivity = Get-Date
		try
		{
			Send_SMS_Key_to_Lanes -StoreNumber "$StoreNumber"
		}
		catch
		{
			Write_Log "Sync Server Key failed: $_" "red"
			try { [System.Windows.Forms.MessageBox]::Show("Sync Server Key failed:`r`n`r`n$_", "Error", "OK", "Error") | Out-Null }
			catch { }
		}
	})
[void]$SubLanePower.DropDownItems.Add($SyncServerKeyItem)

############################################################################
# 17) TRS CLT Reprocess (Build SAL_HDR_SUS@TER) Menu Item
############################################################################
$TrsCltReprocessItem = New-Object System.Windows.Forms.ToolStripMenuItem("Retrieve Transactions from Lanes")
$TrsCltReprocessItem.ToolTipText = "Creates/refreshes SAL_HDR_SUS@TER (per lane) from SAL_HDR CLOSED by date + tran range, using NP/TCP/File protocol logic."
$TrsCltReprocessItem.Add_Click({
		$script:LastActivity = Get-Date
		Retrive_Transactions -StoreNumber $StoreNumber
	})
[void]$SubLaneFiles.DropDownItems.Add($TrsCltReprocessItem)

############################################################################
# 18) Reboot Lane Menu Item
############################################################################
$RebootLaneItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Lane")
$RebootLaneItem.ToolTipText = "Reboot the selected lane/s."
$RebootLaneItem.Add_Click({
		$script:LastActivity = Get-Date
		Reboot_Nodes -StoreNumber "$StoreNumber" -NodeTypes Lane
	})
[void]$SubLanePower.DropDownItems.Add($RebootLaneItem)

############################################################################
# NOTE: Click handler already assigned above. Avoid duplicate Add_Click.
############################################################################

# --- Assembled layout: frequent actions on top, the rest grouped in submenus ---
[void]$ContextMenuLane.Items.Clear()
[void]$ContextMenuLane.Items.Add($LaneDBMaintenanceItem)
[void]$ContextMenuLane.Items.Add($RefreshDatabaseItem)
[void]$ContextMenuLane.Items.Add($LaneServiceStatusItem)
[void]$ContextMenuLane.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ContextMenuLane.Items.Add($SubLaneMaintenance)
[void]$ContextMenuLane.Items.Add($SubLaneSchedule)
[void]$ContextMenuLane.Items.Add($SubLaneFiles)
[void]$ContextMenuLane.Items.Add($SubLanePower)
foreach ($sub in @($SubLaneMaintenance, $SubLaneSchedule, $SubLaneFiles, $SubLanePower)) { & $Style_SubMenu $sub }
$toolTip.SetToolTip($LaneToolsButton, "Click to see Lane-related tools. Right-click shows lane protocol window.")
