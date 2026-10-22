######################################################################################################################
#
# Server Tools Button
#
######################################################################################################################

############################################################################
# Server Tools Button (Green)
############################################################################

# Server Context Menu
$ContextMenuServer = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuServer.ShowItemToolTips = $true
& $Apply_Modern_Menu_Style $ContextMenuServer

# Submenu containers (the assembled layout is at the bottom of this menu section)
$SubServerSchedule = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule")
$SubServerSchedule.ToolTipText = "Scheduled maintenance and backup tasks"
$SubServerSystem = New-Object System.Windows.Forms.ToolStripMenuItem("System Care")
$SubServerSystem.ToolTipText = "Windows and system-level maintenance"

$ServerToolsButton = & $New_Category_Button -Name "ServerToolsButton" -Text "Server Tools" -X 50 -Accent $uiCategoryServer `
	-HoverBack ([System.Drawing.Color]::FromArgb(235, 248, 235)) -DownBack ([System.Drawing.Color]::FromArgb(220, 240, 220)) -HoverFore $uiCategoryServer `
	-Menu $ContextMenuServer -ProtocolMode 'Server' -ToolTipText "Click for server-related tools. Right-click for server DB info."

############################################################################
# 1) Server DB Maintenance
############################################################################
$ServerDBMaintenanceItem = New-Object System.Windows.Forms.ToolStripMenuItem("Server DB Maintenance")
$ServerDBMaintenanceItem.ToolTipText = "Runs maintenance on the store server database."
$ServerDBMaintenanceItem.Add_Click({
		$script:LastActivity = Get-Date
		$confirmation = [System.Windows.Forms.MessageBox]::Show(
			"Do you want to proceed with the server database maintenance?",
			"Confirmation",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Question
		)
		if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			Process_Server -StoresqlFilePath $StoresqlFilePath -PromptForSections
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Operation canceled.",
				"Canceled",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			) | Out-Null
		}
	})

############################################################################
# 2) Schedule the DB maintenance at the lanes
############################################################################
$ServerScheduleRepairItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Server DB Maintenance")
$ServerScheduleRepairItem.ToolTipText = "Schedule a task to run maintenance at the server database."
$ServerScheduleRepairItem.Add_Click({
		$script:LastActivity = Get-Date
		Schedule_Server_DB_Maintenance -StoreNumber $StoreNumber
	})
[void]$SubServerSchedule.DropDownItems.Add($ServerScheduleRepairItem)

############################################################################
# 3) Schedule the local DB backup on the server
############################################################################
$ServerScheduleBackupItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Local DB Backup")
$ServerScheduleBackupItem.ToolTipText = "Schedule a task to run database backup on the server."
$ServerScheduleBackupItem.Add_Click({
		$script:LastActivity = Get-Date
		Schedule_LocalDB_Backup
	})
[void]$SubServerSchedule.DropDownItems.Add($ServerScheduleBackupItem)

############################################################################
# 4) Schedule the Storeman ZIP backup on the server
############################################################################
$ServerScheduleStoremanZipBackupItem = New-Object System.Windows.Forms.ToolStripMenuItem("Schedule Storeman ZIP Backup")
$ServerScheduleStoremanZipBackupItem.ToolTipText = "Schedule a task to back up the Storeman folder to a weekly ZIP archive."
$ServerScheduleStoremanZipBackupItem.Add_Click({
		$script:LastActivity = Get-Date
		Schedule_Storeman_Zip_Backup
	})
[void]$SubServerSchedule.DropDownItems.Add($ServerScheduleStoremanZipBackupItem)

############################################################################
# 5) Manage SQL 'sa' Account Menu Item
############################################################################
$ManageSaAccountItem = New-Object System.Windows.Forms.ToolStripMenuItem("Manage SQL 'sa' Account")
$ManageSaAccountItem.ToolTipText = "Enable or disable the 'sa' account on the local SQL Server with a predefined password."
$ManageSaAccountItem.Add_Click({
		$script:LastActivity = Get-Date
		Manage_Sa_Account
	})

############################################################################
# 6) Repair Windows Menu Item
############################################################################
$RepairWindowsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair Windows")
$RepairWindowsItem.ToolTipText = "Perform repairs on the Windows operating system."
$RepairWindowsItem.Add_Click({
		$script:LastActivity = Get-Date
		Repair_Windows
	})
[void]$SubServerSystem.DropDownItems.Add($RepairWindowsItem)

############################################################################
# 7) Configure System Settings Menu Item
############################################################################
$ConfigureSystemSettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Configure System Settings")
$ConfigureSystemSettingsItem.ToolTipText = "Set High Performance power plan, disable sleep, ensure required services are Automatic/running, and apply visual tweaks."
$ConfigureSystemSettingsItem.Add_Click({
		$script:LastActivity = Get-Date
		$confirmResult = [System.Windows.Forms.MessageBox]::Show(
			"Warning: This will modify power, services, and visual settings. Continue?",
			"Confirm Changes",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Warning
		)
		if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			try { Configure_System_Settings }
			catch
			{
				if (Get-Command Write_Log -ErrorAction SilentlyContinue)
				{
					Write_Log "Configure_System_Settings threw an error: $_" "red"
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Configure System Settings",
						[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				}
			}
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show("Operation canceled.", "Canceled",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
		}
	})
[void]$SubServerSystem.DropDownItems.Add($ConfigureSystemSettingsItem)

############################################################################
# 8) Organize Desktop Items Menu Item
############################################################################
$OrganizeDesktopItemsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Organize Desktop Items")
$OrganizeDesktopItemsItem.ToolTipText = "Move non-system items from the Desktop into the 'Unorganized Items' folder."
$OrganizeDesktopItemsItem.Add_Click({
		$script:LastActivity = Get-Date
		$targetFolderName = "Unorganized Items"
		$confirm = [System.Windows.Forms.MessageBox]::Show(
			"This will move non-system items from the Desktop into '$targetFolderName'. Continue?",
			"Confirm Desktop Organization",
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Warning
		)
		if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes)
		{
			[System.Windows.Forms.MessageBox]::Show("Operation canceled.", "Canceled",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
			return
		}
		try
		{
			if (Get-Command Organize_Desktop_Items -ErrorAction SilentlyContinue)
			{
				Organize_Desktop_Items -UnorganizedFolderName $targetFolderName
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show(
					"Organize_Desktop_Items is not loaded in this session.",
					"Function Not Found",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
			}
		}
		catch
		{
			if (Get-Command Write_Log -ErrorAction SilentlyContinue)
			{
				Write_Log "Organize_Desktop_Items threw an error: $_" "red"
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Organize Desktop Items",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
			}
		}
	})
[void]$SubServerSystem.DropDownItems.Add($OrganizeDesktopItemsItem)

############################################################################
# NOTE: the click / right-click handlers and the control registration for
# $ServerToolsButton were already done when the button was created above -
# do NOT add them again (it opened the menu twice per click).
############################################################################

# --- Assembled layout: frequent actions on top, the rest grouped in submenus ---
[void]$ContextMenuServer.Items.Add($ServerDBMaintenanceItem)
[void]$ContextMenuServer.Items.Add($ManageSaAccountItem)
[void]$ContextMenuServer.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ContextMenuServer.Items.Add($SubServerSchedule)
[void]$ContextMenuServer.Items.Add($SubServerSystem)
foreach ($sub in @($SubServerSchedule, $SubServerSystem)) { & $Style_SubMenu $sub }
