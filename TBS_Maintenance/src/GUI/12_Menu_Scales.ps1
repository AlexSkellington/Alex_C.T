######################################################################################################################
#
# Scales Tools Button
#
######################################################################################################################

############################################################################
# Scales Tools Anchor Button (Purple)
############################################################################
# Scale Context Menu
$ContextMenuScale = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuScale.ShowItemToolTips = $true
& $Apply_Modern_Menu_Style $ContextMenuScale

$ScaleToolsButton = & $New_Category_Button -Name "ScaleToolsButton" -Text "Scale Tools" -X 525 -Accent $uiCategoryScale `
	-HoverBack ([System.Drawing.Color]::FromArgb(245, 230, 250)) -DownBack ([System.Drawing.Color]::FromArgb(230, 215, 240)) -HoverFore $uiCategoryScale `
	-Menu $ContextMenuScale -ToolTipText "Click for scale-related tools."

# Submenu containers (the assembled layout is at the bottom of this menu section)
$SubScaleServices = New-Object System.Windows.Forms.ToolStripMenuItem("Services")
$SubScaleServices.ToolTipText = "BMS / ScaleCommApp services and reboot"
$SubScaleConfig = New-Object System.Windows.Forms.ToolStripMenuItem("Configuration")
$SubScaleConfig.ToolTipText = "Scale configuration and database tables"
$SubScaleFiles = New-Object System.Windows.Forms.ToolStripMenuItem("Files & Shares")
$SubScaleFiles.ToolTipText = "Scale file deployment and C$ shares"

############################################################################
# 1) Repair BMS Service
############################################################################
$repairBMSItem = New-Object System.Windows.Forms.ToolStripMenuItem("Repair BMS Service")
$repairBMSItem.ToolTipText = "Repairs the BMS service for scale deployment."
$repairBMSItem.Add_Click({
		$script:LastActivity = Get-Date
		Repair_BMS
	})
[void]$SubScaleServices.DropDownItems.Add($repairBMSItem)

############################################################################
# 2) Troubleshoot ScaleCommApp
############################################################################
$troubleshootItem = New-Object System.Windows.Forms.ToolStripMenuItem("Troubleshoot ScaleCommApp")
$troubleshootItem.ToolTipText = "Checks and fixes ScaleCommApp config (StoreName, First Scale IP, SQL instance/db, BatchSendFull=10000)."
$troubleshootItem.Add_Click({
		$script:LastActivity = Get-Date
		Troubleshoot_ScaleCommApp -AllMatches -Force
	})
[void]$SubScaleServices.DropDownItems.Add($troubleshootItem)

############################################################################
# 3) Reboot Scales
############################################################################
$Reboot_ScalesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Scales")
$Reboot_ScalesItem.ToolTipText = "Reboot Scale/s."
$Reboot_ScalesItem.Add_Click({
		$script:LastActivity = Get-Date
		Reboot_Nodes -StoreNumber "$StoreNumber" -NodeTypes Scale
	})
[void]$SubScaleServices.DropDownItems.Add($Reboot_ScalesItem)

############################################################################
# 4) Open Scale C$ Share(s)
############################################################################
$OpenScaleCShareItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Scale C$ Share(s)")
$OpenScaleCShareItem.ToolTipText = "Select scales and open their C$ administrative shares as 'bizuser' (bizerba/biyerba)."
$OpenScaleCShareItem.Add_Click({
		$script:LastActivity = Get-Date
		Open_Selected_Node_C_Path -StoreNumber $StoreNumber -NodeTypes Scale
	})
[void]$SubScaleFiles.DropDownItems.Add($OpenScaleCShareItem)

############################################################################
# 5) Edit_TBS_SCL_ver520_Table Menu Item
############################################################################
$OrganizeScaleTableItem = New-Object System.Windows.Forms.ToolStripMenuItem("Edit_TBS_SCL_ver520_Table")
$OrganizeScaleTableItem.ToolTipText = "Organize the Scale SQL table (TBS_SCL_ver520)."
$OrganizeScaleTableItem.Add_Click({
		$script:LastActivity = Get-Date
		Edit_TBS_SCL_ver520_Table
	})
[void]$SubScaleConfig.DropDownItems.Add($OrganizeScaleTableItem)

############################################################################
# 6) Deploy Scale Currency Files
############################################################################
$DeployScaleCurrencyFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Deploy Scale Currency Files")
$DeployScaleCurrencyFilesItem.ToolTipText = "Push currency-configured price files (.txt, .properties) to selected scales (Bizerba only)."
$DeployScaleCurrencyFilesItem.Add_Click({
		$script:LastActivity = Get-Date
		Deploy_Scale_Currency_Files
	})
[void]$SubScaleFiles.DropDownItems.Add($DeployScaleCurrencyFilesItem)

############################################################################
# 7) Edit Ishida SDPs (ALL configs)
############################################################################
$editIshidaSDPsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Edit Ishida SDPs (All Configs)")
$editIshidaSDPsItem.ToolTipText = "Pick subdepartments from SDP_TAB (F04/F1022) and write to IshidaSDPs in ALL ScaleCommApp *.exe.config files."
$editIshidaSDPsItem.Add_Click({
		$script:LastActivity = Get-Date
		Pick_And_Update_IshidaSDPs -AllMatches
	})
[void]$SubScaleConfig.DropDownItems.Add($editIshidaSDPsItem)

############################################################################
# 8) Update Scales Specials
############################################################################
$UpdateScalesSpecialsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Scales Specials")
$UpdateScalesSpecialsItem.ToolTipText = "Update scale specials immediately or schedule as a daily 5AM task."
$UpdateScalesSpecialsItem.Add_Click({
		$script:LastActivity = Get-Date
		Update_Scales_Specials_Interactive
	})
[void]$SubScaleConfig.DropDownItems.Add($UpdateScalesSpecialsItem)

############################################################################
# 9) Update Scale Config and DB (F272 Upsert)
############################################################################
$UpdateScaleConfigAndDBItem = New-Object System.Windows.Forms.ToolStripMenuItem("Update Scale Config & DB (F272 Upsert)")
$UpdateScaleConfigAndDBItem.ToolTipText = "Updates ScaleCommApp configs and upserts F272 in SCL_TAB for POS_TAB F82=1 in item range."
$UpdateScaleConfigAndDBItem.Add_Click({
		$script:LastActivity = Get-Date
		Update_ScaleConfig_And_DB
	})

############################################################################
# 10) Configure Subdepartments & SdpDefault (OBJ.F16 / OBJ.F18 / OBJ.F17)
############################################################################
$sep_ConfigureSdpDefault = New-Object System.Windows.Forms.ToolStripSeparator
#$ContextMenuScale.Items.Add($sep_ConfigureSdpDefault) | Out-Null
$ConfigureSdpDefaultItem = New-Object System.Windows.Forms.ToolStripMenuItem("Configure Sub-Departments for Scales")
$ConfigureSdpDefaultItem.ToolTipText = "Populate subdepartments (Auto or FAM/RPC/CAT) and set SdpDefault to OBJ.F16 / OBJ.F18 / OBJ.F17, or restore SdpDefault=1."
$ConfigureSdpDefaultItem.Add_Click({
		$script:LastActivity = Get-Date
		Configure_Scale_Sub-Departments
	})
[void]$SubScaleConfig.DropDownItems.Add($ConfigureSdpDefaultItem)

############################################################################
# 11) Schedule Duplicate File Monitor
############################################################################
$ScheduleRemoveDupesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Remove duplicate files from (toBizerba)")
$ScheduleRemoveDupesItem.ToolTipText = "Monitor for and auto-delete duplicate files in (toBizerba). Run now or schedule as SYSTEM."
$ScheduleRemoveDupesItem.Add_Click({
		$script:LastActivity = Get-Date
		Remove_Duplicate_Files_From_toBizerba
	})
[void]$SubScaleConfig.DropDownItems.Add($ScheduleRemoveDupesItem)

############################################################################
# 12) Deploy Scale Files (Per-Scale S00 / Alex Profile BMS)
############################################################################
$DeployScaleFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Deploy Scale Files (Per-Scale S00)")
$DeployScaleFilesItem.ToolTipText = "Interactive per-scale Bizerba management (Full, Specials, Delete, Status, Sales)."
$DeployScaleFilesItem.Add_Click({
		$script:LastActivity = Get-Date
		Manage_Bizerba_Scales_Alex_Profile -StoreNumber $StoreNumber
	})

############################################################################
# NOTE: the click handler and the control registration for $ScaleToolsButton
# already happened above - do NOT re-add the button.
############################################################################

# --- Assembled layout: frequent actions on top, the rest grouped in submenus ---
[void]$ContextMenuScale.Items.Add($UpdateScaleConfigAndDBItem)
[void]$ContextMenuScale.Items.Add($DeployScaleFilesItem)
[void]$ContextMenuScale.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ContextMenuScale.Items.Add($SubScaleServices)
[void]$ContextMenuScale.Items.Add($SubScaleConfig)
[void]$ContextMenuScale.Items.Add($SubScaleFiles)
foreach ($sub in @($SubScaleServices, $SubScaleConfig, $SubScaleFiles)) { & $Style_SubMenu $sub }
$toolTip.SetToolTip($ScaleToolsButton, "Click for scale-related tools.")
