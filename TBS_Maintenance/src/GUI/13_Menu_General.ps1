######################################################################################################################
#
# General Tools Button
#
######################################################################################################################

############################################################################
# General Tools Anchor Button (Orange)
############################################################################
# General Context Menu
$ContextMenuGeneral = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuGeneral.ShowItemToolTips = $true
& $Apply_Modern_Menu_Style $ContextMenuGeneral

$GeneralToolsButton = & $New_Category_Button -Name "GeneralToolsButton" -Text "General Tools" -X 775 -Accent $uiCategoryGeneral `
	-HoverBack ([System.Drawing.Color]::FromArgb(255, 243, 224)) -DownBack ([System.Drawing.Color]::FromArgb(250, 228, 195)) -HoverFore ([System.Drawing.Color]::FromArgb(200, 100, 0)) `
	-Menu $ContextMenuGeneral -ToolTipText "Click for general maintenance tools."

# Submenu containers (the assembled layout is at the bottom of this menu section)
$SubGenDiagnostics = New-Object System.Windows.Forms.ToolStripMenuItem("Diagnostics & Repair")
$SubGenDiagnostics.ToolTipText = "Journal, hosts, archive bit and node repairs"
$SubGenTools = New-Object System.Windows.Forms.ToolStripMenuItem("Tools")
$SubGenTools.ToolTipText = "INI editor and activation"

############################################################################
# 1) Activate Windows ("Alex_C.T")
############################################################################
$activateItem = New-Object System.Windows.Forms.ToolStripMenuItem("Activate Windows")
$activateItem.ToolTipText = "Activate Windows using Alex_C.T's method."
$activateItem.Add_Click({
		$script:LastActivity = Get-Date
		Invoke_Secure_Script
	})
[void]$SubGenTools.DropDownItems.Add($activateItem)

############################################################################
# 2) Reboot System
############################################################################
$rebootItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot System")
$rebootItem.ToolTipText = "Reboot the host system immediately."
$rebootItem.Add_Click({
		$script:LastActivity = Get-Date
		$rebootResult = [System.Windows.Forms.MessageBox]::Show(
			"Do you want to reboot now?",
			"Reboot",
			[System.Windows.Forms.MessageBoxButtons]::YesNo
		)
		if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			Restart-Computer -Force
			Delete_Files -Path "$TempDir" -SpecifiedFiles `
						 "Server_Database_Maintenance.sqi", `
						 "Lane_Database_Maintenance.sqi", `
						 "TBS_Maintenance_Script.ps1"
		}
	})

############################################################################
# 3) Copy Files Between Nodes
############################################################################
$Copy_Files_Between_NodesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Files Between Nodes")
$Copy_Files_Between_NodesItem.ToolTipText = "Copy (storeman) subfolders/files from Server or a Lane to selected lanes."
$Copy_Files_Between_NodesItem.Add_Click({
		$script:LastActivity = Get-Date
		Copy_Files_Between_Nodes
	})

############################################################################
# 4) Edit INIs (Setup.ini and others)
############################################################################
$INI_EditorItem = New-Object System.Windows.Forms.ToolStripMenuItem("INI_Editor")
$INI_EditorItem.ToolTipText = "Edit \storeman\<relative>\*.ini (default: office\Setup.ini) on Server or a Lane, then optionally copy to other lanes."
$INI_EditorItem.Add_Click({
		$script:LastActivity = Get-Date
		INI_Editor
	})
[void]$SubGenTools.DropDownItems.Add($INI_EditorItem)

############################################################################
# 5) Fix Journal
############################################################################
$fixJournalItem = New-Object System.Windows.Forms.ToolStripMenuItem("Fix Journal")
$fixJournalItem.ToolTipText = "Fix journal entries for the specified date."
$fixJournalItem.Add_Click({
		$script:LastActivity = Get-Date
		Fix_Journal -StoreNumber $StoreNumber -OfficePath $OfficePath
	})
[void]$SubGenDiagnostics.DropDownItems.Add($fixJournalItem)

############################################################################
# 6) Reboot selected Backoffices
############################################################################
$RebootBackofficesItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reboot Backoffices")
$RebootBackofficesItem.ToolTipText = "Reboot the selected Backoffice/s."
$RebootBackofficesItem.Add_Click({
		$script:LastActivity = Get-Date
		Reboot_Nodes -StoreNumber $StoreNumber -NodeTypes Backoffice
	})
[void]$SubGenDiagnostics.DropDownItems.Add($RebootBackofficesItem)

############################################################################
# 7) Remove Archive Bit
############################################################################
$RemoveArchiveBitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Remove Archive Bit")
$RemoveArchiveBitItem.ToolTipText = "Remove archived bit from all lanes and server. Option to schedule as a repeating task."
$RemoveArchiveBitItem.Add_Click({
		$script:LastActivity = Get-Date
		Remove_Archive_Bit
	})
[void]$SubGenDiagnostics.DropDownItems.Add($RemoveArchiveBitItem)

############################################################################
# 8) Sync Hosts File for Selected Nodes
############################################################################
$SyncHostsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Sync Host Files")
$SyncHostsItem.ToolTipText = "Update the host file with the nodes selected, then copy the local host file to the selected node."
$SyncHostsItem.Add_Click({
		$script:LastActivity = Get-Date
		Sync_Selected_Node_Hosts -StoreNumber $StoreNumber
	})
[void]$SubGenDiagnostics.DropDownItems.Add($SyncHostsItem)

############################################################################
# 9) Fix Deploy CHG
############################################################################
$FixDeployCHGItem = New-Object System.Windows.Forms.ToolStripMenuItem("Fix Deploy_CHG")
$FixDeployCHGItem.ToolTipText = "Restores the deploy line to DEPLOY_CHG.sql for scale management."
$FixDeployCHGItem.Add_Click({
		$script:LastActivity = Get-Date
		[void](Fix_Deploy_CHG)
	})
[void]$SubGenDiagnostics.DropDownItems.Add($FixDeployCHGItem)

############################################################################
# NOTE: the click handler and the control registration for $GeneralToolsButton
# already happened above - a second Add_Click would open the menu twice.
############################################################################

# --- Assembled layout: frequent actions on top, the rest grouped in submenus ---
[void]$ContextMenuGeneral.Items.Add($rebootItem)
[void]$ContextMenuGeneral.Items.Add($Copy_Files_Between_NodesItem)
[void]$ContextMenuGeneral.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$ContextMenuGeneral.Items.Add($SubGenDiagnostics)
[void]$ContextMenuGeneral.Items.Add($SubGenTools)
foreach ($sub in @($SubGenDiagnostics, $SubGenTools)) { & $Style_SubMenu $sub }
$toolTip.SetToolTip($GeneralToolsButton, "Click for general maintenance tools.")
