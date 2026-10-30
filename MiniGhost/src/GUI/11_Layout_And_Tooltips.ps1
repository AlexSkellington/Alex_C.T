# ===================================================================================================
#                              SECTION: GUI - Button Layout and Tooltips
# ---------------------------------------------------------------------------------------------------
# Description:
#   Styles the eight action buttons (created with their click handlers in 02..09) and places them
#   in the three cards of 01_Main_Form, then sets their tooltips.
# ===================================================================================================

# Card -> buttons, top to bottom; Danger buttons get red text
$actionGroups = @(
	@{ Host = $script:CardIdentity; Accent = $uiAccentIdentity; Buttons = @($changeMachineNameButton, $updateStoreNumberButton, $configureNetworkButton) },
	@{ Host = $script:CardDatabase; Accent = $uiAccentDatabase; Buttons = @($updateSQLDatabaseButton, $truncateTablesButton) },
	@{ Host = $script:CardSystem; Accent = $uiAccentSystem; Buttons = @($registryCleanupButton, $summaryButton, $rebootButton) }
)
$dangerButtons = @($truncateTablesButton, $rebootButton)

foreach ($group in $actionGroups)
{
	# Dock=Top stacks the LAST added control at the top, so add bottom-up
	$ordered = @($group.Buttons)
	[array]::Reverse($ordered)
	foreach ($btn in $ordered)
	{
		$kind = if ($dangerButtons -contains $btn) { 'Danger' } else { 'Action' }
		Set_Modern_Button_Style -Button $btn -Kind $kind -AccentColor $group.Accent
		$btn.Dock = 'Fill'
		$btn.Margin = New-Object System.Windows.Forms.Padding(0)
		$btn.MinimumSize = New-Object System.Drawing.Size(0, 0)
		
		# wrapper = button + 8 px gap below it (docked controls ignore Margin)
		$slot = New-Object System.Windows.Forms.Panel
		$slot.Dock = 'Top'
		$slot.Height = 54
		$slot.BackColor = [System.Drawing.Color]::White
		$slot.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
		$slot.Controls.Add($btn)
		$group.Host.Controls.Add($slot)
	}
}

# Tooltips (one per button)
$toolTip.SetToolTip($updateStoreNumberButton, "Change the store number in Startup.ini, Server.ini, SMSStart.ini, System.ini, INFO_*901_WIN.INI and SmsHttps.INI, then update STD_TAB / TER_TAB 901.")
$toolTip.SetToolTip($configureNetworkButton, "Switch the active network adapter to DHCP, or to a static IP in the gateway's subnet (the address is pinged first).")
$changeMachineNameToolTipText = "Rename this computer or re-enter the current machine name to force a full INI + SQL sync."
if ($script:NodeRole -match '^(?i)Lane$')
{
	$changeMachineNameToolTipText += " Lane formats support POS001, POS-001, LN026001, 0384LANE001, or just 001 to reuse the current style."
}
elseif ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$')
{
	$changeMachineNameToolTipText += " Server formats support SERVER001, SERVER-001, or 0231SERVER001."
}
else
{
	$changeMachineNameToolTipText += " Accepted formats depend on node role and are shown in the rename dialog."
}
$toolTip.SetToolTip($changeMachineNameButton, $changeMachineNameToolTipText)
$toolTip.SetToolTip($truncateTablesButton, "Empty the transaction, report and sales tables (TRS_*, RPT_*, SAL_*, REC_*, INV_* ...) - asks first. Cannot be undone.")
$toolTip.SetToolTip($registryCleanupButton, "Remove the 'GT*' counter values under Store Management\Counters (after confirmation).")
$toolTip.SetToolTip($updateSQLDatabaseButton, "Bring TER_TAB, STO_TAB, LNK_TAB, RUN_TAB and STD_TAB in line with this machine's name and store number.")
$toolTip.SetToolTip($summaryButton, "See which operations of this session succeeded, failed or were skipped.")
$toolTip.SetToolTip($rebootButton, "Restart the computer (asks first). Needed after a machine rename.")
