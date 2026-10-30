# Apply consistent styling + layout behavior for all main action buttons
$mainButtons = @(
	$updateStoreNumberButton, $configureNetworkButton, $changeMachineNameButton,
	$truncateTablesButton, $registryCleanupButton, $updateSQLDatabaseButton,
	$summaryButton, $rebootButton
)

foreach ($btn in $mainButtons)
{
	$btn.Dock = 'Fill'
	$btn.Margin = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
	$btn.MinimumSize = New-Object System.Drawing.Size(0, 42)
	$btn.UseVisualStyleBackColor = $true
}

# Tooltips (one per button)
$toolTip.SetToolTip($updateStoreNumberButton, "Update the store number (startup.ini) and refresh the on-screen store info. If supported, also syncs related SQL/INI values.")
$toolTip.SetToolTip($configureNetworkButton, "Configure NIC settings (IP/Subnet/Gateway/DNS) using store/lane conventions and validate the current network configuration.")
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
$toolTip.SetToolTip($truncateTablesButton, "Run SQL cleanup by truncating selected tables (maintenance routine).")
$toolTip.SetToolTip($registryCleanupButton, "Remove 'GT' registry values (after confirmation) to clean up old configuration leftovers.")
$toolTip.SetToolTip($updateSQLDatabaseButton, "Update SQL records to reflect the new machine name/store number for this node.")
$toolTip.SetToolTip($summaryButton, "View a session summary of which operations succeeded, failed, or were skipped.")
$toolTip.SetToolTip($rebootButton, "Restart the computer (optional cleanup before reboot).")

# Add buttons to the Actions grid (3 columns x 3 rows)
# Row 1
$actionsLayout.Controls.Add($changeMachineNameButton, 0, 0)
$actionsLayout.Controls.Add($configureNetworkButton,   1, 0)
$actionsLayout.Controls.Add($updateStoreNumberButton, 2, 0)

# Row 2 (middle cell intentionally left empty to preserve the classic 3-column layout)
$actionsLayout.Controls.Add($truncateTablesButton,     0, 1)
$actionsLayout.Controls.Add($registryCleanupButton,    2, 1)

# Row 3
$actionsLayout.Controls.Add($updateSQLDatabaseButton,  0, 2)
$actionsLayout.Controls.Add($summaryButton,            1, 2)
$actionsLayout.Controls.Add($rebootButton,             2, 2)
