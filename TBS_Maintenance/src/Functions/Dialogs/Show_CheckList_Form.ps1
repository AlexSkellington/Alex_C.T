# ===================================================================================================
#                              FUNCTION: Show_CheckList_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generic modal check-list dialog: a prompt, a CheckedListBox of items, Select All /
#   Deselect All and OK / Cancel. The Select All button turns yellow when some items are
#   checked and green when all are. Used by Show_Table_Selection_Form and
#   Show_Section_Selection_Form.
#
# Parameters:
#   - Title   [string]: window title
#   - Prompt  [string]: line above the list
#   - Items   [string[]]: entries to list (in the given order, none checked)
#   - Width   [int]: list width in pixels (default 410)
#   - Height  [int]: list height in pixels (default 400)
#
# Returns:
#   Array of the checked items, or $null when cancelled.
# ===================================================================================================

function Show_CheckList_Form
{
	param (
		[string]$Title = 'Select',
		[string]$Prompt = '',
		[string[]]$Items = @(),
		[int]$Width = 410,
		[int]$Height = 400
	)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = $Title
	$form.ClientSize = New-Object System.Drawing.Size(($Width + 20), ($Height + 140))
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	$label = New-Object System.Windows.Forms.Label
	$label.Text = $Prompt
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$label.AutoSize = $true
	$form.Controls.Add($label)

	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 40)
	$checkedListBox.Size = New-Object System.Drawing.Size($Width, $Height)
	$checkedListBox.CheckOnClick = $true
	foreach ($item in $Items) { [void]$checkedListBox.Items.Add($item, $false) }
	$form.Controls.Add($checkedListBox)

	# Two button rows, centred under the list
	$rowLeft = [int](($Width + 20 - 380) / 2)
	$newButton = {
		param ([string]$Text, [int]$X, [int]$Y, [int]$W)
		$b = New-Object System.Windows.Forms.Button
		$b.Text = $Text
		$b.Location = New-Object System.Drawing.Point($X, $Y)
		$b.Size = New-Object System.Drawing.Size($W, 32)
		$b.BackColor = [System.Drawing.SystemColors]::Control
		$form.Controls.Add($b)
		return $b
	}
	$btnSelectAll = & $newButton 'Select All' $rowLeft ($Height + 60) 180
	$btnDeselectAll = & $newButton 'Deselect All' ($rowLeft + 200) ($Height + 60) 180
	$btnOK = & $newButton 'OK' ($rowLeft + 40) ($Height + 100) 140
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$btnCancel = & $newButton 'Cancel' ($rowLeft + 200) ($Height + 100) 140
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $form }

	# Select All colour: control (none) / yellow (some) / green (all)
	$updateSelectAll = {
		$checked = $checkedListBox.CheckedItems.Count
		$btnSelectAll.BackColor = if ($checked -gt 0 -and $checked -eq $checkedListBox.Items.Count) { [System.Drawing.Color]::LightGreen }
		elseif ($checked -gt 0) { [System.Drawing.Color]::Yellow }
		else { [System.Drawing.SystemColors]::Control }
	}
	$setAll = {
		param ([bool]$State)
		for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) { $checkedListBox.SetItemChecked($i, $State) }
		& $updateSelectAll
	}

	$btnSelectAll.Add_Click({ & $setAll $true })
	$btnDeselectAll.Add_Click({ & $setAll $false })
	# (ItemCheck fires before the state changes - MouseUp / KeyUp see the new state)
	$checkedListBox.Add_MouseUp({ & $updateSelectAll })
	$checkedListBox.Add_KeyUp({ & $updateSelectAll })

	try
	{
		if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
		return @($checkedListBox.CheckedItems | ForEach-Object { [string]$_ })
	}
	finally { $form.Dispose() }
}
