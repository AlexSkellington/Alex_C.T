# ===================================================================================================
#                              FUNCTION: Show_Table_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Dialog listing all tables discovered by Get_Table_Aliases in a check list box, with
#   Select All / Deselect All buttons; returns the checked table names (with _TAB).
#
# Parameters:
#   - AliasResults  [System.Collections.ArrayList] (required): rows from Get_Table_Aliases
#
# Returns:
#   - Array of checked table names ($null when canceled)
# ===================================================================================================

function Show_Table_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[System.Collections.ArrayList]$AliasResults
	)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Tables to Process"
	$form.Size = New-Object System.Drawing.Size(450, 570)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Please select the tables you want to pump:"
	$label.Location = New-Object System.Drawing.Point(10, 10)
	$label.AutoSize = $true
	$form.Controls.Add($label)

	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Location = New-Object System.Drawing.Point(10, 40)
	$checkedListBox.Size = New-Object System.Drawing.Size(410, 400)
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)

	$distinctTables = $AliasResults | Select-Object -ExpandProperty Table -Unique | Sort-Object
	foreach ($tableName in $distinctTables)
	{
		[void]$checkedListBox.Items.Add($tableName, $false)
	}

	# Buttons styled/positioned like Node form (Y=460 and 500 for a 550-height window)
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Location = New-Object System.Drawing.Point(20, 460)
	$btnSelectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnSelectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnSelectAll)

	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Location = New-Object System.Drawing.Point(220, 460)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnDeselectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnDeselectAll)

	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(60, 500)
	$btnOK.Size = New-Object System.Drawing.Size(140, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(220, 500)
	$btnCancel.Size = New-Object System.Drawing.Size(140, 32)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $form }

	$setBtnColor = {
		param ($btn,
			$state)
		switch ($state)
		{
			1 { $btn.BackColor = [System.Drawing.Color]::Yellow }
			2 { $btn.BackColor = [System.Drawing.Color]::LightGreen }
			Default { $btn.BackColor = [System.Drawing.SystemColors]::Control }
		}
	}

	$isAllChecked = {
		for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
		{
			if (-not $checkedListBox.GetItemChecked($i)) { return $false }
		}
		return $checkedListBox.Items.Count -gt 0
	}
	$isAnyChecked = {
		for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
		{
			if ($checkedListBox.GetItemChecked($i)) { return $true }
		}
		return $false
	}

	$btnSelectAll.Add_Click({
			$allChecked = & $isAllChecked
			if (-not $allChecked)
			{
				for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
				{
					$checkedListBox.SetItemChecked($i, $true)
				}
				& $setBtnColor $btnSelectAll 2
			}
		})
	$btnDeselectAll.Add_Click({
			for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++)
			{
				$checkedListBox.SetItemChecked($i, $false)
			}
			& $setBtnColor $btnSelectAll 0
		})

	$checkedListBox.Add_ItemCheck({
			Start-Sleep -Milliseconds 30
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	$checkedListBox.Add_MouseUp({
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})
	$checkedListBox.Add_KeyUp({
			$allChecked = & $isAllChecked
			$anyChecked = & $isAnyChecked
			if ($allChecked)
			{
				& $setBtnColor $btnSelectAll 2
			}
			elseif ($anyChecked)
			{
				& $setBtnColor $btnSelectAll 1
			}
			else
			{
				& $setBtnColor $btnSelectAll 0
			}
		})

	$dialogResult = $form.ShowDialog()
	if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
	{
		$selectedTables = @()
		foreach ($item in $checkedListBox.CheckedItems)
		{
			$selectedTables += $item
		}
		return $selectedTables
	}
	else
	{
		return $null
	}
}
