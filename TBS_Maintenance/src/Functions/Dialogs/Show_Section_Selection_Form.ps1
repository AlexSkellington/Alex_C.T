# ===================================================================================================
#                              FUNCTION: Show_Section_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Dialog with a check list box of SQL sections (Select All / Deselect All); returns the
#   array of checked section names, or $null when canceled.
#
# Parameters:
#   - SectionNames  [string[]] (required): section names to list
#
# Returns:
#   - Array of checked section names ($null when canceled)
# ===================================================================================================

function Show_Section_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$SectionNames
	)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select SQL Sections"
	$form.StartPosition = "CenterScreen"
	$form.Size = New-Object System.Drawing.Size(550, 440)
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Check the sections you want to run, then click OK."
	$label.AutoSize = $true
	$label.Left = 20
	$label.Top = 10
	$form.Controls.Add($label)
	
	$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
	$checkedListBox.Width = 500
	$checkedListBox.Height = 280
	$checkedListBox.Left = 20
	$checkedListBox.Top = 40
	$checkedListBox.CheckOnClick = $true
	$form.Controls.Add($checkedListBox)
	
	foreach ($name in $SectionNames)
	{
		[void]$checkedListBox.Items.Add($name, $false)
	}
	
	# Place Select All and Deselect All like the Node dialog (Y=310)
	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.Location = New-Object System.Drawing.Point(75, 320)
	$btnSelectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnSelectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnSelectAll)
	
	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.Location = New-Object System.Drawing.Point(275, 320)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnDeselectAll.BackColor = [System.Drawing.SystemColors]::Control
	$form.Controls.Add($btnDeselectAll)
	
	# Place OK/Cancel like Node dialog (Y=350)
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(115, 360)
	$btnOK.Size = New-Object System.Drawing.Size(140, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(275, 360)
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
		$selectedSections = @()
		foreach ($item in $checkedListBox.CheckedItems)
		{
			$selectedSections += $item
		}
		return $selectedSections
	}
	else
	{
		return $null
	}
}
