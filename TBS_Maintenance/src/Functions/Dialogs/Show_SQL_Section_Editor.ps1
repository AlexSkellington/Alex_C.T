# ===================================================================================================
#                              FUNCTION: Show_SQL_Section_Editor
# ---------------------------------------------------------------------------------------------------
# Description:
#   Dialog to pick, edit and add SQL sections before they run: a check list of the sections on the
#   left, the selected section's SQL (editable) on the right, plus custom SQL sections of your own.
#   Edits to the built-in sections apply to this run only; custom sections are saved on this PC
#   and offered again next time.
#
# Parameters:
#   - Sections   [object[]] (required): objects with Name and Sql, in run order
#   - Title      [string]: window title
#   - CustomFile [string]: JSON file for the custom sections (none = custom sections not saved)
#
# Returns:
#   The checked sections, in list order, as objects with Name and Sql (edited text) - $null when
#   cancelled or nothing was checked.
#
# Notes:
#   - Click handlers run in their own scope: all state lives in the hashtable on the form's .Tag
#     (assigning a plain variable inside a handler would only set a local copy).
# ===================================================================================================

function Show_SQL_Section_Editor
{
	param (
		[Parameter(Mandatory = $true)]
		[object[]]$Sections,
		[string]$Title = 'SQL Sections',
		[string]$CustomFile
	)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$toCrLf = { param ([string]$s) if ($null -eq $s) { '' } else { ($s -replace "`r?`n", "`r`n").Trim() } }

	$items = New-Object System.Collections.Generic.List[object]
	foreach ($s in $Sections)
	{
		$sql = & $toCrLf ([string]$s.Sql)
		$items.Add([pscustomobject]@{ Name = [string]$s.Name; Sql = $sql; Original = $sql; Custom = $false })
	}
	# Custom sections saved earlier on this PC
	if ($CustomFile -and (Test-Path -LiteralPath $CustomFile))
	{
		try
		{
			foreach ($c in @(Get-Content -LiteralPath $CustomFile -Raw -ErrorAction Stop | ConvertFrom-Json))
			{
				if ($c -and $c.Name) { $items.Add([pscustomobject]@{ Name = [string]$c.Name; Sql = (& $toCrLf ([string]$c.Sql)); Original = ''; Custom = $true }) }
			}
		}
		catch { try { Write_Log ("Saved custom SQL could not be read (" + $_.Exception.Message + "): " + $CustomFile) "yellow" } catch { } }
	}

	$editorForm = New-Object System.Windows.Forms.Form
	$editorForm.Text = $Title
	$editorForm.ClientSize = New-Object System.Drawing.Size(940, 590)
	$editorForm.MinimumSize = New-Object System.Drawing.Size(760, 480)
	$editorForm.StartPosition = "CenterScreen"
	$editorForm.ShowInTaskbar = $false
	$editorForm.Tag = @{ Items = $items; Current = -1; CustomFile = $CustomFile }

	$prompt = New-Object System.Windows.Forms.Label
	$prompt.Text = "Tick the sections to run. Click a section to see or edit its SQL. Add Custom SQL for your own queries."
	$prompt.Location = New-Object System.Drawing.Point(12, 10)
	$prompt.AutoSize = $true
	$editorForm.Controls.Add($prompt)

	$list = New-Object System.Windows.Forms.CheckedListBox
	$list.Location = New-Object System.Drawing.Point(12, 38)
	$list.Size = New-Object System.Drawing.Size(320, 440)
	$list.Anchor = 'Top,Bottom,Left'
	$list.CheckOnClick = $true
	$list.IntegralHeight = $false
	foreach ($it in $items) { [void]$list.Items.Add($(if ($it.Custom) { '[Custom] ' + $it.Name } else { $it.Name }), $false) }
	$editorForm.Controls.Add($list)

	$sqlLabel = New-Object System.Windows.Forms.Label
	$sqlLabel.Text = "Select a section on the left"
	$sqlLabel.Location = New-Object System.Drawing.Point(344, 38)
	$sqlLabel.Size = New-Object System.Drawing.Size(584, 20)
	$sqlLabel.Anchor = 'Top,Left,Right'
	$sqlLabel.AutoEllipsis = $true
	$editorForm.Controls.Add($sqlLabel)

	$sqlBox = New-Object System.Windows.Forms.TextBox
	$sqlBox.Multiline = $true
	$sqlBox.AcceptsReturn = $true
	$sqlBox.AcceptsTab = $true
	$sqlBox.WordWrap = $false
	$sqlBox.ScrollBars = 'Both'
	$sqlBox.MaxLength = 0
	$sqlBox.Font = New-Object System.Drawing.Font("Consolas", 10)
	$sqlBox.Location = New-Object System.Drawing.Point(344, 60)
	$sqlBox.Size = New-Object System.Drawing.Size(584, 418)
	$sqlBox.Anchor = 'Top,Bottom,Left,Right'
	$sqlBox.Enabled = $false
	# The popup theme timer sets Segoe UI on every text box - SQL reads better monospaced
	$sqlBox.Add_FontChanged({ param ($s, $e) if ($s.Font.Name -ne 'Consolas') { $s.Font = New-Object System.Drawing.Font("Consolas", 10) } })
	$editorForm.Controls.Add($sqlBox)

	$newButton = {
		param ([string]$Text, [int]$X, [int]$Y, [int]$W, [string]$Anchor)
		$b = New-Object System.Windows.Forms.Button
		$b.Text = $Text
		$b.Location = New-Object System.Drawing.Point($X, $Y)
		$b.Size = New-Object System.Drawing.Size($W, 32)
		$b.Anchor = $Anchor
		$editorForm.Controls.Add($b)
		return $b
	}
	$btnSelectAll = & $newButton 'Select All' 12 488 154 'Bottom,Left'
	$btnDeselectAll = & $newButton 'Deselect All' 178 488 154 'Bottom,Left'
	$btnAdd = & $newButton 'Add Custom SQL' 12 526 154 'Bottom,Left'
	$btnRemove = & $newButton 'Remove Custom' 178 526 154 'Bottom,Left'
	$btnReset = & $newButton 'Undo My Edits' 344 488 154 'Bottom,Left'
	$btnOK = & $newButton 'Run Checked' 666 546 128 'Bottom,Right'
	$btnCancel = & $newButton 'Cancel' 800 546 128 'Bottom,Right'
	$btnRemove.Enabled = $false
	$btnReset.Enabled = $false

	$hint = New-Object System.Windows.Forms.Label
	$hint.Text = "Edits to the built-in sections apply to this run only. Custom SQL is saved on this PC. GO lines split batches."
	$hint.Location = New-Object System.Drawing.Point(510, 494)
	$hint.Size = New-Object System.Drawing.Size(418, 40)
	$hint.Anchor = 'Bottom,Left,Right'
	$hint.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
	$editorForm.Controls.Add($hint)

	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$editorForm.AcceptButton = $null # Enter must make new lines in the SQL box
	$editorForm.CancelButton = $btnCancel

	# Keep the text of the section being left, show the new one
	$list.Add_SelectedIndexChanged({
			$st = $editorForm.Tag
			if ($st.Current -ge 0 -and $st.Current -lt $st.Items.Count) { $st.Items[$st.Current].Sql = $sqlBox.Text }
			$i = $list.SelectedIndex
			$st.Current = $i
			if ($i -lt 0) { $sqlBox.Enabled = $false; return }
			$it = $st.Items[$i]
			$sqlBox.Enabled = $true
			$sqlBox.Text = $it.Sql
			$sqlLabel.Text = $(if ($it.Custom) { "Custom SQL: " + $it.Name } else { "SQL for: " + $it.Name })
			$btnRemove.Enabled = [bool]$it.Custom
			$btnReset.Enabled = -not $it.Custom
		})
	$btnReset.Add_Click({
			$st = $editorForm.Tag
			if ($st.Current -ge 0) { $sqlBox.Text = $st.Items[$st.Current].Original }
		})
	$btnSelectAll.Add_Click({ for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) } })
	$btnDeselectAll.Add_Click({ for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) } })
	$btnAdd.Add_Click({
			$st = $editorForm.Tag
			$n = 1
			while (@($st.Items | Where-Object { $_.Name -eq ("Custom SQL " + $n) }).Count -gt 0) { $n++ }
			$st.Items.Add([pscustomobject]@{ Name = ("Custom SQL " + $n); Sql = "-- Your SQL here. It runs against the store database.`r`n"; Original = ''; Custom = $true })
			$ix = $list.Items.Add(("[Custom] Custom SQL " + $n), $true)
			$list.SelectedIndex = $ix
			$sqlBox.Focus()
			$sqlBox.SelectAll()
		})
	$btnRemove.Add_Click({
			$st = $editorForm.Tag
			$i = $st.Current
			if ($i -lt 0 -or -not $st.Items[$i].Custom) { return }
			$st.Current = -1
			$st.Items.RemoveAt($i)
			$list.Items.RemoveAt($i)
			$sqlBox.Text = ''
			$sqlBox.Enabled = $false
			$sqlLabel.Text = "Select a section on the left"
			$btnRemove.Enabled = $false
			$btnReset.Enabled = $false
		})

	try
	{
		if ($editorForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
		$st = $editorForm.Tag
		if ($st.Current -ge 0 -and $st.Current -lt $st.Items.Count) { $st.Items[$st.Current].Sql = $sqlBox.Text }

		# Save the custom sections (checked or not) for next time
		if ($CustomFile)
		{
			try
			{
				$dir = Split-Path -Parent $CustomFile
				if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
				$custom = @($st.Items | Where-Object { $_.Custom -and -not [string]::IsNullOrWhiteSpace($_.Sql) } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Sql = $_.Sql } })
				if ($custom.Count -gt 0) { [System.IO.File]::WriteAllText($CustomFile, (ConvertTo-Json -InputObject $custom -Depth 3), [System.Text.Encoding]::UTF8) }
				elseif (Test-Path -LiteralPath $CustomFile) { Remove-Item -LiteralPath $CustomFile -Force -ErrorAction SilentlyContinue }
			}
			catch { try { Write_Log ("Custom SQL could not be saved: " + $_.Exception.Message) "yellow" } catch { } }
		}

		$chosen = New-Object System.Collections.Generic.List[object]
		for ($i = 0; $i -lt $list.Items.Count; $i++)
		{
			if (-not $list.GetItemChecked($i)) { continue }
			$it = $st.Items[$i]
			if ([string]::IsNullOrWhiteSpace($it.Sql)) { continue }
			$chosen.Add([pscustomobject]@{ Name = $(if ($it.Custom) { '[Custom] ' + $it.Name } else { $it.Name }); Sql = $it.Sql; Edited = ((-not $it.Custom) -and ($it.Sql -ne $it.Original)) })
		}
		if ($chosen.Count -eq 0) { return $null }
		return , $chosen.ToArray()
	}
	finally { $editorForm.Dispose() }
}
