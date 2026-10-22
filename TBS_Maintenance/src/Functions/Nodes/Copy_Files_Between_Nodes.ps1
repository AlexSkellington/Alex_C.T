# ===================================================================================================
#                              FUNCTION: Copy_Files_Between_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Copies \storeman\ folders/files from a chosen source (Server or Lane) to selected lanes
#   through a GUI with a quick-items menu (Htm, BitMaps, Setting.ini, System.ini, ...).
#   Validates the source path before copying; the lane root is resolved as \\<Lane>\Storeman
#   with a \\<Lane>\C$\storeman fallback. Files and folders are supported.
#
# Variables:
#   - $script:FunctionResults: LaneMachines / LaneNumToMachineName (known lanes)
#   - $BasePath: local server storeman root (source)
#
# Writes:
#   - the copied folders/files on the selected lane stores
# ===================================================================================================

function Copy_Files_Between_Nodes
{
	Write_Log "`r`n==================== Starting Copy_Files_Between_Nodes ====================`r`n" "blue"

	# =========================================================================================
	# CONFIG: Quick items + optional extra server roots
	#   Label   -> text shown in the Quick... context menu (tip: start with '\storeman\...')
	#   Rel     -> relative path under the storeman root (NO leading slash)
	#   Type    -> "Folder" or "File" (informational)
	#   LaneOnly-> $true to enable only when Lane is selected as source
	# =========================================================================================
	$QuickItems = @(
		@{ Label = "\storeman\office\Htm"; Rel = "office\Htm"; Type = "Folder"; LaneOnly = $false },
		@{ Label = "\storeman\BitMaps"; Rel = "BitMaps"; Type = "Folder"; LaneOnly = $false },
		@{ Label = "\storeman\office\Setting.ini"; Rel = "office\Setting.ini"; Type = "File"; LaneOnly = $true },
		@{ Label = "\storeman\office\System.ini"; Rel = "office\System.ini"; Type = "File"; LaneOnly = $true }
		# Add more here...
		# @{ Label = "\storeman\XchDev\ApiVerifoneMX.ini"; Rel="XchDev\ApiVerifoneMX.ini"; Type="File"; LaneOnly=$false }
	)

	# Optional extra server roots to try AFTER $BasePath (if present in any scope)
	$ExtraServerRoots = @(
		# "E:\storeman"
	)

	# ---------------- Known lanes for combo (reads your cached maps if present) ----------------
	$knownLanes = @()
	if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'])
	{
		$knownLanes = $script:FunctionResults['LaneMachines'].Values | Where-Object { $_ } | Select-Object -Unique
	}
	elseif ($script:FunctionResults.ContainsKey('LaneNumToMachineName') -and $script:FunctionResults['LaneNumToMachineName'])
	{
		$knownLanes = $script:FunctionResults['LaneNumToMachineName'].Values | Where-Object { $_ } | Select-Object -Unique
	}

	# ========================= UI - DPI-safe layout =========================
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	[void][System.Windows.Forms.Application]::EnableVisualStyles()

	# --- Form shell ---
	$frm = New-Object System.Windows.Forms.Form
	$frm.Text = "Copy Files Between Nodes - Pick Source and Folders"
	$frm.StartPosition = 'CenterScreen'
	$frm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
	$frm.Font = New-Object System.Drawing.Font('Segoe UI', 9)
	$frm.MinimizeBox = $false
	$frm.MaximizeBox = $false
	$frm.FormBorderStyle = 'FixedDialog'
	$frm.Size = New-Object System.Drawing.Size(760, 560)
	$frm.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)

	# --- Main layout (header, source, items, bottom) ---
	$layoutMain = New-Object System.Windows.Forms.TableLayoutPanel
	$layoutMain.Dock = 'Fill'
	$layoutMain.Padding = New-Object System.Windows.Forms.Padding(12)
	$layoutMain.ColumnCount = 1
	$layoutMain.RowCount = 4
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$frm.Controls.Add($layoutMain)

	# Header
	$lbl = New-Object System.Windows.Forms.Label
	$lbl.Text = "Choose the source and which folder(s)/file(s) to copy to the selected lanes:"
	$lbl.Font = New-Object System.Drawing.Font("Segoe UI", [float]11, [System.Drawing.FontStyle]::Bold)
	$lbl.AutoSize = $true
	$lbl.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
	$layoutMain.Controls.Add($lbl, 0, 0)

	# ---------------- Source group ----------------
	$grpSrc = New-Object System.Windows.Forms.GroupBox
	$grpSrc.Text = "Source"
	$grpSrc.Dock = 'Top'
	$grpSrc.AutoSize = $true
	$grpSrc.AutoSizeMode = 'GrowAndShrink'
	$grpSrc.Padding = New-Object System.Windows.Forms.Padding(12)
	$grpSrc.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
	$layoutMain.Controls.Add($grpSrc, 0, 1)

	# Source row: Server (alone), then Lane + combo same row
	$layoutSrc = New-Object System.Windows.Forms.TableLayoutPanel
	$layoutSrc.Dock = 'Top'
	$layoutSrc.AutoSize = $true
	$layoutSrc.AutoSizeMode = 'GrowAndShrink'
	$layoutSrc.ColumnCount = 3
	$layoutSrc.RowCount = 2
	[void]$layoutSrc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutSrc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutSrc.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutSrc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutSrc.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$grpSrc.Controls.Add($layoutSrc)

	$rbServer = New-Object System.Windows.Forms.RadioButton
	$rbServer.Text = "Server (this computer)"
	$rbServer.Checked = $true
	$rbServer.AutoSize = $true
	$layoutSrc.Controls.Add($rbServer, 0, 0)
	$layoutSrc.SetColumnSpan($rbServer, 3)

	$rbLane = New-Object System.Windows.Forms.RadioButton
	$rbLane.Text = "Lane (source)"
	$rbLane.AutoSize = $true
	$layoutSrc.Controls.Add($rbLane, 0, 1)

	$lblLane = New-Object System.Windows.Forms.Label
	$lblLane.Text = "Select lane:"
	$lblLane.AutoSize = $true
	$lblLane.Margin = New-Object System.Windows.Forms.Padding(12, 3, 6, 0)
	$layoutSrc.Controls.Add($lblLane, 1, 1)

	$cboLane = New-Object System.Windows.Forms.ComboBox
	$cboLane.DropDownStyle = 'DropDownList'
	$cboLane.Enabled = $false
	$cboLane.Dock = 'Top'
	$layoutSrc.Controls.Add($cboLane, 2, 1)

	$knownLanes | Sort-Object | ForEach-Object { [void]$cboLane.Items.Add($_) }
	$rbLane.Add_CheckedChanged({
			$cboLane.Enabled = $rbLane.Checked
			if ($rbLane.Checked -and $cboLane.Items.Count -gt 0 -and -not $cboLane.SelectedItem) { $cboLane.SelectedIndex = 0 }
		})

	# ---------------- Item picker group ----------------
	$grpItems = New-Object System.Windows.Forms.GroupBox
	$grpItems.Text = "Folder(s)/File(s) to Copy"
	$grpItems.Dock = 'Fill'
	$grpItems.Padding = New-Object System.Windows.Forms.Padding(12)
	$grpItems.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
	$layoutMain.Controls.Add($grpItems, 0, 2)

	$layoutItems = New-Object System.Windows.Forms.TableLayoutPanel
	$layoutItems.Dock = 'Fill'
	$layoutItems.ColumnCount = 2
	$layoutItems.RowCount = 2
	[void]$layoutItems.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutItems.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$layoutItems.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$layoutItems.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$grpItems.Controls.Add($layoutItems)

	$clb = New-Object System.Windows.Forms.CheckedListBox
	$clb.CheckOnClick = $true
	$clb.Dock = 'Fill'
	$clb.HorizontalScrollbar = $true # show long UNC paths completely
	$layoutItems.Controls.Add($clb, 0, 0)

	$pnlBtns = New-Object System.Windows.Forms.FlowLayoutPanel
	$pnlBtns.FlowDirection = 'TopDown'
	$pnlBtns.WrapContents = $false
	$pnlBtns.AutoSize = $true
	$pnlBtns.Dock = 'Top'
	$pnlBtns.Padding = New-Object System.Windows.Forms.Padding(0)
	$pnlBtns.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
	$layoutItems.Controls.Add($pnlBtns, 1, 0)

	# ---- Add button now says "Add Folder/Files..." and shows a small context menu ----
	$btnAdd = New-Object System.Windows.Forms.Button
	$btnAdd.Text = "Add Folder/Files..."
	$btnAdd.Width = 160; $btnAdd.Height = 30
	$pnlBtns.Controls.Add($btnAdd)

	$btnQuick = New-Object System.Windows.Forms.Button
	$btnQuick.Text = "Quick..."
	$btnQuick.Width = 160; $btnQuick.Height = 30
	$btnQuick.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
	$pnlBtns.Controls.Add($btnQuick)

	$btnRemove = New-Object System.Windows.Forms.Button
	$btnRemove.Text = "Remove"
	$btnRemove.Width = 160; $btnRemove.Height = 30
	$btnRemove.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
	$pnlBtns.Controls.Add($btnRemove)

	$chkMirror = New-Object System.Windows.Forms.CheckBox
	$chkMirror.Text = "Mirror (delete extras on targets)"
	$chkMirror.AutoSize = $true
	$chkMirror.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
	$layoutItems.Controls.Add($chkMirror, 0, 1)

	$toolTip = New-Object System.Windows.Forms.ToolTip
	$toolTip.SetToolTip($chkMirror, "Robocopy /MIR will delete files on the destination that are not present in the source.")

	# ---------------- Bottom buttons ----------------
	$pnlBottom = New-Object System.Windows.Forms.TableLayoutPanel
	$pnlBottom.AutoSize = $true
	$pnlBottom.ColumnCount = 2
	$pnlBottom.RowCount = 1
	[void]$pnlBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$pnlBottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$pnlBottom.Dock = 'Right'
	$layoutMain.Controls.Add($pnlBottom, 0, 3)

	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Width = 110; $btnOK.Height = 32
	$btnOK.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
	$pnlBottom.Controls.Add($btnOK, 0, 0)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Width = 110; $btnCancel.Height = 32
	$btnCancel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
	$pnlBottom.Controls.Add($btnCancel, 1, 0)

	$frm.AcceptButton = $btnOK
	$frm.CancelButton = $btnCancel

	# ===================== QUICK CONTEXT MENU (built once; no auto-reopen) =====================
	$cmsQuick = New-Object System.Windows.Forms.ContextMenuStrip
	foreach ($qi in $QuickItems)
	{
		if (-not $qi) { continue }

		$mi = New-Object System.Windows.Forms.ToolStripMenuItem
		$mi.Text = $qi.Label
		$mi.Tag = $qi # keep metadata with the menu item

		# Click handler MUST use param($s,$e); $s = sender (the menu item)
		$mi.Add_Click({
				param ($s,
					$e)

				# Storeman root of the selected source (shows its own message when unreachable)
				$root = & $resolveInitialRoot
				if (-not $root) { return }
				# Build full path safely (trim any leading slash in Rel so Join-Path doesn't drop root)
				$meta = $s.Tag
				$rel = ([string]$meta.Rel).Trim().TrimStart('\', '/')
				$full = [System.IO.Path]::Combine($root, $rel)

				# Verify exists BEFORE adding
				if (-not (Test-Path -LiteralPath $full))
				{
					[System.Windows.Forms.MessageBox]::Show("That item does not exist on the selected source:`r`n$full",
						"Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
					Write_Log "[Source Missing] $full not found; not added to list." "yellow"
					return
				}

				# Add real full path (avoid duplicates)
				if (-not ($clb.Items -contains $full))
				{
					[void]$clb.Items.Add($full, $true)
					Write_Log "Added item: $full" "gray"
				}
			})

		[void]$cmsQuick.Items.Add($mi)
	}

	# On open: toggle LaneOnly items enabled/disabled, then show menu (no auto-reopen)
	$btnQuick.Add_Click({
			foreach ($mi in $cmsQuick.Items)
			{
				if ($mi -is [System.Windows.Forms.ToolStripMenuItem])
				{
					$meta = $mi.Tag
					if ($meta -and $meta.ContainsKey('LaneOnly') -and $meta.LaneOnly) { $mi.Enabled = $rbLane.Checked }
					else { $mi.Enabled = $true }
				}
			}
			$cmsQuick.Show($btnQuick, 0, $btnQuick.Height)
		})

	# ===================== "Add Folder/Files..." context menu =====================
	$cmsAdd = New-Object System.Windows.Forms.ContextMenuStrip
	$miAddFolder = New-Object System.Windows.Forms.ToolStripMenuItem
	$miAddFolder.Text = "Add Folder..."
	[void]$cmsAdd.Items.Add($miAddFolder)

	$miAddFiles = New-Object System.Windows.Forms.ToolStripMenuItem
	$miAddFiles.Text = "Add File(s)..."
	[void]$cmsAdd.Items.Add($miAddFiles)

	# Storeman root of the selected source (server: $BasePath / extras; lane: \\<lane>\Storeman,
	# then \\<lane>\C$\storeman); shows a message and returns $null when unreachable
	$resolveInitialRoot = {
		$initialPath = $null
		if ($rbServer.Checked)
		{
			$candidates = @($script:BasePath, $global:BasePath, $BasePath) + $ExtraServerRoots
			foreach ($cand in ($candidates | Where-Object { $_ } | Select-Object -Unique))
			{
				if (Test-Path -LiteralPath $cand) { $initialPath = $cand; break }
			}
			if (-not $initialPath)
			{
				[System.Windows.Forms.MessageBox]::Show(
					"Server storeman root not found using `$BasePath (or extras).",
					"Source Unreachable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
				Write_Log "[Source Unreachable] Server storeman root not found via `$BasePath or extras." "red"
				return $null
			}
		}
		else
		{
			if (-not $cboLane.SelectedItem)
			{
				[System.Windows.Forms.MessageBox]::Show("Pick a lane as source first.", "Missing Lane",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				return $null
			}
			$ln = [string]$cboLane.SelectedItem
			$try1 = "\\$ln\Storeman"; if (Test-Path -LiteralPath $try1) { $initialPath = $try1 }
			if (-not $initialPath) { $try2 = "\\$ln\c$\storeman"; if (Test-Path -LiteralPath $try2) { $initialPath = $try2 } }
			if (-not $initialPath)
			{
				[System.Windows.Forms.MessageBox]::Show("\\$ln\Storeman and \\$ln\C$\storeman are not accessible.",
					"Source Unreachable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				Write_Log "[Source Unreachable] $ln has no \\Storeman or \\C$\storeman." "red"
				return $null
			}
		}
		return $initialPath
	}

	# --- Add Folder... (keeps your old logic, just under the new menu) ---
	$miAddFolder.Add_Click({
			$initialPath = & $resolveInitialRoot
			if (-not $initialPath) { return }

			$fd = New-Object System.Windows.Forms.FolderBrowserDialog
			$fd.Description = "Pick a folder under \storeman\ (based on your selected source)."
			$fd.ShowNewFolderButton = $false
			$fd.RootFolder = [System.Environment+SpecialFolder]::Desktop
			$fd.SelectedPath = $initialPath

			$res = $fd.ShowDialog()
			if ($res -ne [System.Windows.Forms.DialogResult]::OK -or -not $fd.SelectedPath) { return }

			if ($fd.SelectedPath -notmatch '(?i)[\\/](storeman)[\\/]')
			{
				[System.Windows.Forms.MessageBox]::Show("Please choose a folder under \storeman\.", "Not Allowed",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				return
			}
			if (-not (Test-Path -LiteralPath $fd.SelectedPath))
			{
				[System.Windows.Forms.MessageBox]::Show("That folder does not exist on the selected source:`r`n$($fd.SelectedPath)",
					"Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
				Write_Log "[Source Missing] $($fd.SelectedPath) not found; not added to list." "yellow"
				return
			}
			if (-not ($clb.Items -contains $fd.SelectedPath))
			{
				[void]$clb.Items.Add($fd.SelectedPath, $true)
				Write_Log "Added item: $($fd.SelectedPath)" "gray"
			}
		})

	# --- Add File(s)... (new) ---
	$miAddFiles.Add_Click({
			$initialPath = & $resolveInitialRoot
			if (-not $initialPath) { return }

			$ofd = New-Object System.Windows.Forms.OpenFileDialog
			$ofd.Title = "Pick file(s) under \storeman\ (based on your selected source)"
			$ofd.InitialDirectory = $initialPath
			$ofd.Multiselect = $true
			$ofd.Filter = "All files (*.*)|*.*"

			$res = $ofd.ShowDialog()
			if ($res -ne [System.Windows.Forms.DialogResult]::OK -or -not $ofd.FileNames -or $ofd.FileNames.Count -eq 0) { return }

			foreach ($f in $ofd.FileNames)
			{
				if ($f -notmatch '(?i)[\\/](storeman)[\\/]')
				{
					[System.Windows.Forms.MessageBox]::Show("Only items under \storeman\ are allowed:`r`n$f",
						"Not Allowed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
					continue
				}
				if (-not (Test-Path -LiteralPath $f))
				{
					[System.Windows.Forms.MessageBox]::Show("That file no longer exists on the selected source:`r`n$f",
						"Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
					Write_Log "[Source Missing] $f not found; not added to list." "yellow"
					continue
				}
				if (-not ($clb.Items -contains $f))
				{
					[void]$clb.Items.Add($f, $true)
					Write_Log "Added item: $f" "gray"
				}
			}
		})

	# Show the mini menu when clicking "Add Folder/Files..." (no re-open behavior)
	$btnAdd.Add_Click({ $cmsAdd.Show($btnAdd, 0, $btnAdd.Height) })

	# Remove entries
	$btnRemove.Add_Click({
			$toRemove = @()
			foreach ($idx in $clb.CheckedIndices) { $toRemove += $idx }
			if (-not $toRemove -and $clb.SelectedIndex -ge 0) { $toRemove = @($clb.SelectedIndex) }
			$toRemove | Sort-Object -Descending | ForEach-Object { $clb.Items.RemoveAt($_) }
		})

	# ---------------- OK / Cancel logic ----------------
	$btnOK.Add_Click({
			$formValid = $true
			if ($rbLane.Checked -and (-not $cboLane.SelectedItem))
			{
				[System.Windows.Forms.MessageBox]::Show("Pick a lane as the source.", "Missing Source",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				$formValid = $false
			}
			if ($clb.CheckedItems.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("Select at least one folder/file to copy.", "Missing Selection",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				$formValid = $false
			}
			if ($formValid -and $chkMirror.Checked)
			{
				$warn = [System.Windows.Forms.MessageBox]::Show(
					"MIRROR mode will delete files in the destination that are not in the source. Continue?",
					"Confirm Mirror",
					[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning
				)
				if ($warn -ne [System.Windows.Forms.DialogResult]::Yes) { $formValid = $false }
			}
			if ($formValid) { $script:__CopyMaps_DialogResult = 'OK'; $frm.Close() }
		})
	$btnCancel.Add_Click({ $script:__CopyMaps_DialogResult = 'Cancel'; $frm.Close() })

	[void]$frm.ShowDialog()
	$dialogResult = $script:__CopyMaps_DialogResult
	Remove-Variable -Name __CopyMaps_DialogResult -Scope Script -ErrorAction SilentlyContinue
	if ($dialogResult -ne 'OK')
	{
		Write_Log "User cancelled source/folder selection." "yellow"
		Write_Log "`r`n==================== Copy_Files_Between_Nodes Function Completed ====================" "blue"
		return
	}

	# ===================== Gather choices =====================
	$useServerAsSource = $rbServer.Checked
	$sourceMachine = if ($useServerAsSource) { $env:COMPUTERNAME }
	else { [string]$cboLane.SelectedItem }
	$itemsToCopy = @(); foreach ($item in $clb.CheckedItems) { $itemsToCopy += [string]$item }
	$doMirror = [bool]$chkMirror.Checked
	Write_Log "Source machine: $sourceMachine  |  Mirror: $doMirror" "cyan"
	Write_Log "Selected item(s): $($itemsToCopy -join ' | ')" "cyan"

	# ===================== Destination selection =====================
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Destination Lanes for Copy" -ExcludedNodes @($sourceMachine)
	if ($null -eq $selection) { Write_Log "Destination selection canceled by user." "yellow"; return }
	$Lanes = @($selection.Lanes)
	if ($Lanes.Count -eq 0) { Write_Log "No destination lanes selected." "yellow"; return }

	# Map lane numbers -> machine names
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "[Mapping Missing] LaneNumToMachineName not available; cannot map lanes to machines." "red"; return
	}

	$destMachines = @()
	foreach ($ln in $Lanes)
	{
		$s = ("$ln").Trim()
		if ($s -match '[A-Za-z]')
		{
			if ($s -match '^(.*?)[\s]*\(\d{3}\)\s*$') { $s = $matches[1].Trim() } # drop "(001)"
			$destMachines += $s; continue
		}
		if ($s -match '^\d+$')
		{
			$pad3 = ('{0:D3}' -f ([int]$s))
			if ($LaneNumToMachineName.ContainsKey($pad3)) { $destMachines += $LaneNumToMachineName[$pad3] }
			elseif ($LaneNumToMachineName.ContainsKey($s)) { $destMachines += $LaneNumToMachineName[$s] }
			else
			{
				$guess = $knownLanes | Where-Object { "$_" -like "*$pad3" } | Select-Object -First 1
				if ($guess) { $destMachines += $guess }
				else { Write_Log "[Unmapped Lane] No machine for lane $pad3." "yellow" }
			}
			continue
		}
		if ($s -match '\((\d{3})\)$' -or $s -match '(\d{3})$')
		{
			$code = $matches[1]
			if ($LaneNumToMachineName.ContainsKey($code)) { $destMachines += $LaneNumToMachineName[$code] }
			else
			{
				$guess = $knownLanes | Where-Object { "$_" -like "*$code" } | Select-Object -First 1
				if ($guess) { $destMachines += $guess }
				else { Write_Log "[Unmapped Lane] No machine for lane $code." "yellow" }
			}
			continue
		}
		Write_Log "[Parse Warning] Could not interpret destination entry: '$s'." "yellow"
	}
	$destMachines = @($destMachines | Where-Object { $_ -and ($_ -ne $sourceMachine) } | Select-Object -Unique)
	if ($destMachines.Count -eq 0)
	{
		[System.Windows.Forms.MessageBox]::Show("No non-source lanes selected. Pick at least one destination lane.", "No Destinations",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
		Write_Log "No destination lanes selected." "yellow"; return
	}
	Write_Log "Destinations ($($destMachines.Count)): $($destMachines -join ', ')" "gray"

	# ===================== Validate destination roots =====================
	$destRootMap = @{ }
	$unreachable = @()
	foreach ($d in $destMachines)
	{
		$root = $null
		$try1 = "\\$d\Storeman"; if (Test-Path -LiteralPath $try1) { $root = $try1 }
		if (-not $root) { $try2 = "\\$d\c$\storeman"; if (Test-Path -LiteralPath $try2) { $root = $try2 } }
		if ($root) { $destRootMap[$d] = $root }
		else { $unreachable += $d }
	}
	if ($unreachable.Count -gt 0)
	{
		Write_Log "Unreachable lane roots (no \\Storeman or \\C$\storeman): $($unreachable -join ', ')" "red"
		$destMachines = @($destMachines | Where-Object { $destRootMap.ContainsKey($_) })
		if ($destMachines.Count -eq 0) { Write_Log "No reachable destinations remain. Aborting." "red"; return }
	}

	# ===================== Copy loop (folder OR single file) =====================
	$failList = New-Object System.Collections.Generic.List[string]
	$okCount = 0

	foreach ($picked in $itemsToCopy)
	{
		# relative path under \storeman\
		$rel = $null
		$m = [regex]::Match($picked, '(?i)[\\/](storeman)[\\/](?<sub>.*)$')
		if ($m.Success) { $rel = $m.Groups['sub'].Value }
		else { Write_Log "[Invalid Path] '$picked' is not under \storeman\." "red"; $failList.Add("$sourceMachine :: $picked (not under \storeman\)"); continue }

		# resolve source full path again (validate)
		if ($useServerAsSource)
		{
			if (-not (Test-Path -LiteralPath $picked)) { Write_Log "[Source Missing] $picked not found on server." "red"; $failList.Add("$sourceMachine :: $picked (missing)"); continue }
			$srcFull = $picked
		}
		else
		{
			$srcRoot = $null
			$tryS1 = "\\$sourceMachine\Storeman"; if (Test-Path -LiteralPath $tryS1) { $srcRoot = $tryS1 }
			if (-not $srcRoot) { $tryS2 = "\\$sourceMachine\c$\storeman"; if (Test-Path -LiteralPath $tryS2) { $srcRoot = $tryS2 } }
			if (-not $srcRoot) { Write_Log "[Source Root Missing] No accessible storeman root on $sourceMachine." "red"; $failList.Add("$sourceMachine :: (no root)"); continue }
			$srcFull = Join-Path $srcRoot $rel
			if (-not (Test-Path -LiteralPath $srcFull)) { Write_Log "[Source Missing] $srcFull not found on lane." "red"; $failList.Add("$sourceMachine :: $srcFull (missing)"); continue }
		}

		$isFile = $false
		try { $isFile = Test-Path -LiteralPath $srcFull -PathType Leaf }
		catch { $isFile = $false }

		foreach ($dest in $destMachines)
		{
			$dstRoot = $destRootMap[$dest]
			if (-not $dstRoot) { $failList.Add("$dest :: (no root)"); continue }

			if ($isFile)
			{
				# Single file copy (no /MIR)
				$dstFull = Join-Path $dstRoot $rel
				$dstDir = Split-Path -Path $dstFull -Parent
				try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
				catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $failList.Add("$dest :: $dstDir (create failed)"); continue }

				$srcDir = Split-Path -Path $srcFull -Parent
				$fileName = Split-Path -Path $srcFull -Leaf
				$null = & robocopy @($srcDir, $dstDir, $fileName, '/R:2', '/W:2', '/NFL', '/NDL', '/NP')
				$exit = $LASTEXITCODE
				$failed = ($exit -ge 8)
				$status = if ($failed) { "FAIL ($exit)" }
				elseif ($exit -ge 4) { "OK* ($exit)" }
				else { "OK ($exit)" }
				if ($failed) { Write_Log "Copy file to $dest FAILED: $status" "red"; $failList.Add("$dest :: $rel (robocopy $exit)") }
				else { Write_Log "Copied file '$rel' to $dest $status" "green"; $okCount++ }
				if ($doMirror) { Write_Log "Note: Mirror ignored for single-file item '$rel'." "gray" }
			}
			else
			{
				# Directory copy
				$dstFull = Join-Path $dstRoot $rel
				try { if (-not (Test-Path -LiteralPath $dstFull)) { New-Item -ItemType Directory -Path $dstFull -Force | Out-Null } }
				catch { Write_Log "[Create Failed] $dstFull : $($_.Exception.Message)" "red"; $failList.Add("$dest :: $dstFull (create failed)"); continue }

				$roboArgs = @($srcFull, $dstFull, '/E', '/COPY:DAT', '/R:2', '/W:2', '/NFL', '/NDL', '/NP', '/MT:8')
				if ($doMirror) { $roboArgs += '/MIR' }
				$null = & robocopy @roboArgs
				$exit = $LASTEXITCODE
				$failed = ($exit -ge 8)
				$status = if ($failed) { "FAIL ($exit)" }
				elseif ($exit -ge 4) { "OK* ($exit)" }
				else { "OK ($exit)" }
				if ($failed) { Write_Log "Copy to $dest FAILED: $status" "red"; $failList.Add("$dest :: $rel (robocopy $exit)") }
				else { Write_Log "Copy to $dest $status" "green"; $okCount++ }
			}
		}
	}

	# ===================== Summary =====================
	$totalOps = ($itemsToCopy.Count) * ($destMachines.Count)
	$failCount = $failList.Count
	Write_Log "`r`n-------------------- Copy Files Summary --------------------" "blue"
	Write_Log "Operations attempted : $totalOps" "gray"
	Write_Log "Successful           : $okCount" "green"
	if ($failCount -gt 0)
	{
		Write_Log "Failed               : $failCount" "red"
		foreach ($f in $failList) { Write_Log "  - $f" "red" }
	}
	else
	{
		Write_Log "Failed               : 0" "green"
	}
	Write_Log "------------------------------------------------------------" "blue"
	Write_Log "`r`n==================== Copy_Files_Between_Nodes Function Completed ====================" "blue"
}
