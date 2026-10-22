# ===================================================================================================
#                              FUNCTION: INI_Editor
# ---------------------------------------------------------------------------------------------------
# Description:
#   GUI INI editor for files under \storeman\<RelativeIniPath>: pick Server/Lane as source,
#   load the INI, select a section, edit all keys (add/edit/delete/reorder) and save. After
#   saving, the INI can be deployed to other lanes either as a merge of the edited section
#   or as a full file replacement. Copy stays disabled until a Save.
#
# Parameters:
#   - RelativeIniPath    [string]: INI path under the storeman root (default office\Setup.ini)
#   - PredefinedIniPaths [string[]]: entries for the source combo
#
# Variables:
#   - $script:FunctionResults: LaneMachines / LaneNumToMachineName (known lanes)
#   - $script:StoremanRoots: candidate local server storeman roots
#
# Writes:
#   - the edited INI (save) and the deployed copy/merge on the selected lanes
# ===================================================================================================

function INI_Editor
{
	param (
		[string]$RelativeIniPath = 'office\Setup.ini',
		[string[]]$PredefinedIniPaths = @('office\Setup.ini', 'office\Setting.ini', 'office\System.ini', 'XchDev\ApiVerifoneMX.ini')
	)
	
	Write_Log "`r`n==================== Starting INI_Editor ====================`r`n" "blue"
	
	# ==============================================================================================
	#                                    HELPERS: paths, parse, write
	# ==============================================================================================
	
	$getKnownLanes = {
		$known = @()
		if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'])
		{ $known = $script:FunctionResults['LaneMachines'].Values | Where-Object { $_ } | Select-Object -Unique }
		elseif ($script:FunctionResults.ContainsKey('LaneNumToMachineName') -and $script:FunctionResults['LaneNumToMachineName'])
		{ $known = $script:FunctionResults['LaneNumToMachineName'].Values | Where-Object { $_ } | Select-Object -Unique }
		return, ($known | Sort-Object)
	}
	
	$getLaneRoot = {
		param ([string]$ComputerName)
		$s1 = "\\$ComputerName\Storeman"; if (Test-Path -LiteralPath $s1) { return $s1 }
		$s2 = "\\$ComputerName\c$\storeman"; if (Test-Path -LiteralPath $s2) { return $s2 }
		return $null
	}
	
	$getServerRoot = {
		foreach ($root in @($script:StoremanRoots))
		{
			if (Test-Path -LiteralPath $root) { return $root }
		}
		return "$($env:SystemDrive)\storeman"
	}
	
	$buildFullPath = {
		param ([string]$Root,
			[string]$Rel)
		$r = ($Rel -as [string]).Trim().TrimStart('\', '/')
		return ([System.IO.Path]::Combine($Root, $r))
	}
	
	# ---------- INI reader (keeps order of sections and keys in memory) ----------
	$readIni = {
		param ([string]$Path)
		if (-not (Test-Path -LiteralPath $Path)) { return $null }
		
		$raw = Get-Content -LiteralPath $Path -Encoding ASCII
		$sections = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]'
		$current = $null
		
		foreach ($line in $raw)
		{
			$t = $line.Trim()
			if ($t -match '^\s*[;#]') { continue }
			if ($t -match '^\[(?<sec>[^\]]+)\]\s*$')
			{
				$sec = $matches['sec']
				if (-not $sections.ContainsKey($sec))
				{
					$od = New-Object System.Collections.Specialized.OrderedDictionary
					$sections.Add($sec, $od)
				}
				$current = $sec
				continue
			}
			if ($current -and $t -match '^(?<k>[^=]+?)\s*=\s*(?<v>.*)$')
			{
				$k = $matches['k'].Trim()
				$v = $matches['v']
				if (-not $sections[$current].Contains($k)) { $sections[$current].Add($k, $v) }
				else { $sections[$current][$k] = $v }
			}
		}
		return @{ RawLines = $raw; Sections = $sections }
	}
	
	# ---------- INI section writer (no extra blank line; preserves key order from OrderedDictionary) ----------
	$writeIniSection = {
		param (
			[string[]]$RawLines,
			[string]$SectionName,
			[System.Collections.Specialized.OrderedDictionary]$NewKeyValues
		)
		$lines = New-Object System.Collections.Generic.List[string]; $lines.AddRange($RawLines)
		$start = -1; $end = $lines.Count
		
		for ($i = 0; $i -lt $lines.Count; $i++)
		{
			if ($lines[$i] -match "^\s*\[$([regex]::Escape($SectionName))\]\s*$") { $start = $i; break }
		}
		
		if ($start -lt 0)
		{
			if ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Trim() -ne '') { $lines.Add('') }
			$lines.Add("[$SectionName]")
			foreach ($k in $NewKeyValues.Keys) { $lines.Add("$k=$($NewKeyValues[$k])") }
			return, $lines.ToArray()
		}
		
		for ($i = $start + 1; $i -lt $lines.Count; $i++)
		{
			if ($lines[$i] -match '^\s*\[.+\]\s*$') { $end = $i; break }
		}
		
		# find last nonblank/noncomment inside section
		$lastContent = $start
		for ($i = [Math]::Min($end, $lines.Count) - 1; $i -gt $start; $i--)
		{
			$t = $lines[$i].Trim()
			if ($t -ne '' -and $t -notmatch '^[;#]') { $lastContent = $i; break }
		}
		$insertAt = $lastContent + 1
		
		# update existing keys
		$updated = @{ }
		for ($i = $start + 1; $i -lt $end; $i++)
		{
			$L = $lines[$i]
			if ($L.Trim() -match '^[;#]') { continue }
			if ($L -match '^(?<k>[^=]+?)\s*=\s*(?<v>.*)$')
			{
				$currKey = $matches['k'].Trim()
				$targetKey = $null
				foreach ($nk in $NewKeyValues.Keys) { if ($nk -ieq $currKey) { $targetKey = $nk; break } }
				if ($targetKey) { $lines[$i] = "$currKey=$($NewKeyValues[$targetKey])"; $updated[$targetKey] = $true }
			}
		}
		
		# add new keys (preserve order from $NewKeyValues)
		$pending = @()
		foreach ($k in $NewKeyValues.Keys) { if (-not $updated.ContainsKey($k)) { $pending += "$k=$($NewKeyValues[$k])" } }
		if ($pending.Count -gt 0)
		{
			$offset = 0
			foreach ($nl in $pending) { $lines.Insert($insertAt + $offset, $nl); $offset++ }
		}
		return, $lines.ToArray()
	}
	
	$lanesToMachines = {
		param ($LaneNums)
		$out = @()
		$map = $script:FunctionResults['LaneNumToMachineName']
		$known = & $getKnownLanes
		foreach ($ln in $LaneNums)
		{
			$pad = $null; try { $pad = "{0:D3}" -f ([int]$ln) }
			catch { $pad = "$ln" }
			$m = $null
			if ($map)
			{
				if ($map.ContainsKey($pad)) { $m = $map[$pad] }
				elseif ($map.ContainsKey([int]$ln)) { $m = $map[[int]$ln] }
				elseif ($map.ContainsKey("$([int]$ln)")) { $m = $map["$([int]$ln)"] }
			}
			if (-not $m)
			{
				$guess = $known | Where-Object { "$_" -like "*$pad" }
				if ($guess) { $m = $guess[0] }
			}
			if ($m) { $out += $m }
		}
		return, (@($out | Where-Object { $_ } | Select-Object -Unique))
	}
	
	# ==============================================================================================
	#                                           UI LAYOUT
	# ==============================================================================================
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	[void][System.Windows.Forms.Application]::EnableVisualStyles()
	
	$frm = New-Object System.Windows.Forms.Form
	$frm.Text = "INI_Editor  -  dynamic section/key editor"
	$frm.Size = New-Object System.Drawing.Size(820, 600) # keep size modest; we'll give the grid more room
	$frm.StartPosition = 'CenterScreen'
	$frm.FormBorderStyle = 'FixedDialog'
	$frm.MaximizeBox = $false
	$frm.MinimizeBox = $false
	$frm.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)
	
	# Root: 1 column; make grid take the most space (Percent=100 row at the end)
	$root = New-Object System.Windows.Forms.TableLayoutPanel
	$root.Dock = 'Fill'
	$root.Padding = New-Object System.Windows.Forms.Padding(8)
	$root.ColumnCount = 1
	$root.RowCount = 6
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # header
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # source group
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # INI row (flow)
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # Section row
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # GRID gets the rest
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # bottom buttons
	$frm.Controls.Add($root)
	
	# ---------- Header (small + tight margins so the grid gains space) ----------
	$hdr = New-Object System.Windows.Forms.Label
	$hdr.Text = "Pick Server/Lane, load an INI, choose/TYPE a section, edit keys, Save, then optionally deploy."
	$hdr.Font = New-Object System.Drawing.Font("Segoe UI", [float]9, [System.Drawing.FontStyle]::Bold)
	$hdr.AutoSize = $true
	$hdr.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 4)
	$root.Controls.Add($hdr, 0, 0)
	
	# ---------- Source group (compacted) ----------
	$grpSrc = New-Object System.Windows.Forms.GroupBox
	$grpSrc.Text = "Source"
	$grpSrc.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
	$grpSrc.Dock = 'Top'
	$root.Controls.Add($grpSrc, 0, 1)
	
	$srcGrid = New-Object System.Windows.Forms.TableLayoutPanel
	$srcGrid.Dock = 'Top'
	$srcGrid.ColumnCount = 2
	$srcGrid.RowCount = 2
	[void]$srcGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$srcGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$srcGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$srcGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$grpSrc.Controls.Add($srcGrid)
	
	$rbSrv = New-Object System.Windows.Forms.RadioButton
	$rbSrv.Text = "Server (this computer)"
	$rbSrv.Checked = $true
	$rbSrv.AutoSize = $true
	$rbSrv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 2)
	$srcGrid.Controls.Add($rbSrv, 0, 0)
	$srcGrid.SetColumnSpan($rbSrv, 2)
	
	$rbLane = New-Object System.Windows.Forms.RadioButton
	$rbLane.Text = "Lane (source)"
	$rbLane.AutoSize = $true
	$rbLane.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 2)
	$srcGrid.Controls.Add($rbLane, 0, 1)
	
	$laneInline = New-Object System.Windows.Forms.FlowLayoutPanel
	$laneInline.FlowDirection = 'LeftToRight'
	$laneInline.WrapContents = $false
	$laneInline.AutoSize = $true
	$laneInline.AutoSizeMode = 'GrowAndShrink'
	$laneInline.Dock = 'Fill'
	$laneInline.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
	$srcGrid.Controls.Add($laneInline, 1, 1)
	
	$lblLane = New-Object System.Windows.Forms.Label
	$lblLane.Text = "Select lane:"
	$lblLane.AutoSize = $true
	$lblLane.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 2)
	$laneInline.Controls.Add($lblLane)
	
	$cboLane = New-Object System.Windows.Forms.ComboBox
	$cboLane.DropDownStyle = 'DropDownList'
	$cboLane.Enabled = $false
	$cboLane.Width = 260
	$cboLane.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
	$laneInline.Controls.Add($cboLane)
	
	(& $getKnownLanes) | ForEach-Object { [void]$cboLane.Items.Add($_) }
	$rbLane.Add_CheckedChanged({
			$cboLane.Enabled = $rbLane.Checked
			if ($rbLane.Checked -and $cboLane.Items.Count -gt 0 -and -not $cboLane.SelectedItem) { $cboLane.SelectedIndex = 0 }
		})
	
	# ---------- INI path row (ONE ROW via FlowLayoutPanel; resizes combo so the button never wraps) ----------
	$iniFlow = New-Object System.Windows.Forms.FlowLayoutPanel
	$iniFlow.FlowDirection = 'LeftToRight'
	$iniFlow.WrapContents = $false
	$iniFlow.Dock = 'Top'
	$iniFlow.AutoSize = $false
	$iniFlow.Height = 38
	$iniFlow.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
	$root.Controls.Add($iniFlow, 0, 2)
	
	$lblRel = New-Object System.Windows.Forms.Label
	$lblRel.Text = "INI path under \storeman\"
	$lblRel.AutoSize = $true
	$lblRel.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
	$iniFlow.Controls.Add($lblRel)
	
	$cboRel = New-Object System.Windows.Forms.ComboBox
	$cboRel.DropDownStyle = 'DropDown' # editable
	$cboRel.Margin = New-Object System.Windows.Forms.Padding(6, 3, 6, 3)
	$cboRel.Width = 520
	foreach ($p in ($PredefinedIniPaths | Select-Object -Unique)) { [void]$cboRel.Items.Add($p) }
	if ($RelativeIniPath -and -not ($cboRel.Items -contains $RelativeIniPath)) { [void]$cboRel.Items.Add($RelativeIniPath) }
	$cboRel.Text = $RelativeIniPath
	$iniFlow.Controls.Add($cboRel)
	
	$btnLoad = New-Object System.Windows.Forms.Button
	$btnLoad.Text = "Load INI"
	$btnLoad.Width = 100
	$btnLoad.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
	$iniFlow.Controls.Add($btnLoad)
	
	# Resize handler to keep all three controls on ONE line
	$iniFlow.Add_Resize({
			$total = $iniFlow.ClientSize.Width
			$wLabel = $lblRel.PreferredSize.Width
			$wBtn = $btnLoad.Width
			$pad = 24 # rough margins between controls
			$newW = [Math]::Max(220, $total - $wLabel - $wBtn - $pad)
			$cboRel.Width = $newW
		})
	
	# ---------- Section row (tight) ----------
	$rowSec = New-Object System.Windows.Forms.TableLayoutPanel
	$rowSec.Dock = 'Top'
	$rowSec.ColumnCount = 4
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 320)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$rowSec.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 2)
	$root.Controls.Add($rowSec, 0, 3)
	
	$lblSec = New-Object System.Windows.Forms.Label
	$lblSec.Text = "Section:"
	$lblSec.AutoSize = $true
	$lblSec.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 0)
	$rowSec.Controls.Add($lblSec, 0, 0)
	
	$cboSection = New-Object System.Windows.Forms.ComboBox
	$cboSection.DropDownStyle = 'DropDown'
	$cboSection.Width = 320
	$cboSection.Margin = New-Object System.Windows.Forms.Padding(0, 3, 6, 3)
	$rowSec.Controls.Add($cboSection, 1, 0)
	
	$lblPath = New-Object System.Windows.Forms.Label
	$lblPath.Text = ""
	$lblPath.AutoSize = $true
	$lblPath.ForeColor = [System.Drawing.Color]::DimGray
	$lblPath.Margin = New-Object System.Windows.Forms.Padding(12, 6, 6, 0)
	$rowSec.Controls.Add($lblPath, 2, 0)
	
	$btnReload = New-Object System.Windows.Forms.Button
	$btnReload.Text = "Reload"
	$btnReload.Width = 90
	$btnReload.Margin = New-Object System.Windows.Forms.Padding(6, 3, 6, 3)
	$rowSec.Controls.Add($btnReload, 3, 0)
	
	# ---------- Grid (make it take the space) ----------
	$grid = New-Object System.Windows.Forms.DataGridView
	$grid.Dock = 'Fill'
	$grid.AllowUserToAddRows = $false
	$grid.AllowUserToDeleteRows = $true
	$grid.AllowDrop = $true
	$grid.AutoSizeColumnsMode = 'Fill'
	$grid.RowHeadersVisible = $false
	$grid.SelectionMode = 'CellSelect'
	$grid.MultiSelect = $false
	$grid.BackgroundColor = [System.Drawing.Color]::White
	$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 243, 245)
	$grid.EnableHeadersVisualStyles = $false
	$root.Controls.Add($grid, 0, 4)
	
	$dt = New-Object System.Data.DataTable
	[void]$dt.Columns.Add('Key', [string])
	[void]$dt.Columns.Add('Value', [string])
	$grid.DataSource = $dt
	$grid.EditMode = 'EditOnEnter'
	$grid.Columns['Key'].SortMode = 'NotSortable'
	$grid.Columns['Value'].SortMode = 'NotSortable'
	
	# ---------- Bottom buttons ----------
	$btnRow = New-Object System.Windows.Forms.FlowLayoutPanel
	$btnRow.FlowDirection = 'RightToLeft'
	$btnRow.Dock = 'Bottom'
	$btnRow.AutoSize = $true
	$btnRow.Padding = New-Object System.Windows.Forms.Padding(0, 6, 6, 6)
	$frm.Controls.Add($btnRow)
	
	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = "Close"
	$btnClose.Width = 90
	$btnRow.Controls.Add($btnClose)
	
	$btnCopy = New-Object System.Windows.Forms.Button
	$btnCopy.Text = "Copy to other lanes..."
	$btnCopy.Width = 170
	$btnCopy.Enabled = $false # stays off until a successful Save
	$btnRow.Controls.Add($btnCopy)
	
	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Text = "Save to source"
	$btnSave.Width = 130
	$btnRow.Controls.Add($btnSave)
	
	$frm.AcceptButton = $btnSave
	$frm.CancelButton = $btnClose
	
	# ==============================================================================================
	#                                   CONTEXT MENU: Add/Delete
	# ==============================================================================================
	
	$ctx = New-Object System.Windows.Forms.ContextMenuStrip
	$miAdd = $ctx.Items.Add("Add key")
	$miDel = $ctx.Items.Add("Delete key")
	$grid.ContextMenuStrip = $ctx
	
	$script:_ctxRow = -1
	$grid.Add_CellMouseDown({
			param ($sender,
				$e)
			if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right)
			{
				if ($e.RowIndex -ge 0)
				{
					$grid.ClearSelection()
					$script:_ctxRow = $e.RowIndex
					$col = if ($e.ColumnIndex -ge 0) { $e.ColumnIndex }
					else { 0 }
					$grid.CurrentCell = $grid.Rows[$e.RowIndex].Cells[$col]
					$grid.Rows[$e.RowIndex].Selected = $true
				}
				else { $script:_ctxRow = -1 }
			}
		})
	
	# Add key at END; start editing; mark dirty (Copy disabled)
	$miAdd.Add_Click({
			$insertAt = $dt.Rows.Count
			$newRow = $dt.NewRow()
			$newRow['Key'] = ''
			$newRow['Value'] = ''
			[void]$dt.Rows.Add($newRow)
			
			$grid.Focus()
			try { $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $insertAt) }
			catch { }
			$grid.CurrentCell = $grid.Rows[$insertAt].Cells[0]
			$grid.BeginEdit($true)
			
			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})
	
	$miDel.Add_Click({
			$rowIdx = -1
			if ($grid.CurrentCell) { $rowIdx = $grid.CurrentCell.RowIndex }
			if ($rowIdx -lt 0 -and $script:_ctxRow -ge 0) { $rowIdx = $script:_ctxRow }
			if ($rowIdx -lt 0 -or $rowIdx -ge $grid.Rows.Count) { return }
			
			$ans = [System.Windows.Forms.MessageBox]::Show(
				"Delete key '" + [string]$dt.Rows[$rowIdx]['Key'] + "'?",
				"Confirm delete", [System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question)
			if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
			
			$dt.Rows.RemoveAt($rowIdx)
			
			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})
	
	# Any edit => dirty => Copy disabled
	$grid.Add_CellValueChanged({ $state.IsDirty = $true; $btnCopy.Enabled = $false })
	$grid.Add_UserDeletedRow({ $state.IsDirty = $true; $btnCopy.Enabled = $false })
	
	# ==============================================================================================
	#                               DRAG & DROP ROW REORDERING
	# ==============================================================================================
	
	$script:_dragRow = -1
	$script:_dragging = $false
	$script:_dragStart = [System.Drawing.Point]::Empty
	
	$grid.Add_MouseDown({
			param ($s,
				$e)
			$script:_dragStart = [System.Drawing.Point]::new($e.X, $e.Y)
			$hit = $grid.HitTest($e.X, $e.Y)
			if ($hit.RowIndex -ge 0 -and $hit.RowIndex -lt $dt.Rows.Count) { $script:_dragRow = $hit.RowIndex }
			else { $script:_dragRow = -1 }
		})
	
	$grid.Add_MouseMove({
			param ($s,
				$e)
			if ($e.Button -band [System.Windows.Forms.MouseButtons]::Left -and $script:_dragRow -ge 0 -and -not $script:_dragging)
			{
				$dx = [Math]::Abs($e.X - $script:_dragStart.X)
				$dy = [Math]::Abs($e.Y - $script:_dragStart.Y)
				if ($dx -gt 4 -or $dy -gt 4)
				{
					$script:_dragging = $true
					$null = $grid.DoDragDrop("row-move", [System.Windows.Forms.DragDropEffects]::Move)
					$script:_dragging = $false
				}
			}
		})
	
	$grid.Add_DragOver({
			param ($s,
				$e)
			if ($script:_dragRow -lt 0) { $e.Effect = [System.Windows.Forms.DragDropEffects]::None; return }
			$pt = $grid.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
			$hit = $grid.HitTest($pt.X, $pt.Y)
			$target = if ($hit.RowIndex -lt 0) { $dt.Rows.Count - 1 }
			else { $hit.RowIndex }
			$e.Effect = if ($target -ne $script:_dragRow) { [System.Windows.Forms.DragDropEffects]::Move }
			else { [System.Windows.Forms.DragDropEffects]::None }
		})
	
	$grid.Add_DragDrop({
			param ($s,
				$e)
			if ($script:_dragRow -lt 0) { return }
			$source = $script:_dragRow
			$script:_dragRow = -1
			
			$pt = $grid.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
			$hit = $grid.HitTest($pt.X, $pt.Y)
			$target = $hit.RowIndex
			if ($target -lt 0) { $target = $dt.Rows.Count - 1 }
			if ($target -gt $dt.Rows.Count - 1) { $target = $dt.Rows.Count - 1 }
			if ($target -eq $source) { return }
			
			$keyVal = [string]$dt.Rows[$source]['Key']
			$valueVal = [string]$dt.Rows[$source]['Value']
			
			$dt.Rows.RemoveAt($source)
			if ($target -gt $source) { $target-- }
			
			$newRow = $dt.NewRow()
			$newRow['Key'] = $keyVal
			$newRow['Value'] = $valueVal
			if ($target -ge 0 -and $target -lt $dt.Rows.Count) { $dt.Rows.InsertAt($newRow, $target) }
			else { [void]$dt.Rows.Add($newRow); $target = $dt.Rows.Count - 1 }
			
			try { $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $target) }
			catch { }
			$grid.CurrentCell = $grid.Rows[$target].Cells[0]
			$grid.Rows[$target].Selected = $true
			
			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})
	
	# ==============================================================================================
	#                                STATE + LOAD / SAVE / COPY
	# ==============================================================================================
	
	$state = @{
		SourceIsServer = $true
		SourceLane	   = $null
		FullPath	   = $null
		IniObj		   = $null
		DidSave	       = $false
		DidDeploy	   = $false
	}
	
	$loadSectionFromCombo = {
		$dt.Rows.Clear()
		if (-not $state.IniObj) { return }
		$secRequested = ($cboSection.Text -as [string]).Trim()
		if (-not $secRequested) { return }
		$secActual = $null
		foreach ($k in $state.IniObj.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
		if (-not $secActual) { return }
		$od = $state.IniObj.Sections[$secActual]
		foreach ($k in $od.Keys) { [void]$dt.Rows.Add($k, [string]$od[$k]) }
		$state.IsDirty = $false
		$btnCopy.Enabled = $false
	}
	
	$doLoad = {
		$rel = ($cboRel.Text -as [string]).Trim()
		if (-not $rel)
		{
			[System.Windows.Forms.MessageBox]::Show("Enter or select a relative INI path (e.g. office\Setup.ini).", "Missing path",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
			return
		}
		
		if ($rbSrv.Checked)
		{
			$state.SourceIsServer = $true
			$state.SourceLane = $null
			$rootSrv = & $getServerRoot
			$full = & $buildFullPath $rootSrv $rel
		}
		else
		{
			if (-not $cboLane.SelectedItem)
			{
				[System.Windows.Forms.MessageBox]::Show("Select a lane.", "Missing lane",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				return
			}
			$state.SourceIsServer = $false
			$state.SourceLane = [string]$cboLane.SelectedItem
			$rootLane = & $getLaneRoot $state.SourceLane
			if (-not $rootLane)
			{
				[System.Windows.Forms.MessageBox]::Show("Lane storeman root not accessible (\\$($state.SourceLane)\Storeman or \\$($state.SourceLane)\C$\storeman).", "Path error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				return
			}
			$full = & $buildFullPath $rootLane $rel
		}
		
		$state.FullPath = $full
		$lblPath.Text = $state.FullPath
		
		$ini = & $readIni $state.FullPath
		if (-not $ini)
		{
			Write_Log "[Info] INI not found, will be created on save: $($state.FullPath)" "yellow"
			$ini = @{
				RawLines = @("; Created by INI_Editor", "")
				Sections = (New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]')
			}
		}
		$state.IniObj = $ini
		
		$cboSection.Items.Clear()
		foreach ($secName in $state.IniObj.Sections.Keys) { [void]$cboSection.Items.Add($secName) }
		if ($cboSection.Items.Count -gt 0) { $cboSection.SelectedIndex = 0 }
		else { $cboSection.Text = "" }
		
		$loadSectionFromCombo.Invoke()
	}
	
	$btnLoad.Add_Click({ & $doLoad })
	$btnReload.Add_Click({ & $doLoad })
	$cboSection.Add_SelectedIndexChanged({ $loadSectionFromCombo.Invoke() })
	$cboSection.Add_TextChanged({ $loadSectionFromCombo.Invoke() })
	
	# ---------- Save ----------
	$btnSave.Add_Click({
			if (-not $state.FullPath)
			{
				[System.Windows.Forms.MessageBox]::Show("Load an INI first.", "No file",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			if (-not $state.IniObj) { return }
			
			$secRequested = ($cboSection.Text -as [string]).Trim()
			if (-not $secRequested)
			{
				[System.Windows.Forms.MessageBox]::Show("Enter or select a section name first.", "No section",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$kv = New-Object System.Collections.Specialized.OrderedDictionary
			foreach ($row in $dt.Rows)
			{
				$k = [string]$row['Key']; $v = [string]$row['Value']
				if ([string]::IsNullOrWhiteSpace($k)) { continue }
				if ($kv.Contains($k)) { $kv.Remove($k) } # last one wins
				$kv.Add($k, $v)
			}
			
			$secActual = $null
			foreach ($k in $state.IniObj.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
			if (-not $secActual)
			{
				$secActual = $secRequested
				$od = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($k in $kv.Keys) { $od.Add($k, $kv[$k]) }
				$state.IniObj.Sections.Add($secActual, $od)
			}
			else
			{
				$od = $state.IniObj.Sections[$secActual]
				foreach ($k in @($od.Keys)) { $od.Remove($k) }
				foreach ($k in $kv.Keys) { $od.Add($k, $kv[$k]) }
			}
			
			$state.IniObj.RawLines = & $writeIniSection $state.IniObj.RawLines $secActual $kv
			
			try
			{
				$dir = Split-Path -Parent $state.FullPath
				if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
				Set-Content -LiteralPath $state.FullPath -Value $state.IniObj.RawLines -Encoding ASCII -Force
				
				Write_Log "Saved [$secActual] to $($state.FullPath)" "green"
				[System.Windows.Forms.MessageBox]::Show("Saved section [$secActual] to:`r`n$($state.FullPath)", "Saved",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				
				$state.DidSave = $true
				$state.IsDirty = $false
				$btnCopy.Enabled = $true
			}
			catch
			{
				Write_Log "Save failed for $($state.FullPath): $($_.Exception.Message)" "red"
				[System.Windows.Forms.MessageBox]::Show("Save failed:`r`n$($_.Exception.Message)", "Error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
			}
		})
	
	# ---------- Copy (blocked if dirty or not saved) ----------
	$btnCopy.Add_Click({
			if ($state.IsDirty)
			{
				[System.Windows.Forms.MessageBox]::Show("Please Save first. Copy is disabled until the INI is saved.", "Save required",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			if (-not $state.FullPath -or -not (Test-Path -LiteralPath $state.FullPath))
			{
				[System.Windows.Forms.MessageBox]::Show("Nothing to copy. Save the INI first.", "No file",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$secRequested = ($cboSection.Text -as [string]).Trim()
			
			# deploy dialog
			$dlg = New-Object System.Windows.Forms.Form
			$dlg.Text = "Deploy options"
			$dlg.StartPosition = 'CenterParent'
			$dlg.FormBorderStyle = 'FixedDialog'
			$dlg.MaximizeBox = $false
			$dlg.MinimizeBox = $false
			$dlg.ClientSize = New-Object System.Drawing.Size(460, 160)
			
			$lbl = New-Object System.Windows.Forms.Label
			$lbl.AutoSize = $true
			$lbl.MaximumSize = New-Object System.Drawing.Size(440, 0)
			$lbl.Location = New-Object System.Drawing.Point(10, 10)
			$lbl.Text = "How would you like to deploy to other lanes?`r`n`r`n" +
			"* Merge - update only section [$secRequested] keys on targets.`r`n" +
			"* Copy  - replace the entire INI file on targets."
			$dlg.Controls.Add($lbl)
			
			$btnMerge = New-Object System.Windows.Forms.Button
			$btnMerge.Text = "Merge"; $btnMerge.Size = New-Object System.Drawing.Size(90, 28)
			$btnMerge.Location = New-Object System.Drawing.Point(140, 110)
			$btnMerge.DialogResult = [System.Windows.Forms.DialogResult]::Yes
			
			$btnCopyAll = New-Object System.Windows.Forms.Button
			$btnCopyAll.Text = "Copy"; $btnCopyAll.Size = New-Object System.Drawing.Size(90, 28)
			$btnCopyAll.Location = New-Object System.Drawing.Point(240, 110)
			$btnCopyAll.DialogResult = [System.Windows.Forms.DialogResult]::No
			
			$btnCancel = New-Object System.Windows.Forms.Button
			$btnCancel.Text = "Cancel"; $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
			$btnCancel.Location = New-Object System.Drawing.Point(340, 110)
			$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			
			$dlg.Controls.AddRange(@($btnMerge, $btnCopyAll, $btnCancel))
			$dlg.AcceptButton = $btnMerge
			$dlg.CancelButton = $btnCancel
			if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $dlg }
			
			$choice = $dlg.ShowDialog()
			if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
			$mergeOnly = ($choice -eq [System.Windows.Forms.DialogResult]::Yes)
			
			$excluded = @()
			if (-not $state.SourceIsServer -and $state.SourceLane) { $excluded = @($state.SourceLane) }
			
			$StoreNumber = $script:FunctionResults['StoreNumber']
			$sel = $null
			try { $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select destination lanes" -ExcludedNodes $excluded }
			catch { $sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select destination lanes" }
			if (-not $sel) { Write_Log "Copy aborted: no lanes selected." "yellow"; return }
			
			$laneNums = @()
			if ($sel -is [System.Collections.IDictionary])
			{
				if ($sel.Contains('Lanes') -and $sel['Lanes']) { $laneNums = @($sel['Lanes']) }
				elseif ($sel.Contains('LaneNumbers') -and $sel['LaneNumbers']) { $laneNums = @($sel['LaneNumbers']) }
			}
			elseif ($sel -is [System.Collections.IEnumerable] -and -not ($sel -is [string])) { $laneNums = @($sel) }
			elseif ($sel -is [string]) { $laneNums = @($sel -split '[,\s]+' | Where-Object { $_ }) }
			$laneNums = @($laneNums | ForEach-Object {
					try { "{0:D3}" -f ([int]$_) }
					catch { "$_" }
				}) | Where-Object { $_ -match '^\d{3}$' }
			if (-not $laneNums -or $laneNums.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("No destination lanes selected.", "No selection",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$destMachines = & $lanesToMachines $laneNums
			if (-not $state.SourceIsServer -and $state.SourceLane) { $destMachines = $destMachines | Where-Object { $_ -ne $state.SourceLane } }
			$destMachines = $destMachines | Select-Object -Unique
			if ($destMachines.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("No valid destination machines resolved.", "No destinations",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			
			$relPath = ($cboRel.Text -as [string]).Trim()
			$srcDir = Split-Path -Parent $state.FullPath
			$srcFile = Split-Path -Leaf $state.FullPath
			$ok = 0; $fail = 0
			
			if ($mergeOnly)
			{
				$kv = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($row in $dt.Rows)
				{
					$k = [string]$row['Key']; $v = [string]$row['Value']
					if ([string]::IsNullOrWhiteSpace($k)) { continue }
					if ($kv.Contains($k)) { $kv.Remove($k) }
					$kv.Add($k, $v)
				}
				
				foreach ($m in $destMachines)
				{
					$rootLane = & $getLaneRoot $m
					if (-not $rootLane) { Write_Log "[Dest Unreachable] $m has no \\Storeman nor \\C$\storeman." "red"; $fail++; continue }
					$dstFull = & $buildFullPath $rootLane $relPath
					$dstDir = Split-Path -Parent $dstFull
					try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
					catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $fail++; continue }
					
					$destIni = & $readIni $dstFull
					if (-not $destIni)
					{
						$destIni = @{ RawLines = @("; Created by INI_Editor (merge)", ""); Sections = (New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]') }
					}
					$secActual = $null
					foreach ($k in $destIni.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
					if (-not $secActual) { $secActual = $secRequested }
					
					$destIni.RawLines = & $writeIniSection $destIni.RawLines $secActual $kv
					try
					{
						Set-Content -LiteralPath $dstFull -Value $destIni.RawLines -Encoding ASCII -Force
						Write_Log "Merged [$secActual] into $dstFull on $m" "green"; $ok++
					}
					catch
					{
						Write_Log "Merge failed on $m ($dstFull): $($_.Exception.Message)" "red"; $fail++
					}
				}
			}
			else
			{
				foreach ($m in $destMachines)
				{
					$rootLane = & $getLaneRoot $m
					if (-not $rootLane) { Write_Log "[Dest Unreachable] $m has no \\Storeman nor \\C$\storeman." "red"; $fail++; continue }
					$dstFull = & $buildFullPath $rootLane $relPath
					$dstDir = Split-Path -Parent $dstFull
					try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
					catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $fail++; continue }
					
					Write_Log "Replacing INI on $m ($dstDir\$srcFile)" "gray"
					$args = @($srcDir, $dstDir, $srcFile, '/COPY:DAT', '/R:2', '/W:2', '/NFL', '/NDL', '/NP', '/MT:8')
					$null = & robocopy @args
					$code = $LASTEXITCODE
					if ($code -ge 8) { Write_Log "Copy to $m FAILED (robocopy $code)" "red"; $fail++ }
					else { Write_Log "Copy to $m OK (robocopy $code)" "green"; $ok++ }
				}
			}
			if ($ok -gt 0) { $state.DidDeploy = $true } # <- at least one target updated OK
			[System.Windows.Forms.MessageBox]::Show("Deploy complete.`r`nOK: $ok   Fail: $fail", "Done",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
		})
	
	# run no-op/activity log when the form actually closes (covers X, Alt+F4, Close button)	$btnClose.Add_Click({
	$frm.Add_FormClosed({
			if (-not $state.DidSave -and -not $state.DidDeploy)
			{
				Write_Log "[No operations] Editor closed without saving or deploying." "gray"
			}
			else
			{
				# Optional: a tiny summary
				$acts = @()
				if ($state.DidSave) { $acts += "save" }
				if ($state.DidDeploy) { $acts += "deploy" }
				Write_Log "[Activity] Completed: $($acts -join ', ')." "gray"
			}
		})
	
	$btnClose.Add_Click({ $frm.Close() })
	
	# ==============================================================================================
	#                                         STARTUP
	# ==============================================================================================
	& $doLoad
	$iniFlow.PerformLayout() # ensure combo gets sized on first draw
	if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $frm }
	[void]$frm.ShowDialog()
	
	Write_Log "`r`n==================== INI_Editor Completed ====================`r`n" "blue"
}
