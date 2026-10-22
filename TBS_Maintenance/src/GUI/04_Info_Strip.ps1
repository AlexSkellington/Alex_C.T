# ===================================================================================================
# ENHANCED INFO STRIP - Modern card with visual status indicators
# ===================================================================================================
$topStripCard = New-Object System.Windows.Forms.Panel
$topStripCard.BackColor = $uiCardBg
$topStripCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$topStripCard.Height = 82  # Slightly taller for better spacing
$topStripCard.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$form.Controls.Add($topStripCard)

# Subtle shadow effect (bottom line)
$topStripCard.add_Paint({
		param ($s, $e)
		try
		{
			$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220, 220, 220))
			$e.Graphics.DrawLine($pen, 1, $s.Height - 1, $s.Width - 1, $s.Height - 1)
			$pen.Dispose()
		}
		catch { }
	})

# Modern inner layout: TableLayoutPanel
$topStripTable = New-Object System.Windows.Forms.TableLayoutPanel
$topStripTable.Dock = 'Fill'
$topStripTable.Padding = New-Object System.Windows.Forms.Padding(16, 10, 16, 10)
$topStripTable.ColumnCount = 3
$topStripTable.RowCount = 2
$topStripTable.BackColor = [System.Drawing.Color]::Transparent

# Columns: left / center / right
[void]$topStripTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$topStripTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 34)))
[void]$topStripTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))

# Rows: two equal rows
[void]$topStripTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$topStripTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$topStripCard.Controls.Add($topStripTable)

# ------------------------------------------------------------------------------------------------
# Shared tooltip (create once)  [?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?][?]
# ------------------------------------------------------------------------------------------------
if (-not $toolTip)
{
	$toolTip = New-Object System.Windows.Forms.ToolTip
	$toolTip.AutoPopDelay = 8000
	$toolTip.InitialDelay = 300
	$toolTip.ReshowDelay = 100
}

# Helper as scriptblock (not a function): make a label look like a link (hand + accent + underline on hover)
$applyLinkStyle = {
	param ([System.Windows.Forms.Label]$Label)
	$Label.Cursor = [System.Windows.Forms.Cursors]::Hand
	if (-not $Label.Tag) { $Label.Tag = [pscustomobject]@{ Color = $Label.ForeColor; Font = $Label.Font } }

	$Label.add_MouseEnter({
			param ($s,
				$e)
			try
			{
				$s.ForeColor = $uiAccent
				$s.Font = New-Object System.Drawing.Font($s.Font, ($s.Font.Style -bor [System.Drawing.FontStyle]::Underline))
			}
			catch { }
		})
	$Label.add_MouseLeave({
			param ($s,
				$e)
			try
			{
				if ($s.Tag -and $s.Tag.Font) { $s.Font = $s.Tag.Font }
				if ($s.Tag -and $s.Tag.Color) { $s.ForeColor = $s.Tag.Color }
			}
			catch { }
		})
}

# ------------------------------------------------------------------------------------------------
# Labels (same functionality, modern font + positioned via table layout)
# ------------------------------------------------------------------------------------------------

# ===================================================================================================
# ENHANCED INFO LABELS - Modern typography with status indicators
# ===================================================================================================

# Helper for info label styling
$applyInfoLabelStyle = {
	param ([System.Windows.Forms.Label]$Label)
	$Label.Cursor = [System.Windows.Forms.Cursors]::Hand
	$Label.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
	$Label.ForeColor = $uiTextPrimary
	if (-not $Label.Tag) { $Label.Tag = [pscustomobject]@{ Color = $Label.ForeColor; Font = $Label.Font } }

	$Label.add_MouseEnter({
			param ($s, $e)
			try { $s.ForeColor = $uiAccent; $s.Font = New-Object System.Drawing.Font($s.Font, ($s.Font.Style -bor [System.Drawing.FontStyle]::Underline)) }
			catch { }
		})
	$Label.add_MouseLeave({
			param ($s, $e)
			try
			{
				if ($s.Tag -and $s.Tag.Font) { $s.Font = $s.Tag.Font }
				if ($s.Tag -and $s.Tag.Color) { $s.ForeColor = $s.Tag.Color }
			}
			catch { }
		})
}

# SMS Version (left, row 1)
$smsVersionLabel = New-Object System.Windows.Forms.Label
$smsVersionLabel.Name = "SMSVersionLabel"
$smsVersionLabel.Text = "SMS Version: N/A"
$smsVersionLabel.AutoSize = $true
$smsVersionLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
$smsVersionLabel.ForeColor = $uiTextPrimary
$smsVersionLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$smsVersionLabel.Anchor = 'Left'
$topStripTable.Controls.Add($smsVersionLabel, 0, 0) | Out-Null
# Vertically center in the 30px cell: Dock=Fill + TextAlign.
# (The old ControlAlignment line used [System.Windows.Forms.TableLayoutPanelCellAlignment],
#  a type that does not exist in .NET, so it silently failed on every run.)
$smsVersionLabel.AutoSize = $false
$smsVersionLabel.Dock = 'Fill'
$smsVersionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$toolTip.SetToolTip($smsVersionLabel, "Shows current SMS version. Click to bring SMSStart to front (or launch it).")
& $applyInfoLabelStyle $smsVersionLabel

# Click: bring SMSStart to front or launch it
$smsVersionLabel.Add_Click({
		$script:LastActivity = Get-Date
		$orig = $smsVersionLabel.ForeColor
		$smsVersionLabel.ForeColor = $uiAccent
		try
		{
			@"
using System;
using System.Runtime.InteropServices;
public static class NativeWin {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
}
"@ | ForEach-Object {
				if (-not ("NativeWin" -as [type])) { Add-Type -TypeDefinition $_ -ErrorAction SilentlyContinue }
			}

			$p = Get-Process -Name 'SMSStart' -ErrorAction SilentlyContinue |
			Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
			if ($p)
			{
				$h = $p.MainWindowHandle
				if ([NativeWin]::IsIconic($h)) { [NativeWin]::ShowWindow($h, 9) | Out-Null }
				[NativeWin]::SetForegroundWindow($h) | Out-Null
				if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "Brought SMSStart to the foreground." "green" }
				return
			}

			$exe = Join-Path $BasePath 'SMSStart.exe'
			$proc = $null
			if (Test-Path -LiteralPath $exe)
			{
				$proc = Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe -Parent) -PassThru -ErrorAction Stop
			}
			else
			{
				$lnk = Get-ChildItem -LiteralPath $BasePath -Filter '*SMSStart*.lnk' -ErrorAction SilentlyContinue | Select-Object -First 1
				if ($lnk) { $proc = Start-Process -FilePath $lnk.FullName -PassThru -ErrorAction Stop }
			}

			if (-not $proc)
			{
				if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "SMSStart not found under $BasePath" "yellow" }
				return
			}

			$null = $proc.WaitForInputIdle(5000)
			for ($i = 0; $i -lt 20 -and $proc.MainWindowHandle -eq 0; $i++)
			{
				Start-Sleep -Milliseconds 150
				$proc.Refresh()
			}
			if ($proc.MainWindowHandle -ne 0)
			{
				if ([NativeWin]::IsIconic($proc.MainWindowHandle)) { [NativeWin]::ShowWindow($proc.MainWindowHandle, 9) | Out-Null }
				[NativeWin]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
			}
			if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "Launched SMSStart from $BasePath and brought to front." "green" }
		}
		finally
		{
			$smsVersionLabel.ForeColor = $orig
		}
	})

# Store Name (centered, row 1)
$storeNameLabel = New-Object System.Windows.Forms.Label
$storeNameLabel.Name = "StoreNameLabel"
$storeNameLabel.Text = "Store Name: N/A"
$storeNameLabel.AutoSize = $true
$storeNameLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
$storeNameLabel.ForeColor = $uiTextPrimary
$storeNameLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$storeNameLabel.Anchor = 'None'
$topStripTable.Controls.Add($storeNameLabel, 1, 0) | Out-Null
# Vertically center in the cell (real fix - see note at smsVersionLabel above)
$storeNameLabel.AutoSize = $false
$storeNameLabel.Dock = 'Fill'
$storeNameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$toolTip.SetToolTip($storeNameLabel, "Shows store name. Click to ping all nodes (Lanes > Scales > Backoffices).")
& $applyInfoLabelStyle $storeNameLabel

# Click: ping all nodes
$storeNameLabel.Add_Click({
		$script:LastActivity = Get-Date
		$orig = $storeNameLabel.ForeColor
		$origText = $storeNameLabel.Text
		$storeNameLabel.ForeColor = $uiAccent
		try
		{
			$storeNameLabel.Text = "Store Name: Pinging..."
			[System.Windows.Forms.Application]::DoEvents()
			Ping_All_Nodes -NodeType "Lane" -StoreNumber $StoreNumber
			Ping_All_Nodes -NodeType "Scale" -StoreNumber $StoreNumber
			Ping_All_Nodes -NodeType "Backoffice" -StoreNumber $StoreNumber
		}
		finally
		{
			# Put the store name back (only the colour used to be restored: "Pinging..." stuck)
			$storeNameLabel.Text = $origText
			$storeNameLabel.ForeColor = $orig
		}
	})

# Store Number (right, row 1)
$storeNumberLabel = New-Object System.Windows.Forms.Label
$storeNumberLabel.Name = "StoreNumberLabel"
$storeNumberLabel.Text = "Store Number: N/A"
$storeNumberLabel.AutoSize = $true
$storeNumberLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
$storeNumberLabel.ForeColor = $uiTextPrimary
$storeNumberLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$storeNumberLabel.Anchor = 'Right'
$topStripTable.Controls.Add($storeNumberLabel, 2, 0) | Out-Null
# Vertically center in the cell (real fix - see note at smsVersionLabel above)
$storeNumberLabel.AutoSize = $false
$storeNumberLabel.Dock = 'Fill'
$storeNumberLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

# keep your original script-scoped reference for other code
$script:storeNumberLabel = $storeNumberLabel

# Hover: refresh tooltip with actual BasePath (+ test-environment context)
$storeNumberLabel.Add_MouseHover({
		$p = if ($script:BasePath) { $script:BasePath }
		elseif ($BasePath) { $BasePath }
		else { $null }
		if ($script:IsTestEnvironment)
		{
			$modePrefix = if ($script:StoremanKeyStale) { "KEY EXPIRED" } else { "TEST ENVIRONMENT" }
			$folderWord = if ($script:StoremanKeyStale) { "storeman folder" } else { "test storeman folder" }
			# NOTE: ${} delimiters - a bare '$folderWord:' would parse as a scope reference
			$toolTip.SetToolTip($storeNumberLabel, "${modePrefix} - $($script:StoremanSourceReason). Open ${folderWord}: ${p}")
		}
		else
		{
			$toolTip.SetToolTip($storeNumberLabel, $(if ($p) { "Open Storeman folder: $p" }
					else { "Storeman folder not detected yet." }))
		}
	})

# Click: open Storeman folder (inline)
$storeNumberLabel.Add_Click({
		$script:LastActivity = Get-Date
		$path = if ($script:BasePath) { $script:BasePath }
		elseif ($BasePath) { $BasePath }
		else { $null }
		if ($path -and (Test-Path -LiteralPath $path))
		{
			try { Start-Process -FilePath $path }
			catch
			{
				[System.Windows.Forms.MessageBox]::Show(
					"Couldn't open: $path`r`n$($_.Exception.Message)",
					"Open Storeman Folder",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Error
				) | Out-Null
			}
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Storeman folder not found.`r`nCurrent value: " + [string]$path,
				"Open Storeman Folder",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			) | Out-Null
		}
	})

# ===================================================================================================
# SECOND ROW - Node counts with modern badges
# ===================================================================================================

# Number of Lanes (left, row 2)
$NodesStore = New-Object System.Windows.Forms.Label
$NodesStore.Name = "NodesStoreLabel"
$NodesStore.Text = "Lanes: N/A"
$NodesStore.AutoSize = $true
$NodesStore.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
$NodesStore.ForeColor = $uiCategoryLane
$NodesStore.Margin = New-Object System.Windows.Forms.Padding(0)
$NodesStore.Anchor = 'Left'
$topStripTable.Controls.Add($NodesStore, 0, 1) | Out-Null
# Vertically center in the cell (real fix - see note at smsVersionLabel above)
$NodesStore.AutoSize = $false
$NodesStore.Dock = 'Fill'
$NodesStore.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# keep your original script-scoped reference for other code
$script:NodesStore = $NodesStore

$toolTip.SetToolTip($NodesStore, "Shows count of Lanes. Click to ping Lane nodes.")
& $applyInfoLabelStyle $NodesStore
$NodesStore.Add_Click({
		$script:LastActivity = Get-Date
		$orig = $NodesStore.ForeColor
		$NodesStore.ForeColor = $uiAccent
		try { Ping_All_Nodes -NodeType "Lane" -StoreNumber $StoreNumber }
		finally { $NodesStore.ForeColor = $uiCategoryLane }
	})

# Number of Scales (center, row 2)
$scalesLabel = New-Object System.Windows.Forms.Label
$scalesLabel.Name = "ScalesLabel"
$scalesLabel.Text = "Scales: N/A"
$scalesLabel.AutoSize = $true
$scalesLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
$scalesLabel.ForeColor = $uiCategoryScale
$scalesLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$scalesLabel.Anchor = 'None'
$topStripTable.Controls.Add($scalesLabel, 1, 1) | Out-Null
# Vertically center in the cell (real fix - see note at smsVersionLabel above)
$scalesLabel.AutoSize = $false
$scalesLabel.Dock = 'Fill'
$scalesLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

# keep your original script-scoped reference for other code
$script:scalesLabel = $scalesLabel

$toolTip.SetToolTip($scalesLabel, "Shows count of Scales. Click to ping Scale nodes.")
& $applyInfoLabelStyle $scalesLabel
$scalesLabel.Add_Click({
		$script:LastActivity = Get-Date
		$orig = $scalesLabel.ForeColor
		$scalesLabel.ForeColor = $uiAccent
		try { Ping_All_Nodes -NodeType "Scale" -StoreNumber $StoreNumber }
		finally { $scalesLabel.ForeColor = $uiCategoryScale }
	})

# Number of Backoffices (right, row 2)
$NodesBackoffices = New-Object System.Windows.Forms.Label
$NodesBackoffices.Name = "BackofficesLabel"
$NodesBackoffices.Text = "Backoffices: N/A"
$NodesBackoffices.AutoSize = $true
$NodesBackoffices.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
$NodesBackoffices.ForeColor = $uiTextSecondary
$NodesBackoffices.Margin = New-Object System.Windows.Forms.Padding(0)
$NodesBackoffices.Anchor = 'Right'
$topStripTable.Controls.Add($NodesBackoffices, 2, 1) | Out-Null
# Vertically center in the cell (real fix - see note at smsVersionLabel above)
$NodesBackoffices.AutoSize = $false
$NodesBackoffices.Dock = 'Fill'
$NodesBackoffices.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

$toolTip.SetToolTip($NodesBackoffices, "Shows count of Backoffices. Click to ping Backoffice nodes.")
& $applyInfoLabelStyle $NodesBackoffices
$NodesBackoffices.Add_Click({
		$script:LastActivity = Get-Date
		$orig = $NodesBackoffices.ForeColor
		$NodesBackoffices.ForeColor = $uiAccent
		try { Ping_All_Nodes -NodeType "Backoffice" -StoreNumber $StoreNumber }
		finally { $NodesBackoffices.ForeColor = $uiTextSecondary }
	})

# ------------------------------------------------------------------------------------------------
# Top-strip layout scriptblock (kept name for compatibility; now it sizes cards cleanly)
# ------------------------------------------------------------------------------------------------
$layoutTopStrip = {
	$form = $script:MainForm # may run from a timer tick during a dialog
	if (-not $form -or $form.IsDisposed) { return }

	try
	{
		# Fit the top strip "card" to the form width with your side padding
		$rowW = [math]::Max(200, $form.ClientSize.Width - (2 * $padSide))
		$topStripCard.Left = $padSide
		$topStripCard.Top = $bannerPanel.Bottom + $padTopStrip
		$topStripCard.Width = $rowW
	}
	catch { }
}

# Run once + re-run on any resize (still compatible with your other layout blocks)
& $layoutTopStrip
$form.add_Resize({ & $layoutTopStrip })
