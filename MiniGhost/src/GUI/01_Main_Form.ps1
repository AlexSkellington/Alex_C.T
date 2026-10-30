# ===================================================================================================
#                                  SECTION: GUI - Main Window
# ---------------------------------------------------------------------------------------------------
# Description:
#   First of the GUI files (src\GUI\01..11 run in order and share script-scope variables): the main
#   $form with its header bar, the "Current Info" tiles, the three action cards the buttons go into
#   (filled by 11_Layout_And_Tooltips) and the footer.
#
# Variables (read by the later GUI files and by functions that update the window):
#   - $form, $script:MainForm, $toolTip, $script:MiniGhostIcon
#   - $script:machineNameLabel, $script:storeNameLabel, $script:storeNumberLabel,
#     $script:ipAddressLabel        value labels of the info tiles (text = the value only)
#   - $script:CardIdentity, $script:CardDatabase, $script:CardSystem   button hosts of the cards
#   - $ui* colours
#
# Notes:
#   - Dialog functions name their own window $form: anything that can run while a dialog is open
#     (handlers, timers) must use $script:MainForm.
#   - Paint handlers only use their sender ($s) and colours stored on .Tag - they run in their
#     own scope.
# ===================================================================================================

# ---- Palette (same family as the TBS Maintenance Script) -----------------------------------------
$uiBg = [System.Drawing.Color]::FromArgb(243, 243, 243)
$uiCard = [System.Drawing.Color]::White
$uiBorder = [System.Drawing.Color]::FromArgb(229, 229, 229)
$uiText = [System.Drawing.Color]::FromArgb(32, 32, 32)
$uiTextMuted = [System.Drawing.Color]::FromArgb(110, 110, 110)
$uiAccentIdentity = [System.Drawing.Color]::FromArgb(0, 103, 192)  # blue
$uiAccentDatabase = [System.Drawing.Color]::FromArgb(126, 87, 214) # violet
$uiAccentSystem = [System.Drawing.Color]::FromArgb(230, 126, 34)   # orange

# ---- Main form -----------------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$script:MainForm = $form
$form.Text = "MiniGhost $VersionNumber"
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.MinimumSize = New-Object System.Drawing.Size(820, 560)
$form.Size = New-Object System.Drawing.Size(940, 620)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.BackColor = $uiBg
$form.ForeColor = $uiText
try
{
	$prop = $form.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
	if ($prop) { $prop.SetValue($form, $true, $null) }
}
catch { }

# Window / taskbar icon drawn from the logo
$script:MiniGhostIcon = $null
try
{
	$iconBmp = New-Object System.Drawing.Bitmap(32, 32)
	$iconG = [System.Drawing.Graphics]::FromImage($iconBmp)
	Draw_MiniGhost_Logo -Graphics $iconG -X 1 -Y 1 -Size 30
	$iconG.Dispose()
	$script:MiniGhostIcon = [System.Drawing.Icon]::FromHandle($iconBmp.GetHicon())
	$form.Icon = $script:MiniGhostIcon
}
catch { }

# Global tooltip provider (per-button tooltips)
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 20000
$toolTip.InitialDelay = 350
$toolTip.ReshowDelay = 150
$toolTip.ShowAlways = $true

# ---- Root layout: header / info / actions / footer -----------------------------------------------
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 4
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$mainLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$mainLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
[void]$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 72)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 92)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$form.Controls.Add($mainLayout)

# ---- Header bar ----------------------------------------------------------------------------------
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Fill'
$headerPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 36)
$headerPanel.add_Paint({
		param ($s, $e)
		try
		{
			$rect = New-Object System.Drawing.Rectangle(0, 0, $s.Width, $s.Height)
			$bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, [System.Drawing.Color]::FromArgb(26, 26, 32), [System.Drawing.Color]::FromArgb(48, 44, 66), [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
			$e.Graphics.FillRectangle($bg, $rect)
			$bg.Dispose()
			$stripe = New-Object System.Drawing.Rectangle(0, 0, 4, $s.Height)
			$sb = New-Object System.Drawing.Drawing2D.LinearGradientBrush($stripe, [System.Drawing.Color]::FromArgb(126, 87, 214), [System.Drawing.Color]::FromArgb(0, 120, 215), [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
			$e.Graphics.FillRectangle($sb, $stripe)
			$sb.Dispose()
		}
		catch { }
	})

$logoPanel = New-Object System.Windows.Forms.Panel
$logoPanel.Location = New-Object System.Drawing.Point(20, 16)
$logoPanel.Size = New-Object System.Drawing.Size(40, 40)
$logoPanel.BackColor = [System.Drawing.Color]::Transparent
$logoPanel.add_Paint({
		param ($s, $e)
		try { Draw_MiniGhost_Logo -Graphics $e.Graphics -X 1 -Y 1 -Size ([float]([Math]::Min($s.Width, $s.Height) - 2)) }
		catch { }
	})
$headerPanel.Controls.Add($logoPanel)

$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Text = "MiniGhost"
$bannerLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
$bannerLabel.ForeColor = [System.Drawing.Color]::White
$bannerLabel.BackColor = [System.Drawing.Color]::Transparent
$bannerLabel.AutoSize = $true
$bannerLabel.Location = New-Object System.Drawing.Point(70, 10)
$headerPanel.Controls.Add($bannerLabel)

$bannerSubLabel = New-Object System.Windows.Forms.Label
$bannerSubLabel.Text = "Re-identify a cloned TBS machine: name, store, network, INIs and SQL"
$bannerSubLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$bannerSubLabel.ForeColor = [System.Drawing.Color]::FromArgb(185, 185, 200)
$bannerSubLabel.BackColor = [System.Drawing.Color]::Transparent
$bannerSubLabel.AutoSize = $true
$bannerSubLabel.Location = New-Object System.Drawing.Point(72, 42)
$headerPanel.Controls.Add($bannerSubLabel)

# Right side: node role + version pills
$pillPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$pillPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$pillPanel.WrapContents = $false
$pillPanel.AutoSize = $true
$pillPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$pillPanel.BackColor = [System.Drawing.Color]::Transparent
$pillPanel.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Right'
$pillPanel.Location = New-Object System.Drawing.Point(600, 23)

$versionPill = New-Object System.Windows.Forms.Label
$versionPill.Text = "v$VersionNumber"
$versionPill.AutoSize = $true
$versionPill.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$versionPill.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 230)
$versionPill.BackColor = [System.Drawing.Color]::FromArgb(62, 60, 80)
$versionPill.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$versionPill.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

$rolePill = New-Object System.Windows.Forms.Label
$roleText = switch -Regex ([string]$script:NodeRole)
{
	'^(?i)Lane$' { 'LANE'; break }
	'^(?i)StoreServer$' { 'STORE SERVER'; break }
	'^(?i)HostServer$' { 'HOST SERVER'; break }
	default { 'ROLE UNKNOWN' }
}
$rolePill.Text = $roleText
$rolePill.AutoSize = $true
$rolePill.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$rolePill.ForeColor = [System.Drawing.Color]::White
$rolePill.BackColor = switch -Regex ([string]$script:NodeRole)
{
	'^(?i)Lane$' { [System.Drawing.Color]::FromArgb(0, 103, 192); break }
	'^(?i)(StoreServer|HostServer)$' { [System.Drawing.Color]::FromArgb(16, 124, 64); break }
	default { [System.Drawing.Color]::FromArgb(110, 110, 120) }
}
$rolePill.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$rolePill.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$toolTip.SetToolTip($rolePill, "Node role from Startup.ini DBNAME (LANESQL = lane, STORESQL = store server, HOSTSQL = host server). It decides which machine-name formats are accepted.")

$pillPanel.Controls.Add($versionPill)
$pillPanel.Controls.Add($rolePill)
$headerPanel.Controls.Add($pillPanel)
# keep the pills right-aligned when the window is resized (Anchor alone does not move an AutoSize flow panel reliably)
$headerPanel.add_Resize({
		param ($s, $e)
		try
		{
			foreach ($c in @($s.Controls))
			{
				if ($c -is [System.Windows.Forms.FlowLayoutPanel]) { $c.Left = [Math]::Max(300, $s.ClientSize.Width - $c.Width - 20); $c.Top = [int](($s.ClientSize.Height - $c.Height) / 2) }
			}
		}
		catch { }
	})
$mainLayout.Controls.Add($headerPanel, 0, 0)

# ---- Current Info: four tiles (caption + value) --------------------------------------------------
$infoCard = New_MiniGhost_Card -Margin (New-Object System.Windows.Forms.Padding(16, 14, 16, 0))
$infoCard.Dock = 'Fill'

$infoLayout = New-Object System.Windows.Forms.TableLayoutPanel
$infoLayout.Dock = 'Fill'
$infoLayout.ColumnCount = 4
$infoLayout.RowCount = 2
$infoLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$infoLayout.Padding = New-Object System.Windows.Forms.Padding(8, 10, 8, 8)
$infoLayout.BackColor = [System.Drawing.Color]::White
foreach ($pct in @(25, 25, 22, 28)) { [void]$infoLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, $pct))) }
[void]$infoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
[void]$infoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$infoCard.Controls.Add($infoLayout)

$currentIP = if ($currentConfigs -and @($currentConfigs).Count -gt 0) { [string](@(@($currentConfigs)[0].IPv4Address)[0].IPAddress) }
else { "Not found" }
if ([string]::IsNullOrWhiteSpace($currentIP)) { $currentIP = "Not found" }

$infoTiles = @(
	@{ Caption = 'MACHINE NAME'; Value = [string]$currentMachineName; Var = 'machineNameLabel' },
	@{ Caption = 'STORE NAME'; Value = [string]$storeName; Var = 'storeNameLabel' },
	@{ Caption = 'STORE NUMBER'; Value = [string]$currentStoreNumber; Var = 'storeNumberLabel' },
	@{ Caption = 'IP ADDRESS'; Value = $currentIP; Var = 'ipAddressLabel' }
)
for ($col = 0; $col -lt $infoTiles.Count; $col++)
{
	$tile = $infoTiles[$col]
	$cap = New-Object System.Windows.Forms.Label
	$cap.Text = $tile.Caption
	$cap.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8)
	$cap.ForeColor = $uiTextMuted
	$cap.Dock = 'Fill'
	$cap.Margin = New-Object System.Windows.Forms.Padding(10, 0, 6, 0)
	$cap.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
	$infoLayout.Controls.Add($cap, $col, 0)
	
	$val = New-Object System.Windows.Forms.Label
	$val.Text = $(if ([string]::IsNullOrWhiteSpace($tile.Value)) { 'N/A' } else { $tile.Value })
	$val.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
	$val.ForeColor = $uiText
	$val.Dock = 'Fill'
	$val.AutoSize = $false
	$val.AutoEllipsis = $true
	$val.Margin = New-Object System.Windows.Forms.Padding(10, 0, 6, 0)
	$val.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
	$infoLayout.Controls.Add($val, $col, 1)
	Set-Variable -Scope Script -Name $tile.Var -Value $val
}
# the full text is in the tooltip when a long value is cut with "..."
foreach ($v in @($script:machineNameLabel, $script:storeNameLabel, $script:storeNumberLabel, $script:ipAddressLabel))
{
	$v.add_TextChanged({ param ($s, $e) try { $toolTip.SetToolTip($s, $s.Text) } catch { } })
	$toolTip.SetToolTip($v, $v.Text)
}
$mainLayout.Controls.Add($infoCard, 0, 1)

# ---- Actions: three cards (Identity / Database / System) -----------------------------------------
$actionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionsLayout.Dock = 'Fill'
$actionsLayout.ColumnCount = 3
$actionsLayout.RowCount = 1
$actionsLayout.Margin = New-Object System.Windows.Forms.Padding(10, 8, 10, 6)
$actionsLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
foreach ($pct in @(33.33, 33.33, 33.34)) { [void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, $pct))) }
[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainLayout.Controls.Add($actionsLayout, 0, 2)

$cardIdentityPanel = New_MiniGhost_Action_Card -Title 'Identity' -Subtitle 'Machine name, store number, network' -Accent $uiAccentIdentity
$cardDatabasePanel = New_MiniGhost_Action_Card -Title 'Database' -Subtitle 'SQL tables for this store / node' -Accent $uiAccentDatabase
$cardSystemPanel = New_MiniGhost_Action_Card -Title 'System' -Subtitle 'Registry, results, restart' -Accent $uiAccentSystem
$actionsLayout.Controls.Add($cardIdentityPanel, 0, 0)
$actionsLayout.Controls.Add($cardDatabasePanel, 1, 0)
$actionsLayout.Controls.Add($cardSystemPanel, 2, 0)
$script:CardIdentity = $cardIdentityPanel.Tag
$script:CardDatabase = $cardDatabasePanel.Tag
$script:CardSystem = $cardSystemPanel.Tag

# ---- Footer --------------------------------------------------------------------------------------
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = 'Fill'
$footerPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$footerPanel.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 238)

$footerLeft = New-Object System.Windows.Forms.Label
$footerLeft.Dock = 'Left'
$footerLeft.AutoSize = $false
$footerLeft.Width = 420
$footerLeft.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$footerLeft.Padding = New-Object System.Windows.Forms.Padding(16, 0, 0, 0)
$footerLeft.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$footerLeft.ForeColor = $uiTextMuted
$footerLeft.Text = "Every action asks before it changes anything."

$footerRight = New-Object System.Windows.Forms.Label
$footerRight.Dock = 'Fill'
$footerRight.AutoSize = $false
$footerRight.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$footerRight.Padding = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
$footerRight.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$footerRight.ForeColor = $uiTextMuted
$footerRight.Text = "Created by Alex_C.T   |   Version $VersionNumber   |   Revised $VersionDate   |   PowerShell $PowerShellVersion"

$footerPanel.Controls.Add($footerRight)
$footerPanel.Controls.Add($footerLeft)
$mainLayout.Controls.Add($footerPanel, 0, 3)
