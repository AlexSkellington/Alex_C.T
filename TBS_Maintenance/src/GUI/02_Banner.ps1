# ===================================================================================================
# ENHANCED HEADER - Modern app bar with version, logo area, and status indicators
# ===================================================================================================
$bannerPanel = New-Object System.Windows.Forms.Panel
$bannerPanel.Dock = 'Top'
$bannerPanel.Height = 56
$bannerPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bannerPanel.Padding = New-Object System.Windows.Forms.Padding(0)

# Gradient paint (sleek dark gradient with accent stripe)
$bannerPanel.add_Paint({
		param ($s, $e)
		try
		{
			$rect = New-Object System.Drawing.Rectangle(0, 0, $s.Width, $s.Height)
			
			# Main gradient
			$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
				$rect,
				[System.Drawing.Color]::FromArgb(25, 25, 30),
				[System.Drawing.Color]::FromArgb(45, 45, 55),
				[System.Drawing.Drawing2D.LinearGradientMode]::Horizontal
			)
			$e.Graphics.FillRectangle($brush, $rect)
			$brush.Dispose()
			
			# Left accent stripe (colorful indicator)
			$accentRect = New-Object System.Drawing.Rectangle(0, 0, 4, $s.Height)
			$accentBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
				$accentRect,
				[System.Drawing.Color]::FromArgb(0, 122, 204),
				[System.Drawing.Color]::FromArgb(0, 180, 255),
				[System.Drawing.Drawing2D.LinearGradientMode]::Vertical
			)
			$e.Graphics.FillRectangle($accentBrush, $accentRect)
			$accentBrush.Dispose()
			
			# Subtle bottom line
			$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 60, 60))
			$e.Graphics.DrawLine($pen, 4, $s.Height - 1, $s.Width, $s.Height - 1)
			$pen.Dispose()
		}
		catch { }
	})

$form.Controls.Add($bannerPanel)

# ===================================================================================================
# Banner Left Section - App Icon + Title
# ===================================================================================================
# App icon circle
$appIconPanel = New-Object System.Windows.Forms.Panel
$appIconPanel.Location = New-Object System.Drawing.Point(12, 12)
$appIconPanel.Size = New-Object System.Drawing.Size(32, 32)
$appIconPanel.BackColor = [System.Drawing.Color]::Transparent
$appIconPanel.add_Paint({
		param ($s, $e)
		try
		{
			$e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
			$cx = $s.Width / 2.0
			$cy = $s.Height / 2.0
			
			# Blue disc (vertical gradient for a little depth)
			$disc = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
				(New-Object System.Drawing.Rectangle(1, 1, $s.Width - 2, $s.Height - 2)),
				[System.Drawing.Color]::FromArgb(0, 160, 235),
				[System.Drawing.Color]::FromArgb(0, 105, 190),
				0.0)
			$e.Graphics.FillEllipse($disc, 1, 1, $s.Width - 2, $s.Height - 2)
			$disc.Dispose()
			
			# White hex nut (flat-top hexagon)
			$r = 7.5
			$hex = New-Object System.Drawing.Drawing2D.GraphicsPath
			$hex.AddPolygon(@(
				(New-Object System.Drawing.PointF($cx + $r, $cy)),
				(New-Object System.Drawing.PointF($cx + $r / 2, $cy + ($r * 0.866))),
				(New-Object System.Drawing.PointF($cx - $r / 2, $cy + ($r * 0.866))),
				(New-Object System.Drawing.PointF($cx - $r, $cy)),
				(New-Object System.Drawing.PointF($cx - $r / 2, $cy - ($r * 0.866))),
				(New-Object System.Drawing.PointF($cx + $r / 2, $cy - ($r * 0.866)))
			))
			$white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
			$e.Graphics.FillPath($white, $hex)
			$hex.Close()
			$hex.Dispose()
			
			# Punch out the center so it reads as a nut
			$punch = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 122, 204))
			$e.Graphics.FillEllipse($punch, $cx - 3.2, $cy - 3.2, 6.4, 6.4)
			$white.Dispose()
			$punch.Dispose()
		}
		catch { }
	})
$bannerPanel.Controls.Add($appIconPanel)

# Title label
$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Name = "BannerTitle"
$bannerLabel.Text = "TBS Maintenance Script"
$bannerLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]13, [System.Drawing.FontStyle]::Bold)
$bannerLabel.ForeColor = [System.Drawing.Color]::White
$bannerLabel.BackColor = [System.Drawing.Color]::Transparent
$bannerLabel.AutoSize = $true
$bannerLabel.Location = New-Object System.Drawing.Point(54, 16)
$bannerPanel.Controls.Add($bannerLabel)

# ===================================================================================================
# Banner Right Section - Help Button (right-aligned)
# ===================================================================================================
$helpButton = New-Object System.Windows.Forms.Button
$helpButton.Name = "HelpButton"
$helpButton.Text = "?"
$helpButton.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$helpButton.Size = New-Object System.Drawing.Size(32, 28)
$helpButton.Location = New-Object System.Drawing.Point(($bannerPanel.Width - 60), 14)
$helpButton.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 60)
$helpButton.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$helpButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$helpButton.FlatAppearance.BorderSize = 0
$helpButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$helpButton.Margin = New-Object System.Windows.Forms.Padding(0)
$helpButton.UseVisualStyleBackColor = $false
$helpButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
$helpButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 90, 150)
$helpButton.Add_MouseEnter({ $helpButton.ForeColor = [System.Drawing.Color]::White })
$helpButton.Add_MouseLeave({ $helpButton.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180) })
$toolTip.SetToolTip($helpButton, "Open TBS Helpdesk (helpdesk.tecnicasystems.com)")

# Hover effect
$helpButton.add_MouseEnter({
		param ($s, $e)
		try { $s.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204) }
		catch { }
	})
$helpButton.add_MouseLeave({
		param ($s, $e)
		try { $s.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 60) }
		catch { }
	})
$helpButton.add_Click({
		$script:LastActivity = Get-Date
		try { Start-Process "https://helpdesk.tecnicasystems.com" }
		catch {
			[System.Windows.Forms.MessageBox]::Show(
				"Couldn't open helpdesk.tecnicasystems.com`r`n$($_.Exception.Message)",
				"Open Link",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			) | Out-Null
		}
	})

$bannerPanel.Controls.Add($helpButton)

# ===================================================================================================
# Banner Right - Live Clock
# ===================================================================================================
$clockLabel = New-Object System.Windows.Forms.Label
$clockLabel.Name = "BannerClock"
$clockLabel.Text = (Get-Date -Format "h:mm:ss tt")
$clockLabel.Font = New-Object System.Drawing.Font("Consolas", [float]9, [System.Drawing.FontStyle]::Regular)
$clockLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$clockLabel.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 55)
$clockLabel.AutoSize = $true
$clockLabel.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
$clockLabel.Location = New-Object System.Drawing.Point(($bannerPanel.Width - 180), 16)
$clockLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$bannerPanel.Controls.Add($clockLabel)

# Clock update timer
$clockTimer = New-Object System.Windows.Forms.Timer
$clockTimer.Interval = 1000
$clockTimer.add_Tick({
		try
		{
			if ($clockLabel -and -not $clockLabel.IsDisposed)
			{
				$clockLabel.Text = (Get-Date -Format "h:mm:ss tt")
			}
		}
		catch { }
	})
$clockTimer.Start()

# ===================================================================================================
# Banner Center - Version + Status
# ===================================================================================================
$bannerStatusPanel = New-Object System.Windows.Forms.Panel
$bannerStatusPanel.Location = New-Object System.Drawing.Point(350, 10)
$bannerStatusPanel.Size = New-Object System.Drawing.Size(250, 36)
$bannerStatusPanel.BackColor = [System.Drawing.Color]::Transparent

# Single clean line: version + build date (the old layout stacked a tiny
# 7pt "Build: <date>" line under the version, which looked out of place)
$bannerDateLabel = New-Object System.Windows.Forms.Label
$bannerDateLabel.Name = "BannerDate"
# Store identity (the version + build date are shown in the status bar at
# the bottom - the banner shows which store this console manages instead).
# Get_Store_And_Database_Info updates this label once the store data resolves.
$bannerDateLabel.Text = "Store Maintenance Console"
try
{
	$bn0 = $script:FunctionResults['StoreNumber']
	if ($bn0 -and "$bn0" -ne 'N/A') { $bannerDateLabel.Text = "Store $bn0" }
}
catch { }
$bannerDateLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]9, [System.Drawing.FontStyle]::Bold)
$bannerDateLabel.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$bannerDateLabel.BackColor = [System.Drawing.Color]::Transparent
$bannerDateLabel.AutoSize = $true
$bannerDateLabel.Location = New-Object System.Drawing.Point(0, 10)
$bannerStatusPanel.Controls.Add($bannerDateLabel)

# Environment badge - visible when the production storeman share was NOT used
# with a fresh key. Two states (see the "Locate Base Path" detection block at
# the top of the script):
#   KEY EXPIRED (red)  - store data present, key older than 14 days, re-key needed
#   TEST      (orange) - plain test/fallback environment
$bannerTestBadge = New-Object System.Windows.Forms.Label
$bannerTestBadge.Name = "BannerTestBadge"
if ($script:StoremanKeyStale)
{
	$bannerTestBadge.Text = " KEY EXPIRED "
	$bannerTestBadge.ForeColor = [System.Drawing.Color]::White
	$bannerTestBadge.BackColor = [System.Drawing.Color]::FromArgb(178, 44, 34)
}
else
{
	$bannerTestBadge.Text = " TEST "
	$bannerTestBadge.ForeColor = [System.Drawing.Color]::FromArgb(40, 25, 0)
	$bannerTestBadge.BackColor = [System.Drawing.Color]::FromArgb(255, 165, 0)
}
$bannerTestBadge.Font = New-Object System.Drawing.Font("Segoe UI", [float]8, [System.Drawing.FontStyle]::Bold)
$bannerTestBadge.AutoSize = $true
$bannerTestBadge.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
$bannerTestBadge.Visible = [bool]$script:IsTestEnvironment
$toolTip.SetToolTip($bannerTestBadge, $(if ($script:StoremanKeyStale) { "KEY EXPIRED - " + $script:StoremanSourceReason + ". Using: $BasePath" } elseif ($script:IsTestEnvironment) { "TEST ENVIRONMENT - " + $script:StoremanSourceReason + ". Using: $BasePath" } else { "Production storeman share in use." }))
$toolTip.SetToolTip($bannerDateLabel, "Storeman source: " + $script:StoremanSource + " - " + $script:StoremanSourceReason)
$bannerStatusPanel.Controls.Add($bannerTestBadge)

$bannerPanel.Controls.Add($bannerStatusPanel)

# ===================================================================================================
# Banner layout helper
#   The right-side banner controls (clock, help button) are created before the docked banner
#   has its final width, so their creation-time positions were wrong (help button ended up
#   over the title, clock over the icon). This helper pins them to the banner's real width
#   and vertically centers EVERY banner element (icon, title, version, clock, help) on the
#   banner's center line so they all sit on the same plane. Runs again on every resize.
# ===================================================================================================
$layoutBanner = {
	if (-not $form -or $form.IsDisposed) { return }
	if (-not $bannerPanel -or $bannerPanel.IsDisposed) { return }
	try
	{
		$w = [Math]::Max($bannerPanel.Width, $form.ClientSize.Width)
		$h = $bannerPanel.Height
		# y-offset that centers a control on the banner's center line
		$centerY = {
				param ($ctrl)
				$ch = $ctrl.Height
				if ($ch -lt 1) { $ch = [System.Windows.Forms.TextRenderer]::MeasureText($ctrl.Text, $ctrl.Font).Height }
				[Math]::Max(0, [int](($h - $ch) / 2))
			}
		$bannerLabel.Location = New-Object System.Drawing.Point(54, (& $centerY $bannerLabel))
		$helpButton.SetBounds($w - 48, (& $centerY $helpButton), 32, $helpButton.Height)
		$clockLabel.Location = New-Object System.Drawing.Point(($w - 158), (& $centerY $clockLabel))
		# Keep the version block right next to the title
		$bannerStatusPanel.Location = New-Object System.Drawing.Point(($bannerLabel.Right + 16), 0)
		$bannerStatusPanel.Height = $h
		$bannerDateLabel.Location = New-Object System.Drawing.Point(0, (& $centerY $bannerDateLabel))
		# Always position the TEST badge next to the version (its visibility
		# is controlled separately; before the form is shown, .Visible
		# misreports False, so it must not gate positioning)
		$bannerTestBadge.Location = New-Object System.Drawing.Point(($bannerDateLabel.Right + 8), (& $centerY $bannerTestBadge))
		$appIconPanel.Location = New-Object System.Drawing.Point(12, (& $centerY $appIconPanel))
	}
	catch { }
}
$form.add_Resize({ & $layoutBanner })
# Re-apply once the form is actually visible (control states like .Visible
# and measured sizes are only reliable after the first Show)
$form.add_Shown({ & $layoutBanner })

# Apply the banner layout now (force the form's layout pass first so the docked
# banner already has its real width)
try { $form.PerformLayout() } catch { }
& $layoutBanner
