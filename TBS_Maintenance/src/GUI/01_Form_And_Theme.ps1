# ===================================================================================================
#                                  SECTION: GUI - Form and Theme
# ---------------------------------------------------------------------------------------------------
# Description:
#   First of the GUI files (src\GUI\01..14 run in order and share script-scope variables):
#   visual styles, layout constants, the colour palette, the shared tooltip, the main $form,
#   the app logo/icon, and the idle/busy state the watchdog in 06 uses.
#
# Variables (read by the later GUI files):
#   - $form, $toolTip, $script:TbsAppIcon, $script:DrawTbsLogo
#   - $padSide, $padTopStrip, $gapCards, $logInnerPad, $minLogH   layout constants
#   - $ui*                                                         colour palette
#   - $script:LastActivity / IsBusy / BusyReason / IdleIgnoreJobNames / IdleHardCloseAfterMinutes
#
# Notes:
#   - The process is NOT DPI-aware (Windows scales the window on high-DPI screens). The layout
#     code below uses fixed pixel sizes; making the process DPI-aware needs that reworked first.
# ===================================================================================================

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ---- Layout constants (shared by the layout blocks in 04, 05 and 14) ---------------------------
if (-not $padSide) { $padSide = 50 } # left/right margin
$padTopStrip = 10 # spacing below header
$gapCards = 10 # vertical gap between cards
$logInnerPad = 10 # padding inside the log frame
$minLogH = 220 # minimum log height

# ===================================================================================================
# Colour palette
# ===================================================================================================
$uiAccent = [System.Drawing.Color]::FromArgb(0, 122, 204)          # Windows 11 accent blue
$uiAccentLight = [System.Drawing.Color]::FromArgb(230, 243, 255)   # Light blue hover
$uiAccentDark = [System.Drawing.Color]::FromArgb(0, 90, 158)       # Dark blue for pressed

$uiBg = [System.Drawing.Color]::FromArgb(242, 242, 242)            # Windows 11 light gray bg
$uiCardBg = [System.Drawing.Color]::White
$uiCardBorder = [System.Drawing.Color]::FromArgb(229, 229, 229)    # Subtle border
$uiHeaderText = [System.Drawing.Color]::White
$uiTextPrimary = [System.Drawing.Color]::FromArgb(32, 32, 32)      # Dark primary text
$uiTextSecondary = [System.Drawing.Color]::FromArgb(102, 102, 102) # Secondary text
$uiTextMuted = [System.Drawing.Color]::FromArgb(153, 153, 153)     # Muted text

# Status colors
$uiStatusSuccess = [System.Drawing.Color]::FromArgb(16, 124, 64)    # Green
$uiStatusWarning = [System.Drawing.Color]::FromArgb(184, 134, 11)   # Amber
$uiStatusError = [System.Drawing.Color]::FromArgb(200, 40, 40)      # Red
$uiStatusInfo = [System.Drawing.Color]::FromArgb(0, 122, 204)       # Blue

# Tool category colors (visual distinction)
$uiCategoryServer = [System.Drawing.Color]::FromArgb(76, 175, 80)    # Green
$uiCategoryLane = [System.Drawing.Color]::FromArgb(33, 150, 243)     # Blue
$uiCategoryScale = [System.Drawing.Color]::FromArgb(156, 39, 176)    # Purple
$uiCategoryGeneral = [System.Drawing.Color]::FromArgb(255, 152, 0)   # Orange

# ------------------------------------------------------------------------------------------------
# Shared ToolTip
# ------------------------------------------------------------------------------------------------
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000
$toolTip.InitialDelay = 300
$toolTip.ReshowDelay = 200
$toolTip.ShowAlways = $true
$toolTip.BackColor = [System.Drawing.Color]::White

# ------------------------------------------------------------------------------------------------
# Main Form (modern palette + DPI scaling)
# ------------------------------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "TBS Maintenance Script" # Modern: keep title clean; details remain visible in header/subtext
$form.Size = New-Object System.Drawing.Size(1006, 570)
$form.MinimumSize = New-Object System.Drawing.Size(800, 520)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor = $uiBg
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

# Reduce flicker on resize / heavy updates (modern smoothness)
try
{
	$prop = $form.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
	if ($prop) { $prop.SetValue($form, $true, $null) }
}
catch { }

# ----------------------------------------------------------------------------------------------
# App logo (blue gradient disc + white hex nut): title bar / taskbar icon and the banner badge
# ----------------------------------------------------------------------------------------------
# Draws the logo into the square (X, Y, Size) of any Graphics. Uses ::new(...) on purpose:
# in command mode 'New-Object PointF($a + $b, $c)' parses as $a + @($b, $c) and throws, which
# is why neither the icon nor the banner badge was ever drawn before v2.7.0.
$script:DrawTbsLogo = {
	param ([System.Drawing.Graphics]$Graphics, [float]$X, [float]$Y, [float]$Size)

	$Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$disc = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
		[System.Drawing.RectangleF]::new($X, $Y, $Size, $Size),
		[System.Drawing.Color]::FromArgb(0, 160, 235),
		[System.Drawing.Color]::FromArgb(0, 105, 190),
		[float]90)
	$Graphics.FillEllipse($disc, $X, $Y, $Size, $Size)
	$disc.Dispose()

	# flat-top hexagon, then punch the center so it reads as a nut
	$cx = $X + $Size / 2
	$cy = $Y + $Size / 2
	$r = $Size * 0.25
	$k = 0.866
	$hex = New-Object System.Drawing.Drawing2D.GraphicsPath
	$hex.AddPolygon([System.Drawing.PointF[]]@(
			[System.Drawing.PointF]::new($cx + $r, $cy),
			[System.Drawing.PointF]::new($cx + $r / 2, $cy + $r * $k),
			[System.Drawing.PointF]::new($cx - $r / 2, $cy + $r * $k),
			[System.Drawing.PointF]::new($cx - $r, $cy),
			[System.Drawing.PointF]::new($cx - $r / 2, $cy - $r * $k),
			[System.Drawing.PointF]::new($cx + $r / 2, $cy - $r * $k)))
	$white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
	$Graphics.FillPath($white, $hex)
	$white.Dispose()
	$hex.Dispose()

	$hole = $Size * 0.21
	$punch = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 122, 204))
	$Graphics.FillEllipse($punch, $cx - $hole / 2, $cy - $hole / 2, $hole, $hole)
	$punch.Dispose()
}

try
{
	$__iconBmp = New-Object System.Drawing.Bitmap(32, 32)
	$__g = [System.Drawing.Graphics]::FromImage($__iconBmp)
	& $script:DrawTbsLogo $__g 1 1 30
	$__g.Dispose()
	$form.Icon = [System.Drawing.Icon]::FromHandle($__iconBmp.GetHicon())
	$script:TbsAppIcon = $form.Icon
	$__iconBmp.Dispose()
}
catch { }

# ------------------------------------------------------------------------------------------------
# -------------------- Idle Close (configurable) + Busy-Safe Watchdog (FULL FIX) --------------------
# ------------------------------------------------------------------------------------------------

# Config / state
if (-not $script:LastActivity) { $script:LastActivity = Get-Date }
$script:IsBusy = $false
$script:BusyReason = $null

# Jobs you want the idle system to IGNORE (supports wildcards)
# IMPORTANT: XE watcher job is named "XEWatcherJob_<StoreNumber>", so we ignore "XEWatcherJob_*"
$script:IdleIgnoreJobNames = @(
	'XEWatcherJob_*',
	'ClearXEFolderJob'
)

# Optional: hard-close after X minutes even if still busy (0 = disabled)
$script:IdleHardCloseAfterMinutes = 0

# Make sure the form sees keystrokes even when a control has focus
try { $form.KeyPreview = $true }
catch { }
