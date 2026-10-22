# ===================================================================================================
#                                       SECTION: Initialize GUI  (Modernized)
# ---------------------------------------------------------------------------------------------------
# Description:
#   Initializes and configures the graphical user interface components for the Store/Lane SQL Execution Tool.
#   Host mode is removed; strictly Store/Server/Lane (Store Mode only).
#
# Modern UI changes (behavior preserved):
#   - Gradient header bar + centered clickable title (still text-only clickable, still opens Helpdesk)
#   - Top "card" strip for the clickable labels (clean spacing, modern font)
#   - Framed log area (card border + borderless RichTextBox)
#   - Status bar for Busy + Idle countdown (new polish; does not affect logic)
#   - FIX: consolidated FormClosing so we DO NOT dispose timers/jobs if user cancels exit
# ===================================================================================================

# ------------------------------------------------------------------------------------------------
# WinForms basics (High DPI + Visual styles)
# ------------------------------------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# Make the process DPI-aware BEFORE any window is created so the UI renders
# crisply on high-DPI displays (no-op on 100% scaling).
try
{
	$tbsDpiCode = @'
[DllImport("user32.dll")]
public static extern bool SetProcessDPIAware();
'@
	if (-not ('TbsDpiHelper' -as [type]))
	{
		Add-Type -Namespace Tbs -Name DpiHelper -MemberDefinition $tbsDpiCode -ErrorAction SilentlyContinue
	}
	if ('TbsDpiHelper' -as [type]) { [void][Tbs.DpiHelper]::SetProcessDPIAware() }
}
catch { }

# ---- UI constants (kept as simple variables so other layout blocks can reuse them) -------------
# NOTE: Your other layout code (like $layoutBottom) often assumes $padSide exists.
if (-not $padSide) { $padSide = 50 } # left/right margin (keeps your original feel)
$padTopStrip = 10 # spacing below header
$cardPad = 12 # inner padding for "cards"
$gapCards = 10 # vertical gap between cards
$logInnerPad = 10 # padding inside the log frame
$minLogH = 220 # minimum log height (used by other layout)
# ===================================================================================================
# ENHANCED COLOR SCHEME (TBS Maintenance Script v2.6.7+ - Modern UI Theme)
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
# Legacy full-form refresh timer is left disabled by default because it can
# cause noticeable UI stalls on slower machines.
# ------------------------------------------------------------------------------------------------
$refreshTimer = New-Object System.Windows.Forms.Timer
$refreshTimer.Interval = 5000
$refreshTimer.add_Tick({
		try
		{
			if (-not $form -or $form.IsDisposed -or -not $form.Visible) { return }
			if ($statusStrip -and -not $statusStrip.IsDisposed) { $statusStrip.Refresh() }
		}
		catch { }
	})

# ------------------------------------------------------------------------------------------------
# ToolTip (kept, slightly refined)
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
# Runtime-generated app icon (blue disc + hex nut) for the title bar / taskbar
# ----------------------------------------------------------------------------------------------
try
{
	$__iconBmp = New-Object System.Drawing.Bitmap(32, 32)
	$__g = [System.Drawing.Graphics]::FromImage($__iconBmp)
	$__g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$__disc = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
		(New-Object System.Drawing.Rectangle(1, 1, 30, 30)),
		[System.Drawing.Color]::FromArgb(0, 160, 235),
		[System.Drawing.Color]::FromArgb(0, 105, 190), 0.0)
	$__g.FillEllipse($__disc, 1, 1, 30, 30)
	$__disc.Dispose()
	$__cx = 16.0; $__cy = 16.0; $__r = 8.0
	$__hex = New-Object System.Drawing.Drawing2D.GraphicsPath
	$__hex.AddPolygon(@(
		(New-Object System.Drawing.PointF($__cx + $__r, $__cy)),
		(New-Object System.Drawing.PointF($__cx + $__r / 2, $__cy + ($__r * 0.866))),
		(New-Object System.Drawing.PointF($__cx - $__r / 2, $__cy + ($__r * 0.866))),
		(New-Object System.Drawing.PointF($__cx - $__r, $__cy)),
		(New-Object System.Drawing.PointF($__cx - $__r / 2, $__cy - ($__r * 0.866))),
		(New-Object System.Drawing.PointF($__cx + $__r / 2, $__cy - ($__r * 0.866)))
	))
	$__white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
	$__g.FillPath($__white, $__hex)
	$__hex.Close()
	$__hex.Dispose()
	$__punch = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 122, 204))
	$__g.FillEllipse($__punch, $__cx - 3.4, $__cy - 3.4, 6.8, 6.8)
	$__white.Dispose()
	$__punch.Dispose()
	$__g.Dispose()
	$form.Icon = [System.Drawing.Icon]::FromHandle($__iconBmp.GetHicon())
	$script:TbsAppIcon = [System.Drawing.Icon]::Clone($form.Icon)
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
