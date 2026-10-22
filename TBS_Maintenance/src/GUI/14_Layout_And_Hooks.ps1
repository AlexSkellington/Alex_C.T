######################################################################################################################
#
# Bottom buttons + logbox layout  (FIXED for StatusStrip + modern top strip)
#
######################################################################################################################

# -------- Bottom buttons + logbox layout (expand buttons, tight gap) --------
$padSide = 50 # left/right margin for logbox & button row
$padBottom = 10 # distance from buttons to bottom edge (ABOVE status strip)
$gapAboveButtons = 10 # vertical gap between logbox and buttons
$btnH = 48 # button height (slightly shorter = more modern)
$btnMinW = 180 # minimum button width so labels never wrap
$gapHPreferred = 18 # preferred horizontal gap between buttons
$gapHMin = 10 # absolute minimum horizontal gap
$minLogH = 140 # don't let the logbox collapse
$logInnerPad = 10 # if you use a log frame panel, this pads the inner log box

# keep WinForms from fighting our math
$ServerToolsButton.Anchor = 'Bottom,Left'
$LaneToolsButton.Anchor = 'Bottom,Left'
$ScaleToolsButton.Anchor = 'Bottom,Left'
$GeneralToolsButton.Anchor = 'Bottom,Left'

# Log frame resizes with window
if ($logFrame) { $logFrame.Anchor = 'Top,Left,Right,Bottom' }

# If you have a log frame panel in your newer UI, it will be resized too.
# If not, we will just resize $logBox directly.
if ($logBox) { $logBox.Anchor = 'Top,Left,Right,Bottom' }

$layoutBottom = {
	if (-not $form -or $form.IsDisposed) { return }

	# --- Compute shared row width (log + button row align) ---------------------
	$rowW = [math]::Max(200, $form.ClientSize.Width - (2 * $padSide))

	# --- Account for StatusStrip (THIS fixes the "buttons cut off" issue) -----
	$bottomReserve = 0
	try
	{
		if ($statusStrip -and -not $statusStrip.IsDisposed) { $bottomReserve = $statusStrip.Height }
	}
	catch { $bottomReserve = 0 }

	# --- Determine log top dynamically (banner + card + margin) -------------
	$logTop = $bannerPanel.Bottom + $topStripCard.Height + 10
	try
	{
		if ($topStripCard -and -not $topStripCard.IsDisposed)
		{
			$logTop = $topStripCard.Bottom + 10
		}
	}
	catch { }

	# --- Button row Y (above status strip) ------------------------------------
	$btnY = $form.ClientSize.Height - $bottomReserve - $btnH - $padBottom

	# --- Lay out log area (with header) ---------------------------------------
	$hasLogFrame = $false
	try { $hasLogFrame = ($logFrame -and -not $logFrame.IsDisposed) }
	catch { $hasLogFrame = $false }

	# Header height
	$lgHeaderH = 28
	$lgPadding = 10

	if ($hasLogFrame)
	{
		# Frame uses the full row width
		$logFrame.Left = $padSide
		$logFrame.Top = $logTop
		$logFrame.Width = $rowW

		# Height ensures a constant gap above buttons (accounting for header)
		$logFrame.Height = [math]::Max($minLogH, $btnY - $logFrame.Top - $gapAboveButtons)

		# Inner logBox inset for clean "card" look (accounting for header + padding)
		$logBox.Left = $lgPadding
		$logBox.Top = $lgHeaderH + $lgPadding
		$logBox.Width = [math]::Max(50, $logFrame.Width - (2 * $lgPadding))
		$logBox.Height = [math]::Max(50, $logFrame.Height - $lgHeaderH - (2 * $lgPadding))
	}
	else
	{
		# Plain logBox layout
		$logBox.Left = $padSide
		$logBox.Top = $logTop
		$logBox.Width = $rowW
		$logBox.Height = [math]::Max($minLogH, $btnY - $logBox.Top - $gapAboveButtons)
	}

	# --- Compute dynamic button width so the four buttons fill the row --------
	$gapH = $gapHPreferred
	$btnW = [math]::Floor(($rowW - (3 * $gapH)) / 4)

	# If too small, clamp width and reduce gaps (not below gapHMin)
	if ($btnW -lt $btnMinW)
	{
		$btnW = $btnMinW
		$gapH = [math]::Floor(($rowW - (4 * $btnW)) / 3)
		if ($gapH -lt $gapHMin) { $gapH = $gapHMin }
	}
	else
	{
		# Distribute leftover pixels (from integer math) into gaps
		$used = (4 * $btnW) + (3 * $gapH)
		$extra = $rowW - $used
		if ($extra -gt 0) { $gapH += [math]::Floor($extra / 3) }
	}

	# --- Position buttons left->right; group width == logbox width ------------
	$x = $padSide
	$ServerToolsButton.SetBounds($x, $btnY, $btnW, $btnH); $x += $btnW + $gapH
	$LaneToolsButton.SetBounds($x, $btnY, $btnW, $btnH); $x += $btnW + $gapH
	$ScaleToolsButton.SetBounds($x, $btnY, $btnW, $btnH); $x += $btnW + $gapH
	$GeneralToolsButton.SetBounds($x, $btnY, $btnW, $btnH)

	# Keep right-side labels aligned with the logbox edge
	& $layoutTopStrip

	# Ensure status strip stays visible
	try { if ($statusStrip -and -not $statusStrip.IsDisposed) { $statusStrip.BringToFront() } }
	catch { }
}

# prevent shrinking so far that 4*minW + 3*minGap won't fit
$minW = (2 * $padSide) + (4 * $btnMinW) + (3 * $gapHMin)
if ($form.MinimumSize.Width -lt $minW)
{
	$form.MinimumSize = New-Object System.Drawing.Size($minW, [math]::Max($form.MinimumSize.Height, 560))
}

# run now + on resize
& $layoutBottom
$form.add_Resize({ & $layoutBottom })

######################################################################################################################
#
# Activity hooks for controls OUTSIDE the main form
#   06_Busy_Idle_Activity already hooks the form, every child control and any control added to
#   the form later (ControlAdded). The menus and the protocol window are separate top-level
#   windows, so they get their own hooks here (the form, log box and buttons used to be hooked
#   a second and third time from this list).
#
######################################################################################################################
foreach ($__c in @($ContextMenuServer, $ContextMenuLane, $ContextMenuScale, $ContextMenuGeneral, $global:ProtocolForm, $global:ProtocolGrid))
{
	if ($__c -isnot [System.Windows.Forms.Control]) { continue }
	try
	{
		$__c.Add_MouseMove({ $script:LastActivity = Get-Date })
		$__c.Add_KeyDown({ $script:LastActivity = Get-Date })
		$__c.Add_Click({ $script:LastActivity = Get-Date })
	}
	catch { }
}
