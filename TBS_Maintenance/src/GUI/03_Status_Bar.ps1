# ===================================================================================================
# ENHANCED STATUS BAR - Modern flat design with more information
# ===================================================================================================
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.SizingGrip = $false
$statusStrip.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$statusStrip.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
$statusStrip.Padding = New-Object System.Windows.Forms.Padding(10, 2, 10, 2)

# Left: Status indicator with icon area
$stLeft = New-Object System.Windows.Forms.ToolStripStatusLabel
$stLeft.Text = "Ready"
$stLeft.Spring = $false
$stLeft.AutoSize = $true
$stLeft.ForeColor = $uiStatusSuccess
$stLeft.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
$stLeft.Padding = New-Object System.Windows.Forms.Padding(0, 0, 20, 0)

# Center: Current operation
$stOperation = New-Object System.Windows.Forms.ToolStripStatusLabel
$stOperation.Name = "StatusOperation"
$stOperation.Text = "Waiting for action..."
$stOperation.ForeColor = $uiTextSecondary
$stOperation.Padding = New-Object System.Windows.Forms.Padding(0, 0, 20, 0)
$stOperation.AutoSize = $true

# Right: Busy indicator
$stBusy = New-Object System.Windows.Forms.ToolStripStatusLabel
$stBusy.Name = "StatusBusy"
$stBusy.Text = ""
$stBusy.ForeColor = $uiStatusWarning
$stBusy.Padding = New-Object System.Windows.Forms.Padding(0, 0, 20, 0)
$stBusy.Visible = $false

# Far Right: Idle timer
$stIdle = New-Object System.Windows.Forms.ToolStripStatusLabel
$stIdle.Name = "StatusIdle"
$stIdle.Text = ""
$stIdle.ForeColor = $uiTextMuted
$stIdle.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$stIdle.Visible = $false

# Spacer (spring) keeps the status on the left and the info items on the right
$stSpacer = New-Object System.Windows.Forms.ToolStripStatusLabel
$stSpacer.Spring = $true
$stSpacer.Text = ""

# Version info gets its own slot so the busy scanner's status updates never wipe it out
$stVersion = New-Object System.Windows.Forms.ToolStripStatusLabel
$stVersion.Name = "StatusVersion"
# NOTE: separator is ASCII '|' on purpose - this file is UTF-8 without BOM and
# Windows PowerShell 5.1 reads BOM-less .ps1 files as ANSI (Windows-1252), which
# rendered the old UTF-8 middle-dot separator as a garbled two-char sequence.
# "Created by" credit restored (exact original wording) - this item is protected
# from the busy scanner, so it stays on screen permanently.
$stVersion.Text = "Created by: Alex_C.T   |   Version: $VersionNumber   |   Revised: $VersionDate   |   PowerShell: $PowerShellVersion"
$stVersion.ForeColor = $uiTextMuted
$stVersion.AutoSize = $true
$stVersion.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)

# Environment indicator (left, next to the status) - only visible when the
# production storeman share was not used with a fresh key.
# KEY EXPIRED (red) when the key is stale; TEST MODE (orange) otherwise.
$stMode = New-Object System.Windows.Forms.ToolStripStatusLabel
$stMode.Name = "StatusMode"
if ($script:StoremanKeyStale)
{
	$stMode.Text = "KEY EXPIRED"
	$stMode.ForeColor = [System.Drawing.Color]::FromArgb(220, 60, 45)
}
else
{
	$stMode.Text = "TEST MODE"
	$stMode.ForeColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
}
$stMode.Font = New-Object System.Drawing.Font("Segoe UI", [float]8.5, [System.Drawing.FontStyle]::Bold)
$stMode.Padding = New-Object System.Windows.Forms.Padding(8, 0, 20, 0)
$stMode.AutoSize = $true
$stMode.Visible = [bool]$script:IsTestEnvironment
# NOTE: ToolStripStatusLabel is a ToolStripItem (not a Control), so it uses
# its own ToolTipText property instead of the shared ToolTip object
$stModePrefix = if ($script:StoremanKeyStale) { "KEY EXPIRED" } else { "TEST ENVIRONMENT" }
$stMode.ToolTipText = "$stModePrefix - " + $script:StoremanSourceReason + ". Using storeman root: $BasePath"

[void]$statusStrip.Items.Add($stLeft)
[void]$statusStrip.Items.Add($stMode)
[void]$statusStrip.Items.Add($stSpacer)
[void]$statusStrip.Items.Add($stOperation)
[void]$statusStrip.Items.Add($stBusy)
[void]$statusStrip.Items.Add($stIdle)
[void]$statusStrip.Items.Add($stVersion)
$form.Controls.Add($statusStrip)

# Helper function to update status bar
$updateStatusBar = {
	param(
		[string]$Status = "Ready",
		[System.Drawing.Color]$StatusColor = $uiStatusSuccess,
		[string]$Operation = "",
		[string]$BusyInfo = ""
	)
	try
	{
		if ($stLeft -and -not $stLeft.IsDisposed) { $stLeft.Text = $Status; $stLeft.ForeColor = $StatusColor }
		if ($stOperation -and -not $stOperation.IsDisposed) { $stOperation.Text = $Operation }
		if ($stBusy -and -not $stBusy.IsDisposed) {
			if ($BusyInfo) { $stBusy.Text = $BusyInfo; $stBusy.Visible = $true; $stBusy.ForeColor = $uiStatusWarning }
			else { $stBusy.Visible = $false }
		}
	}
	catch { }
}
