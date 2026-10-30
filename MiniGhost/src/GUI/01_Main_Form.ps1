# ===================================================================================================
#                                       SECTION: Initialize GUI Components
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and initializes the main graphical user interface (GUI) form and its components.
# ===================================================================================================

# Create the main form (resizable, DPI-aware)
$form = New-Object System.Windows.Forms.Form
$form.Text = "Created by Alex_C.T | Version: $VersionNumber | Revised: $VersionDate"
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.MinimumSize = New-Object System.Drawing.Size(760, 520)
$form.Size = New-Object System.Drawing.Size(880, 580)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.BackColor = [System.Drawing.Color]::White

# Global tooltip provider (per-button tooltips)
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 20000
$toolTip.InitialDelay = 350
$toolTip.ReshowDelay = 150
$toolTip.ShowAlways = $true

# ---------------------------------------------------------------------------------------------------
# Main layout (Banner + Info + Actions)
# ---------------------------------------------------------------------------------------------------
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 3
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(12)
$mainLayout.Margin  = New-Object System.Windows.Forms.Padding(0)
$mainLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize

[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$form.Controls.Add($mainLayout)

# Banner Label
$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Text = "PowerShell Script - Mini Ghost"
$bannerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$bannerLabel.AutoSize = $true
$bannerLabel.TextAlign = 'MiddleCenter'
$bannerLabel.Dock = 'Fill'
$bannerLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)

$mainLayout.Controls.Add($bannerLabel, 0, 0)

# Display current machine name and store number in labels
$script:machineNameLabel = New-Object System.Windows.Forms.Label
$script:machineNameLabel.Text = "Current Machine Name: $currentMachineName"
$script:machineNameLabel.AutoSize = $true
$script:machineNameLabel.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 2)

$script:storeNameLabel = New-Object System.Windows.Forms.Label
$script:storeNameLabel.Text = "Store Name: $storeName"
$script:storeNameLabel.AutoSize = $true
$script:storeNameLabel.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 2)

$script:storeNumberLabel = New-Object System.Windows.Forms.Label
$script:storeNumberLabel.Text = "Store Number: $currentStoreNumber"
$script:storeNumberLabel.AutoSize = $true
$script:storeNumberLabel.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 2)

# Display current IP address in a label
$currentIP = if ($currentConfigs -and $currentConfigs.Count -gt 0) { $currentConfigs[0].IPv4Address.IPAddress }
else { "IP Not Found" }

$script:ipAddressLabel = New-Object System.Windows.Forms.Label
$script:ipAddressLabel.Text = "Current IP Address: $currentIP"
$script:ipAddressLabel.AutoSize = $true
$script:ipAddressLabel.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 6)

# Info group
$infoGroup = New-Object System.Windows.Forms.GroupBox
$infoGroup.Text = "Current Info"
$infoGroup.Dock = 'Top'
$infoGroup.AutoSize = $true
$infoGroup.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$infoGroup.Padding = New-Object System.Windows.Forms.Padding(10)
$infoGroup.Margin  = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)

$infoLayout = New-Object System.Windows.Forms.TableLayoutPanel
$infoLayout.Dock = 'Fill'
$infoLayout.ColumnCount = 1
$infoLayout.RowCount = 4
$infoLayout.AutoSize = $true
$infoLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$infoLayout.Margin = New-Object System.Windows.Forms.Padding(0)

$infoLayout.Controls.Add($machineNameLabel, 0, 0)
$infoLayout.Controls.Add($storeNameLabel, 0, 1)
$infoLayout.Controls.Add($storeNumberLabel, 0, 2)
$infoLayout.Controls.Add($ipAddressLabel, 0, 3)

$infoGroup.Controls.Add($infoLayout)

$mainLayout.Controls.Add($infoGroup, 0, 1)

# ---------------------------------------------------------------------------------------------------
# Actions group (buttons live here; populated after button creation)
# ---------------------------------------------------------------------------------------------------
$actionsGroup = New-Object System.Windows.Forms.GroupBox
$actionsGroup.Text = "Actions"
$actionsGroup.Dock = 'Fill'
$actionsGroup.Padding = New-Object System.Windows.Forms.Padding(10)
$actionsGroup.Margin  = New-Object System.Windows.Forms.Padding(0)

$actionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionsLayout.Dock = 'Fill'
$actionsLayout.ColumnCount = 3
$actionsLayout.RowCount = 3
$actionsLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize

[void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.34)))

[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.34)))

$actionsGroup.Controls.Add($actionsLayout)

$mainLayout.Controls.Add($actionsGroup, 0, 2)
