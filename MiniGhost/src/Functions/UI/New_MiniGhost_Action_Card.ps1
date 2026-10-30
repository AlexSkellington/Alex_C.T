# ===================================================================================================
#                                 FUNCTION: New_MiniGhost_Action_Card
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns an action card for the main window: coloured dot, title, subtitle, and a host panel
#   for its buttons (in the card's .Tag).
#
# Parameters:
#   - Title, Subtitle [string]: card header text
#   - Accent          [Color]: colour of the header dot
# ===================================================================================================

function New_MiniGhost_Action_Card
{
	param ([string]$Title, [string]$Subtitle, [System.Drawing.Color]$Accent)

	$card = New_MiniGhost_Card -Margin (New-Object System.Windows.Forms.Padding(6))
	$card.Dock = 'Fill'

	# Buttons go into the host panel (11_Layout_And_Tooltips); the header docks above it.
	# ($host is a read-only automatic variable - hence $buttonHost)
	$buttonHost = New-Object System.Windows.Forms.Panel
	$buttonHost.Dock = 'Fill'
	$buttonHost.BackColor = [System.Drawing.Color]::White
	$buttonHost.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 12)
	$card.Controls.Add($buttonHost)

	$head = New-Object System.Windows.Forms.Panel
	$head.Dock = 'Top'
	$head.Height = 58
	$head.BackColor = [System.Drawing.Color]::White
	$head.Tag = $Accent
	$head.add_Paint({
			param ($s, $e)
			try
			{
				$e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
				$dot = New-Object System.Drawing.SolidBrush($s.Tag)
				$e.Graphics.FillEllipse($dot, 14, 16, 10, 10)
				$dot.Dispose()
			}
			catch { }
		})
	$t = New-Object System.Windows.Forms.Label
	$t.Text = $Title
	$t.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
	$t.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
	$t.AutoSize = $true
	$t.Location = New-Object System.Drawing.Point(30, 9)
	$head.Controls.Add($t)
	$st = New-Object System.Windows.Forms.Label
	$st.Text = $Subtitle
	$st.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
	$st.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
	$st.AutoSize = $true
	$st.Location = New-Object System.Drawing.Point(31, 32)
	$head.Controls.Add($st)
	$card.Controls.Add($head)

	$card.Tag = $buttonHost
	return $card
}
