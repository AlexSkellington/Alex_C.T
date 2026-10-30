# ===================================================================================================
#                                    FUNCTION: New_MiniGhost_Card
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns a white card panel with a 1 px light border (main-window sections).
# ===================================================================================================

function New_MiniGhost_Card
{
	param ([System.Windows.Forms.Padding]$Margin)
	$card = New-Object System.Windows.Forms.Panel
	$card.BackColor = [System.Drawing.Color]::White
	$card.Margin = $Margin
	$card.Padding = New-Object System.Windows.Forms.Padding(1)
	$card.add_Paint({
			param ($s, $e)
			try
			{
				$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(229, 229, 229))
				$e.Graphics.DrawRectangle($pen, 0, 0, ($s.Width - 1), ($s.Height - 1))
				$pen.Dispose()
			}
			catch { }
		})
	return $card
}
