# ===================================================================================================
#                                FUNCTION: Set_Modern_Button_Style
# ---------------------------------------------------------------------------------------------------
# Description:
#   Gives a button the flat MiniGhost look: Primary (filled accent), Secondary (white with a border),
#   Action (main-window card button: left-aligned, coloured stripe) or Danger (Action in red text).
#
# Parameters:
#   - Button      [Button]: the button to style
#   - Kind        [string]: Primary | Secondary | Action | Danger (default Secondary)
#   - AccentColor [Color]: stripe colour for Action / Danger (default accent blue)
#
# Notes:
#   - Hover / pressed colours come from FlatAppearance, so no mouse handlers are needed.
#   - The stripe is drawn by a Paint handler that reads the colour from the button's .Tag (handlers
#     run in their own scope - nothing from this function is visible there).
# ===================================================================================================

function Set_Modern_Button_Style
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.Button]$Button,
		[Parameter(Mandatory = $false)]
		[ValidateSet('Primary', 'Secondary', 'Action', 'Danger')]
		[string]$Kind = 'Secondary',
		[Parameter(Mandatory = $false)]
		[System.Drawing.Color]$AccentColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
	)
	
	$Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$Button.UseVisualStyleBackColor = $false
	$Button.Cursor = [System.Windows.Forms.Cursors]::Hand
	$Button.FlatAppearance.BorderSize = 1
	
	switch ($Kind)
	{
		'Primary'
		{
			$Button.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
			$Button.ForeColor = [System.Drawing.Color]::White
			$Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
			$Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
			$Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(25, 117, 197)
			$Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 84, 158)
		}
		'Secondary'
		{
			$Button.BackColor = [System.Drawing.Color]::White
			$Button.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
			$Button.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
			$Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(209, 209, 209)
			$Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
			$Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
		}
		default
		{
			# Action / Danger: card buttons in the main window
			$Button.BackColor = [System.Drawing.Color]::White
			$Button.ForeColor = if ($Kind -eq 'Danger') { [System.Drawing.Color]::FromArgb(196, 43, 28) } else { [System.Drawing.Color]::FromArgb(32, 32, 32) }
			$Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
			$Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
			$Button.Padding = New-Object System.Windows.Forms.Padding(14, 0, 8, 0)
			$Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
			$Button.FlatAppearance.MouseOverBackColor = if ($Kind -eq 'Danger') { [System.Drawing.Color]::FromArgb(253, 240, 238) } else { [System.Drawing.Color]::FromArgb(240, 246, 253) }
			$Button.FlatAppearance.MouseDownBackColor = if ($Kind -eq 'Danger') { [System.Drawing.Color]::FromArgb(249, 224, 220) } else { [System.Drawing.Color]::FromArgb(224, 236, 250) }
			$Button.Tag = $AccentColor
			$Button.Add_Paint({
					param ($s, $e)
					try
					{
						if ($s.Tag -is [System.Drawing.Color])
						{
							$stripe = New-Object System.Drawing.SolidBrush($s.Tag)
							$e.Graphics.FillRectangle($stripe, 0, 0, 4, $s.Height)
							$stripe.Dispose()
						}
					}
					catch { }
				})
		}
	}
}
