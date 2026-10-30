# ===================================================================================================
#                                 FUNCTION: Draw_MiniGhost_Logo
# ---------------------------------------------------------------------------------------------------
# Description:
#   Draws the MiniGhost logo (violet gradient disc with a white ghost) into the square (X, Y, Size)
#   of any Graphics - the window icon and the header badge use it.
#
# Notes:
#   - Coordinates are computed in expression mode and passed to .NET methods, never as
#     'New-Object T($a + $b, ...)' (command mode turns that into $a + @($b, ...)).
# ===================================================================================================

function Draw_MiniGhost_Logo
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Drawing.Graphics]$Graphics,
		[float]$X = 0,
		[float]$Y = 0,
		[float]$Size = 32
	)
	
	$Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$disc = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
		[System.Drawing.RectangleF]::new($X, $Y, $Size, $Size),
		[System.Drawing.Color]::FromArgb(126, 87, 214),
		[System.Drawing.Color]::FromArgb(79, 70, 189),
		[float]90)
	$Graphics.FillEllipse($disc, $X, $Y, $Size, $Size)
	$disc.Dispose()
	
	# Ghost: round head, straight sides, three scallops along the bottom
	$w = $Size * 0.48
	$h = $Size * 0.56
	$left = $X + ($Size - $w) / 2
	$top = $Y + $Size * 0.2
	$bump = $w / 3
	$ghost = New-Object System.Drawing.Drawing2D.GraphicsPath
	$ghost.AddArc($left, $top, $w, $w, [float]180, [float]180)
	$ghost.AddLine(($left + $w), ($top + $w / 2), ($left + $w), ($top + $h - $bump / 2))
	for ($i = 0; $i -lt 3; $i++)
	{
		$ghost.AddArc(($left + $w - ($i + 1) * $bump), ($top + $h - $bump), $bump, $bump, [float]0, [float]180)
	}
	$ghost.CloseFigure()
	$white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
	$Graphics.FillPath($white, $ghost)
	$white.Dispose()
	$ghost.Dispose()
	
	$eye = $Size * 0.085
	$eyeY = $top + $w * 0.38
	$dark = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(79, 70, 189))
	$Graphics.FillEllipse($dark, ($left + $w * 0.3 - $eye / 2), $eyeY, $eye, ($eye * 1.3))
	$Graphics.FillEllipse($dark, ($left + $w * 0.7 - $eye / 2), $eyeY, $eye, ($eye * 1.3))
	$dark.Dispose()
}
