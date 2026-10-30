# ===================================================================================================
#                                FUNCTION: Set_Modern_Dialog_Style
# ---------------------------------------------------------------------------------------------------
# Description:
#   Restyles a dialog and every control in it to match the main window: white background, Segoe UI,
#   flat buttons (the given primary button filled with the accent colour), bordered text boxes.
#
# Parameters:
#   - Form          [Form]: the dialog
#   - PrimaryButton [Button]: the default action (OK / Reboot Now); other buttons get Secondary
#
# Notes:
#   - Layout is not changed: positions and sizes stay as the dialog defined them.
#   - Monospaced fonts (report text) are kept.
# ===================================================================================================

function Set_Modern_Dialog_Style
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.Form]$Form,
		[Parameter(Mandatory = $false)]
		[System.Windows.Forms.Button]$PrimaryButton
	)
	
	$Form.BackColor = [System.Drawing.Color]::White
	$Form.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
	$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
	if ($script:MiniGhostIcon) { try { $Form.Icon = $script:MiniGhostIcon; $Form.ShowIcon = $true } catch { } }
	
	$pending = New-Object System.Collections.Generic.Queue[System.Windows.Forms.Control]
	foreach ($c in @($Form.Controls)) { $pending.Enqueue($c) }
	while ($pending.Count -gt 0)
	{
		$c = $pending.Dequeue()
		foreach ($child in @($c.Controls)) { $pending.Enqueue($child) }
		try
		{
			if ($c -is [System.Windows.Forms.Button])
			{
				if ($PrimaryButton -and [object]::ReferenceEquals($c, $PrimaryButton)) { Set_Modern_Button_Style -Button $c -Kind Primary }
				else { Set_Modern_Button_Style -Button $c -Kind Secondary }
			}
			elseif ($c -is [System.Windows.Forms.TextBox])
			{
				if ($c.Font.Name -notmatch '(?i)consolas|courier|lucida console') { $c.Font = New-Object System.Drawing.Font("Segoe UI", 10) }
				if ($c.ReadOnly -and $c.BorderStyle -eq [System.Windows.Forms.BorderStyle]::None)
				{
					$c.BackColor = [System.Drawing.Color]::White
				}
				else
				{
					$c.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
				}
			}
			elseif ($c -is [System.Windows.Forms.Label])
			{
				$c.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
				$c.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
			}
			elseif ($c -is [System.Windows.Forms.FlowLayoutPanel] -and $c.Dock -eq [System.Windows.Forms.DockStyle]::Bottom)
			{
				# button bar: light footer like Windows 11 dialogs
				$c.BackColor = [System.Drawing.Color]::FromArgb(243, 243, 243)
			}
		}
		catch { }
	}
}
