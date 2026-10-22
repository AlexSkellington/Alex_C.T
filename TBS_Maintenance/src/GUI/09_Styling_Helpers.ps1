# ===================================================================================================
#                                 SECTION: GUI - Styling Helpers
# ---------------------------------------------------------------------------------------------------
# Description:
#   Shared look for menus and popup dialogs (scriptblocks, PS 5.1 friendly):
#     $Apply_Modern_Menu_Style   style a ContextMenuStrip (the four tool menus, grid menus)
#     $Style_SubMenu             style a submenu's dropdown the same way
#     $Apply_Modern_Popup_Style  restyle a popup Form and all of its controls
#   plus the popup theme timer, which applies $Apply_Modern_Popup_Style automatically to every
#   new window the script opens (dialogs inside the functions need no styling code of their own).
#
# Notes:
#   - The menu renderer (src\Resources\CSharp\ModernToolStripRenderer.cs) is compiled once here.
#     Before v2.7.0 it never compiled (an unused class raised a warning-as-error and the WinForms
#     references were missing) and an invalid ItemMargin assignment aborted the styling first,
#     so the menus always fell back to the default renderer.
# ===================================================================================================

# Fallback colours (the palette in 01 defines these; same variables, script scope)
if (-not $script:UiAccent) { $script:UiAccent = [System.Drawing.Color]::DodgerBlue }
if (-not $script:UiCardBorder) { $script:UiCardBorder = [System.Drawing.Color]::FromArgb(220, 225, 232) }
if (-not $script:UiCardBg) { $script:UiCardBg = [System.Drawing.Color]::White }
if (-not $script:UiTextDark) { $script:UiTextDark = [System.Drawing.Color]::FromArgb(28, 28, 30) }

# ----------------------------------------------------------------------------
# Menu renderer (compiled once per process)
# ----------------------------------------------------------------------------
if (-not ('ModernToolStripRenderer' -as [type]))
{
	try
	{
		Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -ErrorAction Stop -TypeDefinition @'
#TBS:EMBED Resources/CSharp/ModernToolStripRenderer.cs
'@
	}
	catch
	{
		try { Write_Log ("Menu renderer could not be compiled - using the default menu look: " + $_.Exception.Message) 'gray' -FileOnly }
		catch { }
	}
}

# ----------------------------------------------------------------------------
# Context menu styling
# ----------------------------------------------------------------------------
$Apply_Modern_Menu_Style = {
	param ([System.Windows.Forms.ContextMenuStrip]$Menu)

	if (-not $Menu) { return }
	try
	{
		$Menu.ShowImageMargin = $false
		$Menu.ShowCheckMargin = $false
		$Menu.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
		$Menu.BackColor = [System.Drawing.Color]::White
		$Menu.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
		$Menu.Padding = New-Object System.Windows.Forms.Padding(2)
		$Menu.GripMargin = New-Object System.Windows.Forms.Padding(0)
		if ('ModernToolStripRenderer' -as [type]) { $Menu.Renderer = New-Object ModernToolStripRenderer }
		else { $Menu.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::Professional }
	}
	catch { }
}

# Submenu dropdowns (a ToolStripMenuItem's DropDown, not a ContextMenuStrip)
$Style_SubMenu = {
	param ($Sub)
	try
	{
		$dd = $Sub.DropDown
		$dd.BackColor = [System.Drawing.Color]::White
		$dd.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
		$dd.Padding = New-Object System.Windows.Forms.Padding(2)
		$dd.ShowImageMargin = $false
		$dd.ShowCheckMargin = $false
		$dd.GripMargin = New-Object System.Windows.Forms.Padding(0)
		if ('ModernToolStripRenderer' -as [type]) { $dd.Renderer = New-Object ModernToolStripRenderer }
	}
	catch { }
}

# ----------------------------------------------------------------------------
# Popup forms: restyle the form and every control in it so older dialogs match the main app
# ----------------------------------------------------------------------------
$Apply_Modern_Popup_Style = {
	param ([System.Windows.Forms.Form]$PopupForm)

	if (-not $PopupForm -or $PopupForm.IsDisposed) { return }

	$stylePopupControls = $null
	$stylePopupControls = {
		param ($Control)

		if (-not $Control) { return }

		try
		{
			if ($Control -is [System.Windows.Forms.Button])
			{
				$measured = [System.Windows.Forms.TextRenderer]::MeasureText($Control.Text, (New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)))
				$minWidth = [Math]::Max([int]$Control.Width, [int]($measured.Width + 22))
				$minHeight = [Math]::Max([int]$Control.Height, [int]($measured.Height + 12), 32)
				$preserveAccentBack = ($Control.BackColor -eq [System.Drawing.Color]::Yellow -or $Control.BackColor -eq [System.Drawing.Color]::LightGreen)

				$Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
				$Control.UseVisualStyleBackColor = $false
				if (-not $preserveAccentBack)
				{
					$Control.BackColor = $script:UiCardBg
				}
				$Control.ForeColor = $script:UiTextDark
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
				$Control.Cursor = [System.Windows.Forms.Cursors]::Hand
				$Control.FlatAppearance.BorderSize = 1
				$Control.FlatAppearance.BorderColor = $script:UiCardBorder
				$Control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
				$Control.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(235, 242, 252)
				if ($Control.Width -lt $minWidth) { $Control.Width = $minWidth }
				if ($Control.Height -lt $minHeight) { $Control.Height = $minHeight }
			}
			elseif ($Control -is [System.Windows.Forms.Label] -or
				$Control -is [System.Windows.Forms.CheckBox] -or
				$Control -is [System.Windows.Forms.RadioButton])
			{
				$Control.ForeColor = $script:UiTextDark
				$Control.BackColor = [System.Drawing.Color]::Transparent
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
			}
			elseif ($Control -is [System.Windows.Forms.TextBox] -or
				$Control -is [System.Windows.Forms.RichTextBox] -or
				$Control -is [System.Windows.Forms.MaskedTextBox])
			{
				$Control.BackColor = [System.Drawing.Color]::White
				$Control.ForeColor = $script:UiTextDark
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
				try { $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle }
				catch { }
			}
			elseif ($Control -is [System.Windows.Forms.ComboBox])
			{
				$Control.BackColor = [System.Drawing.Color]::White
				$Control.ForeColor = $script:UiTextDark
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
				try { $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
				catch { }
			}
			elseif ($Control -is [System.Windows.Forms.NumericUpDown] -or
				$Control -is [System.Windows.Forms.DateTimePicker])
			{
				$Control.BackColor = [System.Drawing.Color]::White
				$Control.ForeColor = $script:UiTextDark
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
			}
			elseif ($Control -is [System.Windows.Forms.ListBox] -or
				$Control -is [System.Windows.Forms.ListView] -or
				$Control -is [System.Windows.Forms.CheckedListBox])
			{
				$Control.BackColor = [System.Drawing.Color]::White
				$Control.ForeColor = $script:UiTextDark
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
			}
			elseif ($Control -is [System.Windows.Forms.GroupBox])
			{
				$Control.ForeColor = $script:UiTextDark
				$Control.BackColor = $script:UiCardBg
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
			}
			elseif ($Control -is [System.Windows.Forms.Panel] -or
				$Control -is [System.Windows.Forms.TabPage] -or
				$Control -is [System.Windows.Forms.FlowLayoutPanel] -or
				$Control -is [System.Windows.Forms.TableLayoutPanel])
			{
				if ($Control.BackColor -eq [System.Drawing.Color]::Empty -or $Control.BackColor.A -eq 0)
				{
					$Control.BackColor = [System.Drawing.Color]::Transparent
				}
				elseif ($Control.BackColor -ne $script:UiCardBg)
				{
					$Control.BackColor = $script:UiCardBg
				}
			}
			elseif ($Control -is [System.Windows.Forms.TabControl])
			{
				$Control.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
			}
			elseif ($Control -is [System.Windows.Forms.DataGridView])
			{
				$Control.BackgroundColor = $script:UiCardBg
				$Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
				$Control.GridColor = $script:UiCardBorder
				$Control.EnableHeadersVisualStyles = $false
				$Control.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
				$Control.ColumnHeadersDefaultCellStyle.ForeColor = $script:UiTextDark
				$Control.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
				$Control.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
				$Control.DefaultCellStyle.ForeColor = $script:UiTextDark
				$Control.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(221, 235, 255)
				$Control.DefaultCellStyle.SelectionForeColor = $script:UiTextDark
				$Control.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 253)
			}

			if ($Control.ContextMenuStrip)
			{
				& $Apply_Modern_Menu_Style $Control.ContextMenuStrip
			}
		}
		catch { }

		try
		{
			foreach ($child in @($Control.Controls))
			{
				& $stylePopupControls $child
			}
		}
		catch { }
	}

	try
	{
		$PopupForm.BackColor = $uiBg
		$PopupForm.ForeColor = $script:UiTextDark
		$PopupForm.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
		$PopupForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
		try
		{
			$prop = $PopupForm.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
			if ($prop) { $prop.SetValue($PopupForm, $true, $null) }
		}
		catch { }

		if ($PopupForm.MainMenuStrip)
		{
			$PopupForm.MainMenuStrip.BackColor = $script:UiCardBg
			$PopupForm.MainMenuStrip.ForeColor = $script:UiTextDark
			$PopupForm.MainMenuStrip.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
		}

		& $stylePopupControls $PopupForm
	}
	catch { }
}

if (-not $script:PopupThemeTimer)
{
	$script:PopupThemeTimer = New-Object System.Windows.Forms.Timer
	$script:PopupThemeTimer.Interval = 300
	$script:PopupThemeTimer.Add_Tick({
			try
			{
				if ([System.Windows.Forms.Application]::OpenForms.Count -le 1 -and $script:StyledPopupForms.Count -eq 0) { return }

				$activeHashes = @{ }
				foreach ($openForm in [System.Windows.Forms.Application]::OpenForms)
				{
					if (-not $openForm -or $openForm.IsDisposed) { continue }
					# $script:MainForm, not $form: during a dialog $form is that dialog, and the main
					# window then got restyled (dark text on the dark banner - the title, store label,
					# clock and badge "disappeared" after any action with a dialog)
					if ([object]::ReferenceEquals($openForm, $script:MainForm)) { continue }

					$formHash = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($openForm)
					$activeHashes[$formHash] = $true

					if (-not $script:StyledPopupForms.ContainsKey($formHash))
					{
						& $Apply_Modern_Popup_Style $openForm
						$script:StyledPopupForms[$formHash] = $true
					}
				}

				foreach ($staleHash in @($script:StyledPopupForms.Keys))
				{
					if (-not $activeHashes.ContainsKey($staleHash))
					{
						[void]$script:StyledPopupForms.Remove($staleHash)
					}
				}
			}
			catch { }
		})
	$script:PopupThemeTimer.Start()
}

# ----------------------------------------------------------------------------
# Category button (Server / Lane / Scale / General): white flat button with a
# colored bottom bar. Click opens its menu; right-click opens the protocol
# window when ProtocolMode is set. Handlers read everything from .Tag (no
# closures - see CLAUDE.md), so one definition serves all four buttons.
# ----------------------------------------------------------------------------
$New_Category_Button = {
	param (
		[string]$Name,
		[string]$Text,
		[int]$X,
		[System.Drawing.Color]$Accent,
		[System.Drawing.Color]$HoverBack,
		[System.Drawing.Color]$DownBack,
		[System.Drawing.Color]$HoverFore,
		[System.Windows.Forms.ContextMenuStrip]$Menu,
		[string]$ProtocolMode,
		[string]$ToolTipText
	)

	$b = New-Object System.Windows.Forms.Button
	$b.Name = $Name
	$b.Text = $Text
	$b.Location = New-Object System.Drawing.Point($X, 475)
	$b.Size = New-Object System.Drawing.Size(200, 52)
	$b.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Regular)
	$b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$b.UseVisualStyleBackColor = $false
	$b.BackColor = [System.Drawing.Color]::White
	$b.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
	$b.Cursor = [System.Windows.Forms.Cursors]::Hand
	$b.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$b.FlatAppearance.BorderSize = 0
	$b.FlatAppearance.MouseOverBackColor = $HoverBack
	$b.FlatAppearance.MouseDownBackColor = $DownBack
	$b.Tag = [pscustomobject]@{
		CategoryColor = $Accent
		HoverBack	  = $HoverBack
		HoverFore	  = $HoverFore
		Menu		  = $Menu
		ProtocolMode  = $ProtocolMode
	}

	# Bottom accent bar
	$b.add_Paint({
			param ($s, $e)
			try
			{
				$pen = New-Object System.Drawing.Pen($s.Tag.CategoryColor, 3)
				$e.Graphics.DrawLine($pen, 4, $s.Height - 4, $s.Width - 4, $s.Height - 4)
				$pen.Dispose()
			}
			catch { }
		})
	$b.add_MouseEnter({
			param ($s, $e)
			try { $s.BackColor = $s.Tag.HoverBack; $s.ForeColor = $s.Tag.HoverFore }
			catch { }
		})
	$b.add_MouseLeave({
			param ($s, $e)
			try { $s.BackColor = [System.Drawing.Color]::White; $s.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32) }
			catch { }
		})
	$b.add_Click({
			param ($s, $e)
			$script:LastActivity = Get-Date
			$s.Tag.Menu.Show($s, 0, $s.Height)
		})
	if ($ProtocolMode)
	{
		$b.add_MouseDown({
				param ($s, $e)
				if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right)
				{
					$script:ProtocolPopupMode = $s.Tag.ProtocolMode
					$global:ProtocolForm.Show()
					$global:ProtocolForm.BringToFront()
				}
			})
	}
	$toolTip.SetToolTip($b, $ToolTipText)
	$form.Controls.Add($b)
	$b
}
