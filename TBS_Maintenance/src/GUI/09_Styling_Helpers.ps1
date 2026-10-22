######################################################################################################################
# 
# Modern UI Styling Helpers (NO functions - scriptblocks only)
#
# - Keeps code PS 5.1 friendly
# - Applies consistent modern styling to Buttons + ContextMenus
######################################################################################################################

# Safe fallback colors (in case these aren't defined elsewhere yet)
if (-not $script:UiAccent) { $script:UiAccent = [System.Drawing.Color]::DodgerBlue }
if (-not $script:UiCardBorder) { $script:UiCardBorder = [System.Drawing.Color]::FromArgb(220, 225, 232) }
if (-not $script:UiCardBg) { $script:UiCardBg = [System.Drawing.Color]::White }
if (-not $script:UiTextDark) { $script:UiTextDark = [System.Drawing.Color]::FromArgb(28, 28, 30) }

# ----------------------------------------------------------------------------
# Apply modern button styling (flat, clean border, hover accent)
# ----------------------------------------------------------------------------
$Apply_Modern_Button_Style = {
	param ([System.Windows.Forms.Button]$Btn)
	
	if (-not $Btn) { return }
	
	try
	{
		# Base look - modern flat design with subtle elevation
		$Btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
		$Btn.UseVisualStyleBackColor = $false
		$Btn.BackColor = $script:UiCardBg
		$Btn.ForeColor = $script:UiTextDark
		$Btn.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Regular)
		$Btn.Cursor = [System.Windows.Forms.Cursors]::Hand
		$Btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		$Btn.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)
		
		# Border + hover states (elevated card feel)
		$Btn.FlatAppearance.BorderSize = 0
		$Btn.FlatAppearance.MouseOverBackColor = $script:UiAccentLight
		$Btn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(220, 235, 250)
		
		# Subtle shadow effect (bottom border)
		$Btn.add_Paint({
				param ($s, $e)
				try
				{
					# Bottom shadow line
					$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 200, 200))
					$e.Graphics.DrawLine($pen, 2, $s.Height - 1, $s.Width - 2, $s.Height - 1)
					$pen.Dispose()
				}
				catch { }
			})
		
		# Store original colors for restore
		if (-not $Btn.Tag -or ($Btn.Tag -isnot [pscustomobject]))
		{
			$Btn.Tag = [pscustomobject]@{
				OrigBack = $Btn.BackColor
				OrigFore = $Btn.ForeColor
			}
		}
		
		# Hover effects (lift + accent)
		$Btn.add_MouseEnter({
				param ($s, $e)
				try
				{
					$s.ForeColor = $script:UiAccentDark
					$s.BackColor = $script:UiAccentLight
					# Subtle elevation effect
					$s.Padding = New-Object System.Windows.Forms.Padding(10, 2, 10, 6)
				}
				catch { }
			})
		$Btn.add_MouseLeave({
				param ($s, $e)
				try
				{
					if ($s.Tag -and $s.Tag.OrigFore) { $s.ForeColor = $s.Tag.OrigFore }
					if ($s.Tag -and $s.Tag.OrigBack) { $s.BackColor = $s.Tag.OrigBack }
					# Restore padding
					$s.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)
				}
				catch { }
			})
		$Btn.add_MouseDown({
				param ($s, $e)
				try
				{
					$s.BackColor = [System.Drawing.Color]::FromArgb(230, 242, 255)
				}
				catch { }
			})
		$Btn.add_MouseUp({
				param ($s, $e)
				try
				{
					if ($s.Bounds.Contains($s.PointToClient([System.Windows.Forms.Cursor]::Position)))
					{
						$s.BackColor = $script:UiAccentLight
					}
				}
				catch { }
			})
	}
	catch { }
}

# ===================================================================================================
# ENHANCED MODERN CONTEXT MENU STYLING - Clean, modern look with icons
# ===================================================================================================
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
		$Menu.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::Professional
		$Menu.Padding = New-Object System.Windows.Forms.Padding(2)
		$Menu.GripMargin = New-Object System.Windows.Forms.Padding(0)
		$Menu.ItemMargin = New-Object System.Windows.Forms.Padding(2)
		
		# Custom renderer for modern menu
		# (NOTE: do NOT assign a ProfessionalColorTable to $Menu.Renderer here -
		#  it throws and the custom renderer below would silently never be applied)
		if (-not ('ModernToolStripRenderer' -as [type])) {
		function global:CreateModernMenuRenderer {
			Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
			$rendererCode = @'
using System;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Drawing;
using System.Drawing.Drawing2D;

public class ModernMenuRenderer : ProfessionalColorTable, IRenderer {
    public Color MenuBorder { get { return Color.FromArgb(229, 229, 229); } }
    public Color MenuItemBorder { get { return Color.Transparent; } }
}

public interface IRenderer { }

public class ModernToolStripRenderer : ToolStripProfessionalRenderer {
    public ModernToolStripRenderer() : base(new ModernColorTable()) { }
    
    protected override void OnRenderItemBackground(ToolStripItemRenderEventArgs e) {
        if (e.Item.Selected) {
            using (var brush = new SolidBrush(Color.FromArgb(230, 243, 255))) {
                e.Graphics.FillRectangle(brush, new Rectangle(0, 0, e.Item.Width - 1, e.Item.Height - 1));
            }
        }
        else {
            base.OnRenderItemBackground(e);
        }
    }
    
    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {
        if (e.Item.Selected) {
            using (var brush = new SolidBrush(Color.FromArgb(230, 243, 255))) {
                e.Graphics.FillRectangle(brush, new Rectangle(0, 0, e.Item.Width - 1, e.Item.Height - 1));
            }
            using (var pen = new Pen(Color.FromArgb(0, 122, 204), 1)) {
                e.Graphics.DrawLine(pen, 0, 0, 0, e.Item.Height - 1);
            }
        }
        else {
            base.OnRenderMenuItemBackground(e);
        }
    }
    
    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { }
    
    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) {
        using (var pen = new Pen(Color.FromArgb(229, 229, 229))) {
            e.Graphics.DrawLine(pen, 0, e.Item.Height / 2, e.Item.Width, e.Item.Height / 2);
        }
    }
}

public class ModernColorTable : ProfessionalColorTable {
    public override Color MenuBorder { get { return Color.FromArgb(229, 229, 229); } }
    public override Color MenuItemBorder { get { return Color.Transparent; } }
    public override Color ToolStripBorder { get { return Color.FromArgb(229, 229, 229); } }
}
'@
			Add-Type -TypeDefinition $rendererCode -ErrorAction SilentlyContinue
		}
		CreateModernMenuRenderer
		}
		if ('ModernToolStripRenderer' -as [type]) {
			$Menu.Renderer = [ModernToolStripRenderer]::new()
		}
	}
	catch { }
}

# ----------------------------------------------------------------------------
# Style a submenu dropdown to match the modern menu theme (the function
# above is typed for ContextMenuStrip, so submenus are styled separately)
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
		if ('ModernToolStripRenderer' -as [type]) { $dd.Renderer = [ModernToolStripRenderer]::new() }
	}
	catch { }
}

# Apply modern styling to popup forms so older dialogs match the main app.
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
					if ($form -and [object]::ReferenceEquals($openForm, $form)) { continue }
					
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
