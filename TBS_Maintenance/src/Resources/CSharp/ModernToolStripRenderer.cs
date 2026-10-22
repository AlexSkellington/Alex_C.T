// Menu renderer for the four tool menus and their submenus (GUI\09_Styling_Helpers).
// Embedded at build time and compiled once with Add-Type (references System.Windows.Forms and
// System.Drawing). Light-blue hover with a 1 px accent bar on the left, thin light separators,
// light borders, no image margin.
using System;
using System.Drawing;
using System.Windows.Forms;

public class ModernColorTable : ProfessionalColorTable
{
    public override Color MenuBorder { get { return Color.FromArgb(229, 229, 229); } }
    public override Color MenuItemBorder { get { return Color.Transparent; } }
    public override Color ToolStripBorder { get { return Color.FromArgb(229, 229, 229); } }
}

public class ModernToolStripRenderer : ToolStripProfessionalRenderer
{
    public ModernToolStripRenderer() : base(new ModernColorTable()) { }

    protected override void OnRenderItemBackground(ToolStripItemRenderEventArgs e)
    {
        if (e.Item.Selected)
        {
            using (var brush = new SolidBrush(Color.FromArgb(230, 243, 255)))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(0, 0, e.Item.Width - 1, e.Item.Height - 1));
            }
        }
        else
        {
            base.OnRenderItemBackground(e);
        }
    }

    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
    {
        if (e.Item.Selected && e.Item.Enabled)
        {
            using (var brush = new SolidBrush(Color.FromArgb(230, 243, 255)))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(0, 0, e.Item.Width - 1, e.Item.Height - 1));
            }
            using (var pen = new Pen(Color.FromArgb(0, 122, 204), 1))
            {
                e.Graphics.DrawLine(pen, 0, 0, 0, e.Item.Height - 1);
            }
        }
        else
        {
            base.OnRenderMenuItemBackground(e);
        }
    }

    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { }

    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
    {
        using (var pen = new Pen(Color.FromArgb(229, 229, 229)))
        {
            e.Graphics.DrawLine(pen, 0, e.Item.Height / 2, e.Item.Width, e.Item.Height / 2);
        }
    }
}
