using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.ComponentModel;

namespace ExcelScore.Custom_Controls
{
    public class MenuRenderer : ToolStripProfessionalRenderer
    {
        // Fields
        private Color primaryColor;
        private Color textColor;
        private int arrowThickness;

        // Property to expose ForeColor in Properties Tab
        [Category("Appearance")]
        [Description("Sets the text color for menu items.")]
        public Color MenuTextColor
        {
            get => textColor;
            set
            {
                textColor = value;
                if (textColor == Color.Empty)
                    textColor = Color.Gainsboro;
            }
        }

        // Constructor
        public MenuRenderer(bool isMainMenu, Color primaryColor, Color textColor)
            : base(new MenuColorTable(isMainMenu, primaryColor))
        {
            this.primaryColor = primaryColor;
            this.textColor = textColor == Color.Empty ? (isMainMenu ? Color.Gainsboro : Color.DimGray) : textColor;
            arrowThickness = isMainMenu ? 3 : 2;
        }

        // Overrides
        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            e.TextColor = e.Item.Selected ? primaryColor : textColor;
            base.OnRenderItemText(e);
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            var graph = e.Graphics;
            var arrowSize = new Size(5, 12);
            var arrowColor = e.Item.Selected ? Color.White : primaryColor;
            var rect = new Rectangle(e.ArrowRectangle.Location.X,
                                     (e.ArrowRectangle.Height - arrowSize.Height) / 2,
                                     arrowSize.Width, arrowSize.Height);
            using (GraphicsPath path = new GraphicsPath())
            using (Pen pen = new Pen(arrowColor, arrowThickness))
            {
                graph.SmoothingMode = SmoothingMode.AntiAlias;
                path.AddLine(rect.Left, rect.Top, rect.Right, rect.Top + rect.Height / 2);
                path.AddLine(rect.Right, rect.Top + rect.Height / 2, rect.Left, rect.Top + rect.Height);
                graph.DrawPath(pen, path);
            }
        }
    }
}
