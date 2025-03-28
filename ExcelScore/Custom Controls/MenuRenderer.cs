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
            Graphics graph = e.Graphics;
            graph.SmoothingMode = SmoothingMode.AntiAlias;

            // Define arrow size
            int arrowWidth = 8;
            int arrowHeight = 12;
            Point arrowPoint = e.ArrowRectangle.Location;
            arrowPoint.Offset(0, (e.ArrowRectangle.Height - arrowHeight) / 2);

            // Define points for a modern chevron-style arrow
            Point[] chevronArrow = new Point[]
            {
        new Point(arrowPoint.X, arrowPoint.Y),                       // Top point
        new Point(arrowPoint.X + arrowWidth, arrowPoint.Y + (arrowHeight / 2)),  // Middle
        new Point(arrowPoint.X, arrowPoint.Y + arrowHeight)          // Bottom point
            };

            // Smooth gradient color effect (optional)
            using (LinearGradientBrush brush = new LinearGradientBrush(e.ArrowRectangle,
                                                                       e.Item.Selected ? Color.White : primaryColor,
                                                                       Color.Transparent,
                                                                       45f))
            {
                graph.FillPolygon(brush, chevronArrow);
            }
        }

    }
}
