using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelScore.Custom_Controls
{
    public class MenuColorTable : ProfessionalColorTable
    {
        // Fields
        private Color backColor;
        private Color leftColumnColor;
        private Color borderColor;
        private Color menuItemBorderColor;
        private Color menuItemSelectedColor;

        // Constructor
        public MenuColorTable(bool isMainMenu, Color primaryColor)
        {
            if (isMainMenu)
            {
                backColor = Color.FromArgb(37, 39, 60);
                leftColumnColor = Color.FromArgb(32, 33, 51);
                borderColor = Color.FromArgb(32, 33, 51);
                menuItemBorderColor = primaryColor;
                menuItemSelectedColor = primaryColor;
            }
            else
            {
                backColor = Color.White;
                leftColumnColor = Color.White;
                borderColor = Color.LightGray;
                menuItemBorderColor = primaryColor;
                menuItemSelectedColor = primaryColor;
            }
        }

        // Overrides
        public override Color ToolStripDropDownBackground => backColor;
        public override Color MenuBorder => borderColor;
        public override Color MenuItemBorder => menuItemSelectedColor;
        public override Color MenuItemSelectedGradientBegin => menuItemSelectedColor;
        public override Color MenuItemSelectedGradientEnd => menuItemSelectedColor;

        // Remove the gray strip
        public override Color ImageMarginGradientBegin => backColor;
        public override Color ImageMarginGradientMiddle => backColor;
        public override Color ImageMarginGradientEnd => backColor;
    }
}
