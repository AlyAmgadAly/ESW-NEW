using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Custom_Controls
{
    public class RJDropdownMenu : ContextMenuStrip
    {
        // Fields
        private bool isMainMenu;
        private int menuItemHeight = 25;
        private Color menuItemTextColor = Color.Empty;
        private Color primaryColor = Color.Empty;
        private Bitmap menuItemHeaderSize;

        // Event for when an item is selected
        public event EventHandler<string> ItemSelected;

        // Constructor
        public RJDropdownMenu(IContainer container) : base(container) { }

        // Properties
        [Browsable(false)]
        public bool IsMainMenu
        {
            get { return isMainMenu; }
            set { isMainMenu = value; }
        }

        [Browsable(false)]
        public int MenuItemHeight
        {
            get { return menuItemHeight; }
            set { menuItemHeight = value; }
        }

        [Browsable(false)]
        public Color MenuItemTextColor
        {
            get { return menuItemTextColor; }
            set { menuItemTextColor = value; }
        }

        [Browsable(false)]
        public Color PrimaryColor
        {
            get { return primaryColor; }
            set { primaryColor = value; }
        }

        // Private methods
        private void LoadMenuItemHeight()
        {
            if (isMainMenu)
                menuItemHeaderSize = new Bitmap(25, 45);
            else
                menuItemHeaderSize = new Bitmap(20, menuItemHeight);

            foreach (ToolStripMenuItem menuItemL1 in this.Items)
            {
                menuItemL1.ImageScaling = ToolStripItemImageScaling.None;
                if (menuItemL1.Image == null) menuItemL1.Image = menuItemHeaderSize;

                foreach (ToolStripMenuItem menuItemL2 in menuItemL1.DropDownItems)
                {
                    menuItemL2.ImageScaling = ToolStripItemImageScaling.None;
                    if (menuItemL2.Image == null) menuItemL2.Image = menuItemHeaderSize;

                    foreach (ToolStripMenuItem menuItemL3 in menuItemL2.DropDownItems)
                    {
                        menuItemL3.ImageScaling = ToolStripItemImageScaling.None;
                        if (menuItemL3.Image == null) menuItemL3.Image = menuItemHeaderSize;

                        foreach (ToolStripMenuItem menuItemL4 in menuItemL3.DropDownItems)
                        {
                            menuItemL4.ImageScaling = ToolStripItemImageScaling.None;
                            if (menuItemL4.Image == null) menuItemL4.Image = menuItemHeaderSize;
                        }
                    }
                }
            }
        }

        // Attach click events to final menu items
        public void AttachClickEvents()
        {
            AssignClickEvents(this.Items);
        }

        private void AssignClickEvents(ToolStripItemCollection menuItems)
        {
            foreach (ToolStripItem item in menuItems)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    if (menuItem.DropDownItems.Count == 0) // Final-level item
                    {
                        menuItem.Click += MenuItem_Click;
                    }
                    else
                    {
                        AssignClickEvents(menuItem.DropDownItems);
                    }
                }
            }
        }

        private void MenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem clickedItem)
            {
                // Trigger event with selected item text
                ItemSelected?.Invoke(this, clickedItem.Text);
            }
        }

        // Overrides
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            if (!this.DesignMode)
            {
                this.Renderer = new MenuRenderer(isMainMenu, primaryColor, menuItemTextColor);
                LoadMenuItemHeight();
                AttachClickEvents();  // Attach event handlers after menu is created
            }
        }
    }
}
