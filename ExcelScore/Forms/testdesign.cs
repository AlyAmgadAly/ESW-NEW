using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class testdesign : Form
    {
        private CircularToggle toggleSwitch;
        public testdesign()
        {
            InitializeComponent();
            InitializeCustomToggle();

        }
        private void InitializeCustomToggle()
        {
            toggleSwitch = new CircularToggle
            {
                Location = new Point(50, 50),
                Size = new Size(50, 50) // Keep it circular
            };

            this.Controls.Add(toggleSwitch);
        }

        public class CircularToggle : CheckBox
        {
            public CircularToggle()
            {
                this.Appearance = Appearance.Button;
                this.FlatStyle = FlatStyle.Flat;
                this.FlatAppearance.BorderSize = 0;
                this.BackColor = Color.Transparent;
                this.Size = new Size(50, 50);
            }

            protected override void OnPaint(PaintEventArgs pevent)
            {
                base.OnPaint(pevent);
                Graphics g = pevent.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                int toggleSize = 22;  // Small circle size
                int padding = 4;      // Space around inner toggle

                // Define Colors
                Color backgroundColor = this.Checked ? Color.Green : Color.Gray;
                Color toggleColor = Color.White;

                // Draw Outer Circle (Toggle Background)
                using (SolidBrush brush = new SolidBrush(backgroundColor))
                {
                    g.FillEllipse(brush, 0, 0, this.Width, this.Height);
                }

                // Calculate Inner Circle Position
                int xPos = this.Checked ? this.Width - toggleSize - padding : padding;
                int yPos = (this.Height - toggleSize) / 2; // Center vertically

                // Draw Inner Circle (Movable Knob)
                using (SolidBrush brush = new SolidBrush(toggleColor))
                {
                    g.FillEllipse(brush, xPos, yPos, toggleSize, toggleSize);
                }
            }

            protected override void OnClick(EventArgs e)
            {
                this.Checked = !this.Checked; // Toggle the state
                this.Invalidate(); // Redraw control
                base.OnClick(e);
            }
        }
        private void testdesign_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            rjDropdownMenu1.Show(button1, button1.Width, 0);
        }

        private void rjDropdownMenu1_VisibleChanged(object sender, EventArgs e)
        {

        }
    }
}
