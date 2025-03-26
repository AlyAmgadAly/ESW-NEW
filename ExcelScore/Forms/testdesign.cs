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
        private Panel togglePanel;
        private Button toggleButton;
        private bool isOn = false;
        public testdesign()
        {
            InitializeComponent();
            InitializeCircularToggleSwitch();
        }
        private void InitializeCircularToggleSwitch()
        {
            // Toggle Panel (Oval Background)
            togglePanel = new Panel
            {
                Size = new Size(60, 30),  // Oval shape
                Location = new Point(50, 50),
                BackColor = Color.Gray,
                BorderStyle = BorderStyle.FixedSingle
            };
            togglePanel.Region = new Region(new System.Drawing.Drawing2D.GraphicsPath());
            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, togglePanel.Width, togglePanel.Height);
            togglePanel.Region = new Region(path);

            // Toggle Button (Circular Knob)
            toggleButton = new Button
            {
                Size = new Size(28, 28), // Perfect circle
                Location = new Point(2, 1), // Starts on left
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            toggleButton.FlatAppearance.BorderSize = 0;
            toggleButton.Region = new Region(new System.Drawing.Drawing2D.GraphicsPath());
            System.Drawing.Drawing2D.GraphicsPath buttonPath = new System.Drawing.Drawing2D.GraphicsPath();
            buttonPath.AddEllipse(0, 0, toggleButton.Width, toggleButton.Height);
            toggleButton.Region = new Region(buttonPath);

            toggleButton.Click += ToggleButton_Click;

            // Add controls to form
            togglePanel.Controls.Add(toggleButton);
            this.Controls.Add(togglePanel);
        }
        private void ToggleButton_Click(object sender, EventArgs e)
        {
            isOn = !isOn;

            if (isOn)
            {
                toggleButton.Location = new Point(30, 1); // Move Right (On)
                togglePanel.BackColor = Color.Green;
            }
            else
            {
                toggleButton.Location = new Point(2, 1); // Move Left (Off)
                togglePanel.BackColor = Color.Gray;
            }
        }
        private void testdesign_Load(object sender, EventArgs e)
        {

        }
    }
}
