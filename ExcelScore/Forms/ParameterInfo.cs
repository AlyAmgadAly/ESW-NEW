using Accord.Statistics.Distributions.Univariate;
using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class ParameterInfo : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        public ParameterInfo()
        {
            InitializeComponent();
        }

        public void GetParameterData()
        {

        }
        public void UpdateParameter(string paraName)
        {
            lbl_ParameterName.Text = $"Parameter: {paraName}"; // Update label

            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Normal;
            }

            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(
                Screen.PrimaryScreen.WorkingArea.Width - this.Width - 200, // 10px padding from the right edge
                (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2 // Center vertically
            );

            this.BringToFront(); // Bring the form to front if minimized
            this.Activate();
        }

        public Worksheet Sheet1 = ExcelFunctions.worksheet;

        public Worksheet Sheet2 = ExcelFunctions.Sheet2;

        public string ParameterName_ParaInfoFrm { get; set; }


        private void ParameterInfo_Load(object sender, EventArgs e)
        {
            
        }

        private void pic_Minimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void pic_CloseChooseTable_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
    }
}
