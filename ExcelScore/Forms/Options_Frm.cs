using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
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
    public partial class Options_Frm : Form
    {
        public Options_Frm()
        {
            InitializeComponent();
        }
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public bool Row_Percent_Options { get; set; }
        public bool Maha_Options { get; set; }
        public bool Total_Column_Options { get; set; }
        public void Options_Frm_Load(object sender, EventArgs e)
        {
            check_Perc_Row_new.Checked = ComparativeGroups.Row_Percent_Comparative;
            check_Maha_new.Checked = ComparativeGroups.Maha_Comparative;
            check_TotalColumn_new.Checked = ComparativeGroups.Total_Column_Comparative;

        }
        public void DrawFrmBorderLines(PaintEventArgs e)
        {



            Pen YLeftPen = new Pen(Color.Black, 2);
            Pen YRightPen = new Pen(Color.Black, 2);
            Pen XUpPen = new Pen(Color.Black, 2);
            Pen XDownPen = new Pen(Color.Black, 2);

            Pen XInLabelPen = new Pen(Color.Black, 2);

            e.Graphics.DrawLine(YLeftPen, 4, 31, 4, 400);
            e.Graphics.DrawLine(YRightPen, this.Width-4, 31, this.Width - 4, 400);

            e.Graphics.DrawLine(XUpPen, 4, 31, this.Width - 4, 31);
            e.Graphics.DrawLine(XDownPen, 4, 400, 172, 400);

            e.Graphics.DrawLine(XInLabelPen, this.Width - 4, 400, 180, 400);

            YLeftPen.Dispose();
            YRightPen.Dispose();
            XUpPen.Dispose();
            XDownPen.Dispose();
            XInLabelPen.Dispose();


        }
        private void pic_DoneOptions_Click(object sender, EventArgs e)
        {
            


            Row_Percent_Options = check_Perc_Row_new.Checked;
            Maha_Options = check_Maha_new.Checked;
            Total_Column_Options = check_TotalColumn_new.Checked;

            // Update the static ComparativeGroups class
            ComparativeGroups.Row_Percent_Comparative = Row_Percent_Options;
            ComparativeGroups.Maha_Comparative = Maha_Options;
            ComparativeGroups.Total_Column_Comparative = Total_Column_Options;

            this.Hide();
        }
        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void Options_Frm_Paint(object sender, PaintEventArgs e)
        {
            DrawFrmBorderLines(e);
        }
    }
}
