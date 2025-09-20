using ExcelScore.App_Forms.Main_Menu_Forms;
using ExcelScore.App_Forms.Primary_Forms.Child_Forms;
using ExcelScore.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.App_Forms.Primary_Forms
{
    public partial class SelectionMainForm : Form
    {
        public SelectionMainForm()
        {
            InitializeComponent();
        }

        private void btn_Comparative_Click(object sender, EventArgs e)
        {
            Child_Forms.Comparative_Child childForm = new Comparative_Child();
            pnl_test.Controls.Clear();
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;   // ensures it fills panel
            pnl_test.Controls.Add(childForm);

            // Force resize right away
            childForm.WindowState = FormWindowState.Normal;
            childForm.AutoScaleMode = AutoScaleMode.None;
            childForm.Show();
        }
    }
}
