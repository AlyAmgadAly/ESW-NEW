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
    public partial class ChooseFrm : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public DataGridView Dgv { get; set; }
        public ChooseFrm()
        {
            InitializeComponent();
        }

        private void ChooseFrm_Load(object sender, EventArgs e)
        {
            //workbook = ExcelFunctionsobj.getWorkbook();
        }

        private void btn_Questionnare_Click(object sender, EventArgs e)
        {
            NewAddDomain addDomains = new NewAddDomain();
            addDomains.Dgv = Dgv;
            addDomains.Show();
            this.Hide();
        }

        

        private void btn_Comparative_Click_1(object sender, EventArgs e)
        {
            ComparativeFrm comparativeFrm = new ComparativeFrm();
            comparativeFrm.Dgv = Dgv;
            comparativeFrm.Show();
            this.Hide();

        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            LoadExcel loadExcel = new LoadExcel();
            loadExcel.Show();
            this.Hide();    
        }

        private void btn_Descriptive_Click(object sender, EventArgs e)
        {
            DescriptiveFrm descriptive = new DescriptiveFrm();
            descriptive.Dgv = Dgv;
            descriptive.Show();
            this.Hide();
        }

        private void btn_letters_Click(object sender, EventArgs e)
        {
            Letters letters = new Letters();
            letters.Dgv = Dgv;
            letters.Show();
            this.Hide();
        }

        private void btn_Correlation_Click(object sender, EventArgs e)
        {
            CorrelationFrm correlation = new CorrelationFrm();
            correlation.Dgv = Dgv;
            correlation.Show();
            this.Hide();
        }

        private void btn_Regression_Click(object sender, EventArgs e)
        {

        }

        private void btn_Relations_Click(object sender, EventArgs e)
        {
            RelationsFrm relationsFrm = new RelationsFrm();
            relationsFrm.Dgv = Dgv;
            relationsFrm.Show();
            this.Hide();
        }
    }
}
