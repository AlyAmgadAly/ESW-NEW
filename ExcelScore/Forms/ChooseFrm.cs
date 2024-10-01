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

        public void ReturnToChooseFrm()
        {
            this.Show();
        }

        
        private void ChooseFrm_Load(object sender, EventArgs e)
        {
            

        }


        //instance for compform
        private ComparativeGroups comparativeGroupsInstance;

        private void btn_Questionnare_Click(object sender, EventArgs e)
        {
            NewAddDomain addDomains = new NewAddDomain();
            addDomains.Dgv = Dgv;
            addDomains.Show();
            this.Hide();
        }






        private void btn_Comparative_Click_1(object sender, EventArgs e)
        {

            //transferring to compform
            if (comparativeGroupsInstance == null || comparativeGroupsInstance.IsDisposed)
            {
                // Create a new instance of ComparativeGroups if it doesn't exist or was closed
                comparativeGroupsInstance = new ComparativeGroups(this);
            }

            // Hide ChooseFrm and show ComparativeGroups
            comparativeGroupsInstance.Dgv = Dgv;
            comparativeGroupsInstance.NormalityParaNameList_ComparaGroups = NormalityParaNameList_ChooseFrm;
            this.Hide();
            
            comparativeGroupsInstance.Show();

            //ComparativeGroups comparativeGroups = new ComparativeGroups();
            //comparativeGroups.Dgv = Dgv;
            //comparativeGroups.Show();
            //this.Hide();

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
            Regression_Frm regression_ = new Regression_Frm();
            regression_.Dgv = Dgv;
            regression_.Show();
            this.Hide();
        }

        private void btn_Relations_Click(object sender, EventArgs e)
        {
            RelationsFrm relationsFrm = new RelationsFrm();
            relationsFrm.Dgv = Dgv;
            relationsFrm.Show();
            this.Hide();
        }
        Workbook NormalityWorkbook;
        Worksheet NormalityWorkSheet;

        public Dictionary<string, string> NormalityParaNameList_ChooseFrm {get;set;} = new Dictionary<string, string>();
        public void GetNormality()
        {
            if(NormalityParaNameList_ChooseFrm != null)
            {
                NormalityParaNameList_ChooseFrm.Clear();
            }
            
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            if (op.ShowDialog() == DialogResult.OK)
            {
                string filepath  = op.FileName;
                NormalityWorkbook = new Aspose.Cells.Workbook(filepath);

                // Accessing the first worksheet in the Excel file
                NormalityWorkSheet = NormalityWorkbook.Worksheets[0];
            }

            if(NormalityWorkbook != null)
            {
                for (int row = 3; row < NormalityWorkSheet.Cells.MaxDataRow + 1; row++)
                {
                    if (NormalityWorkSheet.Cells[row, 0].Value != null && NormalityWorkSheet.Cells[row, 1].Value != null)
                    {
                        NormalityParaNameList_ChooseFrm.Add(NormalityWorkSheet.Cells[row, 1].Value.ToString(), NormalityWorkSheet.Cells[row, 0].Value.ToString());

                    }

                }
            }


            if (comparativeGroupsInstance != null && !comparativeGroupsInstance.IsDisposed)
            {
                comparativeGroupsInstance.UpdateData(NormalityParaNameList_ChooseFrm);
            }


        }
        private void btn_GetNormality_Click(object sender, EventArgs e)
        {
            GetNormality();


        }

        private void btn_Coding_Click(object sender, EventArgs e)
        {
            Coding_Frm coding_ = new Coding_Frm();
            coding_.Dgv = Dgv;
            coding_.Show();
            this.Hide();
        }
    }
}
