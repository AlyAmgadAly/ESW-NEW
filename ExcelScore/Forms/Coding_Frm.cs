using Aspose.Cells;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
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
    public partial class Coding_Frm : Form
    {
        public Coding_Frm()
        {
            InitializeComponent();
        }
        public DataGridView Dgv { get; set; }

        static ExcelFunctions excelFunctions = new ExcelFunctions();

        public Aspose.Cells.Workbook OrginalWorkbook = excelFunctions.GetWorkbook();

        public Aspose.Cells.Worksheet Original_Sheet1 = excelFunctions.GetWorksheet();

        public Aspose.Cells.Worksheet Original_Sheet2 = excelFunctions.GetSheet2();


        public static Aspose.Cells.Workbook New_Workbook = excelFunctions.GetNewWorkBook();


        public Aspose.Cells.Worksheet New_Sheet1 = New_Workbook.Worksheets[0];
        public string CheckNominalOrScale(int col)
        {
            string Nominal_Or_Scale = "";

            if (Original_Sheet2.Cells[1, col].Value != null)
            {
                Nominal_Or_Scale = Original_Sheet2.Cells[1, col].Value.ToString();
            }
            return Nominal_Or_Scale;
        }

        public void PerformScaleCheck(int col)
        {
            if (Original_Sheet1.Cells[0 , col].Value.ToString() == New_Sheet1.Cells[0 , col].Value.ToString())
            {
                for (int row = 0; row <= Original_Sheet2.Cells.MaxDataRow; row++)
                {

                }

            }

            
        }

        public void CheckData()
        {
            for (int col = 0; col <= Original_Sheet2.Cells.MaxDataColumn; col++)
            {
                string ReturnedType = CheckNominalOrScale(col);
                
                if(ReturnedType == "Scale")
                {

                }
            }
        }

        private void Coding_Frm_Load(object sender, EventArgs e)
        {
            
        }
    }
}
