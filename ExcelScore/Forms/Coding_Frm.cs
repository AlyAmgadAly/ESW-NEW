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


        public static string New_Workbook_Filepath = excelFunctions.GetNewWorkBook();

        public  static Aspose.Cells.Workbook New_Workbook = new Aspose.Cells.Workbook(New_Workbook_Filepath);

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
            bool hasfalse = false;

            if (Original_Sheet1.Cells[0, col].Value != null && New_Sheet1.Cells[0, col].Value != null)
            {
                if (Original_Sheet1.Cells[0, col].Value.ToString() == New_Sheet1.Cells[0, col].Value.ToString())
                {
                    string ParameterCheck = $"Checking Parameter : {Original_Sheet1.Cells[0, col].Value.ToString()}";
                    
                    for (int row = 1; row <= Original_Sheet1.Cells.MaxDataRow; row++)
                    {
                        string OriginalSheetValue = Original_Sheet1.Cells[row, col].Value.ToString();
                        string NewSheetValue = New_Sheet1.Cells[row, col].Value.ToString();

                        if (!(OriginalSheetValue == NewSheetValue))
                        {
                            MessageBox.Show($"Data Not the same at Parameter : {Original_Sheet1.Cells[0, col].Value.ToString()} at row {row + 1}\nData is : {OriginalSheetValue} at coding\nData is : {NewSheetValue} at Master Sheet");
                            hasfalse = true;
                        }
                    }

                    if (hasfalse == false)
                    {
                        MessageBox.Show($"{ParameterCheck} \n All correct");
                    }

                }
            }
            else if(Original_Sheet1.Cells[0, col].Value == null || New_Sheet1.Cells[0, col].Value == null)
            {
                MessageBox.Show($"Data Column Missing at Parameter : {Original_Sheet1.Cells[0, col].Value.ToString()}");
            }
            

            
        }

        public void CheckData()
        {
            for (int col = 0; col <= Original_Sheet2.Cells.MaxDataColumn; col++)
            {
                string ReturnedType = CheckNominalOrScale(col);
                
                if(ReturnedType == "Scale")
                {
                    PerformScaleCheck(col);
                }
            }
        }

        public void UpdateWorkbook(string Workbook_filepath)
        {
            //Update Master Sheet

            Aspose.Cells.Workbook New_Workbook = new Aspose.Cells.Workbook(Workbook_filepath);
            New_Sheet1 = New_Workbook.Worksheets[0];



            //Update Coding Sheet
            string filepath = ExcelFunctions.filepath;
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);
            Original_Sheet1 = workbook.Worksheets[0];
            Original_Sheet2 = workbook.Worksheets[1];


            MessageBox.Show("File Updated");



        }
        private void Coding_Frm_Load(object sender, EventArgs e)
        {
            //UpdateWorkbook(New_Workbook_Filepath);
            
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
            ChooseFrm chooseFrm = new ChooseFrm();
            chooseFrm.Dgv = Dgv;
            chooseFrm.Show();
        }

        private void btn_Update_Click(object sender, EventArgs e)
        {
            UpdateWorkbook(New_Workbook_Filepath);
        }

        private void btn_Done_Click(object sender, EventArgs e)
        {
            CheckData();
        }
    }
}
