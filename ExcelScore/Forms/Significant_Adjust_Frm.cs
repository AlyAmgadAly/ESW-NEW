using Aspose.Cells;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Windows.Forms;
using Accord.IO;
using System.IO;
using DocumentFormat.OpenXml;
using System.Diagnostics;
using MathNet.Numerics.Statistics;
using DocumentFormat.OpenXml.Drawing.Charts;
using Humanizer;
using static Humanizer.On;

namespace ExcelScore.Forms
{
    public partial class Significant_Adjust_Frm : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        

        public Significant_Adjust_Frm()
        {
            InitializeComponent();
            
            
        }

        public ComparativeTable comparativeTable_ToModify_SignifAdj { get; set; }

        static ExcelFunctions excelFunctionsobj = new ExcelFunctions();
        Aspose.Cells.Workbook workbook = excelFunctionsobj.GetWorkbook();
        Aspose.Cells.Worksheet worksheet = excelFunctionsobj.GetWorksheet();

        private void CenterOnScreen()
        {
            var screen = Screen.FromControl(this);
            int x = (screen.Bounds.Width - this.Width) / 2 + screen.Bounds.X;
            int y = (screen.Bounds.Height - this.Height) / 2 + screen.Bounds.Y;
            this.Location = new System.Drawing.Point(x, y);
        }
        public void FillComboBox()
        {

            foreach (var parameter in comparativeTable_ToModify_SignifAdj.Parameters)
            {
                if(parameter.IsGroup)
                {
                    foreach (var kvp in parameter.DIC_LablesIfNomainal)
                    {
                        cmb_TotalOrGroup.Items.Add(kvp.Value);
                    }
                    continue;
                }
                cmb_ParameterResult.Items.Add(parameter.Name);
            }
        }
        private void Significant_Adjust_Frm_Load(object sender, EventArgs e)
        {
            CenterOnScreen();
            data_allPara.CellValueChanged -= data_allPara_CellValueChanged;
            FillDatagridview();
            FillComboBox();
            data_allPara.CellValueChanged += data_allPara_CellValueChanged;
            
        }

        public List<double> GetColData(int col)
        {
            List<double> Values = new List<double>();

            int rowCount = worksheet.Cells.MaxDataRow;


            for (int i = 1; i <= rowCount; i++)
            {
                if(worksheet.Cells[i, col].Value.ToString() == ".")
                {
                    continue;
                }
                string value = worksheet.Cells[i, col].Value.ToString();
                Values.Add(double.Parse(value));
            }

            return Values;

        }
        public void GetN(int ParameterColumnIndex , int GroupColIndex , Classes.Parameter GroupParameter)
        {
            int rowCount = worksheet.Cells.MaxDataRow;

            List<double> ParameterValues = new List<double>();
            string ParameterName = cmb_ParameterResult.Text;
            int Count = -1;



            if (cmb_TotalOrGroup.Text == "Total")
            {
                ParameterValues = GetColData(ParameterColumnIndex);
                Count = ParameterValues.Count;

                
                lbl_N.Text = "(n = " + Count + " )";
            }
            else
            {
                // A group selected
                string GroupName = cmb_TotalOrGroup.Text;
                double requiredGroupValue = -1;

                foreach (var kvp in GroupParameter.DIC_LablesIfNomainal)
                {
                    if(kvp.Value == GroupName)
                    {
                        requiredGroupValue = kvp.Key;
                    }
                }

                for (int i = 1; i <= rowCount; i++)
                {
                    string GroupValue = worksheet.Cells[i, GroupColIndex].Value.ToString();
                    string ParameterValue = worksheet.Cells[i, ParameterColumnIndex].Value.ToString();

                    if ((GroupValue == ".") || (ParameterValue == "."))
                    {
                        continue;
                    }
                    if(double.Parse(GroupValue) == requiredGroupValue)
                    {
                        ParameterValues.Add(double.Parse(ParameterValue));
                    }
                    
                }
                Count = ParameterValues.Count;
                lbl_N.Text = "(n = " + Count + " )";

            }





        }
        public (int , int) GetColumnIndexes(string AParameterName , string AGroupName)
        {
            int groupColumnIndex = -1;
            int ParameterColumnIndex = -1;
            for (int excelcol = 0; excelcol <= worksheet.Cells.MaxDataColumn; excelcol++)
            {
                if (worksheet.Cells[0, excelcol].Value?.ToString() == AGroupName)
                {
                    groupColumnIndex = excelcol;
                }
                else if(worksheet.Cells[0, excelcol].Value?.ToString() == AParameterName)
                {
                    ParameterColumnIndex = excelcol;
                }

                if((groupColumnIndex != -1) && (ParameterColumnIndex != -1))
                {
                    break;
                }
            }


            return (ParameterColumnIndex, groupColumnIndex);
        }

        public string ConvertDoubletoStringDec(double number)
        {
            string strnum = number.ToString("0.00");
            return strnum;
        }
        public double CalculateMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;
            int middle = n / 2;

            if (n == 1)
            {
                return values[0];
            }
            else if (n % 2 == 0)
            {

                // Even number of elements, average the middle two
                return (values[middle - 1] + values[middle]) / 2.0;
            }

            else
            {

                // Odd number of elements, return the middle one
                return values[middle];
            }
        }
        public double CalculateLowerMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;


            if (n % 2 == 0)
            {
                int middle = n / 2;
                // Even number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(0, middle));
            }
            else
            {
                //The fix was here we get same range from zero to middle but middle index is different

                int middle = (n + 1) / 2;

                return CalculateMedian(values.GetRange(0, middle));
            }

        }

        public double CalculateUpperMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;


            if (n % 2 == 0)
            {
                int middle = n / 2;
                // Even number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle));
            }
            else if (n == 1)
            {
                return values[0];
            }
            else if (n == 3)
            {
                return (values[1] + values[2]) / 2.0;
            }

            else
            {
                int middle = (n + 1) / 2;
                // Odd number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle - 1));

            }
        }
        public void DescriptiveEquationInsert(List<double> ParameterValues)
        {
            double Min = -1;
            double Max = -1;

            double Mean = -1;
            double SD = -1;
            double medianValue = -1;
            double perc25th = -1;
            double perc75th = -1;


            Min = ParameterValues.Min();
            Max = ParameterValues.Max();
            Mean = ParameterValues.Average();
            SD = ParameterValues.StandardDeviation();



            int middleIndex = ParameterValues.Count / 2;
            if (ParameterValues.Count % 2 == 0)
            {
                // For even count of elements, take the average of the two middle values
                double middleValue1 = ParameterValues.OrderBy(x => x).ElementAt(middleIndex - 1);
                double middleValue2 = ParameterValues.OrderBy(x => x).ElementAt(middleIndex);
                medianValue = (middleValue1 + middleValue2) / 2.0;
            }
            else
            {
                // For odd count of elements, directly take the middle value
                medianValue = ParameterValues.OrderBy(x => x).ElementAt(middleIndex);
            }

            perc25th = CalculateLowerMedian(ParameterValues);
            perc75th = CalculateUpperMedian(ParameterValues);

            txt_MinMax.Text = ConvertDoubletoStringDec(Min) + " - " + ConvertDoubletoStringDec(Max);
            txt_MeanSD.Text = ConvertDoubletoStringDec(Mean) + " ± " + ConvertDoubletoStringDec(SD);
            txt_MedianIQR.Text = ConvertDoubletoStringDec(medianValue) + "( " + ConvertDoubletoStringDec(perc25th) + " - " + ConvertDoubletoStringDec(perc75th) + " )";
        }
        public void GetDescriptive(int ParameterColIndex, int GroupColIndex, Classes.Parameter GroupParameter)
        {
            int rowCount = worksheet.Cells.MaxDataRow;

            List<double> ParameterValues = new List<double>();

            

            if (cmb_TotalOrGroup.Text == "Total")
            {
                ParameterValues = GetColData(ParameterColIndex);

                DescriptiveEquationInsert(ParameterValues);


            }

            else
            {
                string GroupName = cmb_TotalOrGroup.Text;
                double requiredGroupValue = -1;

                foreach (var kvp in GroupParameter.DIC_LablesIfNomainal)
                {
                    if (kvp.Value == GroupName)
                    {
                        requiredGroupValue = kvp.Key;
                    }
                }

                for (int i = 1; i <= rowCount; i++)
                {
                    string GroupValue = worksheet.Cells[i, GroupColIndex].Value.ToString();
                    string ParameterValue = worksheet.Cells[i, ParameterColIndex].Value.ToString();

                    if ((GroupValue == ".") || (ParameterValue == "."))
                    {
                        continue;
                    }
                    if (double.Parse(GroupValue) == requiredGroupValue)
                    {
                        ParameterValues.Add(double.Parse(ParameterValue));
                    }
                }


                DescriptiveEquationInsert(ParameterValues);



            }
        }

        public void GetTest(Classes.Parameter Parameter , Classes.Parameter GroupParameter , int ParameterColIndex , int GroupColIndex )
        {
            List<double> Groups = GetColData(GroupColIndex);

            int GroupCount = Groups.GroupBy(g => g)
                                   .Where(g => g.Count() > 1)
                                   .Count();

            if (Parameter.NominalOrScale== "Scale")
            {
                if(Parameter.NormalOrAbnormal == "Normal")
                {



                }
            }
        }

        public void CalculateData()
        {
            try
            {
                if (cmb_ParameterResult.SelectedIndex != -1)
                {
                    if (cmb_TotalOrGroup.SelectedIndex != -1)
                    {

                        Classes.Parameter GroupParameter = null;

                        string ParameterName = cmb_ParameterResult.Text;
                        


                        int GroupColIndex = -1;
                        int ParameterColIndex = -1;


                        foreach (var parameter in comparativeTable_ToModify_SignifAdj.Parameters)
                        {
                            if (parameter.IsGroup)
                            {
                                GroupParameter = parameter;
                                break;
                            }
                        }

                        Classes.Parameter MyParameter = comparativeTable_ToModify_SignifAdj.Parameters.SingleOrDefault(p => p.Name == ParameterName);

                        (ParameterColIndex, GroupColIndex) = GetColumnIndexes(ParameterName, GroupParameter.Name);

                        GetN(ParameterColIndex, GroupColIndex, GroupParameter);

                        GetDescriptive(ParameterColIndex, GroupColIndex, GroupParameter);

                        GetTest(MyParameter, GroupParameter , ParameterColIndex , GroupColIndex);



                    }
                }
            }
            catch(IOException)
            {
                MessageBox.Show("Close Excel File and save again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            
        }

        public void AddDataGridColumns()
        {
            const int columnWidth = 70;
            const int maxDataGridWidth = 500;

            var groupParameter = comparativeTable_ToModify_SignifAdj.Parameters.FirstOrDefault(p => p.IsGroup);
            data_allPara.Columns.Add(groupParameter.Name, groupParameter.Name);

            foreach (var parameter in comparativeTable_ToModify_SignifAdj.Parameters)
            {
                if (!parameter.IsGroup)
                {
                    data_allPara.Columns.Add(parameter.Name, parameter.Name);
                }
            }

            data_allPara.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            int totalWidth = comparativeTable_ToModify_SignifAdj.Parameters.Count * columnWidth;
            data_allPara.Width = totalWidth > maxDataGridWidth ? maxDataGridWidth : totalWidth;
            data_allPara.Height = 620;

            data_allPara.ScrollBars = ScrollBars.Both;

            for (int i = 0; i < data_allPara.Columns.Count; i++)
            {
                data_allPara.Columns[i].Width = columnWidth;
            }

            data_allPara.CellPainting += DataGridView_CellPainting;
            this.Controls.Add(data_allPara);
        }

        

        public void DeleteSheetsExcept(string filePath)
        {
            try
            {
                var sheetsToKeep = new[] { "Rawdata", "ListVariables", "Source" };
                var sheetsToDelete = new List<Aspose.Cells.Worksheet>();

                foreach (var sheet in workbook.Worksheets)
                {
                    if (!sheetsToKeep.Contains(sheet.Name))
                    {
                        sheetsToDelete.Add(sheet);
                    }
                }

                foreach (var sheetName in sheetsToDelete)
                {
                    workbook.Worksheets.RemoveAt(workbook.Worksheets.IndexOf(sheetName));
                }

                workbook.Save(filePath);
            }
            catch (IOException)
            {
                MessageBox.Show("Close Excel File and save again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                e.Paint(e.CellBounds, DataGridViewPaintParts.All);
                using (Pen pen = new Pen(System.Drawing.Color.Black))
                {
                    e.Graphics.DrawLine(pen, e.CellBounds.Right - 1, e.CellBounds.Top,
                                        e.CellBounds.Right - 1, e.CellBounds.Bottom);
                }
                e.Handled = true;
            }
        }

        public void AddDataGridRows()
        {
            data_allPara.Rows.Clear();
            var groupParameter = comparativeTable_ToModify_SignifAdj.Parameters.FirstOrDefault(p => p.IsGroup);
            int rowCount = worksheet.Cells.MaxDataRow;

            for (int i = 0; i < rowCount; i++)
            {
                data_allPara.Rows.Add();
            }

            if (groupParameter != null)
            {
                int groupColumnIndex = -1;
                for (int excelcol = 0; excelcol <= worksheet.Cells.MaxDataColumn; excelcol++)
                {
                    if (worksheet.Cells[0, excelcol].Value?.ToString() == groupParameter.Name)
                    {
                        groupColumnIndex = excelcol;
                        break;
                    }
                }

                if (groupColumnIndex != -1)
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        data_allPara.Rows[i - 1].Cells[0].Value = worksheet.Cells[i, groupColumnIndex].Value?.ToString();
                    }
                }
            }

            int insertDataGridColumn = 1;
            foreach (var parameter in comparativeTable_ToModify_SignifAdj.Parameters)
            {
                if (parameter.IsGroup) continue;

                for (int excelcol = 0; excelcol <= worksheet.Cells.MaxDataColumn; excelcol++)
                {
                    if (worksheet.Cells[0, excelcol].Value?.ToString() == parameter.Name)
                    {
                        for (int i = 1; i <= rowCount; i++)
                        {
                            data_allPara.Rows[i - 1].Cells[insertDataGridColumn].Value = worksheet.Cells[i, excelcol].Value?.ToString();
                        }
                        break;
                    }
                }
                insertDataGridColumn++;
            }
        }

        public void FillDatagridview()
        {
            AddDataGridColumns();
            AddDataGridRows();
            foreach (DataGridViewColumn column in data_allPara.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            data_allPara.SelectionMode = DataGridViewSelectionMode.CellSelect;

            // Optionally, disable row header selection to make it clear that only cells can be selected
            data_allPara.RowHeadersVisible = false;
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Update_Click_1(object sender, EventArgs e)
        {
            string filepath = ExcelFunctions.filepath;
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);
            worksheet = workbook.Worksheets[0];
            MessageBox.Show("File Updated");
            AddDataGridRows();
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void data_allPara_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void data_allPara_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // Retrieve the new value
                var newValue = data_allPara.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                

                // Update Excel cell
                UpdateExcelCell(e.RowIndex, e.ColumnIndex, newValue);

                

            }
        }

        private void UpdateExcelCell(int rowIndex, int columnIndex, object newValue)
        {
            // Locate the corresponding Excel column by name
            string dataGridColumnName = data_allPara.Columns[columnIndex].Name;
            int excelCol = -1;
            for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
            {
                if (worksheet.Cells[0, col].Value?.ToString() == dataGridColumnName)
                {
                    excelCol = col;
                    break;
                }
            }

            // If the corresponding Excel column was found, update the cell
            if (excelCol != -1)
            {
                try
                {
                    int excelRow = rowIndex + 1;


                    if (double.TryParse((string)newValue, out double parsedValue))
                    {
                        worksheet.Cells[excelRow, excelCol].PutValue(parsedValue);
                    }
                    else if ((string)newValue == ".")
                    {
                        worksheet.Cells[excelRow, excelCol].PutValue(".");
                    }
                    else
                    {
                        MessageBox.Show("Invalid input. Please enter a valid number or '.'", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    workbook.Save(ExcelFunctions.filepath);


                    CalculateData();
                }
                catch (IOException)
                {
                    MessageBox.Show("Close Excel File and save again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        private void btn_Done_Click(object sender, EventArgs e)
        {
            DeleteSheetsExcept(ExcelFunctions.filepath);
        }

        private void cmb_ParameterResult_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cmb_ParameterResult.SelectedIndex != - 1)
            {
                if(cmb_TotalOrGroup.SelectedIndex != -1)
                {
                    CalculateData();
                }
            }
        }

        private void cmb_TotalOrGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_ParameterResult.SelectedIndex != -1)
            {
                if (cmb_TotalOrGroup.SelectedIndex != -1)
                {
                    CalculateData();
                }
            }
        }
    }
}
