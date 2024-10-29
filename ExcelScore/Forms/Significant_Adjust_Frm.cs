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

        private void Significant_Adjust_Frm_Load(object sender, EventArgs e)
        {
            CenterOnScreen();
            data_allPara.CellValueChanged -= data_allPara_CellValueChanged;
            FillDatagridview();
            InitializeOldValues(); // Initialize old values here
            data_allPara.CellValueChanged += data_allPara_CellValueChanged;
            
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

        private void InitializeOldValues()
        {
            for (int i = 0; i < data_allPara.Rows.Count; i++)
            {
                for (int j = 0; j < data_allPara.Columns.Count; j++)
                {
                    
                }
            }
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
                    worksheet.Cells[excelRow, excelCol].PutValue(Convert.ToDouble(newValue));
                    workbook.Save(ExcelFunctions.filepath);
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

        

        
    }
}
