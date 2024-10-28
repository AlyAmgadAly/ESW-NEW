using Aspose.Cells;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
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
    public partial class Significant_Adjust_Frm : Form
    {
        public Significant_Adjust_Frm()
        {
            InitializeComponent();
        }

        public ComparativeTable comparativeTable_ToModify_SignifAdj { get; set; }

        static ExcelFunctions excelFunctionsobj = new ExcelFunctions();

        Worksheet worksheet = excelFunctionsobj.GetWorksheet();
        private void CenterOnScreen()
        {
            // Get the bounds of the screen where the form will be centered
            var screen = Screen.FromControl(this);
            int x = (screen.Bounds.Width - this.Width) / 2 + screen.Bounds.X;
            int y = (screen.Bounds.Height - this.Height) / 2 + screen.Bounds.Y;

            // Set the form's location
            this.Location = new System.Drawing.Point(x, y);
        }
        private void Significant_Adjust_Frm_Load(object sender, EventArgs e)
        {
            CenterOnScreen();
            FillDatagridview();
        }

        public void AddDataGridColumns()
        {
            var groupParameter = comparativeTable_ToModify_SignifAdj.Parameters.FirstOrDefault(p => p.IsGroup);
            data_allPara.Columns.Add(groupParameter.Name, groupParameter.Name);

            foreach (var parameter in comparativeTable_ToModify_SignifAdj.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }
                else
                {
                    data_allPara.Columns.Add(parameter.Name, parameter.Name);
                }
            }
            data_allPara.Width = comparativeTable_ToModify_SignifAdj.Parameters.Count * 70;

            for (int i = 1; i <= comparativeTable_ToModify_SignifAdj.Parameters.Count; i++)
            {
                data_allPara.Columns[i - 1].Width = 70;
            }

            data_allPara.CellPainting += DataGridView_CellPainting;

            this.Controls.Add(data_allPara);
        }

        private void DataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // Check if the cell is not a header cell
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // Draw the cell content
                e.Paint(e.CellBounds, DataGridViewPaintParts.All);

                // Draw the right border for the cell
                using (Pen pen = new Pen(System.Drawing.Color.Black))
                {
                    // Draw the line at the right edge of the cell
                    e.Graphics.DrawLine(pen, e.CellBounds.Right - 1, e.CellBounds.Top,
                                        e.CellBounds.Right - 1, e.CellBounds.Bottom);
                }

                // Mark the event as handled
                e.Handled = true;
            }
        }
        public void AddDataGridRows()
        {
            // Get the group parameter
            var groupParameter = comparativeTable_ToModify_SignifAdj.Parameters.FirstOrDefault(p => p.IsGroup);

            if (groupParameter != null)
            {
                data_allPara.Rows.Clear();
                for (int excelcol = 0; excelcol < worksheet.Cells.MaxDataColumn; excelcol++) 
                {
                    if (worksheet.Cells[0 , excelcol].Value != null)
                    {
                        if (worksheet.Cells[0, excelcol].Value.ToString() == groupParameter.Name)
                        {
                            for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                            {
                                int rowIndex = data_allPara.Rows.Add();
                                // Set the value for the first cell in the newly added row
                                data_allPara.Rows[rowIndex].Cells[0].Value = worksheet.Cells[i, excelcol].Value.ToString();
                            }
                            break;
                        }

                    }
                    
                }
            }
            
        }
        public void FillDatagridview()
        {
            AddDataGridColumns();
            AddDataGridRows();
        }
        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Update_Click_1(object sender, EventArgs e)
        {
            string filepath = ExcelFunctions.filepath;
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);

            // Accessing the first worksheet in the Excel file
            worksheet = workbook.Worksheets[0];

            MessageBox.Show("File Updated");

            AddDataGridRows();
        }
    }
}
