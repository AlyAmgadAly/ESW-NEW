using ExcelScore.App_UI;
using ExcelScore.Classes;
using ExcelScore.FormsDesigns;
using ExcelScore.Properties;
using ExcelScore.StatClasses;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    public partial class Analytical_Child : Form
    {
        public Analytical_Child()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }
        public List<StatParameter> SpssParameters;
        private void Analytical_Child_Load(object sender, EventArgs e)
        {
            SpssParameters = FormDataTransfer.Get<List<StatParameter>>("SPSS_Parameters");
            PopDataGrid_AllSpssParam();
        }

        private void btn_ChooseFormat_Click(object sender, EventArgs e)
        {
            FormManager.ShowStandaloneForm<ChooseAnalyticFormat_Child>();
        }

        private void list_Nominal_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        public void PopDataGrid_AllSpssParam()
        {
            //Load Parameters from SPSS
            data_allPara.Rows.Clear();
            foreach (var spssparam in SpssParameters)
            {
                data_allPara.Rows.Add(spssparam.Name);
            }

            //Make Sorting Off 
            foreach (DataGridViewColumn column in data_allPara.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            //Scroll Wheel and disabled scroll bars
            data_allPara.MouseWheel += myDataGridView_MouseWheel;

            //Making Col Resizable
            data_allPara.AllowUserToResizeColumns = false; 
            data_allPara.Columns[0].Resizable = DataGridViewTriState.True;

            //Insert Picture Type
            InsertPicture_Type_new();
        }


        private void myDataGridView_MouseWheel(object sender, MouseEventArgs e)
        {
            int currentIndex = data_allPara.FirstDisplayedScrollingRowIndex;
            int scrollLines = SystemInformation.MouseWheelScrollLines;

            if (e.Delta > 0) // Scroll up
            {
                int newIndex = Math.Max(0, currentIndex - scrollLines);
                data_allPara.FirstDisplayedScrollingRowIndex = newIndex;
            }
            else if (e.Delta < 0) // Scroll down
            {
                int newIndex = Math.Min(data_allPara.RowCount - 1, currentIndex + scrollLines);
                data_allPara.FirstDisplayedScrollingRowIndex = newIndex;
            }
        }

        public void InsertPicture_Type_new()
        {
            // Step 1: Create a dictionary for quick row lookup
            var paraNameToRowMap = new Dictionary<string, DataGridViewRow>();

            // Fill the dictionary with the paraName and corresponding row
            foreach (DataGridViewRow row in data_allPara.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    string paraName = row.Cells[0].Value.ToString();
                    if (!paraNameToRowMap.ContainsKey(paraName))
                    {
                        paraNameToRowMap[paraName] = row;
                    }
                }
            }

            // Step 2: Process the Excel sheet and update the DataGridView rows
            foreach(var spssparam in SpssParameters)
            {
                
                string paraName = spssparam.Name;

                if (paraNameToRowMap.TryGetValue(paraName, out var targetRow))
                {
                    // Insert the image based on the type
                    if (spssparam.Type == "Nominal")
                    {
                        targetRow.Cells["ColMeasure"].Value = Resource.Final_Nominal_Color;
                    }
                    else if (spssparam.Type == "Scale")
                    {
                        targetRow.Cells["ColMeasure"].Value = Resource.FinalScale_Color;
                    }
                }

                
                
            }
        }

        private void txt_ParaName_TextChanged(object sender, EventArgs e)
        {
            string filterText = txt_ParaName.Text.ToLower(); // Convert to lowercase for case-insensitive comparison

            // Iterate through all rows in the DataGridView
            foreach (DataGridViewRow row in data_allPara.Rows)
            {
                // Skip the new row placeholder if it's visible
                if (row.IsNewRow) continue;

                // Get the value of the cell in the first column
                string cellValue = row.Cells[0].Value.ToString().ToLower(); // Convert to lowercase for case-insensitive comparison

                // Check if the cell value contains the filter text
                bool shouldShow = cellValue.Contains(filterText);

                // Set the row's visibility
                row.Visible = shouldShow;
            }
        }

        private void btn_NameLabel_Click(object sender, EventArgs e)
        {
            
        }

        private void pic_previous_Click(object sender, EventArgs e)
        {
            SelectionMainForm selectionMainForm = new SelectionMainForm();
            selectionMainForm.Show();
            this.Hide();
        }

        private void btn_Select_Click(object sender, EventArgs e)
        {
            SelectFrm_Child selectFrm_Child = new SelectFrm_Child();
            selectFrm_Child.Show();
        }
    }
}
