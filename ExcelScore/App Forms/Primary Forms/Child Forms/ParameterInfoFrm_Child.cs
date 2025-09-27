using ExcelScore.App_UI;
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
    public partial class ParameterInfoFrm_Child : Form
    {
        public ParameterInfoFrm_Child()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }

        public void UpdateParameter(StatParameter viewedParameter)
        {
            data_allPara.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            data_allPara.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            data_allPara.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            data_allPara.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            data_allPara.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            data_allPara.Rows.Clear();
            list_Codes.Items.Clear();

            lbl_ParameterName.Text = $"{viewedParameter.Name}"; // Update label
            lbl_Measure.Text = $"Measure: {viewedParameter.Type}";


            
            List<int> SerialList= Enumerable.Range(1, viewedParameter.RawValues.Count).ToList();
            for (int i = 1; i <= SerialList.Count; i++)
            {
                data_allPara.Rows.Add(i);
                //data_allPara.Rows[i-1].Cells[1].Value = viewedParameter.RawValues[i-1];
            }


            //foreach (var value in viewedParameter.RawValues)
            //{
            //    data_allPara.Rows.Add(value);

            //}

            

            foreach (var ValueLabel in viewedParameter.ValueLabels)
            {
                string value_Label = $"                         {ValueLabel.Key} = {ValueLabel.Value}";
                list_Codes.Items.Add(value_Label);
            }

            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Normal;
            }

            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(
                Screen.PrimaryScreen.WorkingArea.Width - this.Width - 200, // 10px padding from the right edge
                (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2 // Center vertically
            );

            this.BringToFront(); // Bring the form to front if minimized
            this.Activate();
        }

        private void ParameterInfoFrm_Child_Load(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in data_allPara.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
    }
}
