using Accord.Statistics.Distributions.Univariate;
using Aspose.Cells;
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
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class ParameterInfo : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        public ParameterInfo()
        {
            InitializeComponent();
        }

        public List<string> GetCodesIfNominal(string parameterName  , Worksheet worksheet)
        {
            List<string> codes = new List<string>();    

            for(int col = 0;col<= worksheet.Cells.MaxDataColumn;col++)
            {
                if (worksheet.Cells[0 , col].Value!=null)
                {
                    string valueCell = worksheet.Cells[0, col].Value.ToString();

                    if(valueCell == parameterName)
                    {
                        if(worksheet.Cells[1, col].Value.ToString() == "Nominal")
                        {
                            for (int row = 2; row <= worksheet.Cells.MaxDataRow; row++)
                            {
                                if (Sheet2.Cells[row, col].Value != null)
                                {
                                    codes.Add(worksheet.Cells[row, col].Value.ToString());
                                }
                                    
                            }
                        }

                        break;


                    }


                    
                }


                
            }




            return codes;


        }
        public void UpdateParameter(string paraName)
        {
            data_allPara.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            data_allPara.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;


            data_allPara.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            data_allPara.Rows.Clear();
            list_Codes.Items.Clear();

            lbl_ParameterName.Text = $"Parameter: {paraName}"; // Update label

            lbl_Measure.Text = Parameter.GetParameterMeasure(paraName, Sheet2);

            List<string> codes = GetCodesIfNominal(paraName, Sheet2);

            if(codes.Count > 0) 
            { 
                foreach (string code in codes) 
                {
                    list_Codes.Items.Add(code);
                }
            }


            List<string> Data = Parameter.GetParameterData(paraName, Sheet1);


            


            foreach (string item in Data)
            {
                data_allPara.Rows.Add(item);
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

        public Worksheet Sheet1 = ExcelFunctions.worksheet;

        public Worksheet Sheet2 = ExcelFunctions.Sheet2;

        public string ParameterName_ParaInfoFrm { get; set; }


        private void ParameterInfo_Load(object sender, EventArgs e)
        {
            
        }

        private void pic_Minimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void pic_CloseChooseTable_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
    }
}
