using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using System.Windows.Forms;
using System.IO; // for file operations
using System.Data.OleDb;
using System.Data.Common;
using System.Collections;
using System.Reflection;
using System.Windows.Media;
using Random = System.Random;
using Color = System.Drawing.Color;


using Aspose.Cells;
using System.Text.RegularExpressions;

using ExcelScore.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using ExcelScore.FormsDesigns;
using ExcelScore.FormsDesigns.Comparative;

namespace ExcelScore
{

  
    public partial class LoadExcel : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        public static System.Data.DataTable MydataTable;
        ExcelFunctions excelFunctions = new ExcelFunctions();
        


        
       
        
        public LoadExcel()
        {
            InitializeComponent();
            



        }      

        //get names of titles to add to combo box
        List <string > DataGridHeaders = new List<string>();
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public void fillDataGrid(System.Data.DataTable dataTable)
        {

            datagrid_excelsheet.DataSource = dataTable;
            for (int i = 0; i < datagrid_excelsheet.Columns.Count; i++)
            {
                string header = datagrid_excelsheet.Columns[i].HeaderText.Trim();
                
                DataGridHeaders.Add(header);
                //cmb_From.Items.Add(header);
                //cmb_To.Items.Add(header);
            }


        }
        DataTable dt;
        private void pic_ImportExcel_Click(object sender, EventArgs e)
        {
            
            dt =  excelFunctions.import();
            fillDataGrid(MydataTable);
        }
        

        private void pic_AddDomain_Click(object sender, EventArgs e)
        {
            if(dt!=null)
            {
                ChooseFrm chooseFrm = new ChooseFrm();
                chooseFrm.Dgv = datagrid_excelsheet;
                chooseFrm.Show();

                this.Hide();
            }
            else
            {
                MessageBox.Show("Import excel file");
            }

            


            //NewAddDomain addDomains = new NewAddDomain();
            //addDomains.Dgv = datagrid_excelsheet;
            //addDomains.Show();
            //addDomain();

            
        }


        



        
      
        //string filepath = "";

        

       
        

        private void pic_appExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            
        }

        private void datagrid_excelsheet_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void pic_exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void panelmove_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void pic_ConvertSpssToExcel_Click(object sender, EventArgs e)
        {
            //string url = "https://secure.ncounter.de/spssconverter";
            //Process.Start(url);

            SpssFileReader spssFileReader = new SpssFileReader();
            spssFileReader.LoadSpssFile();
        }

       

       
       

        private void button1_Click(object sender, EventArgs e)
        {
            MainFormsDesign mainFormsDesign = new MainFormsDesign();
            mainFormsDesign.Show();
        }

        

        private void button1_Click_3(object sender, EventArgs e)
        {
            

            string syntaxPath = @"C:\Users\Win\Desktop\auto_crosstab.sps";
            string savFilePath = @"C:\Users\Win\Desktop\asss.sav";

            string syntax = $@"
GET
  FILE='{savFilePath.Replace(@"\", @"\\")}.'.
DATASET NAME DataSet1 WINDOW=ASIS.
CROSSTABS
  /TABLES=B BY A
  /FORMAT=AVALUE TABLES
  /STATISTICS=CHISQ
  /CELLS=COUNT COLUMN
  /COUNT ROUND CELL
  /METHOD=MC CIN(99) SAMPLES(10000).
";

            File.WriteAllText(syntaxPath, syntax);


            Type spssType = Type.GetTypeFromProgID("SPSS.Application");
            dynamic spssApp = Activator.CreateInstance(spssType);

            // Optional: Show SPSS interface (or hide it by setting to false)
            spssApp.Visible = true;

            // Run the syntax
            spssApp.ExecuteSyntax(syntaxPath);

            dynamic outputDoc = spssApp.GetDesignatedOutputDoc();
            outputDoc.SaveAs(@"C:\Users\Win\Desktop\my_output.spo");



        }
    }
}
