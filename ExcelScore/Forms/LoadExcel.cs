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
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;
using SkiaSharp;
using ExcelScore.Classes;
using static SkiaSharp.HarfBuzz.SKShaper;


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
        List<string> DataGridHeaders = new List<string>();
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

            dt = excelFunctions.import();
            fillDataGrid(MydataTable);
        }


        private void pic_AddDomain_Click(object sender, EventArgs e)
        {
            if (dt != null)
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

            
            //foreach (var item in tests)
            //{
            //    foreach (var row in item.)
            //}
        }






        private void button1_Click(object sender, EventArgs e)
        {
            MainFormsDesign mainFormsDesign = new MainFormsDesign();
            mainFormsDesign.Show();
        }



        private void button1_Click_3(object sender, EventArgs e)
        {
            //string resultText;


            ////string output = SPSS_TestRunner.RunDescriptiveSyntax("Groups", new List<string> { "Age" }, SPSS_TestRunner.DefaultDescriptiveStats, out resultText);
            ////var results = SPSS_TestRunner.ParseDescriptiveOutput(resultText , new List<string> { "Age"} , groupLabels);


            //string percentile = SPSS_TestRunner.RunPercentilesSyntax("Groups", new List<string> { "Age" }, out resultText);
            //var percentiles = SPSS_TestRunner.ParseTukeyTukeyHingesOnly(percentile, new List<string> { "Age" }, groupLabels , "Groups");

            //foreach (var result in percentiles)
            //{
            //    MessageBox.Show($"Variable: {result.VariableName}");

            //    MessageBox.Show("  Total:");
            //    foreach (var kvp in result.TotalPercentiles)
            //        MessageBox.Show($"    {kvp.Key}: {kvp.Value}");

            //    MessageBox.Show("  Groups:");
            //    foreach (var group in result.GroupPercentiles)
            //    {
            //        MessageBox.Show($"    {group.Key}: 25={group.Value["25"]}, 50={group.Value["50"]}, 75={group.Value["75"]}");
            //    }
            //}

            var groupLabels = new List<string> { "Patient", "Control", "3.00", "4.00" , "Total"};

            List<(string Name, string Type)> parameters = new List<(string Name, string Type)>
            {
                
                ("Sex" , "Nominal"),
                ("Age","Scale" ),
                ("Weight","Scale" ),
                //("Height","Scale" )
            };

            
            var finalResults = SPSSUnifiedRunner.RunAllFromUnifiedSyntax("Groups", parameters, groupLabels);

            var scaleResults = finalResults
    .Where(r => r.Type == "Scale" && r.Descriptives != null)
    .ToList();

            foreach (var result in scaleResults)
            {
                var desc = result.Descriptives;
                var msg = $"--- {desc.VariableName} ---\n";

                foreach (var group in desc.Stats_Groups)
                {
                    msg += $"Group: {group.Key}\n";
                    foreach (var stat in group.Value)
                    {
                        msg += $"[{stat.Key}, {stat.Value}]\n";
                    }
                }

                msg += "Total:\n";
                foreach (var stat in desc.Stats_Total)
                {
                    msg += $"[{stat.Key}, {stat.Value}]\n";
                }

                MessageBox.Show(msg, desc.VariableName);
            }



        }





    }
}
