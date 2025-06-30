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
            string resultPath = @"C:\Users\Win\Desktop\result.txt";
            string spoPath = @"C:\Users\Win\Desktop\output.spo";

            if (File.Exists(syntaxPath)) File.Delete(syntaxPath);
            if (File.Exists(resultPath)) File.Delete(resultPath);
            if (File.Exists(spoPath)) File.Delete(spoPath);

            string syntax = $@"
GET FILE='{savFilePath.Replace(@"\", @"\\")}'.
DATASET NAME DataSet1 WINDOW=ASIS.

OMS
  /SELECT TABLES
  /IF SUBTYPES = ['Crosstabulation', 'Chi-Square Tests']
  /DESTINATION FORMAT = TEXT OUTFILE = '{resultPath.Replace(@"\", @"\\")}'.

CROSSTABS
  /TABLES=B BY A
  /FORMAT=AVALUE TABLES
  /STATISTICS=CHISQ
  /CELLS=COUNT COLUMN
  /COUNT ROUND CELL
  /METHOD=MC CIN(99) SAMPLES(10000).

OMSEND.
";
            File.WriteAllText(syntaxPath, syntax);

            Type spssType = Type.GetTypeFromProgID("SPSS.Application");
            dynamic spssApp = Activator.CreateInstance(spssType);
            dynamic syntaxDoc = spssApp.OpenSyntaxDoc(syntaxPath);
            syntaxDoc.Run();

            string outputText = null;
            int waited = 0, maxWaitMs = 3000, intervalMs = 100;

            while (waited < maxWaitMs)
            {
                try
                {
                    if (File.Exists(resultPath))
                    {
                        outputText = File.ReadAllText(resultPath);
                        dynamic outputDoc = spssApp.GetDesignatedOutputDoc();
                        outputDoc.SaveAs(spoPath);
                        if (!string.IsNullOrWhiteSpace(outputText))
                            break;
                    }
                }
                catch { }

                Thread.Sleep(intervalMs);
                waited += intervalMs;
            }

            if (!string.IsNullOrWhiteSpace(outputText))
            {
                var lines = outputText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                var chiResults = new List<ChiSquareResult>();
                var crosstabRows = new List<CrosstabRow>();
                List<string> headers = null;

                bool inCrosstab = false, inChi = false;

                for (int i = 0; i < lines.Length; i++)
                {
                    var line = lines[i].Trim();

                    if (line.Contains("B * A Crosstabulation"))
                    {
                        inCrosstab = true;
                        continue;
                    }

                    if (inCrosstab && line.StartsWith("Chi-Square Tests"))
                    {
                        inCrosstab = false;
                        inChi = true;
                        continue;
                    }

                    if (inCrosstab)
                    {
                        if (line.StartsWith("A"))
                        {
                            headers = Regex.Split(line.Trim(), @"\s+").ToList();
                            headers.Add("Total");
                        }
                        else if (line.Contains("Count") && !line.StartsWith("Total"))
                        {
                            var countParts = Regex.Split(line.Trim(), @"\s+");
                            string groupName = countParts[0];
                            var counts = countParts.Skip(2).ToList();

                            // Look ahead to % row
                            if (i + 1 < lines.Length)
                            {
                                var nextLine = lines[i + 1].Trim();
                                if (nextLine.Contains("% within A"))
                                {
                                    var percentParts = Regex.Split(nextLine.Trim(), @"\s+");
                                    var percentages = percentParts.Skip(2).ToList();

                                    crosstabRows.Add(new CrosstabRow
                                    {
                                        GroupName = groupName,
                                        Counts = counts,
                                        Percentages = percentages
                                    });

                                    i++; // Skip next line
                                }
                            }
                        }
                    }

                    if (inChi)
                    {
                        if (line.StartsWith("Pearson") || line.StartsWith("Likelih") || line.StartsWith("Fisher") || line.StartsWith("Linear"))
                        {
                            var parts = Regex.Split(line.Trim(), @"\s+");

                            chiResults.Add(new ChiSquareResult
                            {
                                Name = parts[0],
                                Value = parts.Length > 1 ? parts[1].Replace("(b)", "").Replace("(c)", "") : null,
                                df = parts.Length > 2 ? parts[2] : null,
                                AsympSig = parts.Length > 3 ? parts[3] : null,
                                MC_Sig2sided = parts.Length > 4 ? parts[4].Replace("(a)", "") : null,
                                MC_CI_Lower = parts.Length > 5 ? parts[5] : null,
                                MC_CI_Upper = parts.Length > 6 ? parts[6] : null
                            });
                        }
                    }
                }

                // Show Crosstab
                string crosstabDisplay = "Crosstab Table:\n";
                if (headers != null)
                    crosstabDisplay += "     | " + string.Join(" | ", headers) + "\n";

                foreach (var row in crosstabRows)
                {
                    crosstabDisplay += $"B = {row.GroupName}\n";
                    crosstabDisplay += "Count     : " + string.Join(" | ", row.Counts) + "\n";
                    crosstabDisplay += "% within A: " + string.Join(" | ", row.Percentages) + "\n\n";
                }
                MessageBox.Show(crosstabDisplay);

                // Show Chi-Square
                string chiDisplay = "Chi-Square Tests:\n";
                foreach (var chi in chiResults)
                {
                    chiDisplay += $"{chi.Name}: χ²={chi.Value}, df={chi.df}, p={chi.AsympSig}, MC p={chi.MC_Sig2sided}, CI=({chi.MC_CI_Lower}, {chi.MC_CI_Upper})\n";
                }
                MessageBox.Show(chiDisplay);
            }
            else
            {
                MessageBox.Show("SPSS did not generate a usable result file in time.");
            }

            try
            {
                spssApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(spssApp);
                spssApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        class ChiSquareResult
        {
            public string Name;
            public string Value;
            public string df;
            public string AsympSig;
            public string MC_Sig2sided;
            public string MC_CI_Lower;
            public string MC_CI_Upper;
        }

        class CrosstabRow
        {
            public string GroupName;         // e.g. 1.00 or 2.00
            public List<string> Counts;      // e.g. [14, 13, 4, 5, 36]
            public List<string> Percentages; // e.g. [40.0%, 65.0%, ...]
        }



    }
}
