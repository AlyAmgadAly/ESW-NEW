using Accord.Statistics.Kernels;
using Aspose.Cells;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Worksheet = Aspose.Cells.Worksheet;

namespace ExcelScore.Forms
{
    public partial class Nursing : Form
    {
        public Nursing()
        {
            InitializeComponent();
        }
        public DataGridView Dgv { get; set; }

        Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(ExcelFunctions.filepath);
        
        public List<Tool> AllTools = new List<Tool>();

        public void AddToolIfNotExists(Tool newTool)
        {
            if (!AllTools.Any(tool => tool.ToolNumber == newTool.ToolNumber || tool.ToolName == newTool.ToolName))
            {
                AllTools.Add(newTool);
            }
        }
        public void ReadToolsAndSubscales()
        {
            Worksheet DetailsSheet = workbook.Worksheets["Details"];

            GetToolsNumberandName(DetailsSheet);

            GetEachToolSubscales(DetailsSheet);


            


            foreach (Tool tool in AllTools)
            {
                MessageBox.Show(tool.ToolName);
                foreach (var subscale in tool.Subscales)
                {
                    MessageBox.Show(subscale.Subscale_Full_Name);
                }
            }

            // getToolName();

        }
        public void GetEachToolSubscales(Worksheet ADetailsSheet)
        {
            try
            {
                foreach (Tool tool in AllTools)
                {
                    int row = 0;
                    int maxRow = ADetailsSheet.Cells.MaxDataRow;

                    while (row <= maxRow) // Ensure we process the last row
                    {
                        string firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                        string secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                        // Check if the row has any data (at least one non-null, non-empty value)
                        bool rowHasData = !(string.IsNullOrWhiteSpace(firstColumn) && string.IsNullOrWhiteSpace(secondColumn));

                        if ((firstColumn != null) && (secondColumn != null))
                        {
                            if (secondColumn == tool.ToolName)
                            {
                                row++;

                                firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                if (firstColumn == "Subscales" && (secondColumn != "0"))
                                {
                                    //add subscales
                                    while (firstColumn != "Likert")
                                    {
                                        row++;
                                        firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                        secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                        if (firstColumn != "Likert")
                                        {
                                            Tool.Subscale mysubscale = new Tool.Subscale();
                                            mysubscale.Subscale_Name = firstColumn;
                                            mysubscale.Subscale_Full_Name = secondColumn;
                                            tool.AddSubscale(mysubscale);
                                        }



                                    }

                                }
                            }
                        }


                        row++; // Move to the next row
                    }
                }
            }
            catch(Exception)
            {
                MessageBox.Show("Error at tool subscale in details");
            }
            
        }
        public void GetToolsNumberandName(Worksheet ADetailsSheet)
        {
            try
            {
                int row = 0;
                int maxRow = ADetailsSheet.Cells.MaxDataRow;

                while (row <= maxRow) // Ensure we process the last row
                {
                    string firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                    string secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                    // Check if the row has any data (at least one non-null, non-empty value)
                    bool rowHasData = !(string.IsNullOrWhiteSpace(firstColumn) && string.IsNullOrWhiteSpace(secondColumn));

                    if ((firstColumn != null) && (secondColumn != null))
                    {
                        if (secondColumn.StartsWith("Tool"))
                        {
                            Tool newtool = new Tool();
                            newtool.ToolNumber = firstColumn;
                            newtool.ToolName = secondColumn;

                            AddToolIfNotExists(newtool);
                        }
                    }


                    row++; // Move to the next row
                }
            }
            catch (Exception) 
            {
                MessageBox.Show("Error at Tool Name And number in details");
            }

            

        }

        public void ReadData()
        {
            Worksheet Sheet1 = workbook.Worksheets[0];

            string pattern = @"Q\d+\.[A-Za-z0-9]+(\.[A-Za-z0-9]+)?\.\d+(\.[A-Za-z0-9]+)?";

            List<Tool> tools = new List<Tool>();

            for (int i = 0;i <= Sheet1.Cells.MaxDataColumn;i++)
            {
                string ExcelString = Sheet1.Cells[0,i]?.Value?.ToString();
                if (!string.IsNullOrEmpty(ExcelString) && Regex.IsMatch(ExcelString, pattern))
                {
                    var parts = ExcelString.Split('.');
                    int itemNumber = int.Parse(parts[0].Substring(1)); // Q1 -> item 1
                    string subscaleName = parts[1]; // A, test, etc.
                    int toolNumber = int.Parse(parts[2]); // 1 for Q1.A.1 -> Tool 1


                }
            }
            
        }
        private void Nursing_Load(object sender, EventArgs e)
        {
            ReadToolsAndSubscales();
            //push test
            //ReadData();

            //ReadToolsAndSubscales();
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
            ChooseFrm chooseFrm = new ChooseFrm();
            chooseFrm.Dgv = Dgv;
            chooseFrm.Show();
        }
    }
}
