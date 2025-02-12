using Accord.Statistics.Kernels;
using Aspose.Cells;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2016.Excel;
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
using static ExcelScore.Classes.Tool;
using static ExcelScore.Forms.Waiting;
using Tool = ExcelScore.Classes.Tool;
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

            GetEachToolLikert(DetailsSheet);

            GetEachToolLevel(DetailsSheet);
            

        }
        

        public void GetEachToolLevel(Worksheet ADetailsSheet)
        {
            
                foreach (Tool tool in AllTools)
                {
                   
                    int row = 0;
                    int maxRow = ADetailsSheet.Cells.MaxDataRow;

                    while (row <= maxRow) // Ensure we process the last row
                    {
                        string firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                        string secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();


                        if ((firstColumn != null) && (secondColumn != null))
                        {
                            if (secondColumn == tool.ToolName)
                            {
                                while (firstColumn != "Level")
                                {


                                    row++;

                                    firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                    secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();
                                }

                                bool toolExists = false;


                                while (!toolExists && row <= maxRow)
                                {
                                    
                                    if (firstColumn == "Level")
                                    {

                                        firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                        secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                        tool.LevelDetermination.Add(secondColumn);
                                        row++;

                                        //Tool.LevelRange ToolLevels = 

                                        List<Tool.LevelRange> levelRanges = new List<Tool.LevelRange>();

                                        while (firstColumn != null && secondColumn != null)
                                        {
                                            
                                            firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                            secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();


                                            Tool.LevelRange NewLevelRange = new Tool.LevelRange();

                                            NewLevelRange.Range = firstColumn;
                                            NewLevelRange.Label = secondColumn;

                                            levelRanges.Add(NewLevelRange);

                                            row++;
                                            firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                            secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();


                                        }

                                        tool.ToolLevels.Add(levelRanges);



                                    }
                                    row++;

                                    firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                    secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                    toolExists = AllTools.Any(t => t.ToolName == secondColumn);

                                    
                                }
                            }
                            
                        }


                        row++;

                    }
                }
            
            
        }

        public void GetEachToolLikert(Worksheet ADetailsSheet)
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


                        if ((firstColumn != null) && (secondColumn != null))
                        {
                            if (secondColumn == tool.ToolName)
                            {
                                while (firstColumn != "Likert")
                                {
                                    row++;

                                    firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                    secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();
                                }

                                if(firstColumn == "Likert")
                                {
                                    row++;

                                    firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                    secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                    while (firstColumn != null)
                                    {
                                        tool.LikertScale.Add(int.Parse(firstColumn), secondColumn);
                                        row++;
                                        firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                        secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                    }
                                }


                            }
                        }
                        row++;

                    }
                }
            }
            catch (Exception) 
            {
                MessageBox.Show("Error at likert at details");
            }
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
                                row = row +2;

                                firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                if (firstColumn == "Scales" && (secondColumn != "0"))
                                {
                                    //add subscales
                                    while (firstColumn != "Likert")
                                    {
                                        row++;
                                        firstColumn = ADetailsSheet.Cells[row, 0]?.Value?.ToString();
                                        secondColumn = ADetailsSheet.Cells[row, 1]?.Value?.ToString();

                                        if (firstColumn != "Likert" && (firstColumn != null || secondColumn != null))
                                        {
                                            if(!firstColumn.Contains("."))
                                            {
                                                Tool.Scale myScale = new Tool.Scale();
                                                myScale.Scale_Name = firstColumn;
                                                myScale.Scale_Full_Name = secondColumn;
                                                tool.AddScale(myScale);
                                            }
                                            else if(firstColumn.Contains("."))
                                            {
                                                string[] Scale_Subscale = new string[2];
                                                Scale_Subscale = firstColumn.Split('.');
                                                string ScaleName = Scale_Subscale[0];
                                                Tool.Scale CurrntScale = tool.Scales.FirstOrDefault(s => s.Scale_Name == ScaleName);

                                                if (CurrntScale != null) 
                                                {
                                                    Tool.Subscale mysubscale = new Tool.Subscale();
                                                    mysubscale.Subscale_Name = firstColumn;
                                                    mysubscale.Subscale_Full_Name = secondColumn;
                                                    CurrntScale.AddSubscale(mysubscale);
                                                }
                                            }
                                            
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
        public Scale GetScaleByToolNumberAndName(string toolNumber, string scaleName)
        {
            return AllTools
                .FirstOrDefault(t => t.ToolNumber == toolNumber)?
                .Scales
                .FirstOrDefault(s => s.Scale_Name == scaleName);
        }

        public void ReadDataOnlyScale(string[] QuestionParts , string FullItemText , Worksheet DataSheet , int itemcol)
        {
            try
            {
                Worksheet ReverseSheet = workbook.Worksheets["Reverse"];

                bool isReverse = false;

                for(int reverserow = 0;reverserow <= ReverseSheet.Cells.MaxDataRow; reverserow++)
                {
                    string cellvalue = ReverseSheet.Cells[reverserow, 0].Value.ToString();

                    if(cellvalue == FullItemText)
                    {
                        isReverse = true;
                        break;
                    }
                }

                

                List<double> itemresponses = new List<double>();

                for (int row = 1; row <= DataSheet.Cells.MaxDataRow;row++)
                {
                    double response = double.Parse(DataSheet.Cells[row,itemcol].Value.ToString());
                    itemresponses.Add(response);
                }

                Tool.Item MyItem = new Tool.Item();

                MyItem.Id = QuestionParts[0];
                MyItem.Text = FullItemText;
                MyItem.IsReverse = isReverse;
                MyItem.ParticipantResponses = itemresponses;

                string scalename = QuestionParts[1];
                string Toolnumber = QuestionParts[2];

                Scale scaletoadditem = GetScaleByToolNumberAndName(Toolnumber, scalename);

                scaletoadditem.AddItem(MyItem);

                

            }
            catch (Exception)
            {
                MessageBox.Show("Error at reading data only Scale");
            }
        }
        public string ReadData()
        {
            Worksheet Sheet1 = workbook.Worksheets[0];

           

            string periodOrDesc = "";

            string pattern = @"Q\d+\.[A-Za-z0-9]+(\.[A-Za-z0-9]+)?\.\d+(\.[A-Za-z0-9]+)?";

            

            for (int i = 0;i <= Sheet1.Cells.MaxDataColumn;i++)
            {
                string ExcelString = Sheet1.Cells[0,i]?.Value?.ToString();
                if (!string.IsNullOrEmpty(ExcelString) && Regex.IsMatch(ExcelString, pattern))
                {
                    string[] parts = ExcelString.Split('.');

                    int partscount = parts.Count();

                    if(partscount > 2)
                    {
                        if(partscount  == 3) 
                        {
                            //Descriptive
                            ReadDataOnlyScale(parts , ExcelString , Sheet1 ,  i);
                            periodOrDesc = "Descriptive";
                        }
                        else if (partscount == 4)
                        {
                            //might have periods

                            
                            if (int.TryParse(parts[3], out int result))
                            {
                                //Descriptive
                                periodOrDesc = "Descriptive";

                            }
                            else
                            {
                                //periods
                                periodOrDesc = "Periods";
                            }


                        }

                        else if (partscount == 5)
                        {
                            // Periods
                            periodOrDesc = "Periods";
                        }



                    }

                }
            }


            return periodOrDesc;


        }
        public void CalculateScore(string PeriodOrDes)
        {
            if(PeriodOrDes == "Descriptive")
            {
                CalculateScoreDescriptive();
            }
        }
        public void CalculateScoreDescriptive()
        {
            //NewNursingExcel.EnsureSheetExists(workbook, "Sheet2", 1);
            //NewNursingExcel.EnsureSheetExists(workbook, "Sheet3", 2);

            Worksheet sheet2 = workbook.Worksheets[1];

            sheet2.Cells[0, 0].Value = 2;

            sheet2.AutoFitColumns();
            workbook.Save(ExcelFunctions.filepath);

        }
        private void Nursing_Load(object sender, EventArgs e)
        {
            ReadToolsAndSubscales();
            
            string PeriodOrDes = ReadData();


            CalculateScore(PeriodOrDes);
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
