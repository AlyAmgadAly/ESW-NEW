using Accord.Statistics.Kernels;
using Aspose.Cells;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelScore.Classes.Tool;
using static ExcelScore.Forms.Waiting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Cell = Aspose.Cells.Cell;
using Color = System.Drawing.Color;
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

        public void UpdatePeriodsTools()
        {

            foreach (var CurrentTool in AllTools)
            {
                foreach (var PeriodTool in CurrentTool.PeriodsTools)
                {
                    CurrentTool.UpdatePeriodTool(CurrentTool.ToolNumber, PeriodTool.ToolName);
                }
            }
        }
        public void GetEachToolPeriod(Worksheet PeriodsSheet)
        {
            try
            {
                if (PeriodsSheet == null)
                {
                    return; // Exit function if the sheet does not exist
                }

                int row = 0;
                int maxRow = PeriodsSheet.Cells.MaxDataRow;

                while (row <= maxRow) // Ensure we process the last row
                {
                    string firstColumn = PeriodsSheet.Cells[row, 0]?.Value?.ToString();

                    foreach (var CurrentTool in AllTools)
                    {
                        if (firstColumn != null)
                        {
                            firstColumn = PeriodsSheet.Cells[row, 0]?.Value?.ToString();
                            if (firstColumn == CurrentTool.ToolName)
                            {

                                //MessageBox.Show(firstColumn);
                                row++;
                                List<Tool> PeriodsTools = new List<Tool>();
                                while(firstColumn != null)
                                {
                                    
                                    firstColumn = PeriodsSheet.Cells[row, 0]?.Value?.ToString();
                                    //MessageBox.Show(firstColumn);

                                    Tool newtool = new Tool();
                                    newtool.ToolNumber = CurrentTool.ToolNumber;
                                    newtool.ToolName = firstColumn;

                                    PeriodsTools.Add(newtool);

                                   

                                    row++;
                                    firstColumn = PeriodsSheet.Cells[row, 0]?.Value?.ToString();
                                   
                                }

                                CurrentTool.PeriodsTools = PeriodsTools;

                            }
                        }

                    }
                    row++;


                }

                UpdatePeriodsTools();
            }
            catch (Exception)
            {
                MessageBox.Show("Error at Tool periods sheet");
            }
        }
        public void ReadToolsAndSubscales()
        {
            Worksheet DetailsSheet = workbook.Worksheets["Details"];
            Worksheet PeriodsSheet = workbook.Worksheets["Periods"];

            GetToolsNumberandName(DetailsSheet);

            GetEachToolSubscales(DetailsSheet);

            GetEachToolLikert(DetailsSheet);

            GetEachToolLevel(DetailsSheet);

            GetEachToolPeriod(PeriodsSheet);
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
                                    tool.hasscales = true;
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
                                                    mysubscale.Subscale_Name = Scale_Subscale[1];
                                                    mysubscale.Subscale_Full_Name = secondColumn;
                                                    CurrntScale.AddSubscale(mysubscale);
                                                }
                                            }
                                            
                                        }



                                    }

                                }
                                else
                                {
                                    tool.hasscales = false;
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



        public Tool GetToolByNumber(string toolNumber)
        {
            return AllTools
                .FirstOrDefault(t => t.ToolNumber == toolNumber);

        }

        public Tool GetToolByName(string toolName)
        {
            return AllTools
                .FirstOrDefault(t => t.ToolName == toolName);

        }
        public void ReadDataOnlyPeriods(string[] QuestionParts, string FullItemText, Worksheet DataSheet, int itemcol)
        {
            Worksheet ReverseSheet = workbook.Worksheets["Reverse"];

            bool isReverse = false;

            for (int reverserow = 0; reverserow <= ReverseSheet.Cells.MaxDataRow; reverserow++)
            {
                string cellvalue = ReverseSheet.Cells[reverserow, 0].Value.ToString();

                if (cellvalue == FullItemText)
                {
                    isReverse = true;
                    break;
                }
            }

            List<double> itemresponses = new List<double>();

            for (int row = 1; row <= DataSheet.Cells.MaxDataRow; row++)
            {
                if (DataSheet.Cells[row, itemcol].Value == null)
                {
                    MessageBox.Show("Empty value at row " + row + " and column " + itemcol);
                }
                double response = double.Parse(DataSheet.Cells[row, itemcol].Value.ToString());
                itemresponses.Add(response);
            }

            Tool.Item MyItem = new Tool.Item();

            MyItem.Id = QuestionParts[0];
            MyItem.Text = FullItemText;
            MyItem.IsReverse = isReverse;
            MyItem.ParticipantResponses = itemresponses;

            if (QuestionParts.Count() == 3)
            {
                //Q1.1.Pre
                string Toolnumber = QuestionParts[1];

                Tool tooltoadditem = GetToolByNumber(Toolnumber);

                Tool PeriodTool = tooltoadditem.GetPeriodTool(Toolnumber, QuestionParts[2]);

                PeriodTool.ToolItems.Add(MyItem);
            }
            else if (QuestionParts.Count() == 4)
            {
                //Q1.A.1.Pre
                string scalename = QuestionParts[1];
                string Toolnumber = QuestionParts[2];
                string PeriodToolName = QuestionParts[3];

                Tool tooltoadditem = GetToolByNumber(Toolnumber);

                Tool PeriodTool = tooltoadditem.GetPeriodTool(Toolnumber, PeriodToolName);

                Scale scaletoadditem = PeriodTool.GetScaleFromPeriodsTool(Toolnumber, PeriodToolName, scalename);


                PeriodTool.ToolItems.Add(MyItem);
                scaletoadditem.Items.Add(MyItem);

            }


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
                    if(DataSheet.Cells[row, itemcol].Value == null)
                    {
                        MessageBox.Show("Empty value at row " + row + " and column " + itemcol);
                    }
                    double response = double.Parse(DataSheet.Cells[row, itemcol].Value.ToString());
                    itemresponses.Add(response);
                }

                Tool.Item MyItem = new Tool.Item();

                MyItem.Id = QuestionParts[0];
                MyItem.Text = FullItemText;
                MyItem.IsReverse = isReverse;
                MyItem.ParticipantResponses = itemresponses;


                if(QuestionParts.Count() == 2)
                {
                    string Toolnumber = QuestionParts[1];

                    Tool tooltoadditem = GetToolByNumber(Toolnumber);

                    tooltoadditem.ToolItems.Add(MyItem);
                }
                else if(QuestionParts.Count() == 3)
                {
                    string scalename = QuestionParts[1];
                    string Toolnumber = QuestionParts[2];

                    Scale scaletoadditem = GetScaleByToolNumberAndName(Toolnumber, scalename);


                    Tool tooltoadditem = GetToolByNumber(Toolnumber);
                    tooltoadditem.ToolItems.Add(MyItem);

                    scaletoadditem.AddItem(MyItem);
                }
                

                

            }
            catch (Exception)
            {
                MessageBox.Show("Error at reading data only Scale");
            }
        }
        

        public string ReadData()
        {
            Worksheet Sheet1 = workbook.Worksheets[0];

            NewNursingExcel.EnsureSheetExists(workbook, "Sheet2", 1);
            NewNursingExcel.EnsureSheetExists(workbook, "Sheet3", 2);

            string periodOrDesc = "";

            //string pattern = @"Q\d+\.[A-Za-z0-9]+(\.[A-Za-z0-9]+)?\.\d+(\.[A-Za-z0-9]+)?";

            string patternnew = @"Q\d+\.(?:[A-Za-z]+(\.[A-Za-z]+)*\.\d+(\.[A-Za-z]+)?|\d+)";


            for (int i = 0;i <= Sheet1.Cells.MaxDataColumn;i++)
            {
                string ExcelString = Sheet1.Cells[0,i]?.Value?.ToString();
                if (!string.IsNullOrEmpty(ExcelString) && Regex.IsMatch(ExcelString, patternnew))
                {
                    string[] parts = ExcelString.Split('.');

                    int partscount = parts.Count();

                    if(partscount > 1)
                    {
                        if((partscount  == 3 && int.TryParse(parts[2], out int result2)) || partscount == 2) 
                        {
                            //Descriptive
                            ReadDataOnlyScale(parts , ExcelString , Sheet1 ,  i);
                            periodOrDesc = "Descriptive";
                        }
                        else if (partscount == 4 || partscount == 3)
                        {
                            //might have periods

                            
                            if (partscount == 4 && int.TryParse(parts[3], out int result))
                            {
                                //Descriptive
                                periodOrDesc = "Descriptive";

                            }
                            else
                            {
                                //periods
                                ReadDataOnlyPeriods(parts, ExcelString, Sheet1, i);
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
                insertvlookTable();
                CalculateScore_Descriptive_NoSubscales();
                
            }

            
        }

        public void StoreScore(int TotalN)
        {
            workbook.CalculateFormula();

            foreach (var CurrentTool in AllTools)
            {
                int Current_col = 3;
                Worksheet CurrenScoreSheet = workbook.Worksheets[CurrentTool.ToolName];

                List<double> Total_Score_Tool = new List<double>();
                List<double> Avg_Score_Tool = new List<double>();
                List<double> Percent_Score_Tool = new List<double>();
                List<List<int>> ComputedLevels_Tool = new List<List<int>>();

                if (CurrentTool.hasscales)
                {
                    foreach (var CurrentScale in CurrentTool.Scales)
                    {
                        Current_col = Current_col + CurrentScale.Items.Count + 2;

                        List<double> Total_Score = new List<double>();
                        List<double> Avg_Score = new List<double>();
                        List<double> Percent_Score = new List<double>();
                        List<List<int>> ComputedLevels = new List<List<int>>();

                        for (int i = 0; i < CurrentTool.ToolLevels.Count; i++)
                        {
                            ComputedLevels.Add(new List<int>());
                        }
                        for (int row = 3;row < TotalN+3;row++)
                        {
                            double CellTotalDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col].Value.ToString());
                            double CellAvgDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col+1].Value.ToString());
                            double CellPercentDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col+2].Value.ToString());

                            Total_Score.Add(CellTotalDouble);
                            Avg_Score.Add(CellAvgDouble);
                            Percent_Score.Add(CellPercentDouble);


                            int levelcount = 0;
                            foreach (var CurrentLevel in CurrentTool.ToolLevels)
                            {
                                int mycol = Current_col + 2 + levelcount + 1;
                                int CellPercentint = int.Parse(CurrenScoreSheet.Cells[row, mycol].Value.ToString());
                                ComputedLevels[levelcount].Add(CellPercentint);
                                levelcount++;
                            }


                            
                        }

                        Current_col = Current_col + 3 + CurrentTool.ToolLevels.Count;

                        CurrentScale.TotalScores = Total_Score;
                        CurrentScale.AverageScores = Avg_Score;
                        CurrentScale.PercentScores = Percent_Score;
                        CurrentScale.ComputedSubLevels = ComputedLevels;

                        
                    }

                    Current_col++;
                    for (int i = 0; i < CurrentTool.ToolLevels.Count; i++)
                    {
                        ComputedLevels_Tool.Add(new List<int>());
                    }

                    for (int row = 3; row < TotalN + 3; row++)
                    {
                        double CellTotalDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col].Value.ToString());
                        double CellAvgDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col + 1].Value.ToString());
                        double CellPercentDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col + 2].Value.ToString());

                        Total_Score_Tool.Add(CellTotalDouble);
                        Avg_Score_Tool.Add(CellAvgDouble);
                        Percent_Score_Tool.Add(CellPercentDouble);


                        int levelcount = 0;
                        foreach (var CurrentLevel in CurrentTool.ToolLevels)
                        {
                            int mycol = Current_col + 2 + levelcount + 1;
                            int CellPercentint = int.Parse(CurrenScoreSheet.Cells[row, mycol].Value.ToString());
                            ComputedLevels_Tool[levelcount].Add(CellPercentint);
                            levelcount++;
                        }



                    }

                    CurrentTool.TotalScores = Total_Score_Tool;
                    CurrentTool.AverageScores = Avg_Score_Tool;
                    CurrentTool.PercentScores = Percent_Score_Tool;
                    CurrentTool.ComputedToolLevels = ComputedLevels_Tool;


                }

                else if(!(CurrentTool.hasscales))
                {
                    Current_col = Current_col + CurrentTool.ToolItems.Count + 2;


                    for (int i = 0; i < CurrentTool.ToolLevels.Count; i++)
                    {
                        ComputedLevels_Tool.Add(new List<int>());
                    }

                    for (int row = 3; row < TotalN + 3; row++)
                    {
                        double CellTotalDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col].Value.ToString());
                        double CellAvgDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col + 1].Value.ToString());
                        double CellPercentDouble = double.Parse(CurrenScoreSheet.Cells[row, Current_col + 2].Value.ToString());

                        Total_Score_Tool.Add(CellTotalDouble);
                        Avg_Score_Tool.Add(CellAvgDouble);
                        Percent_Score_Tool.Add(CellPercentDouble);


                        int levelcount = 0;
                        foreach (var CurrentLevel in CurrentTool.ToolLevels)
                        {
                            int mycol = Current_col + 2 + levelcount + 1;
                            int CellPercentint = int.Parse(CurrenScoreSheet.Cells[row, mycol].Value.ToString());
                            ComputedLevels_Tool[levelcount].Add(CellPercentint);
                            levelcount++;
                        }



                    }

                    CurrentTool.TotalScores = Total_Score_Tool;
                    CurrentTool.AverageScores = Avg_Score_Tool;
                    CurrentTool.PercentScores = Percent_Score_Tool;
                    CurrentTool.ComputedToolLevels = ComputedLevels_Tool;
                }
            }
        }

        public int getN()
        {
            int itemcount = AllTools[0].ToolItems[0].ParticipantResponses.Count;

            return itemcount; 
        }


        static Syncfusion.Drawing.Color  LightGreen = ColorClass.LightOliveGreen();
        System.Drawing.Color LightGreenExcel = System.Drawing.Color.FromArgb(LightGreen.ToArgb());

        public void SetToolNames( Worksheet ScoreSheet, Tool CurrentTool , int TotalN , int current_colctr)
        {
            for (int colorrow = 0; colorrow < TotalN + 3; colorrow++)
            {
                NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, colorrow, current_colctr, "", LightGreenExcel);
            }
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 2, current_colctr, CurrentTool.ToolName, LightGreenExcel);
        }
        public void InsertheaderScale(Worksheet ScoreSheet , Scale MyScale , int current_colctr)
        {
            int mergecountcol = MyScale.Items.Count;
            ScoreSheet.Cells.Merge(1, current_colctr, 1, mergecountcol);
            ScoreSheet.Cells.Merge(0, current_colctr, 1, mergecountcol);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 1, current_colctr, MyScale.Scale_Name, LightGreenExcel);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 0, current_colctr, "", LightGreenExcel);
        }
        public void PutItemData(int current_colctr, Scale Myscale , Worksheet ScoreSheet ,Tool CurrentTool)
        {
            int itemcol = current_colctr;
            int minLikert = CurrentTool.LikertScale.Keys.Min();
            int maxLikert = CurrentTool.LikertScale.Keys.Max();

            foreach (var Myscaleitem in Myscale.Items)
            {
                


                bool reverse = Myscaleitem.IsReverse;
                NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 2, itemcol, Myscaleitem.Id, LightGreenExcel);

                int item_partic_data_row = 3;
                foreach (var participantdata in Myscaleitem.ParticipantResponses)
                {
                    double InsertData = participantdata;
                    if (reverse)
                    {
                        InsertData = (minLikert + maxLikert) - InsertData;
                    }
                    NewNursingExcel.SetCellValueAndCenter_int(ScoreSheet, item_partic_data_row, itemcol, InsertData);
                    item_partic_data_row++;
                }


                itemcol++;
            }
        }

        public void PutItemData_Overall(int current_colctr, Worksheet ScoreSheet, Tool CurrentTool)
        {
            int itemcol = current_colctr;
            int minLikert = CurrentTool.LikertScale.Keys.Min();
            int maxLikert = CurrentTool.LikertScale.Keys.Max();

            foreach (var MyToolitem in CurrentTool.ToolItems)
            {
                NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 0, itemcol, "", LightGreenExcel);
                NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 1, itemcol, "", LightGreenExcel);

                bool reverse = MyToolitem.IsReverse;
                NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 2, itemcol, MyToolitem.Id, LightGreenExcel);

                int item_partic_data_row = 3;
                foreach (var participantdata in MyToolitem.ParticipantResponses)
                {
                    double InsertData = participantdata;
                    if (reverse)
                    {
                        InsertData = (minLikert + maxLikert) - InsertData;
                    }
                    NewNursingExcel.SetCellValueAndCenter_int(ScoreSheet, item_partic_data_row, itemcol, InsertData);
                    item_partic_data_row++;
                }


                itemcol++;
            }
        }

        public static void ProcessScoreColumns(Worksheet ScoreSheet, int startRow, int totalRows, ref int current_colctr, Scale Scale, Color fillColor, Tool CurrentTool)
        {
            // === TOTAL COLUMN ===
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 2, current_colctr, "Total", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 0, current_colctr, "", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 1, current_colctr, "", fillColor);

            for (int row = startRow; row < totalRows + startRow; row++)
            {
                string destinationCell = ScoreSheet.Cells[row, current_colctr].Name;
                string range = ScoreSheet.Cells[row, current_colctr - (Scale.Items.Count + 1)].Name + ":" + ScoreSheet.Cells[row, current_colctr - 2].Name;
                ScoreSheet.Cells[destinationCell].Formula = $"=SUM({range})";

                Style returnedStyle = NewNursingExcel.ScoreStyle(ScoreSheet, fillColor, row, current_colctr);
                ScoreSheet.Cells[row, current_colctr].SetStyle(returnedStyle);
            }
            current_colctr++;

            // === AVERAGE COLUMN ===
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 2, current_colctr, "Avg", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 0, current_colctr, "", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 1, current_colctr, "", fillColor);

            for (int row = startRow; row < totalRows + startRow; row++)
            {
                string destinationCell = ScoreSheet.Cells[row, current_colctr].Name;
                string range = ScoreSheet.Cells[row, current_colctr - (Scale.Items.Count + 2)].Name + ":" + ScoreSheet.Cells[row, current_colctr - 3].Name;
                ScoreSheet.Cells[destinationCell].Formula = $"=AVERAGE({range})";

                Style returnedStyle = NewNursingExcel.ScoreStyle(ScoreSheet, fillColor, row, current_colctr);
                ScoreSheet.Cells[row, current_colctr].SetStyle(returnedStyle);
            }
            current_colctr++;

            // === PERCENT COLUMN ===
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 2, current_colctr, "Percent", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 0, current_colctr, "", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 1, current_colctr, "", fillColor);

            int minLikert = CurrentTool.LikertScale.Keys.Min();
            int maxLikert = CurrentTool.LikertScale.Keys.Max();

            for (int row = startRow; row < totalRows + startRow; row++)
            {
                string destinationCell = ScoreSheet.Cells[row, current_colctr].Name;
                string avgCell = ScoreSheet.Cells[row, current_colctr - 1].Name;

                ScoreSheet.Cells[destinationCell].Formula = $"=({avgCell} - {minLikert}) / ({maxLikert} - {minLikert}) * 100";

                Style returnedStyle = NewNursingExcel.ScoreStyle(ScoreSheet, fillColor, row, current_colctr);
                ScoreSheet.Cells[row, current_colctr].SetStyle(returnedStyle);
            }
            current_colctr++;
        }
        public void insertvlookTable()
        {
            Worksheet vloovkupSheet = workbook.Worksheets["Sheet3"];

            

            int Tablesvlookrow = 0;

            foreach (var ToolVlook in AllTools)
            {
                vloovkupSheet.Cells.Merge(Tablesvlookrow, 0, 1, 2);
                NewNursingExcel.SetCellValueAndCenterText(vloovkupSheet, Tablesvlookrow, 0, ToolVlook.ToolName, LightGreenExcel);

                Tablesvlookrow++;

                int levelID = 1;
                foreach (var Level in ToolVlook.ToolLevels)
                {
                    NewNursingExcel.SetCellValueAndCenterText(vloovkupSheet, Tablesvlookrow, 0, "Level."+ levelID + "."+ ToolVlook.ToolNumber, LightGreenExcel);
                    levelID++;
                    Tablesvlookrow++;

                    int levelvalue = 1;
                    foreach (var Levelrange in Level)
                    {
                        NewNursingExcel.SetCellValueAndCenterText(vloovkupSheet, Tablesvlookrow, 0, Levelrange.Range, LightGreenExcel);
                        NewNursingExcel.SetCellValueAndCenterText(vloovkupSheet, Tablesvlookrow, 1, levelvalue.ToString(), LightGreenExcel);
                        levelvalue++;
                        Tablesvlookrow++;

                    }


                }

            }
        }

        public (int, int) getstartEndColumn(Tool CurrentTool , string LevelID)
        {
            Worksheet vloovkupSheet = workbook.Worksheets["Sheet3"];


            int startrow = 0;
            int Endrow = 0;

          
            int vlookuprow = 0;

            while(vlookuprow <= vloovkupSheet.Cells.MaxDataRow)
            {
                if (vloovkupSheet.Cells[vlookuprow, 0].Value != null)
                {
                    string firstcol = vloovkupSheet.Cells[vlookuprow, 0].Value.ToString();
                    if (firstcol == LevelID)
                    {
                        vlookuprow++;
                        startrow = vlookuprow;
                        while (vloovkupSheet.Cells[vlookuprow, 1].Value != null)
                        {
                            vlookuprow++;
                        }

                        Endrow = vlookuprow-1;
                    }
                }

                vlookuprow++;



            }

           







            return (startrow, Endrow);
        }

        public List<string> getDataRangesforOverall(Worksheet CurrentSheet , Tool CurrentTool , int TotalN , int current_colctr)
        {
            List<string> Ranges = new List<string>();
            int StartCol = 3;

            foreach (var CurrentScale in CurrentTool.Scales)
            {
                string range = CurrentSheet.Cells[3, StartCol + 1].Name + ":" + CurrentSheet.Cells[3, StartCol + CurrentScale.Items.Count].Name;

                Ranges.Add(range);
                StartCol = StartCol + CurrentScale.Items.Count + 1 +3 + CurrentTool.ToolLevels.Count+1;

            }

            return Ranges;

            


        }

        public List<string> getDataRangesforOverall_Noscale(Worksheet CurrentSheet, Tool CurrentTool, int TotalN, int current_colctr)
        {
            List<string> Ranges = new List<string>();
            int StartCol = 3;

            string range = CurrentSheet.Cells[3, StartCol + 1].Name + ":" + CurrentSheet.Cells[3, StartCol + CurrentTool.ToolItems.Count].Name;

            Ranges.Add(range);
         
            return Ranges;




        }
        static List<string> GetIncrementedCombinedRanges(string combinedRange, int targetRow)
        {
            List<string> result = new List<string>();
            string[] ranges = combinedRange.Split(','); // Split into individual ranges
            List<(string col1, int row1, string col2, int row2)> parsedRanges = new List<(string, int, string, int)>();

            // Parse each range
            foreach (string range in ranges)
            {
                Match match = Regex.Match(range.Trim(), @"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");

                if (match.Success)
                {
                    string col1 = match.Groups[1].Value;  // Left column (e.g., "E")
                    int row1 = int.Parse(match.Groups[2].Value);  // Start row (e.g., 4)
                    string col2 = match.Groups[3].Value;  // Right column (e.g., "J")
                    int row2 = int.Parse(match.Groups[4].Value);  // End row (e.g., 4)

                    parsedRanges.Add((col1, row1, col2, row2));
                }
            }

            // Generate incremented ranges until targetRow is reached
            while (parsedRanges.Count > 0)
            {
                List<string> newRanges = new List<string>();
                List<(string, int, string, int)> updatedRanges = new List<(string, int, string, int)>();

                foreach (var (col1, row1, col2, row2) in parsedRanges)
                {
                    if (row1 <= targetRow && row2 <= targetRow)
                    {
                        // Increment row numbers
                        newRanges.Add($"{col1}{row1}:{col2}{row2}");
                        updatedRanges.Add((col1, row1 + 1, col2, row2 + 1)); // Prepare for next iteration
                    }
                }

                if (newRanges.Count > 0)
                {
                    result.Add(string.Join(",", newRanges)); // Combine them properly
                }

                // Update the list for the next iteration
                parsedRanges = updatedRanges;
            }

            return result;
        }
        static string CombineRanges(List<string> ranges)
        {
            return string.Join(",", ranges);
        }
        public void InsertDataOverall(Worksheet CurrentSheet , ref int current_colctr ,Color fillColor , List<string> Ranges , int totalRows , Tool CurrentTool)
        {
            int minLikert = CurrentTool.LikertScale.Keys.Min();
            int maxLikert = CurrentTool.LikertScale.Keys.Max();

            NewNursingExcel.SetCellValueAndCenterText(CurrentSheet, 2, current_colctr, "Total", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(CurrentSheet, 0, current_colctr, "", fillColor);
            

            string Allranges = CombineRanges(Ranges);
            List<string> Myranges = GetIncrementedCombinedRanges(Allranges, totalRows + 3);

            int row = 3;
            Style returnedStyle = NewNursingExcel.ScoreStyle(CurrentSheet, fillColor, row, current_colctr);
            foreach (var CurrentRange in Myranges)
            {
                
                string destinationCellTotal = CurrentSheet.Cells[row, current_colctr].Name;
                CurrentSheet.Cells[destinationCellTotal].Formula = $"=SUM({CurrentRange})";
                CurrentSheet.Cells[row, current_colctr].SetStyle(returnedStyle);

                string destinationCellAvg = CurrentSheet.Cells[row, current_colctr+1].Name;
                CurrentSheet.Cells[destinationCellAvg].Formula = $"=AVERAGE({CurrentRange})";
                CurrentSheet.Cells[row, current_colctr+1].SetStyle(returnedStyle);

                string destinationCellPerc = CurrentSheet.Cells[row, current_colctr + 2].Name;
                string avgCell = CurrentSheet.Cells[row, current_colctr +1].Name;
                CurrentSheet.Cells[destinationCellPerc].Formula = $"=({avgCell} - {minLikert}) / ({maxLikert} - {minLikert}) * 100";
                CurrentSheet.Cells[row, current_colctr+2].SetStyle(returnedStyle);

                row++;

            }

            current_colctr++;
            NewNursingExcel.SetCellValueAndCenterText(CurrentSheet, 2, current_colctr, "Avg", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(CurrentSheet, 0, current_colctr, "", fillColor);
            


            current_colctr++;
            NewNursingExcel.SetCellValueAndCenterText(CurrentSheet, 2, current_colctr, "Percent", fillColor);
            NewNursingExcel.SetCellValueAndCenterText(CurrentSheet, 0, current_colctr, "", fillColor);
            



        }

        public void InsertLevel(Worksheet CurrentScoreSheet , ref int current_colctr , Tool CurrentTool , int TotalN)
        {
            int levelId = 1;
            int leveldeterminectr = 0;
            int VlookUpColumn = 0;

            //datacell of percent or total or average is by level ID

            foreach (var Level in CurrentTool.ToolLevels)
            {
                NewNursingExcel.SetCellValueAndCenterText(CurrentScoreSheet, 0, current_colctr, "", LightGreenExcel);
                NewNursingExcel.SetCellValueAndCenterText(CurrentScoreSheet, 1, current_colctr, "", LightGreenExcel);
                string currentLevelID = "Level." + levelId + "." + CurrentTool.ToolNumber;
                NewNursingExcel.SetCellValueAndCenterText(CurrentScoreSheet, 2, current_colctr, currentLevelID, LightGreenExcel);

                int startrow = 0;
                int endrow = 0;



                (startrow, endrow) = getstartEndColumn(CurrentTool, currentLevelID);
                startrow++;
                endrow++;
                string TableArrayRange = "$A$" + startrow + ":" + "$B$" + endrow;


                if (CurrentTool.LevelDetermination[leveldeterminectr] == "Total")
                {
                    VlookUpColumn = current_colctr - (3 + leveldeterminectr);

                }
                else if (CurrentTool.LevelDetermination[leveldeterminectr] == "Avg")
                {
                    VlookUpColumn = current_colctr - (2 + leveldeterminectr);
                }
                else if (CurrentTool.LevelDetermination[leveldeterminectr] == "Percent")
                {
                    VlookUpColumn = current_colctr - (1 + leveldeterminectr);
                }

                for (int row = 3; row < TotalN + 3; row++)
                {
                    string ScoreCell = CurrentScoreSheet.Cells[row, VlookUpColumn].Name;
                    string destinationCell = CurrentScoreSheet.Cells[row, current_colctr].Name;
                    string formula = "=VLOOKUP(" + ScoreCell + "," + "Sheet3!" + TableArrayRange + ",2" + ")";

                    CurrentScoreSheet.Cells[destinationCell].Formula = formula;
                    Style returnedStyle = NewNursingExcel.ScoreStyle(CurrentScoreSheet, LightGreenExcel, row, current_colctr);
                    CurrentScoreSheet.Cells[row, current_colctr].SetStyle(returnedStyle);
                }

                leveldeterminectr++;
                current_colctr++;
                levelId++;
            }
        }


        public void CalculateScore_Descriptive_NoSubscales()
        {

            int TotalN = getN();


            //Worksheet ScoreSheet = workbook.Worksheets["Sheet2"];
            Worksheet VlookupSheet = workbook.Worksheets["Sheet3"];


            int vlookCtrrow = 0;
            foreach (var CurrentTool in AllTools)
            {


                int current_colctr = 2;

                NewNursingExcel.EnsureSheetExists(workbook, CurrentTool.ToolName, workbook.Worksheets.Count);
                Worksheet CurrentScoreSheet = workbook.Worksheets[CurrentTool.ToolName];

                //insert tool names and highlight them
                SetToolNames(CurrentScoreSheet, CurrentTool, TotalN, current_colctr);

                //leave blank col
                current_colctr += 2;

                //insert items

                //header Merge

                if (CurrentTool.hasscales)
                {

                    foreach (var Scale in CurrentTool.Scales)
                    {
                        InsertheaderScale(CurrentScoreSheet, Scale, current_colctr);

                        PutItemData(current_colctr, Scale, CurrentScoreSheet , CurrentTool);

                        current_colctr = current_colctr + Scale.Items.Count + 1;


                        ProcessScoreColumns(CurrentScoreSheet, 3, TotalN, ref current_colctr, Scale, LightGreenExcel, CurrentTool);


                        InsertLevel(CurrentScoreSheet, ref current_colctr, CurrentTool, TotalN);
                        current_colctr ++ ;

                    }

                    //start inserting overall

                    CurrentScoreSheet.Cells.Merge(1, current_colctr, 1, 3+ CurrentTool.ToolLevels.Count);
                    NewNursingExcel.SetCellValueAndCenterText(CurrentScoreSheet, 1, current_colctr, "Overall", LightGreenExcel);

                    List<string> Ranges = getDataRangesforOverall(CurrentScoreSheet , CurrentTool , TotalN , current_colctr);

                    InsertDataOverall(CurrentScoreSheet, ref current_colctr, LightGreenExcel , Ranges , TotalN , CurrentTool);

                    current_colctr++;

                    InsertLevel(CurrentScoreSheet, ref current_colctr, CurrentTool, TotalN);




                }
                else if(!(CurrentTool.hasscales))
                {
                    

                    PutItemData_Overall(current_colctr, CurrentScoreSheet, CurrentTool);

                    current_colctr = current_colctr + CurrentTool.ToolItems.Count + 1;


                    CurrentScoreSheet.Cells.Merge(1, current_colctr, 1, 3 + CurrentTool.ToolLevels.Count);

                    NewNursingExcel.SetCellValueAndCenterText(CurrentScoreSheet, 1, current_colctr, "Overall", LightGreenExcel);

                    List<string> Ranges = getDataRangesforOverall_Noscale(CurrentScoreSheet, CurrentTool, TotalN, current_colctr);

                    InsertDataOverall(CurrentScoreSheet, ref current_colctr, LightGreenExcel, Ranges, TotalN, CurrentTool);

                   current_colctr++;

                   InsertLevel(CurrentScoreSheet, ref current_colctr, CurrentTool, TotalN);
                }

                CurrentScoreSheet.AutoFitColumns();
                CurrentScoreSheet.AutoFitRows();

            }

            StoreScore(TotalN);



            try
            {
                workbook.Save(ExcelFunctions.filepath);
            }
            catch (Exception)
            {
                MessageBox.Show("Close Excel");
            }
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

        private void button1_Click(object sender, EventArgs e)
        {
            NursingWordNew nursingWordNew = new NursingWordNew();
            nursingWordNew.Alltools_Tables = AllTools;
            nursingWordNew.Show();
            this.Hide();
        }
    }
}
