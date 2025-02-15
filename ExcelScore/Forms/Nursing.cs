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
                        if(partscount  == 3 || partscount == 2) 
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
                insertvlookTable();
                CalculateScore_Descriptive_NoSubscales();
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
        public void PutItemData(int current_colctr, Scale Myscale , Worksheet ScoreSheet )
        {
            int itemcol = current_colctr;

            foreach (var Myscaleitem in Myscale.Items)
            {
                NewNursingExcel.SetCellValueAndCenterText(ScoreSheet, 2, itemcol, Myscaleitem.Id, LightGreenExcel);

                int item_partic_data_row = 3;
                foreach (var participantdata in Myscaleitem.ParticipantResponses)
                {
                    NewNursingExcel.SetCellValueAndCenter_int(ScoreSheet, item_partic_data_row, itemcol, participantdata);
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

                        PutItemData(current_colctr, Scale, CurrentScoreSheet);

                        current_colctr = current_colctr + Scale.Items.Count + 1;



                        ProcessScoreColumns(CurrentScoreSheet, 3, TotalN, ref current_colctr, Scale, LightGreenExcel, CurrentTool);


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
                            string TableArrayRange = "$A$"+startrow + ":" + "$B$"+ endrow;


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

                        // 

                        current_colctr ++ ;

                    }

                    List<string> Ranges = getDataRangesforOverall(CurrentScoreSheet , CurrentTool , TotalN , current_colctr);

                    foreach (var item in Ranges)
                    {
                        MessageBox.Show(item);
                    }

                   //ProcessScoreColumns(CurrentScoreSheet, 3, TotalN, ref current_colctr, Scale, LightGreenExcel, CurrentTool);

                }








                CurrentScoreSheet.AutoFitColumns();


            }



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




       
    }
}
