using BitMiracle.LibTiff.Classic;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
using Python.Runtime;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static alglib;
using static ExcelScore.Classes.Tool;

namespace ExcelScore.Forms
{
    public partial class NursingWordNew : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        public NursingWordNew()
        {
            InitializeComponent();
        }
        WordDocument document;
        WordClass wordObj = new WordClass();

        public List<Tool> Alltools_Tables = new List<Tool>();

        private void NursingWordNew_Load(object sender, EventArgs e)
        {

        }
        public IWTable DesignItemsTable(int ItemsCount , int LikertCount ,  string TableHeader , Tool CurrentTool , int tableorder ,int ToolN , ref bool hastoolTittle)
        {
            IWSection section = wordObj.CreatePortraitSection();

            
            if (!hastoolTittle)
            {
                wordObj.AddToolTitle_Dynamic(section,CurrentTool.ToolName);
                hastoolTittle = true; 
            }

            string title = "Distribution of the studied Nursing students according to " + CurrentTool.ToolName + " items (n = "+ ToolN + ")";

            wordObj.AddTitle_Dynamic(section , title , ref tableorder);

            int WordTableRows = 2 + ItemsCount;

            int WordTableColumns = 2 + (CurrentTool.LikertScale.Keys.Count * 2);


            IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

            wordObj.GeneralTableFormat(table);


            wordObj.ApplyMerges(table, LikertCount);

            wordObj.ApplyBorders(table, WordTableRows, WordTableColumns);

            wordObj.Widths_Descriptive_Items(table, WordTableRows, LikertCount);

            wordObj.Add_Header_Items_Nusring(table, LikertCount, TableHeader, CurrentTool);


            wordObj.AddNo_perc_Center(table, WordTableColumns);

            wordObj.AddQuestionNo(table, ItemsCount);


            return table;
        }
        public void Items(Tool CurrentTool , int ToolN)
        {
            int likertScaleCount = CurrentTool.LikertScale.Keys.Count;


            bool hastitle = false;
            if (CurrentTool.hasscales)
            {
                
                foreach (var CurrentScale in CurrentTool.Scales)
                {
                    int tableorder = 1;

                    IWTable table = DesignItemsTable(CurrentScale.Items.Count , likertScaleCount , CurrentScale.Scale_Full_Name , CurrentTool , tableorder, ToolN ,ref  hastitle);



                    int Insertrow = 2;

                    foreach (var CurrentItem in CurrentScale.Items)
                    {
                        int InsertColumn = 2;

                        foreach (var CurrentLikertScore in CurrentTool.LikertScale.Keys)
                        {
                            int totalLikertFreq = 0;
                            foreach (var ParticpantResponse in CurrentItem.ParticipantResponses)
                            {
                                if(ParticpantResponse == CurrentLikertScore)
                                {
                                    totalLikertFreq++;
                                }
                            }
                            double totalLikertPerc = ((double)totalLikertFreq / ToolN)*100;


                            wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertFreq.ToString());
                            InsertColumn++;
                            wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertPerc.ToString("0.0"));
                            InsertColumn++;
                        }

                        Insertrow++;
                    }



                    if(likertScaleCount > 3)
                    {
                        wordObj.FormatTableCustom(table, 11, 4, 2);
                        wordObj.LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    }
                    else
                    {
                        wordObj.FormatTableCustom(table, 12, 4, 2);
                    }



                }
            }
            else if(!(CurrentTool.hasscales))
            {
                int tableorder = 1;
                IWTable table = DesignItemsTable(CurrentTool.ToolItems.Count, likertScaleCount, CurrentTool.ToolName, CurrentTool , tableorder , ToolN , ref hastitle);

                int Insertrow = 2;

                foreach (var CurrentItem in CurrentTool.ToolItems)
                {
                    int InsertColumn = 2;

                    foreach (var CurrentLikertScore in CurrentTool.LikertScale.Keys)
                    {
                        int totalLikertFreq = 0;
                        foreach (var ParticpantResponse in CurrentItem.ParticipantResponses)
                        {
                            if (ParticpantResponse == CurrentLikertScore)
                            {
                                totalLikertFreq++;
                            }
                        }
                        double totalLikertPerc = ((double)totalLikertFreq / ToolN) * 100;


                        wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertFreq.ToString());
                        InsertColumn++;
                        wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertPerc.ToString("0.0"));
                        InsertColumn++;
                    }

                    Insertrow++;
                }

            }
        }
        public void InsertScale_Names(Tool CurrentTool , IWTable table)
        {
            int currentrow = 2;
            foreach (var CurrentScale in CurrentTool.Scales)
            {
                wordObj.AddPara_NoCenter(table, currentrow, 0, CurrentScale.Scale_Full_Name);
                currentrow++;
            }
            wordObj.AddPara_Center(table, currentrow, 0, "Overall");

        }

        public void InsertScoreRange (IWTable table, Tool CurrentTool)
        {
            int minlikert = CurrentTool.LikertScale.Keys.Min();
            int maxlikert = CurrentTool.LikertScale.Keys.Max();

            int currentrow = 2;
            foreach (var CurrentScale in CurrentTool.Scales)
            {
                int minScore = CurrentScale.Items.Count * minlikert;
                int maxScore = CurrentScale.Items.Count * maxlikert;

                string CurrentScoreRange = "("+ minScore + " – "+ maxScore + ")";
                wordObj.AddPara_Center(table, currentrow, 1, CurrentScoreRange);
                currentrow++;

            }

            int minScoreTool = CurrentTool.ToolItems.Count * minlikert;
            int maxScoreTool = CurrentTool.ToolItems.Count * maxlikert;

            string CurrentScoreRangeTool = "(" + minScoreTool + " – " + maxScoreTool + ")";

            wordObj.AddPara_Center(table, currentrow, 1, CurrentScoreRangeTool);



        }

        public void InsertLevel_Scale_Overall_Data(IWTable table , int LevelID , Tool CurrenTool , int ToolN , List<LevelRange> CurrentLevel)
        {
            int row = 3;
            int Count = CurrentLevel.Count;

            foreach (var CurrentScale in CurrenTool.Scales)
            {
                // Get the computed counts and percentages
                var computedCounts = CurrentScale.ComputedSubLevels[LevelID]
                    .GroupBy(n => n)
                    .ToDictionary(g => g.Key, g => (g.Count(), (g.Count() / (double)ToolN) * 100));

                // Ensure all levels (1 to Count) are included
                Dictionary<int, (int Count, double Percentage)> sortedCounts = Enumerable.Range(1, Count)
                    .ToDictionary(level => level, level =>
                        computedCounts.ContainsKey(level)
                            ? computedCounts[level]
                            : (0, 0.0) // Default to 0 if the level is missing
                    );

                // Sort by level (ascending)
                sortedCounts = sortedCounts.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);


                int Currentcol = 1;

                foreach (var CurrentCount in sortedCounts)
                {
                    double value = CurrentCount.Key;          
                    int count = CurrentCount.Value.Count;      
                    double percentage = CurrentCount.Value.Percentage;

                    wordObj.Addpara_CenterNoBOLD(table, row, Currentcol, count.ToString());
                    wordObj.Addpara_CenterNoBOLD(table, row, Currentcol+1, percentage.ToString("0.0"));

                    Currentcol = Currentcol + 2;
                }
                row++;

            }

            var computedCounts_Tool = CurrenTool.ComputedToolLevels[LevelID]
                    .GroupBy(n => n)
                    .ToDictionary(g => g.Key, g => (g.Count(), (g.Count() / (double)ToolN) * 100));

            // Ensure all levels (1 to Count) are included
            Dictionary<int, (int Count, double Percentage)> sortedCounts_Tool = Enumerable.Range(1, Count)
                .ToDictionary(level => level, level =>
                    computedCounts_Tool.ContainsKey(level)
                        ? computedCounts_Tool[level]
                        : (0, 0.0) // Default to 0 if the level is missing
                );

            // Sort by level (ascending)
            sortedCounts_Tool = sortedCounts_Tool.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            int Currentcol_Tool = 1;

            foreach (var CurrentCount in sortedCounts_Tool)
            {
                double value = CurrentCount.Key;
                int count = CurrentCount.Value.Count;
                double percentage = CurrentCount.Value.Percentage;

                wordObj.AddPara_Center(table, row, Currentcol_Tool, count.ToString());
                wordObj.AddPara_Center(table, row, Currentcol_Tool + 1, percentage.ToString("0.0"));

                Currentcol_Tool = Currentcol_Tool + 2;
            }
           
        }
        public void Level_Scales_Overall(Tool CurrentTool, int ToolN)
        {
            if(CurrentTool.hasscales)
            {
                int LevelID = 0;

                int tableorder = 1; 
                foreach (var CurrentLevel in CurrentTool.ToolLevels)
                {
                    IWSection section = wordObj.CreatePortraitSection();

                    

                    string title = "Distribution of the studied Nursing students according to levels of " + CurrentTool.ToolName + " (n = " + ToolN + ")";

                    wordObj.AddTitle_Dynamic(section, title, ref tableorder);

                    int WordTableRows = 3 + CurrentTool.Scales.Count + 1;

                    int WordTableColumns = get_Level_Scale_Columns(CurrentLevel);


                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.GeneralTableFormat(table);


                    wordObj.Apply_Level_Scale_overall_merge(table, WordTableColumns);

                    wordObj.ApplyLevel_Scale_Borders(table, WordTableRows, WordTableColumns);

                    wordObj.Level_Overall_Widths(table, WordTableRows, WordTableColumns);


                    wordObj.InsertHeaders_Scale_Overall(table, CurrentTool, CurrentLevel, WordTableColumns);

                    
                    InsertLevel_Scale_Overall_Data(table, LevelID, CurrentTool, ToolN , CurrentLevel);
                    LevelID++;

                    wordObj.FormatTableCustom(table, 12, 4, 2);

                }
                
            }

            
        }

        private int get_Level_Scale_Columns(List<LevelRange> Currentlevel)
        {
            int col = 0;

            col = 1 + (Currentlevel.Count * 2);


            return col;


        }

        public void Score_Scales_Overall(Tool CurrentTool, int ToolN)
        {
            if(CurrentTool.hasscales)
            {
                bool TotalScore_MinMax = true;
                bool TotalScore_MeanSD = true;
                bool TotalScore_Median = false;
                bool Avg_MeanSD = true;
                bool Percent_MeanSD = true;
                bool Rank = true;

                int minLikert = CurrentTool.LikertScale.Keys.Min();
                int MaxLikert = CurrentTool.LikertScale.Keys.Max();

                bool[] flags =
                {
            TotalScore_MinMax, TotalScore_MeanSD, TotalScore_Median,Avg_MeanSD, Percent_MeanSD, Rank
            };
                int tableorder = 1;
                IWSection section = wordObj.CreatePortraitSection();

                string title = "Distribution of the studied Nursing students according to " + CurrentTool.ToolName + " score (n = " + ToolN + ")";

                wordObj.AddTitle_Dynamic(section, title, ref tableorder);

                int WordTableRows = 2 + CurrentTool.Scales.Count + 1;

                int WordTableColumns = get_Score_Scale_Columns(flags);

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                wordObj.GeneralTableFormat(table);

                wordObj.Apply_Score_Scale_overall_merge(table, TotalScore_MinMax, TotalScore_MeanSD, TotalScore_Median, Rank, WordTableColumns);
                wordObj.ApplyTotalScoreNoPBorders(table, WordTableRows, WordTableColumns);

                wordObj.InsertGeneral_Score_Scale_overall_header_new(table, TotalScore_MinMax, TotalScore_MeanSD, TotalScore_Median, Avg_MeanSD, Percent_MeanSD, Rank, WordTableColumns, minLikert, MaxLikert);
                wordObj.Score_Overall_Widths(table, WordTableRows, WordTableColumns, Rank);



                InsertScale_Names(CurrentTool,table ) ;


                InsertScoreRange(table, CurrentTool);

                InsertGeneral_Score_Scale_overall_Scores_new2(table, TotalScore_MinMax, TotalScore_MeanSD, TotalScore_Median, Avg_MeanSD, Percent_MeanSD, Rank, CurrentTool);

                wordObj.FormatTableCustom(table, 12, 4, 2);
            }

            

        }
     

        CustomMathClass customMathClass = new CustomMathClass();    
       
       
        public void InsertGeneral_Score_Scale_overall_Scores_new2(
    IWTable table,
    bool TS_Minmax, bool TS_MeanSD, bool TS_Median,
    bool Avg_meanSD, bool Percent_meanSD, bool Rank,
    Tool CurrentTool)
        {
            if (Rank)
            {
                CurrentTool.RankScales();
            }

            int currentRow = 2;

            // Insert scores for each scale
            foreach (var CurrentScale in CurrentTool.Scales)
            {
                InsertScoreData(table, CurrentScale, currentRow, TS_Minmax, TS_MeanSD, TS_Median, Avg_meanSD, Percent_meanSD, Rank, isTool: false);
                currentRow++;
            }

            // Insert scores for the entire tool (without rank)
            InsertScoreData(table, CurrentTool, currentRow, TS_Minmax, TS_MeanSD, TS_Median, Avg_meanSD, Percent_meanSD, Rank, isTool: true);
        }

        private void InsertScoreData(
            IWTable table,
            dynamic DataSource, // Accepts both CurrentScale and CurrentTool
            int row,
            bool TS_Minmax, bool TS_MeanSD, bool TS_Median,
            bool Avg_meanSD, bool Percent_meanSD, bool Rank,
            bool isTool) // New flag to indicate if it's a tool
        {

            List<double> TotalSorted = new List<double>(DataSource.TotalScores);
            List<double> AverageSorted = new List<double>(DataSource.AverageScores);
            List<double> PercentSorted = new List<double>(DataSource.PercentScores);



            Dictionary<string, string> TotalScore = customMathClass.Basic_Calculations(TotalSorted);
            Dictionary<string, string> AvgScore = customMathClass.Basic_Calculations(AverageSorted);
            Dictionary<string, string> PercentScore = customMathClass.Basic_Calculations(PercentSorted);

            int col = 2; // Start at column 2
            
            // Determine which wordObj function to use
            Action<IWTable, int, int, string> AddParaFunc = isTool
                ? wordObj.AddPara_Center
                : wordObj.Addpara_CenterNoBOLD;

            // Insert Total Score Data (if enabled)
            if (TS_Minmax || TS_MeanSD || TS_Median)
            {
                if (TS_Minmax)
                    AddParaFunc(table, row, col++, TotalScore["Min-Max"]);
                if (TS_MeanSD)
                    AddParaFunc(table, row, col++, TotalScore["Mean ± StdDev"]);
                if (TS_Median)
                    AddParaFunc(table, row, col++, TotalScore["Median"]);
            }

            // Insert Average Score (if enabled)
            if (Avg_meanSD)
                AddParaFunc(table, row, col++, AvgScore["Mean ± StdDev"]);

            // Insert Percent Score (if enabled)
            if (Percent_meanSD)
                AddParaFunc(table, row, col++, PercentScore["Mean ± StdDev"]);

            // Insert Rank (if enabled) - Only for Scales, NOT for the Tool
            if (Rank && !isTool)
                wordObj.AddPara_Center(table, row, col++, DataSource.Rank.ToString());

            if (Rank && isTool)
            {
                
                wordObj.AddPara_Center(table, row, col, "");
                table.Rows[row].Cells[col].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;
            }
                

            
        }


        public int get_Score_Scale_Columns(bool[] flags)
        {
            int col = 2;

            int ColtrueCtr = flags.Count(b => b);


            col = col + ColtrueCtr;


            return col;



        }
        public void Correlation_Nursing_Desgin()
        {
            IWSection section = wordObj.CreatePortraitSection();

            int ComparisonRows = ((Alltools_Tables.Count) * (Alltools_Tables.Count - 1)) / 2;

            int WordTableRows =  1 + ComparisonRows;

            int WordTableColumns = 3;

            IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);




        }
        
        public int  ToolN { get; set; }

        public string TableType { get; set; }

        public int GetItemCount_Periods(Scale myscale, Tool CurrentToolPeriod)
        {
            int itemcount = 0;

            foreach (var item in myscale.Items)
            {
                if (item.Text.Contains(CurrentToolPeriod.ToolName))
                {
                    itemcount++;
                }
            }


            return itemcount;
        }
        public void Items_Periods_InsertData_Portrait(Tool CurrentTool , int likertScaleCount , ref int insertcolumn , Tool.Scale CurrentScale)
        {
            int ToolPeriodsCount = CurrentTool.PeriodsTools.Count;

            foreach (var CurrentToolPeriod in CurrentTool.PeriodsTools)
            {
                


            }
            

        }
        public void Items_Periods_portrait(Tool CurrentTool , int ToolN)
        {
            int likertScaleCount = CurrentTool.LikertScale.Keys.Count;

            bool hastitle = false;

            int TestExist = 0;

            int WordTableColumns = 2 + (CurrentTool.PeriodsTools.Count * likertScaleCount * 2) + 2;

            int periodsCount = CurrentTool.PeriodsTools.Count;

            Tool CurrentToolPeriod = CurrentTool.PeriodsTools[0];

            if (CurrentToolPeriod.hasscales)
            {
                foreach (var CurrentScale in CurrentToolPeriod.Scales)
                {
                    int itemcount = GetItemCount_Periods(CurrentScale, CurrentToolPeriod);
                    int WordTableRows = 3 + itemcount;

                    IWSection section = wordObj.CreatePortraitSection();

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.GeneralTableFormat(table);

                    wordObj.Merges_Periods_Items(table, WordTableColumns, likertScaleCount, TestExist);


                    wordObj.Borders_Periods_Items(table, WordTableRows, WordTableColumns);

                    wordObj.Widths_Periods_Items(table, WordTableRows, WordTableColumns, likertScaleCount, TestExist, periodsCount);


                    wordObj.Header_Periods_Items(table, likertScaleCount, CurrentScale, CurrentTool , WordTableColumns ,  itemcount);

                    int insertcolumn = 2;
                    Items_Periods_InsertData_Portrait(CurrentTool, likertScaleCount, ref insertcolumn , CurrentScale);

                    //foreach (var CurrentItem in CurrentScale.Items)
                    //{
                    //    int InsertColumn = 2;

                    //    foreach (var CurrentLikertScore in CurrentTool.LikertScale.Keys)
                    //    {
                    //        int totalLikertFreq = 0;
                    //        foreach (var ParticpantResponse in CurrentItem.ParticipantResponses)
                    //        {
                    //            if (ParticpantResponse == CurrentLikertScore)
                    //            {
                    //                totalLikertFreq++;
                    //            }
                    //        }
                    //        double totalLikertPerc = ((double)totalLikertFreq / ToolN) * 100;


                    //        wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertFreq.ToString());
                    //        InsertColumn++;
                    //        wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertPerc.ToString("0.0"));
                    //        InsertColumn++;
                    //    }

                    //    Insertrow++;
                    //}





                    List<int> columnsToRemove = new List<int> { WordTableColumns - 1, WordTableColumns - 2 };
                    if (!(TestExist > 0))
                    {
                        wordObj.DeleteColumn(columnsToRemove, table);
                    }


                    wordObj.FormatTable_Periods_Items(table, periodsCount, likertScaleCount, TestExist);



                }


            }



        }
        private void button1_Click(object sender, EventArgs e)
        {
            document = wordObj.InitWord();

            if (TableType == "Descriptive")
            {
                foreach (var CurrentTool in Alltools_Tables)
                {
                    //int toolN = CurrentTool.ToolItems[0].ParticipantResponses.Count;


                    Items(CurrentTool, ToolN);
                    Score_Scales_Overall(CurrentTool, ToolN);
                    Level_Scales_Overall(CurrentTool, ToolN);


                }
            }
            else if(TableType == "Periods")
            {
                foreach (var CurrentTool in Alltools_Tables)
                {
                    if(CurrentTool.PeriodsTools.Count == 2)
                    {
                        //any likert not above 3
                        if(CurrentTool.LikertScale.Count <= 3)
                        {
                            Items_Periods_portrait(CurrentTool, ToolN);
                        }
                        //else gonna be the vertical table which starts from 4 likert
                    }
                    else if (CurrentTool.PeriodsTools.Count > 2)
                    {
                        if (CurrentTool.LikertScale.Count < 3)
                        {
                            Items_Periods_portrait(CurrentTool, ToolN);
                        }
                        else if (CurrentTool.LikertScale.Count >= 3)
                        {

                        }

                       
                    }



                }


            }
            



            string filepath = wordObj.SaveWord();

            wordObj.removeHeader(filepath);




        }

        public void ItemsPeriods()
        {

        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        public void InsertSpssScore_Descr()
        {
#nullable enable
            Dictionary<string, (List<double?> Values, string MeasurementLevel, Dictionary<double, string>? ValueLabels)> data =
                new Dictionary<string, (List<double?>, string, Dictionary<double, string>?)>();

            int toolN = Alltools_Tables[0].ToolItems[0].ParticipantResponses.Count;





            // Scale Data (No Value Labels)
            var SerialData = SPSS_Class.ConvertToSPSSFormat(
                SPSS_Class.CreateSerial(toolN).Select(x => (double?)x).ToList(),  // Convert to List<double?>
                "Scale"
            );

            data.Add("Serial", SerialData);


            var GXData = SPSS_Class.ConvertToSPSSFormat(
                SPSS_Class.CreateGX(toolN).Select(x => (double?)x).ToList(),  // Convert to List<double?>
                "Scale"
            );

            data.Add("GX", GXData);


            int emptycount = 1;
            string empty = "VAR0000" + emptycount;
            List<double?> emptyVariable = Enumerable.Repeat<double?>(null, toolN).ToList();

            data.Add(empty, (emptyVariable, "Scale", null));

            emptycount++;




            foreach (var CurrentTool in Alltools_Tables)
            {

                if (CurrentTool.hasscales)
                {

                    foreach (var CurrentScale in CurrentTool.Scales)
                    {
                        int levelId = 0;

                        string total = "Total." + CurrentScale.Scale_Name + "." + CurrentTool.ToolNumber;
                        var TotalScaleData = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentScale.TotalScores), "Scale");
                        data.Add(total, TotalScaleData);


                        string Avg = "Avg." + CurrentScale.Scale_Name + "." + CurrentTool.ToolNumber;
                        var AvgScaleData = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentScale.AverageScores), "Scale");
                        data.Add(Avg, AvgScaleData);

                        string Percent = "Perc." + CurrentScale.Scale_Name + "." + CurrentTool.ToolNumber;
                        var PercScaleData = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentScale.PercentScores), "Scale");
                        data.Add(Percent, PercScaleData);





                        foreach (var level in CurrentTool.ToolLevels)
                        {
                            string Level = "Level." + CurrentScale.Scale_Name + "." + (levelId + 1) + "." + CurrentTool.ToolNumber;

                            int currentlevelcount = level.Count;
                            List<double> Levelnum = SPSS_Class.CreateSerial(currentlevelcount);
                            List<string> levellabel = new List<string>();

                            foreach (var Currentlevel in level)
                            {
                                levellabel.Add(Currentlevel.Label);
                            }

                            Dictionary<double, string> LevelValueLabel = Levelnum.Zip(levellabel, (key, value) => new { key, value })
                                            .ToDictionary(x => x.key, x => x.value);

                            var LevelNominalData = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentScale.ComputedSubLevels[levelId].Select(x => (double)x).ToList()), "Nominal", LevelValueLabel);
                            data.Add(Level, LevelNominalData);

                            levelId++;

                        }

                        empty = "VAR0000" + emptycount;
                        data.Add(empty, (emptyVariable, "Scale", null));

                        emptycount++;

                    }

                    int levelId_tool = 0;

                    string total_tool = "Overall.Total." + CurrentTool.ToolNumber;
                    var TotalData_tool = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentTool.TotalScores), "Scale");
                    data.Add(total_tool, TotalData_tool);


                    string Avg_tool = "Overall.Avg." + CurrentTool.ToolNumber;
                    var AvgData_Tool = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentTool.AverageScores), "Scale");
                    data.Add(Avg_tool, AvgData_Tool);

                    string Percent_tool = "Overall.Perc." + CurrentTool.ToolNumber;
                    var PercData_tool = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentTool.PercentScores), "Scale");
                    data.Add(Percent_tool, PercData_tool);


                    foreach (var level in CurrentTool.ToolLevels)
                    {

                        string Level = "Overall.Level." + (levelId_tool + 1) + "." + CurrentTool.ToolNumber;

                        int currentlevelcount = level.Count;
                        List<double> Levelnum = SPSS_Class.CreateSerial(currentlevelcount);
                        List<string> levellabel = new List<string>();

                        foreach (var Currentlevel in level)
                        {
                            levellabel.Add(Currentlevel.Label);
                        }

                        Dictionary<double, string> LevelValueLabel = Levelnum.Zip(levellabel, (key, value) => new { key, value })
                        .ToDictionary(x => x.key, x => x.value);

                        var LevelNominalData = SPSS_Class.ConvertToSPSSFormat(SPSS_Class.ConvertToNullable(CurrentTool.ComputedToolLevels[levelId_tool].Select(x => (double)x).ToList()), "Nominal", LevelValueLabel);
                        data.Add(Level, LevelNominalData);

                        levelId_tool++;

                    }

                    empty = "VAR0000" + emptycount;
                    data.Add(empty, (emptyVariable, "Scale", null));

                    emptycount++;


                }
            }


            SPSS_Class.InsertMultipleVariablesnewa(data);

        }
        private void button2_Click(object sender, EventArgs e)
        {

            InsertSpssScore_Descr();
        }
        pythonStat pythonStata = new pythonStat();   
        private void button3_Click(object sender, EventArgs e)
        {
            //SPSS_Class.RunSpssSyntax();

            string pythonHome = @"C:\Program Files (x86)\IBM\SPSS\Statistics\23\Python\python27.dll";
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", pythonHome);
            PythonEngine.Initialize();

            using (Py.GIL()) // Acquire the Global Interpreter Lock
            {
                try
                {
                    // Import the SPSS module
                    dynamic spss = Py.Import("spss");

                    // Load the SPSS data file
                    string filePath = @"C:\Users\win\Desktop\aaaa.sav";
                    spss.Submit($"GET FILE='{filePath}'.");

                    // Run DISPLAY DICTIONARY to show variable details
                   

                    // Run DESCRIPTIVES to get basic statistics
                    spss.Submit("DESCRIPTIVES VARIABLES=ALL.");

                    string outputFile = @"C:\Users\win\Desktop\spss_output.spo";
                    spss.Submit($"OUTPUT SAVE OUTFILE='{outputFile}'.");

                    Console.WriteLine("SPSS commands executed successfully!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }

            // Shutdown Python.NET
            PythonEngine.Shutdown();
        }
    }
}
