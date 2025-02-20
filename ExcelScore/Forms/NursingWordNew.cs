using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class NursingWordNew : Form
    {
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
        public IWTable DesignItemsTable(int ItemsCount , int LikertCount ,  string TableHeader , Tool CurrentTool)
        {
            IWSection section = wordObj.CreatePortraitSection();

            int WordTableRows = 2 + ItemsCount;

            int WordTableColumns = 2 + (CurrentTool.LikertScale.Keys.Count * 2);


            IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

            wordObj.GeneralTableFormat(table);


            wordObj.ApplyMerges(table, LikertCount);

            wordObj.ApplyBorders(table, WordTableRows, WordTableColumns);

            wordObj.SetWidths(table, WordTableRows, LikertCount);

            wordObj.Add_Header_Items_Nusring(table, LikertCount, TableHeader, CurrentTool);


            wordObj.AddNo_perc_Center(table, WordTableColumns);

            wordObj.AddQuestionNo(table, ItemsCount);


            return table;
        }
        public void Items(Tool CurrentTool , int ToolN)
        {
            int likertScaleCount = CurrentTool.LikertScale.Keys.Count;

            

            if (CurrentTool.hasscales)
            {

                foreach (var CurrentScale in CurrentTool.Scales)
                {

                    IWTable table = DesignItemsTable(CurrentScale.Items.Count , likertScaleCount , CurrentScale.Scale_Full_Name , CurrentTool);

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




                    wordObj.FormatTableCustom(table, 12, 4, 2);


                }
            }
            else if(!(CurrentTool.hasscales))
            {

                IWTable table = DesignItemsTable(CurrentTool.ToolItems.Count, likertScaleCount, CurrentTool.ToolName, CurrentTool);

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
        public void Score_Scales_Overall(Tool CurrentTool)
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

                IWSection section = wordObj.CreatePortraitSection();

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
        public void GetRank_Scales()
        {
            
        }

        CustomMathClass customMathClass = new CustomMathClass();    
        public void InsertGeneral_Score_Scale_overall_Scores_new(
    IWTable table,
    bool TS_Minmax, bool TS_MeanSD, bool TS_Median,
    bool Avg_meanSD, bool Percent_meanSD, bool Rank
     , Tool CurrentTool)
        {

            int currentRow = 2;
            

            foreach (var CurrentScale in CurrentTool.Scales)
            {

                Dictionary<string, string> TotalScore = customMathClass.Basic_Calculations(CurrentScale.TotalScores);
                Dictionary<string, string> Avg_score = customMathClass.Basic_Calculations(CurrentScale.AverageScores);
                Dictionary<string, string> Percent_score = customMathClass.Basic_Calculations(CurrentScale.PercentScores);

                int col = 2;

                
                if (TS_Minmax || TS_MeanSD || TS_Median)
                {

                    if (TS_Minmax)
                    {
                        wordObj.Addpara_CenterNoBOLD(table, currentRow, col, TotalScore["Min-Max"]);
                        col++;
                    }
                    if (TS_MeanSD)
                    {
                        wordObj.Addpara_CenterNoBOLD(table, currentRow, col, TotalScore["Mean ± StdDev"]);
                        col++;
                    }
                    if (TS_Median)
                    {
                        wordObj.Addpara_CenterNoBOLD(table, currentRow, col, TotalScore["Median"]);
                        col++;
                    }
                }

                // Avg Score Header (if enabled)
                if (Avg_meanSD)
                {
                    wordObj.Addpara_CenterNoBOLD(table, currentRow, col, Avg_score["Mean ± StdDev"]);

                    col++;
                }

                // Percent Score Header (if enabled)
                if (Percent_meanSD)
                {
                    wordObj.Addpara_CenterNoBOLD(table, currentRow, col, Percent_score["Mean ± StdDev"]);
                    col++;
                }

                // Rank Header (if enabled)
                if (Rank)
                {
                    wordObj.Addpara_CenterNoBOLD(table, 0, col, "Rank");
                    col++;
                }

                currentRow++;

            }

            Dictionary<string, string> TotalScore_Tool = customMathClass.Basic_Calculations(CurrentTool.TotalScores);
            Dictionary<string, string> Avg_score_Tool = customMathClass.Basic_Calculations(CurrentTool.AverageScores);
            Dictionary<string, string> Percent_score_Tool = customMathClass.Basic_Calculations(CurrentTool.PercentScores);

            int colTool = 2;


            if (TS_Minmax || TS_MeanSD || TS_Median)
            {

                if (TS_Minmax)
                {
                    wordObj.Addpara_CenterNoBOLD(table, currentRow, colTool, TotalScore_Tool["Min-Max"]);
                    colTool++;
                }
                if (TS_MeanSD)
                {
                    wordObj.Addpara_CenterNoBOLD(table, currentRow, colTool, TotalScore_Tool["Mean ± StdDev"]);
                    colTool++;
                }
                if (TS_Median)
                {
                    wordObj.Addpara_CenterNoBOLD(table, currentRow, colTool, TotalScore_Tool["Median"]);
                    colTool++;
                }
            }

            // Avg Score Header (if enabled)
            if (Avg_meanSD)
            {
                wordObj.Addpara_CenterNoBOLD(table, currentRow, colTool, Avg_score_Tool["Mean ± StdDev"]);

                colTool++;
            }

            // Percent Score Header (if enabled)
            if (Percent_meanSD)
            {
                wordObj.Addpara_CenterNoBOLD(table, currentRow, colTool, Percent_score_Tool["Mean ± StdDev"]);
                colTool++;
            }

            // Rank Header (if enabled)
            if (Rank)
            {
                wordObj.Addpara_CenterNoBOLD(table, 0, colTool, "Rank");
                colTool++;
            }

            





        }
        public void InsertGeneral_Score_Scale_overall_Scores_new2(
    IWTable table,
    bool TS_Minmax, bool TS_MeanSD, bool TS_Median,
    bool Avg_meanSD, bool Percent_meanSD, bool Rank,
    Tool CurrentTool)
        {
            int currentRow = 2;

            // Insert scores for each scale
            foreach (var CurrentScale in CurrentTool.Scales)
            {
                InsertScoreData(table, CurrentScale, currentRow, TS_Minmax, TS_MeanSD, TS_Median, Avg_meanSD, Percent_meanSD, Rank);
                currentRow++;
            }

            // Insert scores for the entire tool
            InsertScoreData(table, CurrentTool, currentRow, TS_Minmax, TS_MeanSD, TS_Median, Avg_meanSD, Percent_meanSD, Rank);
        }

        /// <summary>
        /// Inserts score data into the table for either a single scale or the entire tool.
        /// </summary>
        private void InsertScoreData(
            IWTable table,
            dynamic DataSource, // Accepts both CurrentScale and CurrentTool
            int row,
            bool TS_Minmax, bool TS_MeanSD, bool TS_Median,
            bool Avg_meanSD, bool Percent_meanSD, bool Rank)
        {
            Dictionary<string, string> TotalScore = customMathClass.Basic_Calculations(DataSource.TotalScores);
            Dictionary<string, string> AvgScore = customMathClass.Basic_Calculations(DataSource.AverageScores);
            Dictionary<string, string> PercentScore = customMathClass.Basic_Calculations(DataSource.PercentScores);

            int col = 2; // Start at column 2

            // Insert Total Score Data (if enabled)
            if (TS_Minmax || TS_MeanSD || TS_Median)
            {
                if (TS_Minmax)
                    wordObj.Addpara_CenterNoBOLD(table, row, col++, TotalScore["Min-Max"]);
                if (TS_MeanSD)
                    wordObj.Addpara_CenterNoBOLD(table, row, col++, TotalScore["Mean ± StdDev"]);
                if (TS_Median)
                    wordObj.Addpara_CenterNoBOLD(table, row, col++, TotalScore["Median"]);
            }

            // Insert Average Score (if enabled)
            if (Avg_meanSD)
                wordObj.Addpara_CenterNoBOLD(table, row, col++, AvgScore["Mean ± StdDev"]);

            // Insert Percent Score (if enabled)
            if (Percent_meanSD)
                wordObj.Addpara_CenterNoBOLD(table, row, col++, PercentScore["Mean ± StdDev"]);

            // Insert Rank (if enabled)
            if (Rank)
                wordObj.Addpara_CenterNoBOLD(table, row, col++, "Rank"); // Assuming Rank is available
        }

        public int get_Score_Scale_Columns(bool[] flags)
        {
            int col = 2;

            int ColtrueCtr = flags.Count(b => b);


            col = col + ColtrueCtr;


            return col;



        }

        private void button1_Click(object sender, EventArgs e)
        {
            document = wordObj.InitWord();

            foreach (var CurrentTool in Alltools_Tables)
            {
                int toolN = CurrentTool.ToolItems[0].ParticipantResponses.Count;
                Items(CurrentTool , toolN);
                Score_Scales_Overall(CurrentTool);
            }

            string filepath = wordObj.SaveWord();

            wordObj.removeHeader(filepath);




        }
    }
}
