using Aspose.Cells;
using Aspose.Cells.Drawing;
using ExcelScore.Classes;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using BorderStyle = Syncfusion.DocIO.DLS.BorderStyle;
using Color = Syncfusion.Drawing.Color;

namespace ExcelScore.Forms
{
    public partial class Word : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        pythonStat pyobj = new pythonStat();

        public List<NewDomainClass> AllDomainsWord = new List<NewDomainClass>();
        public Workbook ExcelWorksheet_ToWord;
        public List<float> WordData = new List<float>();
        ExcelFunctions excelobj = new ExcelFunctions(); 
        
        public int startrow = 0;

        public Word()
        {
            InitializeComponent();
        }
        
        

       



        public void Table_totalScore()
        {
            Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
            int rows = Sheet2.Cells.MaxDataRow;
            int cols = Sheet2.Cells.MaxDataColumn;

            float datacount = rows - 1;
            int TotalScoretrackctr = 0;


            foreach (var item in list_Periods.Items)
            {
                string[] domains = item.ToString().Split(' '); // Assuming space was used as separator
                foreach (string Mydomain in domains)
                {
                   
                    string value = Sheet2.Cells[0, TotalScoretrackctr].Value.ToString();
                    

                    if(value == Mydomain)
                    {
                        int index = AllDomainsWord.FindIndex(domain => domain.DomainName == Mydomain);
                        TotalScoretrackctr = TotalScoretrackctr + AllDomainsWord[index].QuestionsSorted_ReverseValue.Count;
                        MessageBox.Show("Found");
                        
                    }
                    else if(value != Mydomain)
                    {
                        int index = AllDomainsWord.FindIndex(domain => domain.DomainName == value);
                        TotalScoretrackctr = TotalScoretrackctr + AllDomainsWord[index].QuestionsSorted_ReverseValue.Count + 4;
                        MessageBox.Show("Not Found");
                    }

                    MessageBox.Show(TotalScoretrackctr.ToString());
                }
            }
        }



        WordClass wordObj = new WordClass();

        public void decide()
        {
            foreach (var item in list_Periods.Items)
            {
                string[] domains = item.ToString().Split(' ');
                int numberofperiods = domains.Length;
                if(numberofperiods == 2)
                {
                    TwoPeriods(domains);
                }
                else if(numberofperiods == 3) 
                {
                    ThreePeriods(domains);
                }
            }
        }

        public void TwoPeriods(string[] Adomains)
        {
            Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
            int rows = Sheet2.Cells.MaxDataRow;
            int cols = Sheet2.Cells.MaxDataColumn;
            float Exceldatacount = rows - 1;



            IWSection section = wordObj.CreatePortraitSection();


            
            string[] domains = Adomains;
            int numberofperiods = domains.Length;

                int ColumnStart = 2;
                wordObj.AddTitle(section, "", 0);

                int index = AllDomainsWord.FindIndex(domain => domain.DomainName == domains[0]);
                string DomainName = AllDomainsWord[index].DomainName;
                int RangeFrom = AllDomainsWord[index].RangeFrom;
                int RangeTo = AllDomainsWord[index].RangeTo;
                int LikertScore = RangeTo - RangeFrom + 1;
                int dataCount = AllDomainsWord[index].QuestionsSorted_ReverseValue.Count;

                
                int WordTableRows = 3 + dataCount;
                int WordTableColumns = (numberofperiods * LikertScore * 2) + 2;
                

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                wordObj.GeneralTableFormat(table);

                wordObj.ApplyTwoPeriodsMerges(table, LikertScore, numberofperiods);

                wordObj.ApplyTwoPeriodsBorders(table ,WordTableRows , WordTableColumns);

                wordObj.SettwoPeriodsWidths(table, WordTableRows, LikertScore, numberofperiods);

                wordObj.Add_Headers_TwoPeriods_Center(table, DomainName, WordTableColumns , LikertScore);

                wordObj.AddNo_perc_Center_periods(table, WordTableColumns);

                wordObj.AddQuestionNo_Periods(table, dataCount);
                

                foreach (string Mydomain in domains)
                {
                    int counter = 0;
                    startrow = 3;
                    int Currentindex = AllDomainsWord.FindIndex(domain => domain.DomainName == Mydomain);
                    string CurrentDomainName = AllDomainsWord[index].DomainName;
                    int CurrentRangeFrom = AllDomainsWord[index].RangeFrom;
                    int CurrentRangeTo = AllDomainsWord[index].RangeTo;
                    int CurrentLikertScore = RangeTo - RangeFrom + 1;
                    int CurrentdataCount = AllDomainsWord[index].QuestionsSorted_ReverseValue.Count;


                if (Currentindex == 0)
                {
                    counter = 0;
                }
                else if(Currentindex > 0)
                {
                    for(int i = 0; i < Currentindex;i++)
                    {
                        counter = counter + AllDomainsWord[i].QuestionsSorted_ReverseValue.Count;
                        counter = counter + 5;
                    }
                    
                }
                    ReadIndexesExcel_periods(Sheet2, table, rows, CurrentdataCount, CurrentRangeFrom, CurrentRangeTo, Exceldatacount, WordTableColumns,ColumnStart , counter);
                    
                    
                    ColumnStart = ColumnStart + LikertScore*2;
                    //trackCtr = trackCtr + dataCount;
                    //trackCtr = trackCtr + 5;
                }

                wordObj.LeftAndRightCellMargin(table);
                wordObj.Font(table, 9.5f);

               

            


        }
        public void Descriptive_Statement()
        {
            foreach(var item in list_descriptive.Items)
            {
                Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
                int rows = Sheet2.Cells.MaxDataRow;
                int cols = Sheet2.Cells.MaxDataColumn;
                float Exceldatacount = rows - 1;

                
                IWSection section = wordObj.CreatePortraitSection();
                int Currentindex = AllDomainsWord.FindIndex(domain => domain.DomainName == item.ToString());

                startrow = 2;
                string CurrentDomainName = AllDomainsWord[Currentindex].DomainName;
                int CurrentRangeFrom = AllDomainsWord[Currentindex].RangeFrom;
                int CurrentRangeTo = AllDomainsWord[Currentindex].RangeTo;
                int LikertScore = CurrentRangeTo - CurrentRangeFrom + 1;
                int dataCount = AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;

                wordObj.AddTitle(section, CurrentDomainName, dataCount);

                int WordTableRows = 2 + dataCount;
                int WordTableColumns = (2 * LikertScore) + 2;

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                wordObj.GeneralTableFormat(table);



                wordObj.ApplyMerges(table, LikertScore);

                wordObj.ApplyBorders(table, WordTableRows, WordTableColumns);

                wordObj.SetWidths(table, WordTableRows, LikertScore);
                wordObj.Add_Headers_Center(table, CurrentDomainName, LikertScore);



                wordObj.AddNo_perc_Center(table, WordTableColumns);

                wordObj.AddQuestionNo(table, dataCount);
                int counter = 0;
                if (Currentindex == 0)
                {
                    counter = 0;
                }
                else if (Currentindex > 0)
                {
                    for (int i = 0; i < Currentindex; i++)
                    {
                        counter = counter + AllDomainsWord[i].QuestionsSorted_ReverseValue.Count;
                        counter = counter + 5;
                    }

                }
                

                ReadIndexesExcel(Sheet2, table, rows, dataCount, CurrentRangeFrom, CurrentRangeTo, Exceldatacount, WordTableColumns , counter);

                


                wordObj.Font(table, 12);
                

                
            }
            

        }





        public void ThreePeriods(string[] Adomains)
        {
            Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
            int rows = Sheet2.Cells.MaxDataRow;
            int cols = Sheet2.Cells.MaxDataColumn;
            float Exceldatacount = rows - 1;



            IWSection section = wordObj.CreateLandscapeSection();



            string[] domains = Adomains;
            int numberofperiods = domains.Length;

            int ColumnStart = 2;
            wordObj.AddTitle(section, "", 0);

            int index = AllDomainsWord.FindIndex(domain => domain.DomainName == domains[0]);
            string DomainName = AllDomainsWord[index].DomainName;
            int RangeFrom = AllDomainsWord[index].RangeFrom;
            int RangeTo = AllDomainsWord[index].RangeTo;
            int LikertScore = RangeTo - RangeFrom + 1;
            int dataCount = AllDomainsWord[index].QuestionsSorted_ReverseValue.Count;


            int WordTableRows = 3 + dataCount;
            int WordTableColumns = (numberofperiods * LikertScore * 2) + 2;


            IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

            wordObj.GeneralTableFormat(table);

            wordObj.ApplyThreePeriodsMerges(table, LikertScore, numberofperiods);

            wordObj.ApplyTwoPeriodsBorders(table, WordTableRows, WordTableColumns);

            wordObj.SetThreePeriodsWidths(table, WordTableRows, LikertScore, numberofperiods);

                wordObj.Add_Headers_ThreePeriods_Center(table, DomainName, WordTableColumns, LikertScore);

                wordObj.AddNo_perc_Center_periods(table, WordTableColumns);

                wordObj.AddQuestionNo_Periods(table, dataCount);


                foreach (string Mydomain in domains)
                {
                int counter = 0;
                    startrow = 3;
                    int Currentindex = AllDomainsWord.FindIndex(domain => domain.DomainName == Mydomain);
                    string CurrentDomainName = AllDomainsWord[index].DomainName;
                    int CurrentRangeFrom = AllDomainsWord[index].RangeFrom;
                    int CurrentRangeTo = AllDomainsWord[index].RangeTo;
                    int CurrentLikertScore = RangeTo - RangeFrom + 1;
                    int CurrentdataCount = AllDomainsWord[index].QuestionsSorted_ReverseValue.Count;
                    if (Currentindex == 0)
                    {
                    counter = 0;
                    }
                    else if (Currentindex > 0)
                    {
                    for (int i = 0; i < Currentindex; i++)
                    {
                        counter = counter + AllDomainsWord[i].QuestionsSorted_ReverseValue.Count;
                        counter = counter + 5;
                    }

                }
                ReadIndexesExcel_periods(Sheet2, table, rows, CurrentdataCount, CurrentRangeFrom, CurrentRangeTo, Exceldatacount, WordTableColumns, ColumnStart, counter);

                ColumnStart = ColumnStart + LikertScore * 2;
                    
                }

                wordObj.LeftAndRightCellMargin(table);
                wordObj.Font(table, 9.5f);
            


        }

        private void Word_Load(object sender, EventArgs e)
        {
            pyobj.InitPython();
            for(int i = 0; i< AllDomainsWord.Count;i++)
            {
                AllDomainsWord[i].LikertScore = "";
                list_allDomains.Items.Add(AllDomainsWord[i].DomainName);
            }

        }


        
        public void ReadIndexesExcel(Worksheet sheet,IWTable table , int rows , int QuestionNum  , int RangeFrom , int RangeTo , float datacount , int WordTableColumns , int counter)
        {
            //indexes.Clear();    
            int arraySize = RangeTo - RangeFrom + 2;
            // Initialize the array
            float[] indexes = new float[arraySize];

            for (int ca = counter; ca < QuestionNum + counter; ca++)
            {
                
                for (int ra = 2; ra < rows + 1; ra++)
                {
                    
                    string value = sheet.Cells[ra, ca].Value.ToString();
                    int valueint = Convert.ToInt32(value);

                    if (valueint >= RangeFrom && valueint <= RangeTo)
                    {
                        // Adjust the index based on the RangeFrom value
                        int adjustedIndex = RangeFrom == 0 ? valueint : valueint - RangeFrom + 1;

                        // Increment the count at the adjusted index
                        indexes[adjustedIndex]++;
                    }

                }

                

                CalculateData(RangeFrom , RangeTo , datacount , indexes);
                AddDatatoWord(table  , WordTableColumns , startrow, indexes);
                startrow++;
            }

            

        }

        public void CalculateData(int RangeFrom, int RangeTo, float datacount, float[]indexes)
        {
            WordData.Clear();
            for (int i = RangeFrom; i <= RangeTo; i++)
            {
                float IndexPerc = indexes[i] / datacount*100;
               
                WordData.Add(indexes[i]);
                WordData.Add(IndexPerc);
            }
        }
        
        
        public void AddDatatoWord(IWTable table ,  int WordTableColumns , int startRow , float[] indexes)
        {
            int ColumnStart = 2;
            

            for (int worddatactr = 0; worddatactr < WordData.Count; worddatactr++)
            {

                
                if (worddatactr % 2 == 0)
                {
                    wordObj.Addpara_CenterNoBOLD(table, startRow, ColumnStart, WordData[worddatactr].ToString());
                   
                }
                else if (worddatactr % 2 != 0)
                {
                    wordObj.Addpara_CenterNoBOLD(table, startRow, ColumnStart, WordData[worddatactr].ToString("0.0")); 
                }
                
                ColumnStart++;
                
            }

            WordData.Clear();
            Array.Clear(indexes, 0, indexes.Length);

            //MessageBox.Show(startRow.ToString());

        }
        public void AddTotalAvgScore_Word(IWTable table  , int itemindex , int InsertColumnPer , string TotalMinMaxValueString ,  string TotalMeanStdValueString , string TotalMedianString , string AvgMeanStdValueString)
        {
            
            int TotRowMinMax = 3 + (itemindex * 6);
            int TotRowMeanStd = 3 + (itemindex * 6) + 1;
            int TotRowMedian = 3 + (itemindex * 6) + 2;
            int AvgRowMeanStd = 3 + (itemindex * 6) + 3;

            wordObj.Addpara_CenterNoBOLD(table, TotRowMinMax, InsertColumnPer, TotalMinMaxValueString);
            wordObj.Addpara_CenterNoBOLD(table, TotRowMeanStd, InsertColumnPer, TotalMeanStdValueString);
            wordObj.Addpara_CenterNoBOLD(table, TotRowMedian, InsertColumnPer, TotalMedianString);
            wordObj.Addpara_CenterNoBOLD(table, AvgRowMeanStd, InsertColumnPer, AvgMeanStdValueString);
        }


        public void AddOverallNoPer_Word(IWTable table, int LevelCount, int InsertColumnPer, string TotalMinMaxValueString, string TotalMeanStdValueString, string TotalMedianString, string AvgMeanStdValueString)
        {
            wordObj.Addpara_CenterNoBOLD(table, LevelCount + 2, InsertColumnPer, TotalMinMaxValueString);
            wordObj.Addpara_CenterNoBOLD(table, LevelCount + 3, InsertColumnPer, TotalMeanStdValueString);
            wordObj.Addpara_CenterNoBOLD(table, LevelCount + 4, InsertColumnPer, TotalMedianString);
            wordObj.Addpara_CenterNoBOLD(table, LevelCount + 5, InsertColumnPer, AvgMeanStdValueString);

        }
        public void AddOverall_Word(IWTable table, int LevelCount, int InsertColumnPer, string TotalMinMaxValueString, string TotalMeanStdValueString, string TotalMedianString, string AvgMeanStdValueString)
        {
            wordObj.Addpara_CenterNoBOLD(table, LevelCount+3, InsertColumnPer, TotalMinMaxValueString);
            wordObj.Addpara_CenterNoBOLD(table, LevelCount + 4, InsertColumnPer, TotalMeanStdValueString);
            wordObj.Addpara_CenterNoBOLD(table, LevelCount + 5, InsertColumnPer, TotalMedianString);
            wordObj.Addpara_CenterNoBOLD(table, LevelCount + 6, InsertColumnPer, AvgMeanStdValueString);

        }
        public double[] OverallLevelNoPerc_NoPeriods(Worksheet sheet, IWTable table, int Excelrows, int WordTableColumns, int counter, int itemindex, int InsertColumnPer, string TotalScoreOrOverall, int LevelCount, List<string> LevelValues)
        {
            ExcelWorksheet_ToWord.CalculateFormula();
            double[] OverallValues = new double[LevelCount];
            int OverallLevelcolumn = counter + 3;
            Range OverallLevelrange = sheet.Cells.CreateRange(2, OverallLevelcolumn, Excelrows, 1);
            double[] OverallLevelValues = new double[OverallLevelrange.RowCount - 1];
            for (int i = 0; i < OverallLevelrange.RowCount - 1; i++)
            {
                Cell OverallLevelcell = OverallLevelrange[i, 0];


                if (OverallLevelcell.IsNumericValue)
                {
                    OverallLevelValues[i] = OverallLevelcell.DoubleValue;

                }

            }

            Dictionary<double, int> valueCounts = new Dictionary<double, int>();

            foreach (var value in OverallLevelValues)
            {
                valueCounts[value] = 0;
            }
            // Count occurrences of each value and store in the dictionary
            foreach (var value in OverallLevelValues)
            {
                if (valueCounts.ContainsKey(value))
                {
                    valueCounts[value]++;
                }
                else
                {
                    valueCounts[value] = 1;
                }


            }

            for (int i = 0; i < LevelValues.Count; i++)
            {
                double CurrentLevel = double.Parse(LevelValues[i]);

                // Check if CurrentLevel exists in valueCounts
                if (valueCounts.ContainsKey(CurrentLevel))
                {
                    OverallValues[i] = valueCounts[CurrentLevel];
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer, valueCounts[CurrentLevel].ToString());
                }
                else
                {
                    // Handle case where CurrentLevel does not exist in valueCounts
                    OverallValues[i] = 0;  // Or any other default value or handling logic
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer, "0");  // Example default output
                }
            }

            int totalCount = OverallLevelValues.Length;

            for (int i = 0; i < LevelValues.Count; i++)
            {
                double CurrentLevel = double.Parse(LevelValues[i]);

                // Check if CurrentLevel exists in valueCounts
                if (valueCounts.ContainsKey(CurrentLevel))
                {
                    double percentage = (double)valueCounts[CurrentLevel] / totalCount * 100;
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer + 1, percentage.ToString("0.0"));
                }
                else
                {
                    // Handle case where CurrentLevel does not exist in valueCounts
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer + 1, "0.0");  // Example default output
                }
            }

            return OverallValues;


            //for (int i = 0; i < LevelValues.Count; i++)
            //{

            //    double CurrentLevel = double.Parse(LevelValues[i]);
            //    OverallValues[i] = valueCounts[CurrentLevel];
            //    //MessageBox.Show(OverallValues[i].ToString());
            //    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer, valueCounts[CurrentLevel].ToString());
            //}

            //int totalCount = OverallLevelValues.Length;

            //for (int i = 0; i < LevelValues.Count; i++)
            //{
            //    double CurrentLevel = double.Parse(LevelValues[i]);
            //    double percentage = (double)valueCounts[CurrentLevel] / totalCount * 100;

            //    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer + 1, percentage.ToString("0.0"));
            //}

            //return OverallValues;

        }
        public double[] OverallLevelNoPerc(Worksheet sheet, IWTable table, int Excelrows, int WordTableColumns, int counter, int itemindex, int InsertColumnPer, string TotalScoreOrOverall, int LevelCount , List<string> LevelValues)
        {
            ExcelWorksheet_ToWord.CalculateFormula();
            double[] OverallValues = new double[LevelCount];
            int OverallLevelcolumn = counter+3;
            Range OverallLevelrange = sheet.Cells.CreateRange(2, OverallLevelcolumn, Excelrows, 1);
            double[] OverallLevelValues = new double[OverallLevelrange.RowCount - 1];
            for (int i = 0; i < OverallLevelrange.RowCount - 1; i++)
            {
                Cell OverallLevelcell = OverallLevelrange[i, 0];
                

                if (OverallLevelcell.IsNumericValue)
                {
                    OverallLevelValues[i] = OverallLevelcell.DoubleValue;

                }  
                
            }

            Dictionary<double, int> valueCounts = new Dictionary<double, int>();

            // Count occurrences of each value and store in the dictionary


            foreach (var value in OverallLevelValues)
            {
                valueCounts[value] = 0;
            }
            // Count occurrences of each value and store in the dictionary
            foreach (var value in OverallLevelValues)
            {
                if (valueCounts.ContainsKey(value))
                {
                    valueCounts[value]++;
                }
                else
                {
                    valueCounts[value] = 1;
                }


            }

            for (int i = 0; i < LevelValues.Count; i++)
            {
                double CurrentLevel = double.Parse(LevelValues[i]);

                // Check if CurrentLevel exists in valueCounts
                if (valueCounts.ContainsKey(CurrentLevel))
                {
                    OverallValues[i] = valueCounts[CurrentLevel];
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer, valueCounts[CurrentLevel].ToString());
                }
                else
                {
                    // Handle case where CurrentLevel does not exist in valueCounts
                    OverallValues[i] = 0;  // Or any other default value or handling logic
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer, "0");  // Example default output
                }
            }

            int totalCount = OverallLevelValues.Length;

            for (int i = 0; i < LevelValues.Count; i++)
            {
                double CurrentLevel = double.Parse(LevelValues[i]);

                // Check if CurrentLevel exists in valueCounts
                if (valueCounts.ContainsKey(CurrentLevel))
                {
                    double percentage = (double)valueCounts[CurrentLevel] / totalCount * 100;
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer + 1, percentage.ToString("0.0"));
                }
                else
                {
                    // Handle case where CurrentLevel does not exist in valueCounts
                    wordObj.Addpara_CenterNoBOLD(table, i + 1, InsertColumnPer + 1, "0.0");  // Example default output
                }
            }

            return OverallValues;

        }
        public void AddDataWord_TotalScoreNoPer(IWTable table, string TotalMinMaxValueString, string TotalMeanStdValueString, string TotalMedianString, string AvgMeanStdValueString , int row)
        {
            wordObj.Addpara_CenterNoBOLD(table, row, 2, TotalMinMaxValueString);
            wordObj.Addpara_CenterNoBOLD(table, row, 3, TotalMeanStdValueString);
            wordObj.Addpara_CenterNoBOLD(table, row, 4, TotalMedianString);
            wordObj.Addpara_CenterNoBOLD(table, row, 5, AvgMeanStdValueString);

        }
        public double[] GetTotal_AVG_Score(Worksheet sheet, IWTable table, int Excelrows, int WordTableColumns, int counter, int itemindex, int InsertColumnPer, string TotalScoreOrOverall , int LevelCount)
        {

            ExcelWorksheet_ToWord.CalculateFormula();
            int TotalScorecolumn = counter;
            int AvgScoreColumn = counter + 1;

            Range TotalScorerange = sheet.Cells.CreateRange(2, TotalScorecolumn, Excelrows, 1);
            Range AvgScorerange = sheet.Cells.CreateRange(2, AvgScoreColumn, Excelrows, 1);

            double[] TotalScorecellValues = new double[TotalScorerange.RowCount - 1];
            double[] AvgScorecellValues = new double[TotalScorerange.RowCount - 1];

            for (int i = 0; i < TotalScorerange.RowCount - 1; i++)
            {
                Cell TotalScorecell = TotalScorerange[i, 0];
                Cell AvgScoreCell = AvgScorerange[i, 0];

                if (TotalScorecell.IsNumericValue)
                {
                    TotalScorecellValues[i] = TotalScorecell.DoubleValue;
                }
                if (AvgScoreCell.IsNumericValue)
                {
                    AvgScorecellValues[i] = AvgScoreCell.DoubleValue;
                }
            }

            double TotminValue = TotalScorecellValues.Min();
            double TotmaxValue = TotalScorecellValues.Max();
            double TotmeanValue = TotalScorecellValues.Average();
            double TotstdDevValue = Math.Sqrt(TotalScorecellValues.Select(x => Math.Pow(x - TotmeanValue, 2)).Sum() / (TotalScorecellValues.Length - 1));
            double TotmedianValue = TotalScorecellValues.OrderBy(x => x).ElementAt(TotalScorecellValues.Length / 2);

            double AvgmeanValue = AvgScorecellValues.Average();
            double AvgstdDevValue = Math.Sqrt(AvgScorecellValues.Select(x => Math.Pow(x - AvgmeanValue, 2)).Sum() / (AvgScorecellValues.Length - 1));




            string TotalMinMaxValueString = FormatMinMaxValue(TotminValue, TotmaxValue);
            string TotalMeanStdValueString = FormatMeanStdValue(TotmeanValue, TotstdDevValue);
            string TotalMedianString = FormatSingleValue(TotmedianValue);

            string AvgMeanStdValueString = FormatMeanStdValue(AvgmeanValue, AvgstdDevValue);


            if (TotalScoreOrOverall == "T")
            {
                AddTotalAvgScore_Word(table, itemindex, InsertColumnPer, TotalMinMaxValueString, TotalMeanStdValueString, TotalMedianString, AvgMeanStdValueString);
            }
            else if(TotalScoreOrOverall == "O")
            {
                AddOverall_Word(table, LevelCount, InsertColumnPer, TotalMinMaxValueString, TotalMeanStdValueString, TotalMedianString, AvgMeanStdValueString);
            }
            //TNO is for Total Score No Periods
            else if(TotalScoreOrOverall == "TNO")
            {
                AddDataWord_TotalScoreNoPer(table, TotalMinMaxValueString, TotalMeanStdValueString, TotalMedianString, AvgMeanStdValueString , itemindex);

            }
            //Overall No periods
            else if(TotalScoreOrOverall == "ONO")
            {
                AddOverallNoPer_Word(table, LevelCount, InsertColumnPer, TotalMinMaxValueString, TotalMeanStdValueString, TotalMedianString, AvgMeanStdValueString);
            }
            


            return TotalScorecellValues;

        }

        
        string FormatMinMaxValue(double minValue, double maxValue)
        {
            string minValueString = minValue.ToString("0.00");
            string maxValueString = maxValue.ToString("0.00");

            // Check and remove trailing "0" if decimals are "00"
            if (minValueString.EndsWith(".00"))
            {
                minValueString = minValueString.Substring(0, minValueString.Length - 1);
            }
            if (maxValueString.EndsWith(".00"))
            {
                maxValueString = maxValueString.Substring(0, maxValueString.Length - 1);
            }

            return minValueString + " – " + maxValueString;
        }
        string FormatMeanStdValue(double meanValue, double stdDevValue)
        {
            string meanValueString = meanValue.ToString("0.00");
            string stdDevValueString = stdDevValue.ToString("0.00");

            // Check and remove trailing "0" if decimals are "00"
            if (meanValueString.EndsWith(".00"))
            {
                meanValueString = meanValueString.Substring(0, meanValueString.Length - 1);
            }
            if (stdDevValueString.EndsWith(".00"))
            {
                stdDevValueString = stdDevValueString.Substring(0, stdDevValueString.Length - 1);
            }

            return meanValueString + " ± " + stdDevValueString;
        }
        string FormatSingleValue(double value)
        {
            string valueString = value.ToString("0.00");

            // Check and remove trailing "0" if decimals are "00"
            if (valueString.EndsWith(".00"))
            {
                valueString = valueString.Substring(0, valueString.Length - 1);
            }

            return valueString;
        }


        public void ReadIndexesExcel_periods(Worksheet sheet, IWTable table, int rows, int QuestionNum, int RangeFrom, int RangeTo, float datacount, int WordTableColumns,int ColumnStart , int counter)
        {
            int arraySize = RangeTo - RangeFrom + 2;
            float[] indexes = new float[arraySize];

            for (int ca = counter; ca < QuestionNum + counter; ca++)
            {
                for (int ra = 2; ra < rows + 1; ra++)
                {

                    string value = sheet.Cells[ra, ca].Value.ToString();
                    int valueint = Convert.ToInt32(value);

                    if (valueint >= RangeFrom && valueint <= RangeTo)
                    {
                        int adjustedIndex = RangeFrom == 0 ? valueint : valueint - RangeFrom + 1;

                        indexes[adjustedIndex]++;
                    }

                }



                CalculateData(RangeFrom, RangeTo, datacount, indexes);
                AddDatatoWord_Periods(table, WordTableColumns, startrow, indexes , ColumnStart);
                if(startrow < QuestionNum+2)
                {
                    
                    startrow++;
                }
                
            }



        }
        public void AddDatatoWord_Periods(IWTable table, int WordTableColumns, int startRow, float[] indexes , int ColumnStart)
        {
            
            for (int worddatactr = 0; worddatactr < WordData.Count; worddatactr++)
            {


                if (worddatactr % 2 == 0)
                {
                    wordObj.Addpara_CenterNoBOLD(table, startRow, ColumnStart, WordData[worddatactr].ToString());

                }
                else if (worddatactr % 2 != 0)
                {
                    wordObj.Addpara_CenterNoBOLD(table, startRow, ColumnStart, WordData[worddatactr].ToString("0.0"));
                }
                ColumnStart++;

            }

            WordData.Clear();
            Array.Clear(indexes, 0, indexes.Length);

        }


        public void OverallPeriods()
        {

            Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
            Worksheet Sheet3 = ExcelWorksheet_ToWord.Worksheets[2];
            int Excelrows = Sheet2.Cells.MaxDataRow;
            int ExcelCols = Sheet2.Cells.MaxDataColumn;
            float Exceldatacount = Excelrows - 1;
            int WordTableColumns = 7;


            List<double[]> TpairedLists = new List<double[]>();
            List<double[]> OverallLists = new List<double[]>();

            for (int index = 0; index < list_Overall.Items.Count; index++)
            {
                TpairedLists.Clear();
                OverallLists.Clear();

                int InsertColumnPer = 1;

                IWSection section = wordObj.CreatePortraitSection();

                wordObj.AddTitle(section, "", 0);


                List<string> LevelValues = new List<string>();

                var item = list_Overall.Items[index];
                string[] domains = item.ToString().Split(' ');
                string CurrenOverall = domains[0];

                LevelValues = excelobj.GetLevelCount(Sheet3, CurrenOverall);
                int DomainGetInfoIndex = AllDomainsWord.FindIndex(domain => domain.DomainName == domains[0]);

                int WordTableRows = 2 + LevelValues.Count + 5;

                int DomainRangeFrom = AllDomainsWord[DomainGetInfoIndex].RangeFrom;
                int DomainRangeTo = AllDomainsWord[DomainGetInfoIndex].RangeTo;

                int TotalScoreFrom = AllDomainsWord[DomainGetInfoIndex].QuestionsSorted_ReverseValue.Count * DomainRangeFrom;
                int TotalScoreTo = AllDomainsWord[DomainGetInfoIndex].QuestionsSorted_ReverseValue.Count * DomainRangeTo;

                int AvgScorefrom = DomainRangeFrom;
                int AvgScoreTo = DomainRangeTo;

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                wordObj.GeneralTableFormat(table);
                wordObj.ApplyMergesOverallPer(table , LevelValues.Count , WordTableRows);
                wordObj.ApplyOverallPeriodsBorders(table, WordTableRows, WordTableColumns , LevelValues.Count);
                wordObj.SetOverallPerWidths(table, WordTableRows, WordTableColumns);
                wordObj.AddHorizontalTitlesOverallPer(table);

                wordObj.AddVerticalTitlesOverallPer(table, LevelValues.Count, TotalScoreFrom, TotalScoreTo, AvgScorefrom, AvgScoreTo);


                for (int i = 0; i < LevelValues.Count; i++)
                {
                    wordObj.Addpara_NoCenterNoBOLD(table, i+2, 0, LevelValues[i]);
                    wordObj.LeftIntendBeforeText(table, i+2, 0 ,14.17f);
                }

                int DomainCount = domains.Length;

                foreach (string Mydomain in domains)
                {
                    int counter = 0;
                    int Currentindex = AllDomainsWord.FindIndex(domain => domain.DomainName == Mydomain);
                    int CurrentdataCount = AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;

                    if (Currentindex == 0)
                    {
                        counter = 0;
                    }
                    else if (Currentindex > 0)
                    {
                        for (int i = 0; i < Currentindex; i++)
                        {
                            counter = counter + AllDomainsWord[i].QuestionsSorted_ReverseValue.Count;
                            counter = counter + 5;
                        }

                    }
                    counter = counter + AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;


                    TpairedLists.Add(GetTotal_AVG_Score(Sheet2, table, Excelrows, WordTableColumns, counter, index, InsertColumnPer , "O" , LevelValues.Count));
                   OverallLists.Add(OverallLevelNoPerc(Sheet2, table, Excelrows, WordTableColumns, counter, index, InsertColumnPer, "O", LevelValues.Count , LevelValues));

                    
                    InsertColumnPer = InsertColumnPer+2;

                }



                
                string[] TestValueFrm = new string[] { };
                int TestInsertRow = LevelValues.Count + 3;
                TestValueFrm = pyobj.TpairedTest(TpairedLists[0].ToList(), TpairedLists[1].ToList());
                
                wordObj.InsertTest_P(table, TestInsertRow, 5, TestValueFrm);    
            }

                
            

            //int WordTableRows = ;


        }

        public void OverallNoPeriods()
        {

            Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
            Worksheet Sheet3 = ExcelWorksheet_ToWord.Worksheets[2];
            int Excelrows = Sheet2.Cells.MaxDataRow;
            int ExcelCols = Sheet2.Cells.MaxDataColumn;
            float Exceldatacount = Excelrows - 1;
            int WordTableColumns = 3;


            List<double[]> TpairedLists = new List<double[]>();
            List<double[]> OverallLists = new List<double[]>();

            for (int index = 0; index < list_Overall.Items.Count; index++)
            {
                TpairedLists.Clear();
                OverallLists.Clear();

                int InsertColumnPer = 1;

                IWSection section = wordObj.CreatePortraitSection();

                wordObj.AddTitle(section, "", 0);


                List<string> LevelValues = new List<string>();

                var item = list_Overall.Items[index];
                string[] domains = item.ToString().Split(' ');
                string CurrenOverall = domains[0];

                LevelValues = excelobj.GetLevelCount(Sheet3, CurrenOverall);
                int DomainGetInfoIndex = AllDomainsWord.FindIndex(domain => domain.DomainName == domains[0]);

                int WordTableRows = 1 + LevelValues.Count + 5;

                int DomainRangeFrom = AllDomainsWord[DomainGetInfoIndex].RangeFrom;
                int DomainRangeTo = AllDomainsWord[DomainGetInfoIndex].RangeTo;

                int TotalScoreFrom = AllDomainsWord[DomainGetInfoIndex].QuestionsSorted_ReverseValue.Count * DomainRangeFrom;
                int TotalScoreTo = AllDomainsWord[DomainGetInfoIndex].QuestionsSorted_ReverseValue.Count * DomainRangeTo;

                int AvgScorefrom = DomainRangeFrom;
                int AvgScoreTo = DomainRangeTo;

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                wordObj.GeneralTableFormat(table);
                wordObj.ApplyMergesOverallNoPer(table, LevelValues.Count, WordTableRows);

                wordObj.ApplyOverallNoPeriodsBorders(table, WordTableRows, WordTableColumns, LevelValues.Count);

                wordObj.SetOverallNoPerWidths(table, WordTableRows, WordTableColumns);

                wordObj.AddHorizontalTitlesOverallNoPer(table);

                wordObj.AddVerticalTitlesOverallNoPer(table, LevelValues.Count, TotalScoreFrom, TotalScoreTo, AvgScorefrom, AvgScoreTo);


                for (int i = 0; i < LevelValues.Count; i++)
                {
                    wordObj.Addpara_NoCenterNoBOLD(table, i + 1, 0, LevelValues[i]);
                    wordObj.LeftIntendBeforeText(table, i + 1, 0, 14.17f);
                }

                int DomainCount = domains.Length;

                foreach (string Mydomain in domains)
                {
                    int counter = 0;
                    int Currentindex = AllDomainsWord.FindIndex(domain => domain.DomainName == Mydomain);
                    int CurrentdataCount = AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;

                    if (Currentindex == 0)
                    {
                        counter = 0;
                    }
                    else if (Currentindex > 0)
                    {
                        for (int i = 0; i < Currentindex; i++)
                        {
                            counter = counter + AllDomainsWord[i].QuestionsSorted_ReverseValue.Count;
                            counter = counter + 5;
                        }

                    }
                    counter = counter + AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;


                    TpairedLists.Add(GetTotal_AVG_Score(Sheet2, table, Excelrows, WordTableColumns, counter, index, InsertColumnPer, "ONO", LevelValues.Count));
                    OverallLists.Add(OverallLevelNoPerc_NoPeriods(Sheet2, table, Excelrows, WordTableColumns, counter, index, InsertColumnPer, "ONO", LevelValues.Count, LevelValues));


                    InsertColumnPer = InsertColumnPer++;
                    wordObj.FormatTable(table, 12);
                }




                //string[] TestValueFrm = new string[] { };
                //int TestInsertRow = LevelValues.Count + 3;
                //TestValueFrm = pyobj.TpairedTest(TpairedLists[0].ToList(), TpairedLists[1].ToList());

                //wordObj.InsertTest_P(table, TestInsertRow, 5, TestValueFrm);
            }


            

            //int WordTableRows = ;


        }

        public void TotalScoreNoPer()
        {
            Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
            Worksheet Sheet3 = ExcelWorksheet_ToWord.Worksheets[2];
            int Excelrows = Sheet2.Cells.MaxDataRow;
            int ExcelCols = Sheet2.Cells.MaxDataColumn;
            float Exceldatacount = Excelrows - 1;

            int WordTableColumns = 6;

            


            
            foreach (var item in list_TotalScorePeriods.Items)
            {

                IWSection section = wordObj.CreatePortraitSection();

                wordObj.AddTitle(section, "", 0);

                string[] domains = item.ToString().Split(' ');
                int NumberOfDomains = domains.Length;

                int WordTableRows = NumberOfDomains + 2;
                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);


                wordObj.GeneralTableFormat(table);

                wordObj.ApplyMergesTotalScoreNoPer(table);
                wordObj.ApplyTotalScoreNoPBorders(table, WordTableRows, WordTableColumns);
                wordObj.SetTotalScoreNoPerWidths(table, WordTableRows, WordTableColumns);



                int DomainGetInfoIndex = AllDomainsWord.FindIndex(domain => domain.DomainName == domains[0]);

                int DomainRangeFrom = AllDomainsWord[DomainGetInfoIndex].RangeFrom;
                int DomainRangeTo = AllDomainsWord[DomainGetInfoIndex].RangeTo;

                

                int AvgScorefrom = DomainRangeFrom;
                int AvgScoreTo = DomainRangeTo;


                wordObj.AddPara_Center(table, 0, 1, "Score Range");
                wordObj.AddPara_Center(table, 0, 2, "Total score");

                wordObj.AddPara_Center(table, 1, 2, "Min. – Max.");

                wordObj.AddPara_Center(table, 1, 3, "Mean ± SD.");

                wordObj.AddPara_Center(table, 1, 4, "Median");

                wordObj.AddPara_Center(table, 0, 5, "Average Score\r\n("+ AvgScorefrom+ " – "+ AvgScoreTo+ ")");
                wordObj.AddPara_Center(table, 1, 5, "Mean ± SD.");

                int insertScoreRangectr = 2;

                for (int k = 0; k < domains.Length; k++)
                {
                    int counter = 0;
                    int Currentindex = AllDomainsWord.FindIndex(domain => domain.DomainName == domains[k]);
                    int CurrentdataCount = AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;


                    int TotalScoreFrom = AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count * DomainRangeFrom;
                    int TotalScoreTo = AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count * DomainRangeTo;


                    if (Currentindex == 0)
                    {
                        counter = 0;
                    }
                    else if (Currentindex > 0)
                    {
                        for (int i = 0; i < Currentindex; i++)
                        {
                            counter = counter + AllDomainsWord[i].QuestionsSorted_ReverseValue.Count;
                            counter = counter + 5;
                        }

                    }
                    counter = counter + AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;


                    string CurrentScoreRange = "(" + TotalScoreFrom + " - " + TotalScoreTo +")";
                    wordObj.InsertScoreRangesTotalScore(table, insertScoreRangectr, CurrentScoreRange);
                    insertScoreRangectr++;

                    GetTotal_AVG_Score(Sheet2, table, Excelrows, WordTableColumns, counter, k+2 , 1, "TNO", 1);


                }






                wordObj.Font(table, 10.5f);

            }
            



        }
        public void TotalScorePeriods()
        {
           
            Worksheet Sheet2 = ExcelWorksheet_ToWord.Worksheets[1];
            Worksheet Sheet3 = ExcelWorksheet_ToWord.Worksheets[2];
            int Excelrows = Sheet2.Cells.MaxDataRow;
            int ExcelCols = Sheet2.Cells.MaxDataColumn;
            float Exceldatacount = Excelrows - 1;


            int WordTableColumns = 5;

            
            int NumberOfDomains = list_TotalScorePeriods.Items.Count;


            //foreach (var item in list_TotalScorePeriods.Items)
            //{
            //    string[] domains = item.ToString().Split(' ');
            //    NumberOfDomains = NumberOfDomains+ domains.Length;
            //}
            
            int  WordTableRows = (NumberOfDomains * 6) + 1;

            

            

            IWSection section = wordObj.CreatePortraitSection();

            wordObj.AddTitle(section, "", 0);


            IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

            wordObj.GeneralTableFormat(table);

            wordObj.ApplyMergesTotalScorePer(table, WordTableRows);

            wordObj.ApplyTotalScorePeriodsBorders(table, WordTableRows, WordTableColumns);

            wordObj.SetTotalScorePerWidths(table, WordTableRows,WordTableColumns);

            wordObj.AddPara_Center(table, 0, 1, "Pre");
            wordObj.AddPara_Center(table, 0, 2, "Post");
            //wordObj.AddPara_Center(table, 0, 3, "t");

            

            wordObj.InsertHighlightTestName(table, 0, 3, "t");

            wordObj.AddPara_Center(table, 0, 4, "p");


            //wordObj.ConstStringsTotalScorePeriods(table, WordTableRows, WordTableColumns);
            List<double[]> TpairedLists = new List<double[]>();  

            for (int index = 0; index < list_TotalScorePeriods.Items.Count; index++)
            {
                TpairedLists.Clear();


                int InsertColumnPer = 1;

                var item = list_TotalScorePeriods.Items[index];
                string[] domains = item.ToString().Split(' ');

                //Here i get the first domain [0] of each one of item in order to pass the data totalscorefrom .. etc to put it inside parantheses in table
                int DomainGetInfoIndex = AllDomainsWord.FindIndex(domain => domain.DomainName == domains[0]);

                int DomainRangeFrom = AllDomainsWord[DomainGetInfoIndex].RangeFrom;
                int DomainRangeTo = AllDomainsWord[DomainGetInfoIndex].RangeTo;

                int TotalScoreFrom = AllDomainsWord[DomainGetInfoIndex].QuestionsSorted_ReverseValue.Count * DomainRangeFrom;
                int TotalScoreTo = AllDomainsWord[DomainGetInfoIndex].QuestionsSorted_ReverseValue.Count * DomainRangeTo;

                int AvgScorefrom = DomainRangeFrom;
                int AvgScoreTo = DomainRangeTo;



                wordObj.ConstStringsTotalScorePeriods(table, WordTableRows, WordTableColumns , TotalScoreFrom , TotalScoreTo , AvgScorefrom , AvgScoreTo ,index);
                int DomainCount = domains.Length;   
                
                foreach (string Mydomain in domains) 
                {
                    int counter = 0;
                    int Currentindex = AllDomainsWord.FindIndex(domain => domain.DomainName == Mydomain);
                    int CurrentdataCount = AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;

                    if (Currentindex == 0)
                    {
                        counter = 0;
                    }
                    else if (Currentindex > 0)
                    {
                        for (int i = 0; i < Currentindex; i++)
                        {
                            counter = counter + AllDomainsWord[i].QuestionsSorted_ReverseValue.Count;
                            counter = counter + 5;
                        }

                    }
                    counter = counter + AllDomainsWord[Currentindex].QuestionsSorted_ReverseValue.Count;


                    TpairedLists.Add(GetTotal_AVG_Score(Sheet2, table, Excelrows, WordTableColumns, counter , index , InsertColumnPer , "T" , 1));
                    InsertColumnPer++;

                }
                string[] TestValueFrm = new string[] { };

                int TestInsertRow = 2 + (index * 6);

                TestValueFrm = pyobj.TpairedTest(TpairedLists[0].ToList(), TpairedLists[1].ToList());

                wordObj.InsertTest_P(table, TestInsertRow, 3, TestValueFrm);


                
            }
            

        }

        



        private void btnDone_Click(object sender, EventArgs e)
        {
            wordObj.InitWord();

            
            decide();
            Descriptive_Statement();

            if(list_TotalScorePeriods.Items.Count > 0)
            {
                TotalScoreNoPer();
                if(periods_check.Checked)
                {
                    TotalScorePeriods();
                }
                
            }
            if(list_Overall.Items.Count >0)
            {
                OverallNoPeriods();
                if (periods_check.Checked)
                {
                    OverallPeriods();
                }  
            }
            
            
            wordObj.SaveWord();

        }

        

        private void pic_exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void list_Periods_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pic_AddPeriodsList_Click(object sender, EventArgs e)
        {
            var selectedDomains = list_allDomains.SelectedItems.Cast<string>();
            string domainsConcatenated = string.Join(" ", selectedDomains); // Space as separator

            if (!string.IsNullOrEmpty(domainsConcatenated))
            {
                list_Periods.Items.Add(domainsConcatenated);
            }
        }

        private void pic_RemovePeriodsList_Click(object sender, EventArgs e)
        {
            if (list_Periods.SelectedIndex != -1)
            {

                list_Periods.Items.Remove(list_Periods.SelectedItem);

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void list_allDomains_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        
        

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void Descriptive_Click(object sender, EventArgs e)
        {

        }

        private void pic_AddDescriptive_Click(object sender, EventArgs e)
        {
            var selectedDomain = list_allDomains.SelectedItem.ToString();
            

            if (!string.IsNullOrEmpty(selectedDomain))
            {
                list_descriptive.Items.Add(selectedDomain);
            }
        }

        private void pic_removeDescriptive_Click(object sender, EventArgs e)
        {
            if (list_descriptive.SelectedIndex != -1)
            {

                list_descriptive.Items.Remove(list_descriptive.SelectedItem);

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_addTotalScoreP_Click(object sender, EventArgs e)
        {
            var selectedDomains = list_allDomains.SelectedItems.Cast<string>();
            string domainsConcatenated = string.Join(" ", selectedDomains); // Space as separator

            if (!string.IsNullOrEmpty(domainsConcatenated))
            {
                list_TotalScorePeriods.Items.Add(domainsConcatenated);
            }
        }

        private void pic_removeTotalScoreP_Click(object sender, EventArgs e)
        {
            if (list_TotalScorePeriods.SelectedIndex != -1)
            {

                list_TotalScorePeriods.Items.Remove(list_TotalScorePeriods.SelectedItem);

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_addOverall_Click(object sender, EventArgs e)
        {
            var selectedDomains = list_allDomains.SelectedItems.Cast<string>();
            string domainsConcatenated = string.Join(" ", selectedDomains); // Space as separator

            if (!string.IsNullOrEmpty(domainsConcatenated))
            {
                list_Overall.Items.Add(domainsConcatenated);
            }
        }

        private void pic_removeOverall_Click(object sender, EventArgs e)
        {
            if (list_Overall.SelectedIndex != -1)
            {

                list_Overall.Items.Remove(list_Overall.SelectedItem);

            }
            else
                MessageBox.Show("Please Select Item!");
        }
    }
}
