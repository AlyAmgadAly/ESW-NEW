using ExcelScore.Classes;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.StatClasses
{
    public static class WordTableStatDesign
    {
        

        private static readonly Dictionary<string, Action<WordTableDesignContext>> _designs
        = new(StringComparer.OrdinalIgnoreCase)
    {
        { "Comparative_Default", WordT_Comparative_Default }
    };

        public static void Execute(WordTableDesignContext context)
        {
            string key = $"{context.Table.TableType}_{context.Table.TableDesignType}";

            if (_designs.TryGetValue(key, out var action))
                action(context);
            else
                throw new Exception($"No design found for: {key}");
        }

        // ================= DESIGNS =================

        // TableDesignType = "Default";
        // TableType = "comparative";
        public static void WordT_Comparative_Default(WordTableDesignContext context)
        {
            var StatTable = context.Table;
            var wordobj = context.Wordobj;

            Dictionary<string, bool> CheckedDataprimary = FormDataTransfer.Get<Dictionary<string, bool>>("nodeCheckedStatusPrimary");
            Dictionary<string, bool> CheckedDataExtra = FormDataTransfer.Get<Dictionary<string, bool>>("nodeCheckedStatusExtra");
            
            bool Total_Column_Comparative = CheckedDataExtra["TotalColumn"];
            bool HasNominal = StatTable.HasNominal();
            bool HasScale = StatTable.HasScale();

            IWSection section = wordobj.CreatePortraitSection();

            StatParameter groupParameter = StatTable.GetGroupParameters().FirstOrDefault();

            //Column count
            int groupcount = groupParameter.ValueLabels.Count;

            int ColCount = Get_Comparative_Default_ColCount(groupcount , StatTable.HasNominal() , Total_Column_Comparative);

            


            //-----------------------------------------
            //Row Count
            List<string> CheckedPrimaryNeeded = new List<string>();
            string[] keysToCheck = { "NumberOfCasesFirst", "MinMaxFirst", "MeanSDFirst", "MedianIQRFirst", "MedianMinMaxSecond" };
            foreach (string key in keysToCheck)
            {
                if (CheckedDataprimary.TryGetValue(key, out bool isChecked) && isChecked)
                {
                    CheckedPrimaryNeeded.Add(key);
                }
            }

            int Variablerows = Get_Comparative_Default_RowCount(StatTable, CheckedPrimaryNeeded);
            (int PairwiseCount, int TotalPairwiseCount) = CountPairwiseRows_Comparative_Default(StatTable);
            int RowCount = 2 + Variablerows + TotalPairwiseCount;


            //-----------------------------------------
            //Table created
            wordobj.AddComparativeTitle(section, "test", 1);
            IWTable table = wordobj.Createtable(section, RowCount, ColCount);
            wordobj.GeneralTableFormat(table);


            //Merges if nominal only
            if (HasNominal)
            {
                wordobj.ApplyGeneralComparativeMerges_NewComparativeGroups_Fn(table, ColCount, groupcount);
            }


            wordobj.ApplyGeneralComparativeBorders_NewComparativeGroups_Fn(table, RowCount, ColCount, groupcount, HasNominal);
            wordobj.Add_GeneralHeaders_Comparative_Center_NewComparativeGroups_Fn(table, RowCount, ColCount, groupcount, HasNominal);

            Set_Comparative_Widths_Default(table, RowCount, ColCount, groupcount, StatTable, wordobj);

            //test can you see this

            ParamaeterBorders_Comparative_Default(table, RowCount, ColCount, StatTable, groupcount, HasScale, HasNominal, CheckedPrimaryNeeded, PairwiseCount, CheckedDataExtra, Total_Column_Comparative,wordobj);



        }


        public static int Get_Comparative_Default_RowCount(StatTable stattable , List<string> CheckedPrimaryNeeded)
        {
            int rowCount = 0;

            bool hasNominal = false;
            foreach (var parameter in stattable.Parameters)
            {
                if(parameter.IsGroup) { continue; }

                if (parameter.Type == "Nominal")
                { 

                    rowCount += parameter.ValueLabels.Keys.Count + 1;
                    hasNominal = true;
                }
                else if (parameter.Type == "Scale")
                {
                    rowCount += CheckedPrimaryNeeded.Count + 1;
                }
            }

            if (!hasNominal)
            {
                rowCount--;
            }

            return rowCount;


        }

        public static (int baseCount, int totalCount) CountPairwiseRows_Comparative_Default(StatTable stattable)
        {
            //we need to get real groups that have already data
            StatParameter groupParameter = stattable.GetGroupParameters().FirstOrDefault();
            int numberofgroups = groupParameter.ValueLabels.Count;
            int baseCount = numberofgroups > 2 ? numberofgroups - 2 : 0;

            int scaleParameterCount = stattable.Parameters
                .Count(p => !p.IsGroup && p.Type == "Scale");

            int totalCount = baseCount * scaleParameterCount;

            return (baseCount, totalCount);
        }


        public static int Get_Comparative_Default_ColCount(int numberofgroups, bool HasNominal, bool hastotal)
        {
            int totalcol = hastotal ? (HasNominal ? 2 : 1) : 0;

            int colCount = HasNominal
                ? 3 + (numberofgroups * 2) + totalcol
                : 3 + numberofgroups + totalcol;

            return colCount;
        }

        public static void Set_Comparative_Widths_Default(
    IWTable table,
    int wordTableRows,
    int wordTableColumns,
    int numberOfGroups,
    StatTable comparativeTable , WordClass wordobj)
        {
            bool hasNominal = comparativeTable.HasNominal();
            
            float firstColumnWidth;
            float middleColumnWidth;
            float testColumnWidth;

            if (numberOfGroups == 2)
            {
                firstColumnWidth = 4f;
                middleColumnWidth = hasNominal ? 1.75f : 3.5f;
                testColumnWidth = 1.85f;
            }
            else if (numberOfGroups == 3)
            {
                firstColumnWidth = 3.25f;
                middleColumnWidth = hasNominal ? 3.1f / 2f : 3.1f;
                testColumnWidth = 1.6f;
            }
            else
            {
                firstColumnWidth = 3.25f;
                middleColumnWidth = hasNominal ? 3.1f / 2f : 3.1f;
                testColumnWidth = 1.6f;
            }

            for (int i = 0; i < wordTableRows; i++)
            {
                // First column
                table.Rows[i].Cells[0].Width =
                    wordobj.SetColumnWidthInCentimeters(firstColumnWidth);

                // Test statistic columns
                table.Rows[i].Cells[wordTableColumns - 1].Width =
                    wordobj.SetColumnWidthInCentimeters(testColumnWidth);

                table.Rows[i].Cells[wordTableColumns - 2].Width =
                    wordobj.SetColumnWidthInCentimeters(testColumnWidth);

                // Group and total columns
                for (int j = 1; j < wordTableColumns - 2; j++)
                {
                    table.Rows[i].Cells[j].Width =
                        wordobj.SetColumnWidthInCentimeters(middleColumnWidth);
                }
            }
        }
        public static void ParamaeterBorders_Comparative_Default(
   IWTable table, int WordTableRows, int WordTableColumns,
   StatTable comparativeTable, int numberofgroups,
   bool HasScale, bool HasNominal, List<string> CheckedScaleDataNeeded, int PairwiseCount, Dictionary<string, bool> CheckedExtraData , bool TotalCol_Comparative_default , WordClass wordobj)
        {
            InsertGroupTitles_Comparative_Default(table , HasNominal , comparativeTable , TotalCol_Comparative_default , wordobj);

        }

        public static void InsertGroupTitles_Comparative_Default(IWTable table, bool HasNominal, StatTable comparativeTable , bool Total_Column_Comparative , WordClass wordobj)
        {
            int startcol = 1;
            int totalcount = 0;

            if (Total_Column_Comparative && HasNominal)
            {
                startcol = 3;
            }
            else if (Total_Column_Comparative && !HasNominal)
            {
                startcol = 2;
            }

            StatParameter groupparameter = comparativeTable.GetGroupParameters().FirstOrDefault();


            groupparameter.CalculateValueFrequencies();


            //inserting group names
            int GroupsNamesInsert = startcol;
            
            foreach (var item in groupparameter.ValueLabels)
            {
                int code = item.Key;
                string label = item.Value;

                if (groupparameter.ValueFrequencies.TryGetValue((double)code, out double count))
                {
                    string GroupNameCount = label + Convert.ToChar(11) + "(n = " + count + ")";
                    wordobj.AddPara_Center(table, 0, GroupsNamesInsert, GroupNameCount);

                    if (HasNominal)
                    {
                        GroupsNamesInsert = GroupsNamesInsert + 2;
                    }
                    else
                    {
                        GroupsNamesInsert++;
                    }
                    totalcount += (int)count;
                }
                
            }
            if (Total_Column_Comparative)
            {
                string TotalCount = "Total" + Convert.ToChar(11) + "(n = " + totalcount + ")";
                wordobj.AddPara_Center(table, 0, 1, TotalCount);
            }
        }




    }
}
