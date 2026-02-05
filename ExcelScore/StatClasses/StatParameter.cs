using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelScore.StatClasses.SpssReaderStat;

namespace ExcelScore.StatClasses
{
    public class StatParameter
    {
        // --- Basic Info from SPSS or manual entry ---
        public string Name { get; set; }
        public string Label { get; set; }
        public string Type { get; set; } // "Nominal" or "Scale"
        public string Normality { get; set; } // "Nominal" or "Scale"
        public List<string> RawValues { get; set; } = new();

        public List<double> RawNonEmptyValues { get; private set; }
        public Dictionary<int, string> ValueLabels { get; set; } = new(); // e.g., 1 = Male

        //manual
        public Dictionary<double, double> ValueFrequencies { get; set; } = new();

        public void BuildRawNonEmptyValues()
        {
            // Remove SPSS missing values "."
            var nonMissing = RawValues
                .Where(v => !string.IsNullOrWhiteSpace(v) && v != ".")
                .ToList();

            // Nothing left → don't create
            if (!nonMissing.Any())
            {
                RawNonEmptyValues = null;
                return;
            }

            var numericValues = new List<double>();

            foreach (var value in nonMissing)
            {
                if (double.TryParse(
                        value,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double parsed))
                {
                    numericValues.Add(parsed);
                }
                else
                {
                    // Found a non-numeric string → categorical variable
                    RawNonEmptyValues = null;
                    return;
                }
            }

            // All non-missing values are numeric → safe to assign
            RawNonEmptyValues = numericValues;
        }

        public void CalculateValueFrequencies()
        {
            if (Type != "Nominal" || RawValues == null || RawValues.Count == 0)
                return;

            ValueFrequencies.Clear();

            foreach (var val in RawValues)
            {
                if (double.TryParse(val, out double numericVal))
                {
                    if (ValueFrequencies.ContainsKey(numericVal))
                        ValueFrequencies[numericVal]++;
                    else
                        ValueFrequencies[numericVal] = 1;
                }
            }
        }

        public bool AllRawValuesIdentical()
        {
            if (RawValues == null || RawValues.Count == 0)
                return true; // or true, depending on your intended logic

            // Check if all values match the first non-null value
            string first = RawValues[0];
            return RawValues.All(v => v == first);
        }
        // --- Grouping info ---
        public bool IsGroup { get; set; }
        public bool IsSubGroup { get; set; }
        public List<string> GroupingVariables { get; set; } = new(); // used in relations

        public Dictionary<double, List<double>> GroupedParameterValues { get; set; } = new Dictionary<double, List<double>>();

        public Dictionary<string , Dictionary<double, List<double>>> GroupedParameterValuesRelation { get; set; } = new Dictionary<string, Dictionary<double, List<double>>>();

        public Dictionary<double, double> EachGroupCount { get; set; } = new();
        public void AssignGroupedValues(List<StatParameter> groupingParameters)
        {
            if (groupingParameters == null || groupingParameters.Count == 0)
                throw new InvalidOperationException("No grouping parameters provided.");

            int totalRows = RawValues.Count;

            foreach (var groupParam in groupingParameters)
            {
                if (groupParam.RawValues.Count != totalRows)
                    throw new InvalidOperationException($"Grouping parameter '{groupParam.Name}' has mismatched row count.");
            }

            // One group parameter (e.g., classic comparative table)
            if (groupingParameters.Count == 1)
            {
                GroupedParameterValues.Clear();
                EachGroupCount.Clear();

                var groupParam = groupingParameters[0];

                // Step 1: Collect all distinct group values
                var distinctGroups = groupParam.RawValues
                    .Where(v => double.TryParse(v, out _))
                    .Select(v => double.Parse(v))
                    .Distinct();

                // Step 2: Initialize all group keys with empty lists
                foreach (var groupVal in distinctGroups)
                {
                    GroupedParameterValues[groupVal] = new List<double>();
                    EachGroupCount[groupVal] = 0;
                }

                // Step 3: Populate actual grouped values
                for (int i = 0; i < totalRows; i++)
                {
                    if (!double.TryParse(groupParam.RawValues[i], out double groupVal)) continue;
                    if (!double.TryParse(this.RawValues[i], out double paramVal)) continue;

                    GroupedParameterValues[groupVal].Add(paramVal);
                    EachGroupCount[groupVal]++;
                }
            }
            else // Multiple group parameters (e.g., relation tables)
            {
                GroupedParameterValuesRelation.Clear();

                foreach (var groupParam in groupingParameters)
                {
                    string groupName = groupParam.Name;

                    if (!GroupedParameterValuesRelation.ContainsKey(groupName))
                        GroupedParameterValuesRelation[groupName] = new Dictionary<double, List<double>>();

                    // Collect all distinct group values
                    var distinctGroups = groupParam.RawValues
                        .Where(v => double.TryParse(v, out _))
                        .Select(v => double.Parse(v))
                        .Distinct();

                    // Initialize all group keys with empty lists
                    foreach (var groupVal in distinctGroups)
                    {
                        GroupedParameterValuesRelation[groupName][groupVal] = new List<double>();
                    }

                    // Populate actual grouped values
                    for (int i = 0; i < totalRows; i++)
                    {
                        if (!double.TryParse(groupParam.RawValues[i], out double groupVal)) continue;
                        if (!double.TryParse(this.RawValues[i], out double paramVal)) continue;

                        GroupedParameterValuesRelation[groupName][groupVal].Add(paramVal);
                    }
                }
            }
        }

        public static int GetParameterRealCount(StatParameter statParameter)
        {
            int realCount = 0;
            foreach (var value in statParameter.RawValues)
            {
                if(value == ".")
                {
                    continue;
                }
                else
                {
                    realCount++;
                }
            }


            return realCount;
        }

        // --- Statistics Results (parsed or calculated) ---
        public SpssSyntaxStat.DescriptiveResult DescriptiveStats { get; set; }

        public Dictionary<string, SpssSyntaxStat.DescriptiveResult> DescriptiveStatsByGroup { get; set; }
    = new Dictionary<string, SpssSyntaxStat.DescriptiveResult>();


        //    public Dictionary<string, SPSSUnifiedRunner.PercentileResult> PercentileStats { get; set; }
        //= new Dictionary<string, SPSSUnifiedRunner.PercentileResult>();

        public SpssSyntaxStat.PercentileResult PercentileStats { get; set; }
    = new SpssSyntaxStat.PercentileResult();
        public SpssSyntaxStat.CrosstabBlockS ChiSquareBlock { get; set; }
        public List<string> Test_PValues { get; set; } = new(); // t, U, F, H results
        public List<string> PostHocResults { get; set; } = new();


        public static bool General_Parameter_Group_Label(string parameterName, string groupLabel, StatTable statTable, StatParameter groupParameter)
        {
            bool hasvalues = false;

            int groupParametercount = statTable.GetGroupParameters().Count;
            StatParameter CurrentParameter = statTable.GetParameterByName(parameterName);
            int GroupLabel_Key;

            // 1️⃣ Try normal label match
            var match = groupParameter.ValueLabels
                .FirstOrDefault(kv => kv.Value.Equals(groupLabel, StringComparison.OrdinalIgnoreCase));

            if (!match.Equals(default(KeyValuePair<int, string>)))
            {
                GroupLabel_Key = match.Key;
            }
            else
            {
                // 2️⃣ Try to parse numeric label
                if (double.TryParse(groupLabel, out double parsedValue))
                {
                    // Handle numeric keys as int (round if needed)
                    int numericKey = (int)Math.Round(parsedValue);

                    if (groupParameter.ValueLabels.ContainsKey(numericKey))
                        GroupLabel_Key = numericKey;
                    else
                        GroupLabel_Key = -1; // not found
                }
                else
                {
                    GroupLabel_Key = -1; // not found
                }
            }

            if (groupLabel == "Total")
            {
                return true;
            }
            if (groupParametercount == 1)
            {
                if (CurrentParameter.GroupedParameterValues
                        .TryGetValue(GroupLabel_Key, out var values) &&
                    values.Count > 0)
                {
                    hasvalues = true;
                }
            }
            else if (groupParametercount > 1)
            {
                if (CurrentParameter.GroupedParameterValuesRelation
                        .TryGetValue(groupParameter.Name, out var groupDict) &&
                    groupDict.TryGetValue(GroupLabel_Key, out var values) &&
                    values.Count > 0)
                {
                    hasvalues = true;
                }
            }




            return hasvalues;
        }

        public int GetGroupKeyfromLabel (string groupLabel)
        {
            // 1️⃣ Try normal label match
            var match = ValueLabels
                .FirstOrDefault(kv =>
                    kv.Value.Equals(groupLabel, StringComparison.OrdinalIgnoreCase));

            if (!match.Equals(default(KeyValuePair<int, string>)))
            {
                return match.Key;
            }

            // 2️⃣ Try numeric label (e.g. "1", "1.00")
            if (double.TryParse(groupLabel, out double parsedValue))
            {
                // Accept only whole numbers
                if (parsedValue % 1 == 0)
                {
                    int numericKey = (int)parsedValue;

                    if (ValueLabels.ContainsKey(numericKey))
                        return numericKey;
                }
            }

            return -1; // not found
        }

        //public static StatParameter FromSpss(SpssParameter spss)
        //{
        //    return new StatParameter
        //    {
        //        Name = spss.Name,
        //        Label = spss.Label,
        //        Type = spss.Type,
        //        RawValues = new List<string>(spss.Values),
        //        ValueLabels = new Dictionary<int, string>(spss.ValueLabels)
        //    };
        //}




    }

}
