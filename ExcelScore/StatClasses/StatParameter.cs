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
        public Dictionary<int, string> ValueLabels { get; set; } = new(); // e.g., 1 = Male

        //manual
        public Dictionary<double, double> ValueFrequencies { get; set; } = new();
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



        // --- Statistics Results (parsed or calculated) ---
        public SPSSUnifiedRunner.DescriptiveResult DescriptiveStats { get; set; }
        public SPSSUnifiedRunner.PercentileResult PercentileStats { get; set; }
        public SPSSUnifiedRunner.CrosstabBlock ChiSquareBlock { get; set; }
        public List<string> Test_PValues { get; set; } = new(); // t, U, F, H results
        public List<string> PostHocResults { get; set; } = new();


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
