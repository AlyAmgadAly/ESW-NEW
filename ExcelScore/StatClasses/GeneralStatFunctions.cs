using DocumentFormat.OpenXml.Spreadsheet;
using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.StatClasses
{
    public class GeneralStatFunctions
    {
        public static Dictionary<string, string> Formatted_Basic_Calculations(List<double> values)
        {
            Dictionary<string, string> ValuesFormatted = new Dictionary<string, string>();
            List<double> NewValues = new List<double>(values);


            double minValue = NewValues.Min();
            double maxValue = NewValues.Max();
            double meanValue = NewValues.Average();
            double stdDevValue = Math.Sqrt(NewValues.Select(x => Math.Pow(x - meanValue, 2)).Sum() / (NewValues.Count - 1));
            double medianValue;
            int middleIndex = NewValues.Count / 2;
            if (NewValues.Count % 2 == 0)
            {
                // For even count of elements, take the average of the two middle values
                double middleValue1 = NewValues.OrderBy(x => x).ElementAt(middleIndex - 1);
                double middleValue2 = NewValues.OrderBy(x => x).ElementAt(middleIndex);
                medianValue = (middleValue1 + middleValue2) / 2.0;
            }
            else
            {
                // For odd count of elements, directly take the middle value
                medianValue = NewValues.OrderBy(x => x).ElementAt(middleIndex);
            }

            double perc25th = CalculateLowerMedian(NewValues);
            double perc75th = CalculateUpperMedian(NewValues);


            string formattedMinMax = FormatMinMaxValue(minValue, maxValue);
            string formattedMeanStd = FormatMeanStdValue(meanValue, stdDevValue);
            string formattedMedian = FormatSingleValue(medianValue);
            string formattedIQR = FormatMinMaxValue(perc25th, perc75th);


            ValuesFormatted["N"] = NewValues.Count.ToString();
            ValuesFormatted["Min"] = minValue.ToString();
            ValuesFormatted["Max"] = maxValue.ToString();
            ValuesFormatted["Mean"] = meanValue.ToString();
            ValuesFormatted["StdDev"] = stdDevValue.ToString();
            ValuesFormatted["perc25th"] = perc25th.ToString();
            ValuesFormatted["perc75th"] = perc75th.ToString();


            ValuesFormatted["Min-Max"] = formattedMinMax;
            ValuesFormatted["Mean ± StdDev"] = formattedMeanStd;
            ValuesFormatted["Median"] = formattedMedian;
            ValuesFormatted["IQR"] = formattedIQR;





            return ValuesFormatted;


        }

        public static string BuildLowCountFilterSyntax(StatTable table, string groupVar, List<int> validGroups)
        {
            var groupCondition = string.Join("  |  ", validGroups.Select(v => $"{groupVar} = {v}"));
            var filterLabel = $"{groupVar} = " + string.Join("  |  ", validGroups);

            // Base condition from user's select (IF or filter variable)
            string selectCond = null;
            if (!string.IsNullOrWhiteSpace(table.SelectStatement))
            {
                if (table.SelectIF)
                {
                    // User typed an IF expression, e.g. Age = 1 & Sex = 2
                    selectCond = table.SelectStatement;
                }
                else
                {
                    // User chose an existing filter variable name, e.g. MyFilterVar
                    // We interpret that as "this var = 1" being selected
                    selectCond = $"{table.SelectStatement} = 1";
                }
            }

            // Combine: (user select) AND (groupVar in validGroups) if select exists
            string combinedCondition = string.IsNullOrWhiteSpace(selectCond)
                ? groupCondition
                : $"({selectCond}) & ({groupCondition})";

            var sb = new StringBuilder();

            // Make sure we start from all cases for this combined filter
            sb.AppendLine("USE ALL.");
            sb.AppendLine($"COMPUTE filter_$=({combinedCondition}).");
            sb.AppendLine($"VARIABLE LABEL filter_$ '{filterLabel} (FILTER)'.");
            sb.AppendLine("VALUE LABELS filter_$  0 'Not Selected' 1 'Selected'.");
            sb.AppendLine("FORMAT filter_$ (f1.0).");
            sb.AppendLine("FILTER BY filter_$.");
            sb.AppendLine("EXECUTE.");

            return sb.ToString();
        }
        public static string BuildSelectSyntax(StatTable table)
        {
            if (string.IsNullOrWhiteSpace(table.SelectStatement))
                return string.Empty;

            var sb = new StringBuilder();

            sb.AppendLine("USE ALL.");

            if (table.SelectIF)
            {
                // User typed an IF condition: Age = 1 & Sex = 2
                sb.AppendLine($"COMPUTE filter_$=({table.SelectStatement}).");
                sb.AppendLine($"VARIABLE LABEL filter_$ '{table.SelectStatement} (FILTER)'.");
                sb.AppendLine("VALUE LABELS filter_$  0 'Not Selected' 1 'Selected'.");
                sb.AppendLine("FORMAT filter_$ (f1.0).");
                sb.AppendLine("FILTER BY filter_$.");
            }
            else
            {
                // User chose an existing filter variable name
                sb.AppendLine($"FILTER BY {table.SelectStatement}.");
            }

            sb.AppendLine("EXECUTE.");
            return sb.ToString();
        }

        public static void ParseUnifiedOutput_Nominal(List<StatTable> tables)
        {
            foreach (var table in tables)
            {
                // collect scale parameter names for this table
                var nominalParameters = table.GetNonGroupParameters()
                    .Where(p => string.Equals(p.Type, "Nominal", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (!nominalParameters.Any())
                    continue;

                var groupParams = table.GetGroupParameters();

                // If there are no group parameters, still call parser once with empty labels
                if (!groupParams.Any())
                {
                    return;
                }

                foreach (var g in groupParams)
                {
                    var groupLabels = new List<string>();

                    if (g.ValueLabels != null && g.ValueLabels.Count > 0)
                    {
                        foreach (var kvp in g.ValueLabels
                                             .OrderBy(k => k.Key)) // sort by numeric key ascending
                        {
                            var key = kvp.Key;
                            var label = kvp.Value;

                            string chosen;

                            // use label only if it's not identical to the numeric key
                            if (!string.Equals(label, key.ToString(), StringComparison.OrdinalIgnoreCase))
                                chosen = label;
                            else
                                chosen = ((double)key).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);

                            groupLabels.Add(chosen);
                        }
                    }

                }
            }
        }
        public static void ParseUnifiedOutput_Descriptives(List<StatTable> tables)
        {
            foreach (var table in tables)
            {
                // collect scale parameter names for this table
                var scaleParamNames = table.GetNonGroupParameters()
                    .Where(p => string.Equals(p.Type, "Scale", StringComparison.OrdinalIgnoreCase))
                    .Select(p => p.Name)
                    .ToList();

                if (!scaleParamNames.Any())
                    continue;

                var groupParams = table.GetGroupParameters();

                // If there are no group parameters, still call parser once with empty labels
                if (!groupParams.Any())
                {
                    return;
                }

                // For each group parameter separately, parse and assign
                foreach (var g in groupParams)
                {
                    var groupLabels = new List<string>();

                    if (g.ValueLabels != null && g.ValueLabels.Count > 0)
                    {
                        foreach (var kvp in g.ValueLabels
                                             .OrderBy(k => k.Key)) // sort by numeric key ascending
                        {
                            var key = kvp.Key;
                            var label = kvp.Value;

                            string chosen;

                            // use label only if it's not identical to the numeric key
                            if (!string.Equals(label, key.ToString(), StringComparison.OrdinalIgnoreCase))
                                chosen = label;
                            else
                                chosen = ((double)key).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);

                            groupLabels.Add(chosen);
                        }
                    }
                    var descResultsForThisGroup = GetDescriptiveResults(scaleParamNames, groupLabels, table, g);

                    // assign each result back into its StatParameter under the group name
                    foreach (var desc in descResultsForThisGroup)
                    {
                        var param = table.GetParameterByName(desc.VariableName);
                        if (param == null) continue;

                        // ensure the per-group container exists on StatParameter
                        if (param.DescriptiveStatsByGroup == null)
                            param.DescriptiveStatsByGroup = new Dictionary<string, SpssSyntaxStat.DescriptiveResult>();

                        // store per group variable (keyed by group parameter name)
                        param.DescriptiveStatsByGroup[g.Name] = desc;

                        // For backward compatibility, if the single DescriptiveStats is empty, set it
                        if (param.DescriptiveStats == null)
                            param.DescriptiveStats = desc;
                    }

                    //ParsePercentile_Output(outputText, scaleParamNames, groupLabels, table, g);
                }


            }

        }
        // Evaluates whether a given row in the table passes the Select condition (if any)
        private static bool RowPassesSelect(StatTable table, int rowIndex)
        {
            // No select → always include
            if (string.IsNullOrWhiteSpace(table.SelectStatement))
                return true;

            // CASE 1: user wrote an IF-like condition, e.g. "Sex = 1 & Age < 30"
            if (table.SelectIF)
                return EvaluateSelectExpression(table, rowIndex, table.SelectStatement);

            // CASE 2: user chose a filter variable name (e.g. "MyFilterVar")
            // Treat it as SPSS filter: 1 = Selected, 0 = Not selected
            var filterParam = table.GetParameterByName(table.SelectStatement);
            if (filterParam == null || rowIndex >= filterParam.RawValues.Count)
                return true; // be permissive if filter variable is missing

            var v = filterParam.RawValues[rowIndex]?.Trim();
            return v == "1" || v == "1.0" || v == "1.00";
        }

        /// <summary>
        /// Evaluates a simple SPSS-like condition (e.g. "Sex = 1 & Age < 30")
        /// on the given row using the table's parameters.
        /// Supports: =, <>, >, <, >=, <= and logical & / && (AND), | / || (OR).
        /// </summary>
        private static bool EvaluateSelectExpression(StatTable table, int rowIndex, string expression)
        {
            // Get ALL SPSS parameters from the form, not just this table's
            var spssParams = FormDataTransfer.Get<List<StatParameter>>("SPSS_Parameters")
                           ?? new List<StatParameter>();

            // Build a dictionary of variable name -> raw string value at this row
            var rowValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var p in spssParams)
            {
                string value = (rowIndex < p.RawValues.Count) ? p.RawValues[rowIndex] : null;
                rowValues[p.Name] = value;
            }

            

            // Normalize logical operators a bit
            string expr = expression.Replace("&&", "&").Replace("||", "|");

            // Split by OR
            var orClauses = expr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var orClause in orClauses)
            {
                // Each OR clause is a series of AND conditions
                var andConditions = orClause.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
                bool allAndTrue = true;

                foreach (var condRaw in andConditions)
                {
                    string cond = condRaw.Trim().Trim('(', ')');
                    if (!EvaluateSimpleCondition(cond, rowValues))
                    {
                        allAndTrue = false;
                        break;
                    }
                }

                if (allAndTrue)
                    return true; // any OR clause true → whole expression true
            }

            return false;
        }

        private static bool EvaluateSimpleCondition(string condition, Dictionary<string, string> rowValues)
        {
            if (string.IsNullOrWhiteSpace(condition))
                return true;

            // Supported operators, longest first
            string[] ops = { ">=", "<=", "<>", "=", ">", "<" };

            string op = null;
            int opPos = -1;

            foreach (var candidate in ops)
            {
                opPos = condition.IndexOf(candidate, StringComparison.Ordinal);
                if (opPos >= 0)
                {
                    op = candidate;
                    break;
                }
            }

            if (op == null)
                return true; // if we can't parse, don't filter out

            string left = condition.Substring(0, opPos).Trim();
            string right = condition.Substring(opPos + op.Length).Trim();

            if (!rowValues.TryGetValue(left, out var rawLeft) || string.IsNullOrWhiteSpace(rawLeft) || rawLeft == ".")
                return false;

            // Try numeric comparison first
            bool leftIsNum = double.TryParse(rawLeft, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double leftNum);
            bool rightIsNum = double.TryParse(right, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double rightNum);

            if (leftIsNum && rightIsNum)
            {
                switch (op)
                {
                    case "=": return leftNum == rightNum;
                    case "<>": return leftNum != rightNum;
                    case ">": return leftNum > rightNum;
                    case "<": return leftNum < rightNum;
                    case ">=": return leftNum >= rightNum;
                    case "<=": return leftNum <= rightNum;
                }
            }
            else
            {
                // String comparison for "=" and "<>"
                switch (op)
                {
                    case "=":
                        return string.Equals(
                            rawLeft.Trim(),
                            right.Trim().Trim('\'', '"'),
                            StringComparison.OrdinalIgnoreCase);

                    case "<>":
                        return !string.Equals(
                            rawLeft.Trim(),
                            right.Trim().Trim('\'', '"'),
                            StringComparison.OrdinalIgnoreCase);

                    default:
                        return false; // >,<,>=,<= on non-numeric → treat as false
                }
            }

            return false;
        }
        public static List<SpssSyntaxStat.DescriptiveResult> GetDescriptiveResults(
    List<string> variableNames,
    List<string> groupLabels,
    StatTable statTable,
    StatParameter groupParameter)
        {
            GeneralStatFunctions generalStatFunctions = new GeneralStatFunctions();

            if (variableNames == null || variableNames.Count == 0)
                return new List<SpssSyntaxStat.DescriptiveResult>();

            var results = variableNames
                .Select(var => new SpssSyntaxStat.DescriptiveResult { VariableName = var })
                .ToList();

            int groupcount = statTable.GetGroupParameters().Count;

            foreach (var variableName in variableNames)
            {
                Dictionary<string, string> StatLabelValues_Total = new Dictionary<string, string>();

                var CurrentParameter = statTable.GetParameterByName(variableName);
                var result = results.First(r => r.VariableName == variableName);

                foreach (var groupLabel in groupLabels)
                {
                    bool hasGroupLabelValues = StatParameter.General_Parameter_Group_Label(
                        variableName,
                        groupLabel,
                        statTable,
                        groupParameter);

                    if (!hasGroupLabelValues)
                        continue;

                    int CurrentGroupKey = groupParameter.GetGroupKeyfromLabel(groupLabel);
                    if (CurrentGroupKey == -1)
                        continue;

                    // === NEW: build filtered list using raw values, group, and table.SelectIF / SelectStatement ===
                    var filteredValues = new List<double>();

                    // assume all parameters share the same row count as the group parameter
                    int rowCount = groupParameter.RawValues.Count;

                    for (int row = 0; row < rowCount; row++)
                    {
                        // 1) value of the current parameter
                        var rawVal = CurrentParameter.RawValues[row];
                        if (string.IsNullOrWhiteSpace(rawVal) || rawVal == ".")
                            continue;

                        // 2) group value must be non-missing and equal to CurrentGroupKey
                        var rawGroup = groupParameter.RawValues[row];
                        if (string.IsNullOrWhiteSpace(rawGroup) || rawGroup == ".")
                            continue;

                        if (!double.TryParse(
                                rawGroup,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out double gVal))
                            continue;

                        if (Math.Abs(gVal - CurrentGroupKey) > double.Epsilon)
                            continue;

                        // 3) apply the select condition on this row (e.g. Sex = 1 & Age < 30)
                        if (!RowPassesSelect(statTable, row))
                            continue;

                        // 4) add numeric parameter value
                        if (double.TryParse(
                                rawVal,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out double v))
                        {
                            filteredValues.Add(v);
                        }
                    }

                    if (filteredValues.Count == 0)
                        continue;

                    Dictionary<string, string> StatLabelValues =
                        GeneralStatFunctions.Formatted_Basic_Calculations(filteredValues);

                    if (!result.Stats_Groups.ContainsKey(groupLabel))
                        result.Stats_Groups[groupLabel] = new Dictionary<string, string>();

                    foreach (var kvp in StatLabelValues)
                    {
                        string Statlabel = kvp.Key;
                        string Statvalue = kvp.Value;

                        if (!result.Stats_Groups[groupLabel].ContainsKey(Statlabel))
                            result.Stats_Groups[groupLabel][Statlabel] = Statvalue;
                    }
                }

                // Total stats: still based on all non-empty values (no select) –
                // if you want totals ALSO filtered by select, we can change this too.
                StatLabelValues_Total = GeneralStatFunctions.Formatted_Basic_Calculations(CurrentParameter.RawNonEmptyValues);

                foreach (var kvp in StatLabelValues_Total)
                {
                    string Statlabel = kvp.Key;
                    string Statvalue = kvp.Value;

                    if (!result.Stats_Total.ContainsKey(Statlabel))
                        result.Stats_Total[Statlabel] = Statvalue;
                }
            }

            return results;
        }

        //before update


        //public static List<SpssSyntaxStat.DescriptiveResult> GetDescriptiveResults(List<string> variableNames, List<string> groupLabels, StatTable statTable, StatParameter groupParameter)
        //{
        //    GeneralStatFunctions generalStatFunctions = new GeneralStatFunctions();

        //    if (variableNames == null || variableNames.Count == 0)
        //        return new List<SpssSyntaxStat.DescriptiveResult>();



        //    var results = variableNames.Select(var => new SpssSyntaxStat.DescriptiveResult { VariableName = var }).ToList();
        //    int groupcount = statTable.GetGroupParameters().Count;

        //    foreach (var variableName in variableNames)
        //    {
        //        Dictionary<string, string> StatLabelValues_Total = new Dictionary<string, string>();
        //        //Only for testing
        //        var CurrentParameter = statTable.GetParameterByName(variableName);

        //        var result = results.First(r => r.VariableName == variableName);

        //        foreach (var groupLabel in groupLabels)
        //        {

        //            bool hasGroupLabelValues = StatParameter.General_Parameter_Group_Label(variableName, groupLabel, statTable, groupParameter);


        //            if (hasGroupLabelValues)
        //            {
        //                int CurrentGroupKey = groupParameter.GetGroupKeyfromLabel(groupLabel);
        //                if (CurrentGroupKey != -1)
        //                {
        //                    //setup filter here for select




        //                    Dictionary<string, string> StatLabelValues = new Dictionary<string, string>();
        //                    if (groupcount == 1)
        //                    {
        //                        StatLabelValues = GeneralStatFunctions.Formatted_Basic_Calculations(CurrentParameter.GroupedParameterValues[CurrentGroupKey]);
        //                    }
        //                    else if (groupcount > 1)
        //                    {
        //                        StatLabelValues = GeneralStatFunctions.Formatted_Basic_Calculations(CurrentParameter.GroupedParameterValuesRelation[groupParameter.Name][CurrentGroupKey]);
        //                    }

        //                    if (!result.Stats_Groups.ContainsKey(groupLabel))
        //                        result.Stats_Groups[groupLabel] = new Dictionary<string, string>();

        //                    foreach (var kvp in StatLabelValues)
        //                    {
        //                        string Statlabel = kvp.Key;
        //                        string Statvalue = kvp.Value;

        //                        if (!result.Stats_Groups[groupLabel].ContainsKey(Statlabel))
        //                            result.Stats_Groups[groupLabel][Statlabel] = Statvalue;
        //                    }


        //                }
        //            }


        //        }

        //        StatLabelValues_Total = GeneralStatFunctions.Formatted_Basic_Calculations(CurrentParameter.RawNonEmptyValues);


        //        foreach (var kvp in StatLabelValues_Total)
        //        {
        //            string Statlabel = kvp.Key;
        //            string Statvalue = kvp.Value;

        //            if (!result.Stats_Total.ContainsKey(Statlabel))
        //                result.Stats_Total[Statlabel] = Statvalue;
        //        }



        //    }
        //    //

        //    return results;

        //    // *********** The way we strore in the class
        //    //if (ProceedLabel == "Total")
        //    //{
        //    //    if (!result.Stats_Total.ContainsKey(statLabel))
        //    //        result.Stats_Total[statLabel] = value;
        //    //}
        //    //else
        //    //{
        //    //    if (!result.Stats_Groups.ContainsKey(ProceedLabel))
        //    //        result.Stats_Groups[ProceedLabel] = new Dictionary<string, string>();

        //    //    if (!result.Stats_Groups[ProceedLabel].ContainsKey(statLabel))
        //    //        result.Stats_Groups[ProceedLabel][statLabel] = value;
        //    //}
        //    //**************************
        //}





        public Dictionary<string, string> Basic_Calculations(List<double> values)
        {
            Dictionary<string, string> Statlabel_Value = new Dictionary<string, string>();
            List<double> NewValues = values;

            double minValue = NewValues.Min();
            double maxValue = NewValues.Max();
            double meanValue = NewValues.Average();
            double stdDevValue = Math.Sqrt(NewValues.Select(x => Math.Pow(x - meanValue, 2)).Sum() / (NewValues.Count - 1));
            double medianValue;
            int middleIndex = NewValues.Count / 2;
            if (NewValues.Count % 2 == 0)
            {
                // For even count of elements, take the average of the two middle values
                double middleValue1 = NewValues.OrderBy(x => x).ElementAt(middleIndex - 1);
                double middleValue2 = NewValues.OrderBy(x => x).ElementAt(middleIndex);
                medianValue = (middleValue1 + middleValue2) / 2.0;
            }
            else
            {
                // For odd count of elements, directly take the middle value
                medianValue = NewValues.OrderBy(x => x).ElementAt(middleIndex);
            }

            double perc25th = CalculateLowerMedian(NewValues);
            double perc75th = CalculateUpperMedian(NewValues);



            Statlabel_Value["N"] = minValue.ToString();
            Statlabel_Value["Min"] = maxValue.ToString();
            Statlabel_Value["Max"] = meanValue.ToString();
            Statlabel_Value["Mean"] = stdDevValue.ToString();
            Statlabel_Value["Median"] = medianValue.ToString();
            Statlabel_Value["25th"] = perc25th.ToString();
            Statlabel_Value["75th"] = perc75th.ToString();


            return Statlabel_Value;


        }


        public static double CalculateLowerMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;


            if (n % 2 == 0)
            {
                int middle = n / 2;
                // Even number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(0, middle));
            }
            else
            {
                //The fix was here we get same range from zero to middle but middle index is different

                int middle = (n + 1) / 2;

                return CalculateMedian(values.GetRange(0, middle));
            }

        }

        public static double CalculateUpperMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;


            if (n % 2 == 0)
            {
                int middle = n / 2;
                // Even number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle));
            }
            else if (n == 1)
            {
                return values[0];
            }
            else if (n == 3)
            {
                return (values[1] + values[2]) / 2.0;
            }

            else
            {
                int middle = (n + 1) / 2;
                // Odd number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle - 1));

            }
        }

        public static double CalculateMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;
            int middle = n / 2;

            if (n == 1)
            {
                return values[0];
            }
            else if (n % 2 == 0)
            {

                // Even number of elements, average the middle two
                return (values[middle - 1] + values[middle]) / 2.0;
            }

            else
            {

                // Odd number of elements, return the middle one
                return values[middle];
            }
        }

        static string FormatMinMaxValue(double minValue, double maxValue)
        {

            string minValueString = "";
            string maxValueString = "";



            minValueString = minValue.ToString("0.00");
            maxValueString = maxValue.ToString("0.00");


            // Check if one value has three numbers (having three digits after the decimal point)
            bool isMinThreeNumbers = minValueString.Split('.')[0].Length == 3;
            bool isMaxThreeNumbers = maxValueString.Split('.')[0].Length == 3;

            // If one value is three numbers, round it to one decimal place
            if (isMinThreeNumbers)
            {
                minValueString = Math.Round(minValue, 1).ToString("0.0");
            }
            if (isMaxThreeNumbers)
            {
                maxValueString = Math.Round(maxValue, 1).ToString("0.0");
            }

            // Remove trailing zeroes if there are two zeroes after the decimal point
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

       static string FormatMeanStdValue(double meanValue, double stdDevValue)
        {
            string meanValueString = meanValue.ToString("0.00");
            string stdDevValueString = stdDevValue.ToString("0.00");


            bool isMeanThreeNumbers = meanValueString.Split('.')[0].Length == 3;
            bool isStdThreeNumbers = stdDevValueString.Split('.')[0].Length == 3;

            // Check if the values are three numbers (having three digits after the decimal point)
            bool isThreeNumbers = meanValueString.Split('.')[0].Length == 3 && stdDevValueString.Split('.')[0].Length == 3;

            // If it's three numbers, round to one decimal place
            if (isMeanThreeNumbers)
            {
                meanValueString = Math.Round(meanValue, 1).ToString("0.0");

            }
            if (isStdThreeNumbers)
            {
                stdDevValueString = Math.Round(stdDevValue, 1).ToString("0.0");
            }

            // Remove trailing zeroes if there are two zeroes after the decimal point
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

        static string FormatSingleValue(double value)
        {
            string valueString = value.ToString("0.00");

            // Check if the value is three numbers (having three digits after the decimal point)
            bool isThreeNumbers = valueString.Split('.')[0].Length == 3;

            // If it's three numbers, round to one decimal place
            if (isThreeNumbers)
            {
                valueString = Math.Round(value, 1).ToString("0.0");
            }

            // Remove trailing zeroes if there are two zeroes after the decimal point
            if (valueString.EndsWith(".00"))
            {
                valueString = valueString.Substring(0, valueString.Length - 1);
            }

            return valueString;
        }
        
    }
}
