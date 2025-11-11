using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.StatClasses
{
    public class GeneralStatFunctions
    {
        //public Dictionary<string, string> Formatted_Basic_Calculations(List<double> values)
        //{
        //    Dictionary<string, string> ValuesFormatted = new Dictionary<string, string>();
        //    List<double> NewValues = values;

        //    double minValue = NewValues.Min();
        //    double maxValue = NewValues.Max();
        //    double meanValue = NewValues.Average();
        //    double stdDevValue = Math.Sqrt(NewValues.Select(x => Math.Pow(x - meanValue, 2)).Sum() / (NewValues.Count - 1));
        //    double medianValue;
        //    int middleIndex = NewValues.Count / 2;
        //    if (NewValues.Count % 2 == 0)
        //    {
        //        // For even count of elements, take the average of the two middle values
        //        double middleValue1 = NewValues.OrderBy(x => x).ElementAt(middleIndex - 1);
        //        double middleValue2 = NewValues.OrderBy(x => x).ElementAt(middleIndex);
        //        medianValue = (middleValue1 + middleValue2) / 2.0;
        //    }
        //    else
        //    {
        //        // For odd count of elements, directly take the middle value
        //        medianValue = NewValues.OrderBy(x => x).ElementAt(middleIndex);
        //    }

        //    double perc25th = CalculateLowerMedian(NewValues);
        //    double perc75th = CalculateUpperMedian(NewValues);


        //    string formattedMinMax = FormatMinMaxValue(minValue, maxValue);
        //    string formattedMeanStd = FormatMeanStdValue(meanValue, stdDevValue);
        //    string formattedMedian = FormatSingleValue(medianValue);
        //    string formattedIQR = FormatMinMaxValue(perc25th, perc75th);


        //    ValuesFormatted["Min-Max"] = formattedMinMax;
        //    ValuesFormatted["Mean ± StdDev"] = formattedMeanStd;
        //    ValuesFormatted["Median"] = formattedMedian;
        //    ValuesFormatted["IQR"] = formattedIQR;





        //    return ValuesFormatted;


        //}
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

        public static List<SpssSyntaxStat.DescriptiveResult> GetDescriptiveResults(List<string> variableNames, List<string> groupLabels, StatTable statTable , StatParameter groupParameter)
        {

            if (variableNames == null || variableNames.Count == 0)
                return new List<SpssSyntaxStat.DescriptiveResult>();



            var results = variableNames.Select(var => new SpssSyntaxStat.DescriptiveResult { VariableName = var }).ToList();
            int groupcount = statTable.GetGroupParameters().Count;

            foreach (var variableName in variableNames)
            {
                var result = results.First(r => r.VariableName == variableName);

                foreach (var groupLabel in groupLabels)
                {
                    bool hasGroupLabelValues = StatParameter.General_Parameter_Group_Label(variableName, groupLabel, statTable, groupParameter);

                    if (groupcount == 1)
                    {
                        //Grouped Normal

                    }
                    else if(groupcount > 1)
                    {
                        // Relation
                    }


                }
                
            }
            //






            return results;

        }

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


        public double CalculateLowerMedian(List<double> values)
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

        public double CalculateUpperMedian(List<double> values)
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

        public double CalculateMedian(List<double> values)
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

        string FormatMinMaxValue(double minValue, double maxValue)
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

        string FormatMeanStdValue(double meanValue, double stdDevValue)
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

        string FormatSingleValue(double value)
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
