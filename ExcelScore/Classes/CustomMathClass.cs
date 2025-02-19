using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{
    public class CustomMathClass
    {



        public Dictionary<string , string> Basic_Calculations(List<double> values)
        {
            Dictionary<string, string> ValuesFormatted = new Dictionary<string, string>();

            double minValue = values.Min();
            double maxValue = values.Max();
            double meanValue = values.Average();
            double stdDevValue = Math.Sqrt(values.Select(x => Math.Pow(x - meanValue, 2)).Sum() / (values.Count - 1));
            double medianValue;
            int middleIndex = values.Count / 2;
            if (values.Count % 2 == 0)
            {
                // For even count of elements, take the average of the two middle values
                double middleValue1 = values.OrderBy(x => x).ElementAt(middleIndex - 1);
                double middleValue2 = values.OrderBy(x => x).ElementAt(middleIndex);
                medianValue = (middleValue1 + middleValue2) / 2.0;
            }
            else
            {
                // For odd count of elements, directly take the middle value
                medianValue = values.OrderBy(x => x).ElementAt(middleIndex);
            }

            double perc25th = CalculateLowerMedian(values);
            double perc75th = CalculateUpperMedian(values);


            string formattedMinMax = FormatMinMaxValue(minValue, maxValue);
            string formattedMeanStd = FormatMeanStdValue(meanValue, stdDevValue);
            string formattedMedian = FormatSingleValue(medianValue);
            string formattedIQR = FormatMinMaxValue(perc25th, perc75th);


            ValuesFormatted["Min-Max"] = formattedMinMax;
            ValuesFormatted["Mean ± StdDev"] = formattedMeanStd;
            ValuesFormatted["Median"] = formattedMedian;
            ValuesFormatted["IQR"] = formattedIQR;





            return ValuesFormatted; 


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
