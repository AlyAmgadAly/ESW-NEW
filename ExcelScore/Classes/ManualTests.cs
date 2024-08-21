using MathNet.Numerics.Distributions;
using MathNet.Numerics.Statistics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Accord.Statistics.Testing;
using System.Windows.Forms;
using Python.Runtime;
using CenterSpace.NMath.Core;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelScore.Classes
{
   

    public class ManualTests
    {
        // chi square : 

        public string[] getchi(Parameter parameter)
        {

            if (parameter.IsGroup || parameter.NominalOrScale != "Nominal")
            {
                return new string[] { "Parameter is not relevant for chi-square test", "" };
            }

            // Identify unique categories across all groups
            var allValues = parameter.GroupedParameterValues.Values.SelectMany(x => x).Distinct().ToList();

            // Prepare contingency table
            var contingencyTable = new List<List<int>>();

            //contingencyTable.Add(new List<int> { 0, 0 });
            //contingencyTable.Add(new List<int> { 0, 1 });
            //contingencyTable.Add(new List<int> { 1, 0 });
            //contingencyTable.Add(new List<int> { 1, 1 });


            foreach (var value in allValues)
            {
                var categoryCounts = new List<int>();
                foreach (var groupData in parameter.GroupedParameterValues.Values)
                {
                    categoryCounts.Add(groupData.Count(x => x == value));
                }
                contingencyTable.Add(categoryCounts);
            }

            int a = 0;
            int b = 0;
            int c = 0;
            int d =0;

            try
            {
               a = contingencyTable[0][0];
            }
            catch(Exception) 
            {
                a = 0;
            }


            try
            {
                b = contingencyTable[0][1];
            }
            catch(Exception) 
            {
                b = 0;
            }

            try
            {
                c = contingencyTable[1][0];
            }
            catch (Exception)
            {
                c = 0;
            }

            try
            {
                d = contingencyTable[1][1];
            }
            catch(Exception)
            { d = 0; }  


            //MessageBox.Show(a.ToString());
            //MessageBox.Show(b.ToString());

            // MessageBox.Show(a.ToString());
            // MessageBox.Show(b.ToString());
            // MessageBox.Show(c.ToString());
            // MessageBox.Show(d.ToString());

            double chiSquareStatistic = CalculateChiSquare(a, b, c, d);
            double pValue;

            if (ShouldUseFisher(a, b, c, d))
            {
                pValue = FisherExactTest(a, b, c, d);
                parameter.Isfisher = true;
            }
            else
            {
                pValue = ChiSquarePValue(chiSquareStatistic, 1); // df = 1 for a 2x2 table
            }


            chiSquareStatistic = Math.Round(chiSquareStatistic, 3);
            pValue = Math.Round(pValue, 3);

            //MessageBox.Show(chiSquareStatistic.ToString());
            //MessageBox.Show(pValue.ToString());

            // Format the p-value
            string chiSquareString = chiSquareStatistic.ToString("0.000");
            string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

            //MessageBox.Show(chiSquareString);
            //MessageBox.Show(pValueString);


            return new string[] { chiSquareString, pValueString };
            
        }

        static double CalculateChiSquare(double a, double b, double c, double d)
        {
            double row1Total = a + b;
            double row2Total = c + d;
            double col1Total = a + c;
            double col2Total = b + d;
            double grandTotal = a + b + c + d;

            double expectedA = (row1Total * col1Total) / grandTotal;
            double expectedB = (row1Total * col2Total) / grandTotal;
            double expectedC = (row2Total * col1Total) / grandTotal;
            double expectedD = (row2Total * col2Total) / grandTotal;

            double chiSquare = ((a - expectedA) * (a - expectedA) / expectedA) +
                               ((b - expectedB) * (b - expectedB) / expectedB) +
                               ((c - expectedC) * (c - expectedC) / expectedC) +
                               ((d - expectedD) * (d - expectedD) / expectedD);

            return chiSquare;
        }
        static double ChiSquarePValue(double chiSquareStatistic, int df)
        {
            var chiSquareDistribution = new ChiSquared(df);
            double pValue = 1 - chiSquareDistribution.CumulativeDistribution(chiSquareStatistic);
            return pValue;
        }
        static bool ShouldUseFisher(int a, int b, int c, int d)
        {
            double row1Total = a + b;
            double row2Total = c + d;
            double col1Total = a + c;
            double col2Total = b + d;
            double grandTotal = a + b + c + d;

            double expectedA = (row1Total * col1Total) / grandTotal;
            double expectedB = (row1Total * col2Total) / grandTotal;
            double expectedC = (row2Total * col1Total) / grandTotal;
            double expectedD = (row2Total * col2Total) / grandTotal;

            return expectedA < 5 || expectedB < 5 || expectedC < 5 || expectedD < 5;
        }

        static double FisherExactTest(int a, int b, int c, int d)
        {


            double pvalue = NMathFunctions.FishersExactTest(a, b, c, d, HypothesisType.TwoSided);



            return pvalue;
        }



        // Student T test 

        public string[] StudentT_Unpaired(List<double> dataGroup1, List<double> dataGroup2)
        {


            bool EqualVariance;
            double[] dataG1_Array = dataGroup1.ToArray();
            double[] dataG2_Array = dataGroup2.ToArray();


            // Combine the groups into a 2D array
            double[][] data = new double[][] { dataG1_Array, dataG2_Array };

            // Create and run the Levene's Test
            var leveneTest = new Accord.Statistics.Testing.LeveneTest(data);

            // Obtain the p-value
            double pValue = leveneTest.PValue;


            if (pValue < 0.05)
            {

                EqualVariance = false;
            }
            else
            {

                EqualVariance = true;
            }

            var TstudentTest = new Accord.Statistics.Testing.TwoSampleTTest(dataG1_Array, dataG2_Array, assumeEqualVariances: EqualVariance);



            double Ttestt = TstudentTest.Statistic;
            double pvalueT = TstudentTest.PValue;


            Ttestt = Math.Round(Ttestt, 3);
            pvalueT = Math.Round(pvalueT, 3);

            string TtestString = Ttestt.ToString("0.000");
            TtestString = TtestString.Replace("-", "");

           // MessageBox.Show(EqualVariance.ToString());
            //MessageBox.Show("T test p " + pvalueT.ToString());

            // Format the p-value
            string pValueString = pvalueT < 0.001 ? "<0.001" : pvalueT.ToString("0.000");

            return new string[] { TtestString, pValueString };

        }

        // U Test
        public string[] UTest(List<double> AdataGroup1, List<double> AdataGroup2)
        {
            double[] group1 = AdataGroup1.ToArray();
            double[] group2 = AdataGroup2.ToArray();


            int totalcount = group1.Length + group2.Length;

            bool CountBool = false;

            if(totalcount > 40)
            {
                CountBool = false;
            }
            else if(totalcount <= 40)
            {
                CountBool = true;
            }
            
            
            var UTest = new Accord.Statistics.Testing.MannWhitneyWilcoxonTest(group1, group2, exact: CountBool);


            double UStat = CalculateUStatistic(group1, group2);
            //double UStat = UTest.Statistic;
            double Upvalue = UTest.PValue;


            UStat = Math.Round(UStat, 3);
            Upvalue = Math.Round(Upvalue, 3);
            //MessageBox.Show("U test : " + UStat.ToString());
           // MessageBox.Show("U test p " + Upvalue.ToString());

            string UtestString = UStat.ToString("0.000");
            string pValueString = Upvalue < 0.001 ? "<0.001" : Upvalue.ToString("0.000");

            string[] TestValue = new string[] { UtestString, pValueString };

            return TestValue;
        }

        static double CalculateUStatistic(double[] sample1, double[] sample2)
        {
            // Combine both samples
            var allData = sample1.Concat(sample2).ToArray();
            var sample1Count = sample1.Length;
            var sample2Count = sample2.Length;

            // Rank data
            var ranks = RankData(allData);

            // Calculate sum of ranks for each sample
            double rankSum1 = sample1.Sum(x => ranks[x]);
            double rankSum2 = sample2.Sum(x => ranks[x]);

            // Calculate U statistic
            double u1 = rankSum1 - sample1Count * (sample1Count + 1) / 2.0;
            double u2 = rankSum2 - sample2Count * (sample2Count + 1) / 2.0;

            // Return the smaller U value
            return Math.Min(u1, u2);
        }

        static Dictionary<double, double> RankData(double[] data)
        {
            var sortedData = data.OrderBy(x => x).ToArray();
            var ranks = new Dictionary<double, double>();

            double rankSum = 1;
            int tieCount = 1;

            for (int i = 1; i < sortedData.Length; i++)
            {
                if (sortedData[i] == sortedData[i - 1])
                {
                    tieCount++;
                    rankSum += i + 1;
                }
                else
                {
                    if (tieCount > 1)
                    {
                        double averageRank = rankSum / tieCount;
                        for (int j = i - tieCount; j < i; j++)
                        {
                            ranks[sortedData[j]] = averageRank;
                        }
                    }
                    else
                    {
                        ranks[sortedData[i - 1]] = i;
                    }
                    rankSum = i + 1;
                    tieCount = 1;
                }
            }

            // Handle the last tie or single element
            if (tieCount > 1)
            {
                double averageRank = rankSum / tieCount;
                for (int j = sortedData.Length - tieCount; j < sortedData.Length; j++)
                {
                    ranks[sortedData[j]] = averageRank;
                }
            }
            else
            {
                ranks[sortedData[sortedData.Length - 1]] = sortedData.Length;
            }

            return ranks;
        }


        //Correlation


        public string[] CalculateCorrelation(double[] x, double[] y, string correlationType)
        {

            string[] result = new string[2];
            double[] tempResult = new double[2];
            if (correlationType == "Pearson")
            {
                // Calculate Pearson's correlation coefficient
                tempResult[0] = PearsonCorrelation(x, y);
                tempResult[1] = PearsonPValue(tempResult[0], Math.Min(x.Length, y.Length));
            }
            else if (correlationType == "Spearman")
            {
                // Calculate Spearman's correlation coefficient
                tempResult[0] = SpearmanRankCorrelation(x, y);
                tempResult[1] = SpearmanPValue(tempResult[0], Math.Min(x.Length, y.Length));
            }

            //MessageBox.Show(tempResult[0].ToString());
            //MessageBox.Show(tempResult[1].ToString());


            tempResult[0] = Math.Round(tempResult[0] , 3);
            tempResult[1] = Math.Round(tempResult[1], 3);

            

            result[0] = tempResult[0].ToString("0.000");
            result[1] = tempResult[1] < 0.001 ? "<0.001" : tempResult[1].ToString("0.000");

            return result;
        }

        static double PearsonCorrelation(double[] x, double[] y)
        {
            // Determine the length of the shorter array
            int length = Math.Min(x.Length, y.Length);

            if (length == 0)
            {
                throw new ArgumentException("Both arrays must have at least one element.");
            }

            double sumX = 0, sumY = 0, sumX2 = 0, sumY2 = 0, sumXY = 0;

            for (int i = 0; i < length; i++)
            {
                sumX += x[i];
                sumY += y[i];
                sumX2 += x[i] * x[i];
                sumY2 += y[i] * y[i];
                sumXY += x[i] * y[i];
            }

            double numerator = (length * sumXY) - (sumX * sumY);
            double denominator = Math.Sqrt((length * sumX2 - sumX * sumX) * (length * sumY2 - sumY * sumY));

            if (denominator == 0)
            {
                return 0; // Avoid division by zero
            }

            return numerator / denominator;
        }
        static double PearsonPValue(double r, int n)
        {
            if (n <= 2)
            {
                throw new ArgumentException("Sample size must be greater than 2.");
            }

            double t = r * Math.Sqrt((n - 2) / (1 - r * r));
            double df = n - 2;

            // Create a Student's t-distribution object with df degrees of freedom
            var studentT = new Accord.Statistics.Distributions.Univariate.TDistribution(df);



            // Calculate the two-tailed p-value
            return 2 * (1 - studentT.DistributionFunction(Math.Abs(t)));
        }


        public static double SpearmanRankCorrelation(double[] x, double[] y)
        {
            int length = Math.Min(x.Length, y.Length);

            if (length <= 1)
            {
                throw new ArgumentException("Both arrays must have at least two elements.");
            }

            // Rank the values
            double[] rankX = Rank(x);
            double[] rankY = Rank(y);

            // Compute Spearman's rank correlation coefficient
            double sumD2 = 0;
            for (int i = 0; i < length; i++)
            {
                double d = rankX[i] - rankY[i];
                sumD2 += d * d;
            }

            double spearmanRho = 1 - (6 * sumD2) / (length * (length * length - 1));
            return spearmanRho;
        }

        public static double[] Rank(double[] values)
        {
            int length = values.Length;
            double[] ranks = new double[length];
            var valueRankPairs = values.Select((v, i) => new { Value = v, Index = i })
                                       .OrderBy(vr => vr.Value)
                                       .ToList();

            int rank = 1;
            for (int i = 0; i < length; i++)
            {
                if (i > 0 && valueRankPairs[i].Value == valueRankPairs[i - 1].Value)
                {
                    ranks[valueRankPairs[i].Index] = ranks[valueRankPairs[i - 1].Index];
                }
                else
                {
                    double sumRanks = rank;
                    int count = 1;

                    // Check for ties
                    while (i + count < length && valueRankPairs[i + count].Value == valueRankPairs[i].Value)
                    {
                        sumRanks += rank + count;
                        count++;
                    }

                    double averageRank = sumRanks / count;
                    for (int j = 0; j < count; j++)
                    {
                        ranks[valueRankPairs[i + j].Index] = averageRank;
                    }

                    rank += count;
                    i += count - 1;
                }
            }

            return ranks;
        }
        static double SpearmanPValue(double rho, int n)
        {
            if (n <= 2)
            {
                throw new ArgumentException("Sample size must be greater than 2.");
            }

            // Calculate the t-statistic
            double t = rho * Math.Sqrt((n - 2) / (1 - rho * rho));
            double df = n - 2;

            // Create a Student's t-distribution object with df degrees of freedom
            var studentT = new Accord.Statistics.Distributions.Univariate.TDistribution(df);



            // Calculate the two-tailed p-value
            return 2 * (1 - studentT.DistributionFunction(Math.Abs(t)));


        }
        public void log(object message)
        {
            MessageBox.Show(message.ToString());
        }
        //F anova

        public AnovaTestResult Fanova(Parameter parameter)
        {
            var groupsData = parameter.GroupedParameterValues.Values;
            var groupLabelsPair = parameter.GroupedParameterValues.Keys;


            double[][] samples = parameter.GroupedParameterValues.Values
            .Select(list => list.ToArray())
            .ToArray();

            var Fanova = new Accord.Statistics.Testing.OneWayAnova(samples);

            double F = Fanova.FTest.Statistic;

            double pValue = Fanova.FTest.PValue;

            string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");


            List<string[]> pairwisecomparisons = new List<string[]>();


            return new AnovaTestResult
            {
                TestValue = F.ToString("0.000"),
                PValue = pValueString,
                PairwiseComparisons = pairwisecomparisons
            };


        }

        public AnovaTestResult Kruskal_H(Parameter parameter) 
        {
            DoubleVector[] samples = parameter.GroupedParameterValues.Values
        .Select(list => new DoubleVector(list.ToArray()))  // Convert each double[] to DoubleVector
        .ToArray();  // Convert IEnumerable<DoubleVector> to DoubleVector[]


            var test = new KruskalWallisTest(samples);

            double HS = test.Statistic;

            double pValue = test.PValue;

            string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");


            List<string[]> pairwisecomparisons = new List<string[]>();


            return new AnovaTestResult
            {
                TestValue = HS.ToString("0.000"),
                PValue = pValueString,
                PairwiseComparisons = pairwisecomparisons
            };



        }


    }

}
