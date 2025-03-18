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
using Accord.Statistics.Analysis;
using Accord.Math;
using Accord.Statistics.Models.Regression.Linear;
using Accord.Statistics.Models.Regression.Fitting;
using Accord.Statistics.Models.Regression;
using Accord;
using DocumentFormat.OpenXml.Drawing;

namespace ExcelScore.Classes
{
   

    public class ManualTests
    {

        //Linear Reg

        public void FRepeated()
        {
            var Fanova = new Accord.Statistics.Analysis.RocAreaMethod();
        }
        public List<double> LinearRegressionn(double[] Indepenent , double[] Dependent)
        {
            List<double> Result = new List<double>();

            try
            {
                var regression = new OrdinaryLeastSquares() { UseIntercept = true };

                SimpleLinearRegression model = regression.Learn(Indepenent, Dependent);

                double[] residuals = new double[Dependent.Length];

                for (int i = 0; i < Dependent.Length; i++)
                {
                    residuals[i] = Dependent[i] - model.Transform(Indepenent[i]);
                }


                double meanSquaredError = residuals.Select(r => r * r).Sum() / (residuals.Length - 2);
                double standardError = Math.Sqrt(meanSquaredError);


                double slopeStandardError = standardError / Math.Sqrt(Indepenent.Select(xi => (xi - Indepenent.Average()) * (xi - Indepenent.Average())).Sum());


                var studentT = new StudentT(0, 1, Dependent.Length - 2);
                double tValue = studentT.InverseCumulativeDistribution(0.975);



                double tStatistic = model.Slope / slopeStandardError;
                double pValue = 2 * (1 - studentT.CumulativeDistribution(Math.Abs(tStatistic)));


                double marginOfError = tValue * slopeStandardError;
                double lowerBound = model.Slope - marginOfError;
                double upperBound = model.Slope + marginOfError;


                Result.Add(model.Slope);
                Result.Add(lowerBound);
                Result.Add(upperBound);
                Result.Add(pValue);

                return Result;
            }
            catch (Exception)
            {
                return new List<double> { 0, 0, 0, 0 };
            }

        }

        public List<double> LogisticRegression(double[] Indepenent, double[] Dependent)
        {
            List<double> Result = new List<double>();

            double[][] inputs = Indepenent.Select(value => new double[] { value }).ToArray();
            var outputs = Dependent;

            var learner = new IterativeReweightedLeastSquares<LogisticRegression>()
            {
                Tolerance = 1e-4,
                MaxIterations = 100,    // Start with a moderate number of iterations
                Regularization = 0
            };

            try
            {
                LogisticRegression regression = learner.Learn(inputs, outputs);
                WaldTest Wald = regression.GetWaldTest(1);
                double B = regression.GetOddsRatio(1);
                double p = Wald.PValue;
                DoubleRange CI = regression.GetConfidenceInterval(1);
                double CI_Lower = CI.Min;
                double CI_Upper = CI.Max;


                Result.Add(B);
                Result.Add(CI_Lower);
                Result.Add(CI_Upper);
                Result.Add(p);


                return Result;

            }
            catch (Exception)
            {
                return new List<double> { 0, 0, 0, 0 };
            }
        }












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

            var matrix = new ConfusionMatrix(new int[,]
            {

                {   a,       b   },
                {  c,       d    },
            });

            double chiSquareStatistic = CalculateChiSquare(a, b, c, d);
            double pValue=0;
            try
            {
                if (ShouldUseFisher(a, b, c, d))
                {
                    pValue = FisherExactTest(a, b, c, d);
                    parameter.Isfisher = true;
                }
                else
                {
                    pValue = ChiSquarePValue(chiSquareStatistic, 1); // df = 1 for a 2x2 table
                }
            }
            catch(Exception)
            {
                if (ShouldUseFisher(a, b, c, d))
                {
                    var fet = new FisherExactTest(matrix, alternate: OneSampleHypothesis.ValueIsDifferentFromHypothesis);
                    pValue = fet.PValue;
                    parameter.Isfisher = true;
                }
                else
                {
                    pValue = ChiSquarePValue(chiSquareStatistic, 1); // df = 1 for a 2x2 table
                }
                
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


        

        public  double CalculateChiSquare(double a, double b, double c, double d)
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

        public  double FisherExactTest(int a, int b, int c, int d)
        {


            double pvalue = NMathFunctions.FishersExactTest(a, b, c, d, HypothesisType.TwoSided);



            return pvalue;
        }



        // Student T test 

        public string[] StudentT_Unpaired(List<double> dataGroup1, List<double> dataGroup2)
        {
            try
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
                double Ttestt = 0;
                double pvalueT = 0;
                try
                {
                    var TstudentTest = new Accord.Statistics.Testing.TwoSampleTTest(dataG1_Array, dataG2_Array, assumeEqualVariances: EqualVariance);
                    Ttestt = TstudentTest.Statistic;
                    pvalueT = TstudentTest.PValue;
                }
                catch (Exception)
                {
                    Ttestt = 0;
                    pvalueT = 0;
                }








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
            catch(Exception)
            {
                return new string[] { "0", "0"};
            }

            

        }

        public string[] UTest(List<double> AdataGroup1, List<double> AdataGroup2)
        {
            try
            {
                using (Py.GIL())
                {
                    // Import necessary Python libraries
                    dynamic scipyStats = Py.Import("scipy.stats");

                    // Define your data groups
                    List<double> dataGroup1 = AdataGroup1;
                    List<double> dataGroup2 = AdataGroup2;

                    // Calculate combined sample size
                    int n1 = dataGroup1.Count;
                    int n2 = dataGroup2.Count;
                    int combinedSize = n1 + n2;



                    Dictionary<int, int> ExactPairs = new Dictionary<int, int>()
                {
                    { 2 , 200 },
                    { 3 , 133 },
                    { 4 , 100 },
                    { 5 , 80 },
                    { 6 , 66 },
                    { 7 , 57 },
                    { 8 , 50 },
                    { 9 , 44 },
                    { 10 , 40 },
                    { 11 , 36 },
                    { 12 , 33 },
                    { 13 , 30 },
                    { 14 , 28 },
                    { 15 , 26 },
                    { 16 , 25 }

                };




                    int minGroup = Math.Min(n1, n2);
                    int MaxGroup = Math.Max(n1, n2);

                    bool exists = ExactPairs.ContainsKey(minGroup);

                    if (exists)
                    {
                        if (ExactPairs.TryGetValue(minGroup, out int value))
                        {
                            if (MaxGroup > value)
                            {
                                exists = false;
                            }
                        }
                    }

                    dynamic result;


                    double UStat = CalculateUStatistic(AdataGroup1.ToArray(), AdataGroup2.ToArray());
                    double Upvalue;

                    if (combinedSize <= 40 || exists)
                    {
                        result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, use_continuity: false, method: "exact", alternative: "two-sided");
                    }

                    // Choose the method based on combined sample size
                    else
                    {
                        result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, use_continuity: false, method: "asymptotic", alternative: "two-sided");
                    }

                    Upvalue = Math.Round(result[1].As<double>(), 3);

                    // Format the p-value
                    string UtestString = UStat.ToString("0.000");
                    string pValueString = Upvalue < 0.001 ? "<0.001" : Upvalue.ToString("0.000");


                    string[] TestValue = new string[] { UtestString, pValueString };

                    return TestValue;
                }
            }
            catch(Exception)
            {
                return new string[] { "–", "–" };
            }
            
        }
        // U Test
        public string[] UTestold(List<double> AdataGroup1, List<double> AdataGroup2)
        {
            try
            {
                double[] group1 = AdataGroup1.ToArray();
                double[] group2 = AdataGroup2.ToArray();

                int totalcount = group1.Length + group2.Length;

                bool exact = totalcount <= 40;

                // Compute the U statistic manually
                double UStat = CalculateUStatistic(group1, group2);
                double Upvalue;

                if (exact)
                {
                    // Use Accord library for exact p-value calculation (small sample case)
                    var UTest = new Accord.Statistics.Testing.MannWhitneyWilcoxonTest(group1, group2, exact: true);
                    Upvalue = UTest.PValue;
                }
                else
                {
                    // Calculate p-value using normal approximation (large sample case)
                    Upvalue = CalculatePValueForLargeSample(UStat, group1.Length, group2.Length);
                }

                // Round the U statistic and p-value for display
                UStat = Math.Round(UStat, 3);
                Upvalue = Math.Round(Upvalue, 3);

                string UtestString = UStat.ToString("0.000");
                string pValueString = Upvalue < 0.001 ? "<0.001" : Upvalue.ToString("0.000");



                return new string[] { UtestString, pValueString };
            }
            catch (Exception)
            {
                return new string[] { "0", "0" };
            }
        }

        // Calculate U statistic as before
        static double CalculateUStatistic(double[] sample1, double[] sample2)
        {
            try
            {
                var allData = sample1.Concat(sample2).ToArray();
                var sample1Count = sample1.Length;
                var sample2Count = sample2.Length;

                var ranks = RankData(allData);

                double rankSum1 = sample1.Sum(x => ranks[x]);
                double rankSum2 = sample2.Sum(x => ranks[x]);

                double u1 = rankSum1 - sample1Count * (sample1Count + 1) / 2.0;
                double u2 = rankSum2 - sample2Count * (sample2Count + 1) / 2.0;

                return Math.Min(u1, u2);
            }
            catch(Exception)
            {
                return 0;
            }
            
        }

        // Function to calculate p-value for large samples using normal approximation
        static double CalculatePValueForLargeSample(double U, int n1, int n2)
        {
            // Mean and standard deviation for U distribution under null hypothesis
            double meanU = (n1 * n2) / 2.0;
            double stdDevU = Math.Sqrt((n1 * n2 * (n1 + n2 + 1)) / 12.0);

            // Z value
            double zValue = (U - meanU) / stdDevU;

            // Calculate p-value using the CDF of the normal distribution (two-tailed)
            double pValue = 2 * (1 - CDFNormal(Math.Abs(zValue)));

            return pValue;
        }

        // CDF function for standard normal distribution (same as before)
        

        // Error function approximation (Erf) for normal CDF calculation (same as before)
        public static double Erf(double x)
        {
            double a1 = 0.254829592;
            double a2 = -0.284496736;
            double a3 = 1.421413741;
            double a4 = -1.453152027;
            double a5 = 1.061405429;
            double p = 0.3275911;

            int sign = (x >= 0) ? 1 : -1;
            x = Math.Abs(x);

            double t = 1.0 / (1.0 + p * x);
            double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-x * x);

            return sign * y;
        }

        private double Erf_New(double x)
        {
            // Approximation of the error function
            double t = 1.0 / (1.0 + 0.5 * Math.Abs(x));
            double tau = t * Math.Exp(-x * x - 1.26551223 + t *
                (1.00002368 + t *
                (0.37409196 + t *
                (0.09678418 + t *
                (-0.18628806 + t *
                (0.27886807 + t *
                (-1.13520398 + t *
                (1.48851587 + t *
                (-0.82215223 + t * 0.17087277)))))))));
            return x >= 0 ? 1 - tau : tau - 1;
        }


        // Ranking function remains unchanged
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


                tempResult = PearsonCorrelation(x, y);
                
            }
            else if (correlationType == "Spearman")
            {
                // Calculate Spearman's correlation coefficient

                //MessageBox.Show(x.Length.ToString());
                //x = x.Except(new double[] { -1 }).ToArray();
                //y = y.Except(new double[] { -1 }).ToArray();

                List<double> newx = new List<double>();
                List<double> newy = new List<double>();
                for (int i = 0; i < x.Length;i++)
                {
                    if (x[i] == -1)
                    {
                        continue;
                    }
                    else if (y[i] == -1)
                    {
                        continue;
                    }
                    else
                    {
                        newx.Add(x[i]);
                        newy.Add(y[i]);
                    }
                }

                double[] Finalx = newx.ToArray();
                double[] Finaly = newy.ToArray();




                //MessageBox.Show(Finalx.Length.ToString());
                //MessageBox.Show(Finaly.Length.ToString());
                //tempResult[0] = ComputeRankCorrelation(Finalx, Finaly);
                tempResult[0] = SpearmanRankCorrelation(Finalx, Finaly);
                tempResult[1] = SpearmanPValue(tempResult[0], Math.Min(Finalx.Length, Finaly.Length));
            }

            //MessageBox.Show(tempResult[0].ToString());
            //MessageBox.Show(tempResult[1].ToString());


            tempResult[0] = Math.Round(tempResult[0] , 4);
            tempResult[1] = Math.Round(tempResult[1], 4);

            

            result[0] = tempResult[0].ToString("0.000");
            result[1] = tempResult[1] < 0.001 ? "<0.001" : tempResult[1].ToString("0.000");

            return result;
        }

        static double[] PearsonCorrelation(double[] x, double[] y)
        {
            int length = Math.Min(x.Length, y.Length);
            int negctr = 0;
            for (int i = 0;i < length; i++)
            {
                if (x[i] == -1)
                {
                    negctr++;
                }
                else if (y[i] == -1)
                {
                    negctr++;
                }

            }

            int finallength = length - negctr;

            //MessageBox.Show(negctr.ToString());
            // Determine the length of the shorter array


            if (length == 0)
            {
                throw new ArgumentException("Both arrays must have at least one element.");
            }

            double sumX = 0, sumY = 0, sumX2 = 0, sumY2 = 0, sumXY = 0;

            for (int i = 0; i < length; i++)
            {
                
                if ((x[i] == -1) || (y[i] == -1))
                {
                    continue;
                }
                else
                {
                    
                    sumX += x[i];
                    sumY += y[i];
                    sumX2 += x[i] * x[i];
                    sumY2 += y[i] * y[i];
                    sumXY += x[i] * y[i];
                }
                
                
            }

            double numerator = (finallength * sumXY) - (sumX * sumY);
            double denominator = Math.Sqrt((finallength * sumX2 - sumX * sumX) * (finallength * sumY2 - sumY * sumY));

            double[] result = new double[2];
            result[0] = numerator / denominator;
            result[1] = PearsonPValue(result[0], finallength);
            //MessageBox.Show(length.ToString());
            return result;
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

            double spearmanRho = 1 - (6.0 * sumD2) / (length * (length * length - 1.0));
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
            try
            {
                var groupsData = parameter.GroupedParameterValues.Values;
                var groupLabelsPair = parameter.GroupedParameterValues.Keys;


                double[][] samples = parameter.GroupedParameterValues.Values
                    .Select(list => list.ToArray())
                    .Where(array => array.Length > 1)
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
            catch(Exception)
            {
                List<string[]> pairwisecomparisons = new List<string[]>();
                return new AnovaTestResult
                {
                    TestValue = "0",
                    PValue = "0",
                    PairwiseComparisons = pairwisecomparisons
                };
            }
            


        }

        public AnovaTestResult Fanova_Relation(Parameter parameter , Parameter groupParameter)
        {
            try
            {
                var groupsData = parameter.GroupedParameterValues_Relation.Values;
                var groupLabelsPair = parameter.GroupedParameterValues_Relation.Keys;


                double[][] samples = parameter.GroupedParameterValues_Relation[groupParameter.Name].Values
                    .Select(list => list.ToArray())
                    .Where(array => array.Length > 1)
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
            catch (Exception)
            {
                List<string[]> pairwisecomparisons = new List<string[]>();
                return new AnovaTestResult
                {
                    TestValue = "0",
                    PValue = "0",
                    PairwiseComparisons = pairwisecomparisons
                };
            }



        }

        public AnovaTestResult Kruskal_H(Parameter parameter) 
        {
            try
            {
                DoubleVector[] samples = parameter.GroupedParameterValues.Values
     .Select(list => list.ToArray())  // Convert each list to double[]
     .Where(array => array.Length > 1)  // Filter out arrays with 1 or fewer elements
     .Select(array => new DoubleVector(array))  // Convert each valid double[] to DoubleVector
     .ToArray(); // Convert IEnumerable<DoubleVector> to DoubleVector[]


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
            catch (Exception)
            {
                List<string[]> pairwisecomparisons = new List<string[]>();
                return new AnovaTestResult
                {
                    TestValue = "0",
                    PValue = "0",
                    PairwiseComparisons = pairwisecomparisons
                };
            }
            



        }

        public AnovaTestResult Kruskal_H_Relation(Parameter parameter , Parameter GroupParameter)
        {
            try
            {
                DoubleVector[] samples = parameter.GroupedParameterValues_Relation[GroupParameter.Name].Values
                        .Select(list => list.ToArray())  // Convert each list to double[]
                        .Where(array => array.Length > 1)  // Filter out arrays with 1 or fewer elements
                        .Select(array => new DoubleVector(array))  // Convert each valid double[] to DoubleVector
                        .ToArray(); // Convert IEnumerable<DoubleVector> to DoubleVector[]


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
            catch (Exception)
            {
                List<string[]> pairwisecomparisons = new List<string[]>();
                return new AnovaTestResult
                {
                    TestValue = "0",
                    PValue = "0",
                    PairwiseComparisons = pairwisecomparisons
                };
            }




        }

        public string[] Tpaired(double[] Para1 , double[] Para2)
        {
            try
            {
                //PairedTTest test = new PairedTTest(Para1, Para2 , TwoSampleHypothesis.ValuesAreDifferent);

                var Nmatht = new CenterSpace.NMath.Core.TwoSamplePairedTTest(Para1, Para2 ,0.05 , HypothesisType.TwoSided);
                double Ttestt = Nmatht.Statistic;
                double pvalueT = Nmatht.P;


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
            catch(Exception)
            {
                return new string[] { "-", "-"};
            }
        }

        public string[] Zpaired(double[] Para1, double[] Para2)
        {
            try
            {
                // Step 1: Calculate the differences between Para1 and Para2
                var differences = Para1.Zip(Para2, (p1, p2) => p1 - p2).ToArray();

                // Step 2: Remove zero differences
                var nonZeroDifferences = differences.Where(d => d != 0).ToArray();

                // Step 3: Rank the absolute values of the differences
                var absDifferences = nonZeroDifferences.Select(Math.Abs).ToArray();

                // Create a list of indexed absolute differences
                var indexedAbsDifferences = absDifferences
                    .Select((value, index) => new { Value = value, Index = index })
                    .OrderBy(x => x.Value)
                    .ToList();

                // Step 3.1: Handle ties by calculating the average rank for tied values
                var ranks = new double[nonZeroDifferences.Length];
                int rankCounter = 1; // Start rank from 1
                int i = 0;
                while (i < indexedAbsDifferences.Count)
                {
                    // Find a group of tied values
                    var tieGroup = indexedAbsDifferences
                        .Where(x => x.Value == indexedAbsDifferences[i].Value)
                        .ToList();

                    // Calculate the average rank for the tied group
                    double averageRank = rankCounter + (tieGroup.Count - 1) / 2.0;

                    // Assign the average rank to all tied values
                    foreach (var item in tieGroup)
                    {
                        ranks[item.Index] = averageRank;
                    }

                    // Increment the rank counter
                    rankCounter += tieGroup.Count;
                    i += tieGroup.Count; // Skip ahead by the number of tied values
                }

                // Step 4: Sum the ranks for negative differences
                double negativeRankSum = 0;
                for (int j = 0; j < nonZeroDifferences.Length; j++)
                {
                    if (nonZeroDifferences[j] < 0)
                    {
                        negativeRankSum += ranks[j];
                    }
                }

                // Step 5: Calculate the Z value
                // Use a normal approximation for large samples
                double n = nonZeroDifferences.Length;
                double mean = n * (n + 1) / 4.0;
                double variance = (n * (n + 1) * (2 * n + 1)) / 24.0;

                // Adjust for ties if necessary
                var tieGroups = absDifferences.GroupBy(x => x).Where(g => g.Count() > 1);
                double tieCorrection = tieGroups.Sum(g => Math.Pow(g.Count(), 3) - g.Count());
                variance -= tieCorrection / 48.0;

                // Continuity correction (optional for small samples)
                double zValue = (negativeRankSum - mean) / Math.Sqrt(variance);

                // Step 6: Calculate the p-value
                double pValue = 2 * (1 - CDFNormal(Math.Abs(zValue))); // Two-tailed test

                // Round values only for display
                string TtestString = zValue.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                // Display the results
                //MessageBox.Show($"Z-value: {TtestString}");
                //MessageBox.Show($"P-value: {pValueString}");

                // Return the result
                return new string[] { TtestString, pValueString };
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"Error: {ex.Message}");
                return new string[] { "-", "-" };
            }
        }


        public static double CDFNormal(double z)
        {
            return 0.5 * (1 + Erf(z / Math.Sqrt(2)));
        }

        // Error function approximation (Erf) for normal CDF calculation



        public string[] UTestNewManual(List<double> AdataGroup1, List<double> AdataGroup2)
        {
            try
            {
                double[] group1 = AdataGroup1.ToArray();
                double[] group2 = AdataGroup2.ToArray();

                int n1 = group1.Length;
                int n2 = group2.Length;
                int totalcount = n1 + n2;

                // Detect ties in the data
                bool hasTies = DetectTies(group1.Concat(group2).ToArray());

                // Determine if exact method should be used
                bool exact = (totalcount <= 40) || (Math.Min(n1, n2) <= 10 && !hasTies);

                // Compute the U statistic
                double UStat = CalculateUStatistic(group1, group2);
                double Upvalue;

                if (exact)
                {
                    // Calculate exact p-value manually for small datasets
                    Upvalue = CalculateExactPValue(UStat, n1, n2);
                }
                else
                {
                    // Calculate p-value using normal approximation for larger datasets
                    Upvalue = CalculatePValueForLargeSample(UStat, n1, n2);
                }

                // Round the U statistic and p-value for display
                UStat = Math.Round(UStat, 3);
                Upvalue = Math.Round(Upvalue, 3);

                string UtestString = UStat.ToString("0.000");
                string pValueString = Upvalue < 0.001 ? "<0.001" : Upvalue.ToString("0.000");

                return new string[] { UtestString, pValueString };
            }
            catch (Exception)
            {
                return new string[] { "0", "0" };
            }
        }

        // Function to calculate exact p-value
        static double CalculateExactPValue(double U, int n1, int n2)
        {
            int totalPermutations = Factorial(n1 + n2) / (Factorial(n1) * Factorial(n2));
            int count = 0;

            // Generate all permutations of ranks
            var allRanks = Enumerable.Range(1, n1 + n2).ToArray();
            foreach (var perm in GetPermutations(allRanks, n1))
            {
                int[] group1Ranks = perm.ToArray();
                int[] group2Ranks = allRanks.Except(group1Ranks).ToArray();

                double rankSum1 = group1Ranks.Sum();
                double rankSum2 = group2Ranks.Sum();

                double u1 = rankSum1 - n1 * (n1 + 1) / 2.0;
                double u2 = rankSum2 - n2 * (n2 + 1) / 2.0;

                if (Math.Min(u1, u2) <= U)
                {
                    count++;
                }
            }

            return (double)count / totalPermutations;
        }

        // Helper function to detect ties
        static bool DetectTies(double[] data)
        {
            // Count the occurrences of each unique value
            var tieGroups = data.GroupBy(x => x).Where(g => g.Count() > 1);

            // Only consider it "ties" if there are meaningful groups affecting the ranks
            // Meaningful ties: at least 2 values tied and account for > 1% of data
            double totalCount = data.Length;
            foreach (var group in tieGroups)
            {
                if (group.Count() > 1 && (group.Count() / totalCount) > 0.01)
                {
                    return true;
                }
            }

            return false;
        }


        // Factorial calculation
        static int Factorial(int n)
        {
            if (n <= 1) return 1;
            return n * Factorial(n - 1);
        }

        // Helper function to generate permutations
        static IEnumerable<IEnumerable<T>> GetPermutations<T>(T[] list, int length)
        {
            if (length == 1) return list.Select(t => new T[] { t });
            return GetPermutations(list, length - 1)
                .SelectMany(t => list.Where(e => !t.Contains(e)),
                            (t1, t2) => t1.Concat(new T[] { t2 }));
        }










    }

}
