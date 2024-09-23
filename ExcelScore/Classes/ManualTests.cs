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
            catch (Exception ex)
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

        // U Test
        public string[] UTest(List<double> AdataGroup1, List<double> AdataGroup2)
        {
            try
            {
                
                double[] group1 = AdataGroup1.ToArray();
                double[] group2 = AdataGroup2.ToArray();


                int totalcount = group1.Length + group2.Length;


                bool CountBool = false;
                if (totalcount > 40)
                {

                    CountBool = false;
                }
                else if (totalcount <= 40)
                {
                    CountBool = true;
                }


                var UTest = new Accord.Statistics.Testing.MannWhitneyWilcoxonTest(group1, group2, exact: CountBool);

                

                double UStat = CalculateUStatistic(group1, group2);

                double Upvalue = UTest.PValue;



                UStat = Math.Round(UStat, 3);
                Upvalue = Math.Round(Upvalue, 3);


                string UtestString = UStat.ToString("0.000");
                string pValueString = Upvalue < 0.001 ? "<0.001" : Upvalue.ToString("0.000");

                string[] TestValue = new string[] { UtestString, pValueString };

                return TestValue;
            }
            catch(Exception)
            {

                string[] TestValue = new string[] { "0", "0"};
                return TestValue;
            }
            
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
                //var Tpaired = new Accord.Statistics.Testing.PairedTTest(Para1, Para2);

                var Nmatht = new TwoSamplePairedTTest(Para1, Para2);
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
                //var Tpaired = new Accord.Statistics.Testing.PairedTTest(Para1, Para2);

                var differences = Para1.Zip(Para2, (p1, p2) => p1 - p2).ToArray();

                // Step 2: Remove zero differences
                var nonZeroDifferences = differences.Where(d => d != 0).ToArray();



                // Step 3: Rank the absolute values of the differences
                var absDifferences = nonZeroDifferences.Select(Math.Abs).ToArray();
                var rankedAbsDifferences = absDifferences.Select((v, i) => new { Value = v, Index = i })
                    .OrderBy(x => x.Value)
                    .Select((x, i) => new { x.Index, Rank = i + 1 })
                    .ToArray();

                var ranks = new double[nonZeroDifferences.Length];
                foreach (var rank in rankedAbsDifferences)
                {
                    ranks[rank.Index] = rank.Rank;
                }

                // Step 4: Sum the ranks for negative differences
                double negativeRankSum = 0;
                for (int i = 0; i < nonZeroDifferences.Length; i++)
                {
                    if (nonZeroDifferences[i] < 0)
                    {
                        negativeRankSum += ranks[i];
                    }
                }

                // Step 5: Calculate the Z value
                // We use a normal approximation for large samples
                double n = nonZeroDifferences.Length;
                double mean = n * (n + 1) / 4;
                double variance = n * (n + 1) * (2 * n + 1) / 24;
                double zValue = (negativeRankSum - mean) / Math.Sqrt(variance);

                // Formatting


                var Nmatht = new Accord.Statistics.Testing.TwoSampleWilcoxonSignedRankTest(Para1, Para2);

                double pvalueT = Nmatht.PValue;



                zValue = Math.Round(zValue, 3);
                pvalueT = Math.Round(pvalueT, 3);

                string TtestString = zValue.ToString("0.000");
                TtestString = TtestString.Replace("-", "");

                // MessageBox.Show(EqualVariance.ToString());
                //MessageBox.Show("T test p " + pvalueT.ToString());

                // Format the p-value
                string pValueString = pvalueT < 0.001 ? "<0.001" : pvalueT.ToString("0.000");


                return new string[] { TtestString, pValueString };
            }
            catch(Exception)
            {
                return new string[] { "-", "-" };
            }
            
        }






    }

}
