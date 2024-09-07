using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Python.Runtime;
using MathNet.Numerics.Statistics;
using Accord.Statistics.Distributions.Univariate;
using MathNet.Numerics.Distributions;
using CenterSpace.NMath.Core;
using System.Data;

namespace ExcelScore.Classes
{
    public class pythonStat
    {
        dynamic np;
        dynamic scipyStats;
        dynamic statsmodels;
        dynamic importlib;




        
        public void InitPython()
        {
            string pythonDll = "";
            string username = Environment.UserName;
            //MessageBox.Show(username);




            if (username != "khaled")
            {
                pythonDll = $@"C:\Users\{username}\AppData\Local\Programs\Python\Python312\python312.dll";
                Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", pythonDll);
                PythonEngine.Initialize();
            }
            else if(username == "khaled")
            {
                pythonDll = @"C:\Users\Khale\AppData\Local\Programs\Python\Python312\python312.dll";
                Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", pythonDll);
                PythonEngine.Initialize();
            }



            using (Py.GIL())
            {
                np = Py.Import("numpy");
                scipyStats = Py.Import("scipy.stats");
                statsmodels = Py.Import("statsmodels.stats.multicomp");
                importlib = Py.Import("importlib");
            }










            //catch (Exception) 
            //{
            //    pythonDll = @"C:\Users\Khale\AppData\Local\Programs\Python\Python312\python38.dll";
            //}


        }

        public AnovaTestResult KruskalWallisNoDunnDynamic(Parameter parameter)
        {
            using (Py.GIL())
            {
                //dynamic np = Py.Import("numpy");
                //dynamic scipyStats = Py.Import("scipy.stats");
                //dynamic statsmodels = Py.Import("statsmodels.sandbox.stats.multicomp");
                //dynamic scikit_posthocs = Py.Import("scikit_posthocs");
                //dynamic pandas = Py.Import("pandas");

                var groupsData = parameter.GroupedParameterValues.Values;
                var groupLabelsPair = parameter.GroupedParameterValues.Keys;

                var pyGroupsData = new PyList();



                foreach (List<double> groupData in groupsData)
                {

                    var pyGroupData = new PyList();
                    foreach (double value in groupData)
                    {
                        pyGroupData.Append(value.ToPython());
                    }
                    pyGroupsData.Append(pyGroupData);
                }



                PyTuple pyGroupsDataTuple = new PyTuple(pyGroupsData.ToArray());
                dynamic groupLabels = parameter.GroupedParameterValues.Keys.ToList().ToPython();






                List<double> allValues = new List<double>();
                List<double> allGroupLabels = new List<double>();

                foreach (var kvp in parameter.GroupedParameterValues)
                {
                    double groupLabel = kvp.Key;
                    List<double> groupValues = kvp.Value;

                    allGroupLabels.AddRange(Enumerable.Repeat(groupLabel, groupValues.Count));
                    allValues.AddRange(groupValues);
                }












                dynamic data = np.array(allValues.ToArray());
                dynamic groupLabelsCombined = np.array(allGroupLabels.ToArray());



                var dfData = new List<List<object>>();
                int numRows = (int)data.shape[0];
                for (int i = 0; i < numRows; i++)
                {
                    List<object> row = new List<object>
                    {
                        data[i],                  // Value
                        groupLabelsCombined[i]    // Group
                    };
                    dfData.Add(row);
                }

                //dynamic df = pandas.DataFrame(dfData, columns: new List<string> { "Value", "Group" });


                string pythonScript = @"
result_kw = scipyStats.kruskal(*groups_data)
";

                dynamic locals = new PyDict();
                dynamic globals = new PyDict();

                globals["scipyStats"] = scipyStats;
                //globals["statsmodels"] = statsmodels;
                globals["groups_data"] = pyGroupsDataTuple;
                globals["group_labels"] = groupLabels;
                globals["np"] = np;
                //globals["df"] = df;
                //globals["scikit_posthocs"] = scikit_posthocs;
                globals["data"] = data;
                globals["group_labels_combined"] = groupLabelsCombined;



                PythonEngine.Exec(pythonScript, locals, globals);

                dynamic kwResult = globals["result_kw"];
                //dynamic result = globals["dunn_result"];





                // Create a pandas DataFrame

                //dynamic dunnResult = globals["dunn_result"];

                double hValue = kwResult[0].As<double>();
                double pValueKruskalWallis = kwResult[1].As<double>();

                

                string pValueString = pValueKruskalWallis <= 0.001 ? "<0.001" : pValueKruskalWallis.ToString("0.000");

                //dynamic pValuesDunn = result;

                //List<string[]> pairwisecomparisons = new List<string[]>();
                //int k = 0;

                //// Assuming groupLabelsCombined contains unique group labels
                //var uniqueGroupLabels = allGroupLabels.Distinct().ToList();

                //for (int i = 0; i < uniqueGroupLabels.Count; i++)
                //{
                //    for (int j = i + 1; j < uniqueGroupLabels.Count; j++)
                //    {
                //        string labelA = uniqueGroupLabels[i].ToString();
                //        string labelB = uniqueGroupLabels[j].ToString();

                //        // Find the corresponding p-value in the correct order
                //        double pAdjValue = pValuesDunn[k].As<double>();

                //        // Handle rounding only if the value is not zero
                //        string pAdjString = pAdjValue != 0.0 ? Math.Round(pAdjValue, 3).ToString("0.000") : "0.000";

                //        // Display or use the p-value as needed
                //        MessageBox.Show(pAdjString);

                //        string comparison = $"{labelA} vs {labelB}";
                //        pairwisecomparisons.Add(new string[] { comparison, pAdjString });

                //        parameter.FPairwise.Add(pAdjString);
                //        k++;
                //    }
                //}


                return new AnovaTestResult
                {
                    TestValue = hValue.ToString("0.000"),
                    PValue = pValueString,
                    //PairwiseComparisons = pairwisecomparisons
                };
            }
        }


        public string[] CalculateCorrelation(double[] x, double[] y, string correlationType)
        {
            string[] result = new string[2];
            double[] tempResult = new double[2];

            if (correlationType == "Pearson")
            {
                // Calculate Pearson's correlation coefficient
                tempResult[0] = Correlation.Pearson(x, y);
            }
            else if (correlationType == "Spearman")
            {
                // Calculate Spearman's correlation coefficient
                tempResult[0] = Correlation.Spearman(x, y);
            }

            // Calculate degrees of freedom (df) for t-distribution
            int n = x.Length;
            int df = n - 2; // degrees of freedom for a two-tailed test

            // Calculate t-statistic
            double t = tempResult[0] * Math.Sqrt(df / (1 - Math.Pow(tempResult[0], 2)));

            // Calculate two-tailed p-value using t-distribution
            var studentT = new StudentT(0, 1, df);
            tempResult[1] = 2 * (1 - studentT.CumulativeDistribution(Math.Abs(t)));
            //tempResult[1] = 2 * (1 - StudentT.CDF(0, 1, df, Math.Abs(t)));

            MessageBox.Show("r " + tempResult[0].ToString());
            MessageBox.Show("p " + tempResult[1].ToString());
            // Format the results
            result[0] = tempResult[0].ToString("0.000");
            result[1] = tempResult[1] <= 0.001 ? "<0.001" : tempResult[1].ToString("0.000");

            return result;
        }
        public AnovaTestResult newANOVAWithTukeyHSDNewDynamic(Parameter parameter)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic scipyStats = Py.Import("scipy.stats");
                dynamic statsmodels = Py.Import("statsmodels.sandbox.stats.multicomp");
                dynamic scikit_posthocs = Py.Import("scikit_posthocs");
                dynamic pandas = Py.Import("pandas");

                var groupsData = parameter.GroupedParameterValues.Values;
                var groupLabelsPair = parameter.GroupedParameterValues.Keys;

                var pyGroupsData = new PyList();



                foreach (List<double> groupData in groupsData)
                {

                    var pyGroupData = new PyList();
                    foreach (double value in groupData)
                    {
                        pyGroupData.Append(value.ToPython());
                    }
                    pyGroupsData.Append(pyGroupData);
                }



                PyTuple pyGroupsDataTuple = new PyTuple(pyGroupsData.ToArray());
                dynamic groupLabels = parameter.GroupedParameterValues.Keys.ToList().ToPython();






                List<double> allValues = new List<double>();
                List<double> allGroupLabels = new List<double>();

                foreach (var kvp in parameter.GroupedParameterValues)
                {
                    double groupLabel = kvp.Key;
                    List<double> groupValues = kvp.Value;

                    allGroupLabels.AddRange(Enumerable.Repeat(groupLabel, groupValues.Count));
                    allValues.AddRange(groupValues);
                }

                dynamic data = np.array(allValues.ToArray());
                dynamic groupLabelsCombined = np.array(allGroupLabels.ToArray());



                var dfData = new List<List<object>>();
                int numRows = (int)data.shape[0];
                for (int i = 0; i < numRows; i++)
                {
                    List<object> row = new List<object>
                    {
                        data[i],                  // Value
                        groupLabelsCombined[i]    // Group
                    };
                    dfData.Add(row);
                }

                dynamic df = pandas.DataFrame(dfData, columns: new List<string> { "Value", "Group" });


                string pythonScript = @"
result_kw = scipyStats.kruskal(*groups_data)
tukey_result = scikit_posthocs.posthoc_tukey(df , val_col=""Value"" , group_col=""Group"")
";

                dynamic locals = new PyDict();
                dynamic globals = new PyDict();

                globals["scipyStats"] = scipyStats;
                globals["statsmodels"] = statsmodels;
                globals["groups_data"] = pyGroupsDataTuple;
                globals["group_labels"] = groupLabels;
                globals["np"] = np;
                globals["df"] = df;
                globals["scikit_posthocs"] = scikit_posthocs;
                globals["data"] = data;
                globals["group_labels_combined"] = groupLabelsCombined;



                PythonEngine.Exec(pythonScript, locals, globals);

                dynamic kwResult = globals["result_kw"];
                dynamic result = globals["tukey_result"];



                double hValue = kwResult[0].As<double>();
                double pValueKruskalWallis = kwResult[1].As<double>();

                if (pValueKruskalWallis < 0.05)
                {
                    parameter.ISFAnovaSig = true;
                }

                string pValueString = pValueKruskalWallis <= 0.001 ? "<0.001" : pValueKruskalWallis.ToString("0.000");

                dynamic pValuestukey = result;




                DataTable dataTable = new DataTable();


                dynamic columns = result.columns;
                foreach (var column in columns)
                {
                    dataTable.Columns.Add(column.ToString());

                }

                foreach (dynamic row in result.itertuples(index: false, name: null))
                {
                    // Create a new row in the DataTable
                    DataRow dataRow = dataTable.NewRow();

                    // Loop through the row data and assign it to the DataTable row
                    int colIndex = 0;
                    foreach (var cell in row)
                    {
                        dataRow[colIndex] = cell;
                        colIndex++;
                    }

                    // Add the row to the DataTable
                    dataTable.Rows.Add(dataRow);
                }

                //foreach (DataColumn column in dataTable.Columns)
                //{
                //    MessageBox.Show(column.ColumnName + "\t");
                //}


                //// Print each row of the DataTable
                //foreach (DataRow row in dataTable.Rows)
                //{
                //    foreach (var item in row.ItemArray)
                //    {
                //        MessageBox.Show(item + "\t");
                //    }

                //}




                List<string[]> pairwisecomparisons = new List<string[]>();
                int k = 0;

                // Assuming groupLabelsCombined contains unique group labels
                var uniqueGroupLabels = allGroupLabels.Distinct().ToList();



                for (int i = 0; i < uniqueGroupLabels.Count; i++)
                {
                    for (int j = i + 1; j < uniqueGroupLabels.Count; j++)
                    {
                        string labelA = uniqueGroupLabels[i].ToString();
                        string labelB = uniqueGroupLabels[j].ToString();

                        // Find the corresponding p-value in the correct order

                        double pvalue;
                        // Handle rounding only if the value is not zero
                        double.TryParse(dataTable.Rows[i][j].ToString(), out pvalue);


                        //MessageBox.Show(pvalue.ToString());

                        string pAdjString = pvalue != 0.0 ? Math.Round(pvalue, 3).ToString("0.000") : "0.000";

                        // Display or use the p-value as needed
                        //MessageBox.Show(pAdjString);

                        string comparison = $"{labelA} vs {labelB}";
                        pairwisecomparisons.Add(new string[] { comparison, pAdjString });

                        parameter.FPairwise.Add(pAdjString);


                        k++;
                    }
                }


                return new AnovaTestResult
                {
                    TestValue = hValue.ToString("0.000"),
                    PValue = pValueString,
                    PairwiseComparisons = pairwisecomparisons
                };
            }
        }
        public AnovaTestResult KruskalWallisWithDunnDynamic(Parameter parameter)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic scipyStats = Py.Import("scipy.stats");
                dynamic statsmodels = Py.Import("statsmodels.sandbox.stats.multicomp");
                dynamic scikit_posthocs = Py.Import("scikit_posthocs");
                dynamic pandas = Py.Import("pandas");

                var groupsData = parameter.GroupedParameterValues.Values;
                var groupLabelsPair = parameter.GroupedParameterValues.Keys;

                var pyGroupsData = new PyList();


                
                foreach (List<double> groupData in groupsData)
                {

                    var pyGroupData = new PyList();
                    foreach (double value in groupData)
                    {
                        pyGroupData.Append(value.ToPython());
                    }
                    pyGroupsData.Append(pyGroupData);
                }

                

                PyTuple pyGroupsDataTuple = new PyTuple(pyGroupsData.ToArray());
                dynamic groupLabels = parameter.GroupedParameterValues.Keys.ToList().ToPython();

                
               



                List<double> allValues = new List<double>();
                List<double> allGroupLabels = new List<double>();

                foreach (var kvp in parameter.GroupedParameterValues)
                {
                    double groupLabel = kvp.Key;
                    List<double> groupValues = kvp.Value;

                    allGroupLabels.AddRange(Enumerable.Repeat(groupLabel, groupValues.Count));
                    allValues.AddRange(groupValues);
                }




                





                

                dynamic data = np.array(allValues.ToArray());
                dynamic groupLabelsCombined = np.array(allGroupLabels.ToArray());



                var dfData = new List<List<object>>();
                int numRows = (int)data.shape[0];
                for (int i = 0; i < numRows; i++)
                {
                    List<object> row = new List<object>
                    {
                        data[i],                  // Value
                        groupLabelsCombined[i]    // Group
                    };
                    dfData.Add(row);
                }

                dynamic df = pandas.DataFrame(dfData, columns: new List<string> { "Value", "Group" });


                string pythonScript = @"
result_kw = scipyStats.kruskal(*groups_data)
dunn_result = scikit_posthocs.posthoc_dunn(df , val_col=""Value"" , group_col=""Group"")
";

                dynamic locals = new PyDict();
                dynamic globals = new PyDict();

                globals["scipyStats"] = scipyStats;
                globals["statsmodels"] = statsmodels;
                globals["groups_data"] = pyGroupsDataTuple;
                globals["group_labels"] = groupLabels;
                globals["np"] = np;
                globals["df"] = df;
                globals["scikit_posthocs"] = scikit_posthocs;
                globals["data"] = data;
                globals["group_labels_combined"] = groupLabelsCombined;



                PythonEngine.Exec(pythonScript, locals, globals);

                dynamic kwResult = globals["result_kw"];
                dynamic result = globals["dunn_result"];

                
                



                
                // Create a pandas DataFrame
                
                //dynamic dunnResult = globals["dunn_result"];

                double hValue = kwResult[0].As<double>();
                double pValueKruskalWallis = kwResult[1].As<double>();

                if (pValueKruskalWallis < 0.05)
                {
                    parameter.ISFAnovaSig = true;
                }

                string pValueString = pValueKruskalWallis <= 0.001 ? "<0.001" : pValueKruskalWallis.ToString("0.000");

                dynamic pValuesDunn = result;

                


                DataTable dataTable = new DataTable();


                dynamic columns = result.columns;
                foreach (var column in columns)
                {
                    dataTable.Columns.Add(column.ToString());
                    
                }

                foreach (dynamic row in result.itertuples(index: false, name: null))
                {
                    // Create a new row in the DataTable
                    DataRow dataRow = dataTable.NewRow();

                    // Loop through the row data and assign it to the DataTable row
                    int colIndex = 0;
                    foreach (var cell in row)
                    {
                        dataRow[colIndex] = cell;
                        colIndex++;
                    }

                    // Add the row to the DataTable
                    dataTable.Rows.Add(dataRow);
                }

                //foreach (DataColumn column in dataTable.Columns)
                //{
                //    MessageBox.Show(column.ColumnName + "\t");
                //}
               

                // Print each row of the DataTable
                //foreach (DataRow row in dataTable.Rows)
                //{
                //    foreach (var item in row.ItemArray)
                //    {
                //        MessageBox.Show(item + "\t");
                //    }
                    
                //}




                List<string[]> pairwisecomparisons = new List<string[]>();
                int k = 0;

                // Assuming groupLabelsCombined contains unique group labels
                var uniqueGroupLabels = allGroupLabels.Distinct().ToList();

                

                for (int i = 0; i < uniqueGroupLabels.Count; i++)
                {
                    for (int j = i + 1; j < uniqueGroupLabels.Count; j++)
                    {
                        string labelA = uniqueGroupLabels[i].ToString();
                        string labelB = uniqueGroupLabels[j].ToString();

                        // Find the corresponding p-value in the correct order

                        double pvalue;
                        // Handle rounding only if the value is not zero
                        double.TryParse(dataTable.Rows[i][j].ToString() , out pvalue);








                        string pAdjString = pvalue != 0.0 ? Math.Round(pvalue, 3).ToString("0.000") : "0.000";

                        // Display or use the p-value as needed
                        //MessageBox.Show(pAdjString);

                        string comparison = $"{labelA} vs {labelB}";
                        pairwisecomparisons.Add(new string[] { comparison, pAdjString });

                        parameter.FPairwise.Add(pAdjString);

                        
                        k++;
                    }
                }


                return new AnovaTestResult
                {
                    TestValue = hValue.ToString("0.000"),
                    PValue = pValueString,
                    PairwiseComparisons = pairwisecomparisons
                };
            }
        }
        public AnovaTestResult ANOVAWithMultipleComparisons(Parameter parameter)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic scipyStats = Py.Import("scipy.stats");
                dynamic statsmodels = Py.Import("statsmodels.stats.multicomp");
                dynamic multitest = Py.Import("statsmodels.stats.multitest");  // Correct import for multipletests

                var groupsData = parameter.GroupedParameterValues.Values;
                var groupLabelsPair = parameter.GroupedParameterValues.Keys;

                var pyGroupsData = new PyList();

                foreach (List<double> groupData in groupsData)
                {
                    var pyGroupData = new PyList();
                    foreach (double value in groupData)
                    {
                        pyGroupData.Append(value.ToPython());
                    }
                    pyGroupsData.Append(pyGroupData);
                }

                PyTuple pyGroupsDataTuple = new PyTuple(pyGroupsData.ToArray());
                dynamic groupLabels = parameter.GroupedParameterValues.Keys.ToList().ToPython();

                List<double> allValues = new List<double>();
                List<double> allGroupLabels = new List<double>();

                foreach (var kvp in parameter.GroupedParameterValues)
                {
                    double groupLabel = kvp.Key;
                    List<double> groupValues = kvp.Value;

                    allGroupLabels.AddRange(Enumerable.Repeat(groupLabel, groupValues.Count));
                    allValues.AddRange(groupValues);
                }

                dynamic data = np.array(allValues.ToArray());
                dynamic groupLabelscombined = np.array(allGroupLabels.ToArray());

                string pythonScript = @"
result_anova = scipyStats.f_oneway(*groups_data)
tukey_result = statsmodels.pairwise_tukeyhsd(data, groups)
bonferroni_result = multipletests.multipletests(tukey_result.pvalues, method='bonferroni')
sidak_result = multipletests.multipletests(tukey_result.pvalues, method='sidak')
";

                dynamic locals = new PyDict();
                dynamic globals = new PyDict();

                globals["scipyStats"] = scipyStats;
                globals["statsmodels"] = statsmodels;
                globals["multitest"] = multitest;  // Import for multipletests
                globals["groups_data"] = pyGroupsDataTuple;
                globals["group_labels"] = groupLabels;
                globals["np"] = np;

                globals["data"] = data;
                globals["groups"] = groupLabelscombined;

                PythonEngine.Exec(pythonScript, locals, globals);

                dynamic fResult = globals["result_anova"];
                dynamic tukeyResult = globals["tukey_result"];
                dynamic bonferroniPValues = globals["bonferroni_pvalues"];
                dynamic sidakPValues = globals["sidak_pvalues"];

                double fValue = fResult[0].As<double>();
                double pValueANOVA = fResult[1].As<double>();

                if (pValueANOVA < 0.05)
                {
                    parameter.ISFAnovaSig = true;
                }

                string pValueString = pValueANOVA <= 0.001 ? "<0.001" : pValueANOVA.ToString("0.000");

                // Process Tukey's HSD results
                dynamic pValuesTukey = tukeyResult.pvalues;

                // Convert numpy.ndarray to C# List<double>
                List<double> pValuesBonferroni = bonferroniPValues.ToArray().ToList();
                List<double> pValuesSidak = sidakPValues.ToArray().ToList();

                List<string[]> pairwiseComparisons = new List<string[]>();
                int k = 0;
                for (int i = 0; i < groupsData.Count; i++)
                {
                    for (int j = i + 1; j < groupsData.Count; j++)
                    {
                        string labelA = parameter.GroupedParameterValues.Keys.ElementAt(i).ToString();
                        string labelB = parameter.GroupedParameterValues.Keys.ElementAt(j).ToString();
                        double pAdjTukey = Math.Round(pValuesTukey[k].As<double>(), 3);
                        double pAdjBonferroni = Math.Round(pValuesBonferroni[k], 3);
                        double pAdjSidak = Math.Round(pValuesSidak[k], 3);

                        string pAdjTukeyString = pAdjTukey <= 0.001 ? "<0.001" : pAdjTukey.ToString("0.000");
                        string pAdjBonferroniString = pAdjBonferroni <= 0.001 ? "<0.001" : pAdjBonferroni.ToString("0.000");
                        string pAdjSidakString = pAdjSidak <= 0.001 ? "<0.001" : pAdjSidak.ToString("0.000");

                        string comparison = $"{labelA} vs {labelB}";
                        pairwiseComparisons.Add(new string[] { comparison, "Tukey: " + pAdjTukeyString, "Bonferroni: " + pAdjBonferroniString, "Sidak: " + pAdjSidakString });

                        parameter.FPairwise.Add(pAdjTukeyString);
                        k++;
                    }
                }

                return new AnovaTestResult
                {
                    TestValue = fValue.ToString("0.000"),
                    PValue = pValueString,
                    PairwiseComparisons = pairwiseComparisons
                };
            }
        }


        public AnovaTestResult ANOVAWithTukeyHSDNewDynamic(Parameter parameter)
        {
            using (Py.GIL())
            {
                //dynamic np = Py.Import("numpy");
                //dynamic scipyStats = Py.Import("scipy.stats");
                //dynamic statsmodels = Py.Import("statsmodels.stats.multicomp");

                //dynamic importlib = Py.Import("importlib");

                var groupsData = parameter.GroupedParameterValues.Values;
                var groupLabelsPair = parameter.GroupedParameterValues.Keys;

                var pyGroupsData = new PyList();

                foreach (List<double> groupData in groupsData)
                {
                    var pyGroupData = new PyList();
                    foreach (double value in groupData)
                    {
                        pyGroupData.Append(value.ToPython());
                    }
                    pyGroupsData.Append(pyGroupData);
                }


                PyTuple pyGroupsDataTuple = new PyTuple(pyGroupsData.ToArray());
                dynamic groupLabels = parameter.GroupedParameterValues.Keys.ToList().ToPython();




                List<double> allValues = new List<double>();
                List<double> allGroupLabels = new List<double>();

                foreach (var kvp in parameter.GroupedParameterValues)
                {
                    double groupLabel = kvp.Key;
                    List<double> groupValues = kvp.Value;

                    // Add group labels to the allGroupLabels list for each value in the current group
                    allGroupLabels.AddRange(Enumerable.Repeat(groupLabel, groupValues.Count));

                    // Add all values in the current group to the allValues list
                    allValues.AddRange(groupValues);
                }


                dynamic data = np.array(allValues.ToArray());
                dynamic groupLabelscombined = np.array(allGroupLabels.ToArray());

                //tukey_result = statsmodels.pairwise_tukeyhsd(, )

                string pythonScript = @"
result_anova = scipyStats.f_oneway(*groups_data)
tukey_result = statsmodels.pairwise_tukeyhsd(data,groups)
";

                dynamic locals = new PyDict();
                dynamic globals = new PyDict();

                globals["scipyStats"] = scipyStats;
                globals["statsmodels"] = statsmodels;
                globals["groups_data"] = pyGroupsDataTuple;
                globals["group_labels"] = groupLabels;
                globals["np"] = np;

                globals["data"] = data;  // Convert List<double> to array
                globals["groups"] = groupLabelscombined;  // Convert List<double> to array

                PythonEngine.Exec(pythonScript, locals, globals);

                dynamic fResult = globals["result_anova"];
                dynamic tukeyResult = globals["tukey_result"];



                double fValue = fResult[0].As<double>();
                double pValueANOVA = fResult[1].As<double>();

                if (pValueANOVA < 0.05)
                {
                    parameter.ISFAnovaSig = true;
                }

                string pValueString = pValueANOVA <= 0.001 ? "<0.001" : pValueANOVA.ToString("0.000");


                dynamic pValuesTukey = tukeyResult.pvalues; // adjust this based on the actual structure of your result
                //double pAdjValue = Math.Round(pValuesTukey[k].As<double>(), 3);

                List<string[]> pairwisecomparisons = new List<string[]>();
                int k = 0;
                for (int i = 0; i < groupsData.Count; i++)
                {
                    for (int j = i + 1; j < groupsData.Count; j++)
                    {
                        string labelA = parameter.GroupedParameterValues.Keys.ElementAt(i).ToString();
                        string labelB = parameter.GroupedParameterValues.Keys.ElementAt(j).ToString();
                        double pAdjValue = Math.Round(pValuesTukey[k].As<double>(), 3);
                        string pAdjString = pAdjValue < 0.001 ? "<0.001" : pAdjValue.ToString("0.000");
                        
                        string comparison = $"{labelA} vs {labelB}";
                        pairwisecomparisons.Add(new string[] { comparison, pAdjString });

                        parameter.FPairwise.Add(pAdjString);
                        k++;
                    }
                }


                //MessageBox.Show(tukeyResult);
                return new AnovaTestResult
                {
                    TestValue = fValue.ToString("0.000"),
                    PValue = pValueString,
                    PairwiseComparisons = pairwisecomparisons
                };
            }
        }





        public string[] NewChiSquare(List<List<int>> contingencyTable , Parameter parameter)
        {
            using (Py.GIL())
            {
               // dynamic scipyStats = Py.Import("scipy.stats");

                // Convert the contingency table to a numpy array
                dynamic tableArray = new PyList();
                foreach (var row in contingencyTable)
                {
                    dynamic rowData = new PyList();
                    foreach (var cell in row)
                    {
                        rowData.append(cell);
                    }
                    tableArray.append(rowData);
                }

                // Declare variables to store the chi-square statistic and p-value
                double chiSquareStatistic = 0.0;
                double pValue = 0.0;

                // Perform chi-square test without correction
                dynamic chiSquareResultNoCorrection = scipyStats.chi2_contingency(tableArray, correction: false);
                chiSquareStatistic = chiSquareResultNoCorrection[0].As<double>();
                pValue = chiSquareResultNoCorrection[1].As<double>();

                // Extract the expected frequencies without correction
                dynamic expectedFreqNoCorrection = chiSquareResultNoCorrection[3];

                // Check if any expected cell count is less than 5
                bool hasSmallExpectedCell = false;
                foreach (var row in expectedFreqNoCorrection)
                {
                    foreach (var cell in row)
                    {
                        if (cell.As<double>() < 5)
                        {
                            hasSmallExpectedCell = true;
                            break;
                        }
                    }
                    if (hasSmallExpectedCell)
                    {
                        break;
                    }
                }

                // If there are small expected cells, perform chi-square test with correction
                if (contingencyTable.Count == 2 && contingencyTable[0].Count == 2 && hasSmallExpectedCell)
                {
                    // Fisher exact test for 2x2 table
                    

                    ManualTests manual = new ManualTests();
                    string[] result = manual.getchi(parameter);
                    chiSquareStatistic = double.Parse(result[0]);
                    pValue = double.Parse(result[1]);
                }

                //else if(contingencyTable.Count > 2 && contingencyTable[0].Count > 2 && hasSmallExpectedCell)
                //{
                //    dynamic fisherExact = Py.Import("FisherExact.fisher_exact");

                //    // Convert the contingency table for Python consumption
                //    dynamic fisherTableArray = new PyList();
                //    foreach (var row in contingencyTable)
                //    {
                //        dynamic rowData = new PyList();
                //        foreach (var cell in row)
                //        {
                //            rowData.append(cell);
                //        }
                //        fisherTableArray.append(rowData);
                //    }

                //    // Perform Fisher's exact test for larger tables
                //    dynamic fisherResult = fisherExact(fisherTableArray);

                //    // Extract the p-value and statistic (if needed)
                //    pValue = fisherResult[1].As<double>();
                //}
                else
                {
                    dynamic chiSquareResult = scipyStats.chi2_contingency(tableArray, correction: false);
                    chiSquareStatistic = chiSquareResult[0].As<double>();
                    pValue = chiSquareResult[1].As<double>();

                }

                // Round the chi-square statistic and p-value
                chiSquareStatistic = Math.Round(chiSquareStatistic, 3);
                pValue = Math.Round(pValue, 3);

                //MessageBox.Show(chiSquareStatistic.ToString());
                //MessageBox.Show(pValue.ToString());
                // Format the p-value
                string chiSquareString = chiSquareStatistic.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                return new string[] { chiSquareString, pValueString };
            }
        }





        public string[] ChiSquareTest(List<List<int>> contingencyTable)
        {
            using (Py.GIL())
            {
                dynamic scipyStats = Py.Import("scipy.stats");

                // Convert the contingency table to a numpy array
                dynamic tableArray = new PyList();
                foreach (var row in contingencyTable)
                {
                    dynamic rowData = new PyList();
                    foreach (var cell in row)
                    {
                        rowData.append(cell);
                    }
                    tableArray.append(rowData);
                }

                // Check if any cell is less than 5
                bool hasSmallCell = false;
                foreach (var row in contingencyTable)
                {
                    foreach (var cell in row)
                    {
                        if (cell < 5)
                        {
                            hasSmallCell = true;
                            break;
                        }
                    }
                    if (hasSmallCell)
                    {
                        break;
                    }
                }

                // Declare variables to store the chi-square statistic and p-value
                double chiSquareStatistic = 0.0;
                double pValue = 0.0;

                // Determine the correction method based on the size of the table and small cells
                if (hasSmallCell && contingencyTable.Count == 2 && contingencyTable[0].Count == 2)
                {
                    // Fisher exact test for 2x2 table
                    dynamic fisherResult = scipyStats.fisher_exact(tableArray);
                    chiSquareStatistic = fisherResult[0].As<double>();
                    pValue = fisherResult[1].As<double>();
                }
                //else if (hasSmallCell)
                //{
                //    // Read the embedded Python script
                //    string script;
                //    using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelScore.ChiMonte.py"))
                //    {
                //        using (StreamReader reader = new StreamReader(stream))
                //        {
                //            script = reader.ReadToEnd();
                //        }
                //    }

                //    // Execute the Python script
                //    using (Py.GIL())
                //    {
                //        dynamic scope = Py.CreateScope();
                //        scope.Exec(script);

                //        // Call the function defined in the Python script
                //        dynamic monteCarloChiSquareFunction = scope.Get("monte_carlo_chi_square");
                //        dynamic result = monteCarloChiSquareFunction(tableArray); // Pass contingency table as argument

                //        // Use the result as needed
                //        chiSquareStatistic = result[0].As<double>();
                //        pValue = result[1].As<double>();
                //    }
                //}


                else
                {
                    // Default chi-square test
                    dynamic chiSquareResult = scipyStats.chi2_contingency(tableArray);
                    chiSquareStatistic = chiSquareResult[0].As<double>();
                    //pValue = chiSquareResult[1].As<double>();
                    pValue = chiSquareResult[1]["Asymp. Sig. (2-sided)"].As<double>();

                    
                }

                // Round the chi-square statistic and p-value
                chiSquareStatistic = Math.Round(chiSquareStatistic, 3);
                pValue = Math.Round(pValue, 3);

                // Format the p-value
                string chiSquareString = chiSquareStatistic.ToString("0.000");
                string pValueString = pValue <= 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { chiSquareString, pValueString };

                return TestValue;
                // Display the results
            }
        }

        public string[] PerformChiSquareTest(Parameter parameter)
        {

            if (parameter.IsGroup || parameter.NominalOrScale != "Nominal")
            {
                return new string[] { "Parameter is not relevant for chi-square test", "" };
            }

            // Identify unique categories across all groups
            var allValues = parameter.GroupedParameterValues.Values.SelectMany(x => x).Distinct().ToList();

            // Prepare contingency table
            var contingencyTable = new List<List<int>>();
            foreach (var value in allValues)
            {
                var categoryCounts = new List<int>();
                foreach (var groupData in parameter.GroupedParameterValues.Values)
                {
                    categoryCounts.Add(groupData.Count(x => x == value));
                }
                contingencyTable.Add(categoryCounts);
            }


            // Perform chi-square test
            return NewChiSquare(contingencyTable , parameter);
        }


        public string[] WilcoxonSignedRankTest(List<double> AdataGroup1, List<double> AdataGroup2)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic scipyStats = Py.Import("scipy.stats");

                // Convert the lists to numpy arrays
                dynamic data1 = np.array(AdataGroup1);
                dynamic data2 = np.array(AdataGroup2);

                // Perform the Wilcoxon signed-rank test
                dynamic result = scipyStats.wilcoxon(data1, data2);

                // Extract the p-value and test statistic from the result
                double pValue = Math.Round(result.pvalue.As<double>(), 3);
                double testStatistic = Math.Round(result.statistic.As<double>(), 3);

                // Format the p-value
                string testStatisticString = testStatistic.ToString("0.000");
                string pValueString = pValue <= 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { testStatisticString, pValueString };

                return TestValue;
            }
        }

        public string[] MannWhitneyUTest(List<double> AdataGroup1, List<double> AdataGroup2)
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

                dynamic result;

                // Choose the method based on combined sample size
                if (combinedSize > 40)
                {
                    // Use the asymptotic method
                    result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, method : "asymptotic", alternative: "two-sided");
                }
                else
                {
                    // Use the exact method
                    result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, method: "exact", alternative: "two-sided");
                }

                // Extract the p-value and U-statistic from the result
                double uStatistic = Math.Round(result[0].As<double>(), 3);
                double pValue = Math.Round(result[1].As<double>(), 3);

                // Format the p-value
                string UtestString = uStatistic.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { UtestString, pValueString };

                return TestValue;
            }
        }


        

        



        public string[] TpairedTest(List<double> AdataGroup1 , List<double> AdataGroup2)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic scipyStats = Py.Import("scipy.stats");

                // Define your paired data
                List<double> dataGroup1 = AdataGroup1;
                List<double> dataGroup2 = AdataGroup2;

                // Convert the lists to numpy arrays
                dynamic data1 = np.array(dataGroup1);
                dynamic data2 = np.array(dataGroup2);

                // Perform the paired t-test
                dynamic result = scipyStats.ttest_rel(data1, data2);

                // Extract the p-value and t-statistic from the result
                double pValue = Math.Round(result[1].As<double>(), 3);
                double tStatistic = Math.Round(result[0].As<double>(), 3);

                // Format the p-value
                string TtestString = tStatistic.ToString("0.000");
                string pValueString = pValue <= 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { TtestString, pValueString };

                return TestValue;
                // Display the results
                
            }

        }
    }
}
