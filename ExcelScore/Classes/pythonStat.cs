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
        




        
        public void InitPython()
        {
            string pythonDll = "";
            string username = Environment.UserName;
            


            try

            {
                
                if (username == "khaled")
                {
                    pythonDll = @"C:\Users\Khale\AppData\Local\Programs\Python\Python312\python312.dll";
                    Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", pythonDll);
                    PythonEngine.Initialize();
                }
                else if (username == "Esraa")
                {
                    //MessageBox.Show(username);
                    pythonDll = @"C:\Users\MMM\AppData\Local\Programs\Python\Python312\python312.dll";
                    Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", pythonDll);
                    PythonEngine.Initialize();
                }
                else if (username != "khaled")
                {
                    pythonDll = $@"C:\Users\{username}\AppData\Local\Programs\Python\Python312\python312.dll";
                    Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", pythonDll);
                    PythonEngine.Initialize();
                }



            }
            catch(Exception)
            {
                MessageBox.Show("Init error");
            }
            



            






            //catch (Exception) 
            //{
            //    pythonDll = @"C:\Users\Khale\AppData\Local\Programs\Python\Python312\python38.dll";
            //}


        }

        


        
        public AnovaTestResult newANOVAWithTukeyHSDNewDynamic(Parameter parameter)
        {
            try
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

                        if (groupValues.Count > 1)
                        {
                            allGroupLabels.AddRange(Enumerable.Repeat(groupLabel, groupValues.Count));
                            allValues.AddRange(groupValues);
                        }
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

                    string pValueString = pValueKruskalWallis < 0.001 ? "<0.001" : pValueKruskalWallis.ToString("0.000");

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

                            MessageBox.Show(pvalue.ToString());

                            string pAdjString = pvalue < 0.001 ? "<0.001" : pvalue.ToString("0.000");

                            
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
            catch (Exception)
            {
                return new AnovaTestResult
                {
                    TestValue = "0",
                    PValue = "0",
                    PairwiseComparisons = new List<string[]> { }
                };
            }

        }
        public AnovaTestResult KruskalWallisWithDunnDynamic(Parameter parameter)
        {
            try
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

                        if (groupValues.Count > 1)
                        {
                            allGroupLabels.AddRange(Enumerable.Repeat(groupLabel, groupValues.Count));
                            allValues.AddRange(groupValues);
                        }
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

                    string pValueString = pValueKruskalWallis < 0.001 ? "<0.001" : pValueKruskalWallis.ToString("0.000");

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
                            double.TryParse(dataTable.Rows[i][j].ToString(), out pvalue);








                            string pAdjString = pvalue < 0.001 ? "<0.001" : pvalue.ToString("0.000");

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
            catch(Exception)
            {
                return new AnovaTestResult
                {
                    TestValue = "0",
                    PValue = "0",
                    PairwiseComparisons = new List<string[]> { }
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
                dynamic np = Py.Import("numpy");
                dynamic scipyStats = Py.Import("scipy.stats");
                dynamic statsmodels = Py.Import("statsmodels.stats.multicomp");

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

                string pValueString = pValueANOVA < 0.001 ? "<0.001" : pValueANOVA.ToString("0.000");


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


                    //
                    //ManualTests manual = new ManualTests();
                    //string[] result = manual.getchi(parameter);
                    //chiSquareStatistic = double.Parse(result[0]);
                    //pValue = double.Parse(result[1]);

                    dynamic fisherResult = scipyStats.fisher_exact(tableArray);
                    pValue = fisherResult[1].As<double>();
                }

                else
                {
                    if(hasSmallExpectedCell)
                    {
                        return new string[] { "FET", "0"};

                        // Perform Monte Carlo simulation
                        //int numSimulations = 10000; // Adjust the number of simulations
                        //double[] simulatedChiSquares = new double[numSimulations];
                        //int rowSum = 0, colSum = 0;
                        //foreach (var row in contingencyTable)
                        //{
                        //    rowSum += row.Sum();
                        //}
                        //for (int i = 0; i < contingencyTable[0].Count; i++)
                        //{
                        //    foreach (var row in contingencyTable)
                        //    {
                        //        colSum += row[i];
                        //    }
                        //}

                        //Random rand = new Random();
                        //for (int i = 0; i < numSimulations; i++)
                        //{
                        //    // Generate a random table based on row and column sums
                        //    dynamic simulatedTable = new PyList();
                        //    for (int r = 0; r < contingencyTable.Count; r++)
                        //    {
                        //        dynamic simulatedRow = new PyList();
                        //        for (int c = 0; c < contingencyTable[0].Count; c++)
                        //        {
                        //            int simulatedValue = rand.Next(1, 10); // Random number generation logic to fill the table
                        //            simulatedRow.append(simulatedValue);
                        //        }
                        //        simulatedTable.append(simulatedRow);
                        //    }

                        //    // Compute chi-square statistic for the simulated table
                        //    dynamic simulatedChiSquare = scipyStats.chi2_contingency(simulatedTable, correction: false);
                        //    simulatedChiSquares[i] = simulatedChiSquare[0].As<double>();
                        //}

                        //// Calculate p-value
                        //double pValueMonteCarlo = simulatedChiSquares.Count(x => x >= chiSquareStatistic) / (double)numSimulations;
                        //pValue = Math.Round(pValueMonteCarlo, 3);

                        //return new string[] { chiSquareStatistic.ToString("0.000"), pValue.ToString()};
                    }
                    else
                    {
                        dynamic chiSquareResult = scipyStats.chi2_contingency(tableArray, correction: false);
                        chiSquareStatistic = chiSquareResult[0].As<double>();
                        pValue = chiSquareResult[1].As<double>();
                    }    
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


        public string chisquareCalcPairwise(List<List<int>> contingencyTable)
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
                    int a = 0;
                    int b = 0;
                    int c = 0;
                    int d = 0;

                    try
                    {
                        a = contingencyTable[0][0];
                    }
                    catch (Exception)
                    {
                        a = 0;
                    }


                    try
                    {
                        b = contingencyTable[0][1];
                    }
                    catch (Exception)
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
                    catch (Exception)
                    { d = 0; }

                    ManualTests manual = new ManualTests();
                    chiSquareStatistic = manual.CalculateChiSquare(a, b, c, d);
                    pValue = manual.FisherExactTest(a, b, c, d);

                }

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

                return pValueString;


            }
        }
        public List<string> ChisquarePairwise(Parameter parameter)
        {
            List<string> pairwise = new List<string>(); 
            if (parameter.IsGroup || parameter.NominalOrScale != "Nominal")
            {
                return new List<string> { "Parameter is not relevant for chi-square test", "" };
            }

            var allValues = parameter.GroupedParameterValues.Values.SelectMany(x => x).Distinct().ToList();
            var contingencyTable = new List<List<int>>();
            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();

            for (int i = 0; i < sortedKeys.Count; i++)
            {

                for (int j = i + 1; j < sortedKeys.Count; j++)
                {
                    double keyI = sortedKeys[i];
                    double keyJ = sortedKeys[j];


                    foreach (var value in allValues)
                    {
                        var categoryCountsi = new List<int>();
                        if (parameter.GroupedParameterValues.TryGetValue(keyI, out var valuesI))
                        {
                            // Count occurrences of 'value' in each list associated with keyI
                            categoryCountsi.Add(valuesI.Count(x => x == value));
                            
                        }
                        contingencyTable.Add(categoryCountsi);
                        

                        var categoryCountsj = new List<int>();
                        if (parameter.GroupedParameterValues.TryGetValue(keyJ, out var valuesJ))
                        {
                            // Count occurrences of 'value' in each list associated with keyJ
                            categoryCountsj.Add(valuesJ.Count(x => x == value));
                        }

                        contingencyTable.Add(categoryCountsj);

                        MessageBox.Show(contingencyTable[0][0].ToString());
                        MessageBox.Show(contingencyTable[0][0].ToString());

                        string pvalue = chisquareCalcPairwise(contingencyTable);

                        pairwise.Add(pvalue);
                    }
                    
                    

                }
            }


            return pairwise;
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
                    result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, use_continuity: true, method : "asymptotic", alternative: "two-sided");
                }
                else
                {
                    // Use the exact method
                    result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, use_continuity: false, method: "exact", alternative: "two-sided");
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


                string TtestString = tStatistic.ToString("0.000");
                TtestString = TtestString.Replace("-", "");

                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { TtestString, pValueString };

                return TestValue;
                // Display the results
                
            }

        }
        
        public string[] WilcoxonTest(List<double> AdataGroup1, List<double> AdataGroup2)
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

                // Perform the Wilcoxon signed-rank test
                dynamic result = scipyStats.wilcoxon(data1, data2);

                // Extract the p-value and test statistic from the result
                double pValue = Math.Round(result[1].As<double>(), 3);
                double wStatistic = Math.Round(result[0].As<double>(), 3);

                string WtestString = wStatistic.ToString("0.000");
                WtestString = WtestString.Replace("-", "");

                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { WtestString, pValueString };

                return TestValue;
            }
        }


        public string[] CochranQTest(List<List<double>> data)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic statsmodels = Py.Import("statsmodels.stats.contingency_tables");

                // Convert the list of lists to a 2D numpy array
                dynamic dataArray = np.array(data);

                // Perform the Cochran's Q test using statsmodels
                dynamic result = statsmodels.cochrans_q(dataArray);

                // Extract the p-value and test statistic from the result
                double pValue = Math.Round(result.pvalue.As<double>(), 3);
                double testStatistic = Math.Round(result.statistic.As<double>(), 3);

                // Format the p-value
                string testStatisticString = testStatistic.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { testStatisticString, pValueString };

                return TestValue;
            }
        }

        public string[] FRepeatedMeasures(List<List<double>> data)
        {
            using (Py.GIL())
            {
                // Import necessary Python libraries
                dynamic np = Py.Import("numpy");
                dynamic pd = Py.Import("pandas");
                dynamic pg = Py.Import("pingouin");

                // Convert the list of lists to a numpy array
                dynamic dataArray = np.array(data);

                // Get the number of subjects and periods
                int numSubjects = dataArray.shape[0];
                int numPeriods = dataArray.shape[1];

                // Dynamically create period labels based on the number of periods
                string[] periodLabels = new string[numPeriods];
                for (int j = 0; j < numPeriods; j++)
                {
                    periodLabels[j] = $"period{j + 1}";  // Labels like "period1", "period2", etc.
                }

                // Flatten data for easier conversion to a DataFrame
                List<double> flattenedData = new List<double>();
                List<int> subjects = new List<int>();
                List<string> periods = new List<string>();

                for (int i = 0; i < numSubjects; i++)
                {
                    for (int j = 0; j < numPeriods; j++)
                    {
                        flattenedData.Add(data[i][j]);
                        subjects.Add(i + 1);
                        periods.Add(periodLabels[j]);
                    }
                }

                // Create DataFrame
                var dataDict = new PyDict();
                dataDict["subject"] = pd.Series(subjects);
                dataDict["period"] = pd.Series(periods);
                dataDict["score"] = pd.Series(flattenedData);

                dynamic df = pd.DataFrame(dataDict);

                // Run the repeated measures ANOVA using pingouin
                dynamic anova = pg.rm_anova(data: df, dv: "score", within: "period", subject: "subject", detailed: true);

                // Extract the F-statistic and p-value
                double fValue = Math.Round(anova["F"][0].As<double>(), 3);
                double pValue = Math.Round(anova["p-unc"][0].As<double>(), 3);

                // Format the output
                string fValueString = fValue.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                return new string[] { fValueString, pValueString };
            }
        }



        public string[] RepeatedMeasuresAnova(List<List<double>> data)
        {
            using (Py.GIL())
            {
                // Import necessary Python libraries
                dynamic np = Py.Import("numpy");
                dynamic pd = Py.Import("pandas");
                dynamic sm = Py.Import("statsmodels.api");
                dynamic AnovaRM = Py.Import("statsmodels.stats.anova");

                // Convert the list of lists to a 2D numpy array
                dynamic dataArray = np.array(data);

                // Prepare the data for ANOVA
                var periods = data[0].Count;
                var subjects = data.Count;
                List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();

                for (int i = 0; i < subjects; i++)
                {
                    for (int j = 0; j < periods; j++)
                    {
                        rows.Add(new Dictionary<string, object>
                {
                    { "Subject", i + 1 },
                    { "Period", $"Period{j + 1}" },
                    { "Score", data[i][j] }
                });
                    }
                }

                dynamic df = pd.DataFrame.from_records(rows);

                // Perform repeated measures ANOVA using AnovaRM
                dynamic aovrm = AnovaRM(data: df, depvar: "Score", subject: "Subject", within: new List<string> { "Period" }).fit();

                // Extract F-statistic and p-value from the ANOVA table
                double fValue = Math.Round(aovrm.anova_table.loc["Period", "F Value"].As<double>(), 3);
                double pValue = Math.Round(aovrm.anova_table.loc["Period", "Pr > F"].As<double>(), 3);

                // Format the p-value
                string fValueString = fValue.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                // Return the results
                return new string[] { fValueString, pValueString };
            }
        }


        public string[] FriedmanTest(List<List<double>> data)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic stats = Py.Import("scipy.stats");

                // Convert the list of lists to a 2D numpy array

                List<double> Pre = data[0];
                List<double> follow = data[1];
                List<double> post = data[2];

                dynamic dataArrayPre = np.array(Pre);
                dynamic dataArrayfollow = np.array(follow);
                dynamic dataArraypost = np.array(post);


                // Perform the Friedman's test using scipy
                dynamic result = stats.friedmanchisquare(dataArrayPre , dataArrayfollow , dataArraypost);

                // Extract the p-value and test statistic from the result
                double pValue = Math.Round(result.pvalue.As<double>(), 3);
                double testStatistic = Math.Round(result.statistic.As<double>(), 3);

                // Format the p-value
                string testStatisticString = testStatistic.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                string[] TestValue = new string[] { testStatisticString, pValueString };

                return TestValue;
            }
        }

        public string[] FRepeatedMeasuresAnovaNEW(List<List<double>> data)
        {
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic pd = Py.Import("pandas");
                dynamic pingouin = Py.Import("pingouin");

                int numSubjects = data[0].Count; // Number of subjects, assuming each inner list has the same number of subjects
                int numPeriods = data.Count; // Number of periods (should be 3 in this case)

                // Create a list of subject IDs
                List<int> subjects = Enumerable.Range(1, numSubjects).ToList();

                // Initialize the DataFrame columns
                var dfData = new Dictionary<string, dynamic>
        {
            { "Subject", np.array(subjects) }, // First column for subject IDs
            { "Period1", np.array(data[0]) },   // Second column for Period 1 data
            { "Period2", np.array(data[1]) },   // Third column for Period 2 data
            { "Period3", np.array(data[2]) }    // Fourth column for Period 3 data
        };

                // Create the DataFrame in wide format
                dynamic df = pd.DataFrame(dfData);

                // Melt the DataFrame to long format for pingouin
                dynamic df_long = pd.melt(df, id_vars: "Subject", var_name: "Time", value_name: "Score");

                // Perform repeated measures ANOVA using pingouin
                dynamic anova_result = pingouin.rm_anova(data: df_long, dv: "Score", within: "Time", subject: "Subject", detailed: true);

                // Extract the Greenhouse-Geisser corrected p-value, test statistic, and F-value
                double testStatistic = Math.Round(anova_result.loc[0, "F"].As<double>(), 3);
                double pValue = Math.Round(anova_result.loc[0, "p-GG-corr"].As<double>(), 3);

                // Format the p-value
                string testStatisticString = testStatistic.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                // Return the formatted test statistic and p-value
                string[] TestValue = new string[] { testStatisticString, pValueString };

                return TestValue;
            }
        }





    }
}
