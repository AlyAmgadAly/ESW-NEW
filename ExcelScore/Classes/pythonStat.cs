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
using Accord.Collections;

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
        public void RunSpssSyntaxWithPythonNet()
        {
            InitPython(); // Ensure Python is initialized

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "SPSS Files (*.sav)|*.sav";
                openFileDialog.Title = "Select an SPSS Data File";

                if (openFileDialog.ShowDialog() != DialogResult.OK) return;

                string spssDataFilePath = openFileDialog.FileName;
                string outputSpoPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "output.spo");

                using (Py.GIL())  // Acquire Python Global Interpreter Lock
                {
                    dynamic spss = Py.Import("spss");      // Import SPSS Python module
                    dynamic spssaux = Py.Import("spssaux"); // Import SPSS auxiliary module

                    // Define SPSS syntax to execute
                    string syntax = $@"
GET FILE='{spssDataFilePath}'.
FREQUENCIES VARIABLES=ALL.
EXECUTE.
";

                    // Run SPSS syntax
                    spss.Submit(syntax);

                    // Save the output as a .spo file
                    spssaux.CreateOutputDoc(outputSpoPath, visible: false);
                }

                MessageBox.Show("SPSS syntax executed successfully! Output saved at: " + outputSpoPath);
            }
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
                    parameter.Isfisher = true;

                    dynamic fisherResult = scipyStats.fisher_exact(tableArray);
                    pValue = fisherResult[1].As<double>();
                }

                else
                {
                    if(hasSmallExpectedCell)
                    {
                        //parameter.IsMonteCarlo = true;

                        //dynamic numpy = Py.Import("numpy");
                        //dynamic special = Py.Import("scipy.special");
                        //Random rnd = new Random();

                        //// Convert .NET contingency table to flat list
                        //int[][] rawArray = contingencyTable.Select(row => row.ToArray()).ToArray();
                        //int R = rawArray.Length;
                        //int C = rawArray[0].Length;
                        //int N = rawArray.Sum(row => row.Sum());

                        //// Row and column margins
                        //int[] rowSum = new int[R];
                        //int[] colSum = new int[C];
                        //for (int r = 0; r < R; r++)
                        //    for (int c = 0; c < C; c++)
                        //    {
                        //        rowSum[r] += rawArray[r][c];
                        //        colSum[c] += rawArray[r][c];
                        //    }

                        //// Helper to compute log-factorial using gammaln
                        //double LogFact(int n) => special.gammaln(n + 1).As<double>();

                        //// Log-probability of a table (hypergeometric constant dropped)
                        //double LogProb(int[,] table)
                        //{
                        //    double sum = 0;
                        //    for (int r = 0; r < R; r++)
                        //        for (int c = 0; c < C; c++)
                        //            sum += LogFact(table[r, c]);
                        //    return sum;
                        //}

                        //// Observed log-probability
                        //int[,] observedTable = new int[R, C];
                        //for (int r = 0; r < R; r++)
                        //    for (int c = 0; c < C; c++)
                        //        observedTable[r, c] = rawArray[r][c];
                        //double logProbObs = LogProb(observedTable);

                        //// Step 1: Create "population" list of cell positions
                        //List<int> rows = new();
                        //List<int> cols = new();
                        //for (int r = 0; r < R; r++)
                        //    for (int c = 0; c < C; c++)
                        //        for (int k = 0; k < rawArray[r][c]; k++)
                        //        {
                        //            rows.Add(r);
                        //            cols.Add(c);
                        //        }

                        //// Step 2: Monte Carlo simulation
                        //int numSimulations = 10000;
                        //int countExtreme = 0;
                        //double probObs = Math.Round(Math.Exp(logProbObs), 6); // SPSS shows test value as actual P(table)
                        //List<double> simProbs = new();

                        //for (int sim = 0; sim < numSimulations; sim++)
                        //{
                        //    // Shuffle column labels
                        //    var shuffledCols = cols.OrderBy(_ => rnd.Next()).ToArray();

                        //    int[,] simTable = new int[R, C];
                        //    for (int i = 0; i < N; i++)
                        //    {
                        //        int r = rows[i];
                        //        int c = shuffledCols[i];
                        //        simTable[r, c]++;
                        //    }

                        //    double logP = LogProb(simTable);
                        //    if (logP <= logProbObs + 1e-10) countExtreme++;
                        //    simProbs.Add(Math.Exp(logP));
                        //}

                        //// Step 3: p-value
                        //double pValuea = Math.Round((double)countExtreme / numSimulations, 3);
                        //string pStr = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                        //// Step 4: Return test value = probability of observed table
                        //string testStr = probObs.ToString("0.000000");

                        return new string[] { "MC", "0"};








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
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                MessageBox.Show(testStatisticString);
                MessageBox.Show(pValueString);

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

                bool exists = ExactPairs.ContainsKey(minGroup);

                dynamic result;


                if (combinedSize <= 40 || exists)
                {
                    result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, use_continuity: false, method: "exact", alternative: "two-sided");
                }

                // Choose the method based on combined sample size
                else
                {
                    // Use the asymptotic method
                    result = scipyStats.mannwhitneyu(dataGroup1, dataGroup2, use_continuity: true, method : "asymptotic", alternative: "two-sided");
                }


                // Extract the p-value and U-statistic from the result
                double uStatistic = Math.Round(result[0].As<double>(), 3);
                double pValue = Math.Round(result[1].As<double>(), 3);

                // Format the p-value
                string UtestString = uStatistic.ToString("0.000");
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                MessageBox.Show(UtestString);
                MessageBox.Show(pValueString);

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

                // Compute the differences between paired values (d = x - y)
                dynamic differences = data1 - data2;

                // Get the absolute values of the differences
                dynamic absDifferences = np.abs(differences);

                // Rank the absolute differences (use scipy.stats.rankdata to handle ties)
                dynamic ranks = scipyStats.rankdata(absDifferences);

                // Calculate the signed rank sum
                double signedRankSum = 0;
                int length = differences.shape[0]; // Get the length of the differences
                for (int i = 0; i < length; i++)
                {
                    double rank = ranks[i].As<double>();  // Convert PyObject to double

                    // If the difference is negative, subtract the rank from the sum (negative ranks)
                    if (differences[i] < 0)
                    {
                        signedRankSum -= rank;
                    }
                    // If the difference is positive, add the rank to the sum (positive ranks)
                    else if (differences[i] > 0)
                    {
                        signedRankSum += rank;
                    }
                }

                // Perform the Wilcoxon signed-rank test using scipy
                dynamic result = scipyStats.wilcoxon(data1, data2);

                // Extract the p-value from the result
                double pValue = Math.Round(result[1].As<double>(), 3);

                // The statistic from the scipy result (wilcoxon result[0]) might be the opposite sign of what you expect.
                double wStatistic = Math.Round(signedRankSum, 3); // Use the signed rank sum for the test statistic

                string WtestString = wStatistic.ToString("0.000");

                // Format the result string: preserving the negative sign if it exists
                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                MessageBox.Show(WtestString);
                MessageBox.Show(pValueString);

                // Return the result as an array
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



        public AnovaTestResult RepeatedMeasuresAnova_WO_GREEN(List<Parameter> parameters)
        {
           // try
            //{
                using (Py.GIL())
                {
                    dynamic np = Py.Import("numpy");
                    dynamic pandas = Py.Import("pandas");
                    dynamic statsmodelsAnova = Py.Import("statsmodels.stats.anova");

                    // Step 1: Prepare data for pandas DataFrame
                    var subjects = Enumerable.Range(1, parameters[0].ParameterValues.Count).ToList(); // Subjects 1, 2, ...
                    var values = new List<double>();
                    var periods = new List<string>();
                    var subjectIds = new List<int>();

                    // Loop through each parameter (Period1, Period2, etc.)
                    foreach (var parameter in parameters)
                    {
                        string periodName = parameter.Name; // Period name (e.g., "Period1")
                        var parameterValues = parameter.ParameterValues;

                        for (int i = 0; i < parameterValues.Count; i++)
                        {
                            values.Add(parameterValues[i]);
                            periods.Add(periodName);
                            subjectIds.Add(subjects[i]);
                        }
                    }

                    // Step 2: Create pandas DataFrame
                    dynamic dataDict = new PyDict
                    {
                        ["Subject"] = subjectIds.ToPython(),
                        ["Period"] = periods.ToPython(),
                        ["Value"] = values.ToPython()
                    };
                    dynamic df = pandas.DataFrame(dataDict);

                    // Step 3: Create the PyList for "within"
                    PyList withinList = new PyList();
                    withinList.Append("Period".ToPython());

                    // Step 4: Run repeated-measures ANOVA
                    dynamic anovaModel = statsmodelsAnova.AnovaRM(
                        data: df,
                        depvar: "Value",
                        subject: "Subject",
                        within: withinList
                    ).fit();

                    dynamic anovaTable = anovaModel.anova_table;

                    // Step 5: Extract results
                    double fValue = anovaTable.loc["Period"]["F Value"].As<double>();
                    double pValue = anovaTable.loc["Period"]["Pr > F"].As<double>();

                    string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                    // Return results

                    MessageBox.Show(fValue.ToString());
                    MessageBox.Show(pValueString);

                    return new AnovaTestResult
                    {
                        TestValue = fValue.ToString("0.000"),
                        PValue = pValueString,
                        PairwiseComparisons = new List<string[]>()
                    };
                }
            //}
            //catch (Exception ex)
            //{
            //    return new AnovaTestResult
            //    {
            //        TestValue = "0",
            //        PValue = "0",
            //        PairwiseComparisons = new List<string[]>()
            //    };
            //}
        }


        public AnovaTestResult RepeatedMeasuresAnovaBoth(List<Parameter> parameters)
        {
            try
            {
                using (Py.GIL())
                {
                    dynamic np = Py.Import("numpy");
                    dynamic pandas = Py.Import("pandas");
                    dynamic pingouin = Py.Import("pingouin");

                    // Step 1: Prepare data for pandas DataFrame
                    var subjects = Enumerable.Range(1, parameters[0].ParameterValues.Count).ToList(); // Subjects 1, 2, ...
                    var values = new List<double>();
                    var periods = new List<string>();
                    var subjectIds = new List<int>();

                    // Loop through each parameter (Period1, Period2, etc.)
                    foreach (var parameter in parameters)
                    {
                        string periodName = parameter.Name; // Period name (e.g., "Period1")
                        var parameterValues = parameter.ParameterValues;

                        for (int i = 0; i < parameterValues.Count; i++)
                        {
                            values.Add(parameterValues[i]);
                            periods.Add(periodName);
                            subjectIds.Add(subjects[i]);
                        }
                    }

                    // Step 2: Create pandas DataFrame
                    dynamic dataDict = new PyDict
                    {
                        ["Subject"] = subjectIds.ToPython(),
                        ["Period"] = periods.ToPython(),
                        ["Value"] = values.ToPython()
                    };
                    dynamic df = pandas.DataFrame(dataDict);

                    // Step 3: Perform Mauchly's Test of Sphericity
                    dynamic mauchlyResult = pingouin.sphericity(data: df, dv: "Value", within: "Period", subject: "Subject");

                    double wValue = mauchlyResult[1].As<double>(); // Extract Mauchly's W
                    double pValueMauchly = mauchlyResult[0].As<double>(); // Extract p-value of Mauchly's Test
                    bool isSphericityViolated = pValueMauchly < 0.05; // Check if sphericity is violated

                    // Decide correction method based on Mauchly's test
                    string correction = isSphericityViolated ? "auto" : null; // Use Greenhouse-Geisser if violated

                    // Output Mauchly's Test Results
                    //MessageBox.Show($"Mauchly's W: {wValue:0.000}");
                    //MessageBox.Show($"Mauchly's p-value: {pValueMauchly:0.000}");
                    //MessageBox.Show(isSphericityViolated
                    //    ? "Sphericity violated, Greenhouse-Geisser correction applied."
                    //    : "Sphericity assumed, no correction applied.");

                    // Step 4: Perform Repeated Measures ANOVA
                    dynamic anovaResults = pingouin.rm_anova(
                        data: df,
                        dv: "Value",            // Dependent variable
                        within: "Period",       // Within-subject variable
                        subject: "Subject",     // Subject identifier
                        correction: correction, // Conditional correction based on Mauchly's test
                        detailed: true          // Include detailed results
                    );

                    // Extract ANOVA results
                    dynamic periodRow = anovaResults.loc[anovaResults["Source"].eq("Period")];
                    double fValue = periodRow["F"].iloc[0].As<double>();
                    double pValue = isSphericityViolated
                        ? periodRow["p-GG-corr"].iloc[0].As<double>() // Greenhouse-Geisser corrected p-value
                        : periodRow["p-unc"].iloc[0].As<double>();    // Sphericity assumed p-value

                    string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                    // Return results
                    //MessageBox.Show($"F Value: {fValue:0.000}");
                    //MessageBox.Show($"p Value: {pValueString}");


                    dynamic uniquePeriods = df["Period"].unique();
                    List<string> periodOrder = new List<string>();

                    foreach (dynamic period in uniquePeriods.tolist())
                    {
                        periodOrder.Add(period.As<string>());
                    }

                    // Step 5: Perform Pairwise Comparisons (Bonferroni adjustment)
                    dynamic pairwiseResults = pingouin.pairwise_ttests(
         data: df,
         dv: "Value",         // Dependent variable
         within: "Period",    // Within-subject variable
         subject: "Subject",  // Subject identifier
         padjust: "bonf"      // Bonferroni correction for confidence intervals
     );

                    // Extract pairwise comparison results
                    var pairwiseComparisons = new List<string[]>();
                    foreach (var row in pairwiseResults.itertuples())
                    {
                        string group1 = row[2].As<string>(); // Group A
                        string group2 = row[3].As<string>(); // Group B
                        double pCorr = row[10].As<double>();  // Corrected p-value ('p-corr')
                                                              //double pCorr = row._p_corr.As<double>();
                        string comparison = pCorr < 0.001 ? "<0.001" : pCorr.ToString("0.000");
                        pairwiseComparisons.Add(new[] { group1, group2, comparison });
                    }

                    var sortedComparisons = new List<string[]>();
                    var addedComparisons = new HashSet<string>();

                    foreach (string period in periodOrder)
                    {
                        var matches = pairwiseComparisons
                            .Where(x => x.Contains(period))
                            .OrderBy(x => periodOrder.IndexOf(x.First(p => p != period)))
                            .ToList();

                        foreach (var match in matches)
                        {
                            string key = string.Join("-", match.OrderBy(p => periodOrder.IndexOf(p)));
                            if (!addedComparisons.Contains(key))
                            {
                                sortedComparisons.Add(match);
                                addedComparisons.Add(key);
                            }
                        }
                    }

                    // Replace the old list with the sorted one
                    


                    return new AnovaTestResult
                    {
                        TestValue = fValue.ToString("0.000"),
                        PValue = pValueString,
                        PairwiseComparisons = sortedComparisons
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


        public AnovaTestResult RepeatedMeasuresAnovaGreen(List<Parameter> parameters)
        {
            
            using (Py.GIL())
            {
                dynamic np = Py.Import("numpy");
                dynamic pandas = Py.Import("pandas");
                dynamic pingouin = Py.Import("pingouin");

                // Step 1: Prepare data for pandas DataFrame
                var subjects = Enumerable.Range(1, parameters[0].ParameterValues.Count).ToList(); // Subjects 1, 2, ...
                var values = new List<double>();
                var periods = new List<string>();
                var subjectIds = new List<int>();

                // Loop through each parameter (Period1, Period2, etc.)
                foreach (var parameter in parameters)
                {
                    string periodName = parameter.Name; // Period name (e.g., "Period1")
                    var parameterValues = parameter.ParameterValues;

                    for (int i = 0; i < parameterValues.Count; i++)
                    {
                        values.Add(parameterValues[i]);
                        periods.Add(periodName);
                        subjectIds.Add(subjects[i]);
                    }
                }

                // Step 2: Create pandas DataFrame
                dynamic dataDict = new PyDict
                {
                    ["Subject"] = subjectIds.ToPython(),
                    ["Period"] = periods.ToPython(),
                    ["Value"] = values.ToPython()
                };
                dynamic df = pandas.DataFrame(dataDict);

                // Step 3: Perform Mauchly's Test of Sphericity
                dynamic mauchlyResult = pingouin.sphericity(data: df, dv: "Value", within: "Period", subject: "Subject");

                double wValue = mauchlyResult[1].As<double>(); // Extract Mauchly's W
                double pValueMauchly = mauchlyResult[0].As<double>(); // Extract p-value of Mauchly's Test
                bool isSphericityViolated = pValueMauchly < 0.05; // Check if sphericity is violated

                // Decide correction method based on Mauchly's test
                string correction = isSphericityViolated ? "GG" : null; // Use Greenhouse-Geisser if violated

                // Output Mauchly's Test Results
                MessageBox.Show($"Mauchly's W: {wValue:0.000}");
                MessageBox.Show($"Mauchly's p-value: {pValueMauchly:0.000}");
                MessageBox.Show(isSphericityViolated
                    ? "Sphericity violated, Greenhouse-Geisser correction applied."
                    : "Sphericity assumed, no correction applied.");

                // Step 4: Perform Repeated Measures ANOVA
                dynamic anovaResults = pingouin.rm_anova(
                    data: df,
                    dv: "Value",            // Dependent variable
                    within: "Period",       // Within-subject variable
                    subject: "Subject",     // Subject identifier
                    correction: correction, // Conditional correction based on Mauchly's test
                    detailed: true          // Include detailed results
                );

                // Extract ANOVA results
                dynamic periodRow = anovaResults.loc[anovaResults["Source"].eq("Period")];
                double fValue = periodRow["F"].iloc[0].As<double>();
                double pValue = isSphericityViolated
                    ? periodRow["p-GG-corr"].iloc[0].As<double>() // Greenhouse-Geisser corrected p-value
                    : periodRow["p-unc"].iloc[0].As<double>();    // Sphericity assumed p-value

                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                // Step 5: Perform Pairwise Comparisons (Bonferroni adjustment)
                dynamic pairwiseResults = pingouin.pairwise_ttests(
     data: df,
     dv: "Value",         // Dependent variable
     within: "Period",    // Within-subject variable
     subject: "Subject",  // Subject identifier
     padjust: "bonf"      // Bonferroni correction for confidence intervals
 );

                // Extract pairwise comparison results
                var pairwiseComparisons = new List<string[]>();
                foreach (var row in pairwiseResults.itertuples())
                {
                    string group1 = row[2].As<string>(); // Group A
                    string group2 = row[3].As<string>(); // Group B
                    double pCorr = row[9].As<double>();  // Corrected p-value ('p-corr')
                    string comparison = $"{group1} vs {group2}: p = {(pCorr < 0.001 ? "<0.001" : pCorr.ToString("0.000"))}";
                    pairwiseComparisons.Add(new[] { group1, group2, comparison });
                }

                // Display pairwise comparison results
                //foreach (var comp in pairwiseComparisons)
                //{
                //    MessageBox.Show(comp[2]); // Display each pairwise comparison result
                //}

                // Return results
                return new AnovaTestResult
                {
                    TestValue = fValue.ToString("0.000"),
                    PValue = pValueString,
                    PairwiseComparisons = pairwiseComparisons
                };
            }
        }

//        public string[] FriedmanTestNew(List<List<double>> data)
//        {
//            using (Py.GIL())
//            {
//                dynamic np = Py.Import("numpy");
//                dynamic stats = Py.Import("scipy.stats");

//                // Convert the List<List<double>> to a Python list of NumPy arrays
//                var pyGroupsData = new PyList();
//                foreach (var groupData in data)
//                {
//                    var pyGroupArray = np.array(groupData.ToArray());
//                    pyGroupsData.Append(pyGroupArray);
//                }

//                // Prepare the Python script
//                string pythonScript = @"
//result_friedman = stats.friedmanchisquare(*groups_data)
//";

//                // Create Python dictionaries for variables
//                dynamic globals = new PyDict();
//                dynamic locals = new PyDict();

//                globals["stats"] = stats;
//                globals["groups_data"] = pyGroupsData;

//                // Execute the Python script
//                PythonEngine.Exec(pythonScript, locals, globals);

//                // Retrieve the results from the Python script
//                dynamic resultFriedman = globals["result_friedman"];
//                double testStatistic = Math.Round(resultFriedman.statistic.As<double>(), 3);
//                double pValue = Math.Round(resultFriedman.pvalue.As<double>(), 3);

//                // Format the output
//                string testStatisticString = testStatistic.ToString("0.000");
//                string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

               

//                return new string[] { testStatisticString, pValueString };
//            }
//        }


        public AnovaTestResult PerformFriedmanWithDunnTest(List<List<double>> data)
        {
            try
            {
                using (Py.GIL())
                {
                    dynamic np = Py.Import("numpy");
                    dynamic stats = Py.Import("scipy.stats");

                    // Convert the List<List<double>> to a Python-compatible NumPy array
                    int nPeriods = data.Count;
                    int nSubjects = data[0].Count; // Assuming all periods have the same number of subjects

                    // Create a 2D NumPy array for periods (rows) and subjects (columns)
                    dynamic pyData = np.array(data.Select(period => np.array(period.ToArray())).ToArray());

                    // Perform the Friedman test
                    string pythonFriedmanScript = @"
result_friedman = stats.friedmanchisquare(*groups_data)
";
                    dynamic globals = new PyDict();
                    dynamic locals = new PyDict();
                    globals["stats"] = stats;
                    globals["groups_data"] = pyData;

                    PythonEngine.Exec(pythonFriedmanScript, locals, globals);

                    // Retrieve Friedman test result
                    dynamic resultFriedman = globals["result_friedman"];
                    double testStatistic = Math.Round(resultFriedman.statistic.As<double>(), 3);
                    double pValue = Math.Round(resultFriedman.pvalue.As<double>(), 3);

                    // If Friedman test is not significant, return the result
                    if (pValue >= 0.05)
                    {
                        return new AnovaTestResult
                        {
                            TestValue = testStatistic.ToString("0.000"),
                            PValue = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000"),
                            PairwiseComparisons = new List<string[]> { }
                        };
                    }

                    // Dunn's test for post-hoc analysis
                    string pythonDunnScript = @"
from scipy.stats import rankdata, norm
import numpy as np

# Rank data for each subject (across periods)
ranks = np.array([rankdata(row) for row in groups_data.T]).T

# Pairwise comparisons between periods
n = groups_data.shape[1]  # Number of subjects
k = groups_data.shape[0]  # Number of periods
comparisons = []
for i in range(k):
    for j in range(i + 1, k):
        # Calculate the difference in mean ranks for the pair
        mean_rank_i = np.mean(ranks[i])
        mean_rank_j = np.mean(ranks[j])
        diff = mean_rank_i - mean_rank_j
        se = np.sqrt(k * (k + 1) / (6.0 * n))  # Standard error for Dunn's test
        z = diff / se
        p_value = 2 * (1 - norm.cdf(abs(z)))  # Two-tailed test
        comparisons.append((i, j, diff, z, p_value))

comparisons
";
                    globals["np"] = np;

                    // Execute the Dunn script
                    PythonEngine.Exec(pythonDunnScript, locals, globals);

                    // Retrieve the pairwise comparison results
                    dynamic comparisons = globals["comparisons"];

                    // Format the output
                    List<string[]> pairwiseResults = new List<string[]>();
                    foreach (dynamic comparison in comparisons)
                    {
                        int i = comparison[0].As<int>();
                        int j = comparison[1].As<int>();
                        double diff = Math.Round(comparison[2].As<double>(), 3);
                        double z = Math.Round(comparison[3].As<double>(), 3);
                        double p = Math.Round(comparison[4].As<double>(), 5); // P-value to 5 decimal places
                        string pairwisep = p < 0.001 ? "<0.001" : p.ToString("0.000");
                        pairwiseResults.Add(new string[] { $"Period {i + 1} vs Period {j + 1}" , pairwisep });
                    }

                    return new AnovaTestResult
                    {
                        TestValue = testStatistic.ToString("0.000"),
                        PValue = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000"),
                        PairwiseComparisons = pairwiseResults
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







        public string[] McNemarTest(List<List<int>> contingencyTable)
        {
            try
            {
                using (Py.GIL())
                {
                    dynamic statsmodels = Py.Import("statsmodels.stats.contingency_tables");
                    dynamic np = Py.Import("numpy");

                    // Ensure the contingency table is 2x2
                    if (contingencyTable.Count != 2 || contingencyTable[0].Count != 2 || contingencyTable[1].Count != 2)
                    {
                        return new string[] { "Invalid table", "0" };
                    }

                    // Extract values from the provided contingency table
                    int a = contingencyTable[0][0];
                    int b = contingencyTable[0][1];
                    int c = contingencyTable[1][0];
                    int d = contingencyTable[1][1];

                    // Create a NumPy 2x2 array in the correct format: { { a, b }, { c, d } }
                    dynamic tableArray = np.array(new int[,] { { a, b }, { c, d } });

                    // Perform McNemar's test
                    dynamic result = statsmodels.mcnemar(tableArray, exact: false, correction: false);

                    // Extract and round the p-value and test statistic
                    double pValue = Math.Round(result.pvalue.As<double>(), 3);
                    double testStatistic = Math.Round(result.statistic.As<double>(), 3);

                    // Format the p-value
                    string testStatisticString = testStatistic.ToString("0.000");
                    string pValueString = pValue < 0.001 ? "<0.001" : pValue.ToString("0.000");

                    return new string[] { testStatisticString, pValueString };
                }
            }
            catch (Exception)
            {
                return new string[] { "NA", "0" };
            }
        }
















    }
}
