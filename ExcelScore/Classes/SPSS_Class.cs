using DocumentFormat.OpenXml.Vml.Office;
using SpssLib.DataReader;
using SpssLib.SpssDataset;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SppComApiLib;
using System.Diagnostics;
using spsswin;
using Python.Runtime;

namespace ExcelScore.Classes
{
    public class SPSS_Class
    {
        public static List<double?> ConvertToNullable(List<double> values)
        {
            return values.Select(x => (double?)x).ToList();
        }
        public static List<double> CreateSerial(int totalN)
        {
            return Enumerable.Range(1, totalN).Select(x => (double)x).ToList();
        }
#nullable enable
        public static (List<double?>, string, Dictionary<double, string>?) ConvertToSPSSFormat(
    List<double?> values,  // ✅ Change to List<double?>
    string measurementLevel,
    Dictionary<double, string>? valueLabels = null)
        {
            // Ensure that the dictionary keeps double keys instead of double?
            Dictionary<double, string>? convertedLabels = valueLabels != null
                ? new Dictionary<double, string>(valueLabels)
                : null;

            return (values, measurementLevel, convertedLabels);
        }



        public static List<double> CreateGX(int totalN)
        {
            return Enumerable.Repeat(1.0, totalN).ToList();
        }

        static string GetSaveFilePath()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "SPSS Files (*.sav)|*.sav";
                saveFileDialog.Title = "Save SPSS File";
                saveFileDialog.DefaultExt = "sav";
                saveFileDialog.AddExtension = true;
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            return ""; 
        }
        public static void InsertMultipleVariablesnew(Dictionary<string, (List<double?>, string, Dictionary<double, string>?)> variablesData)
        {
            // Open Save File Dialog
            string outputFile = GetSaveFilePath();
            if (string.IsNullOrEmpty(outputFile)) return; // Exit if no file selected

            // Find the maximum row count across all variables
            int maxRows = variablesData.Max(entry => entry.Value.Item1.Count);  // Corrected access to Values

            // Normalize all lists to match maxRows by padding with null
            var adjustedData = variablesData.ToDictionary(
                entry => entry.Key,
                entry => (
                    entry.Value.Item1.Concat(Enumerable.Repeat<double?>(null, maxRows - entry.Value.Item1.Count)).ToList(),
                    entry.Value.Item2,
                    entry.Value.Item3
                )
            );

            // Convert input data into SPSS-compatible variables
            var variables = new List<Variable>();
            

            foreach (var entry in adjustedData)
            {

                var (values, measurementLevel, valueLabels) = entry.Value;
                variables.Add(new Variable
                {
                    Name = entry.Key,
                    Type = DataType.Numeric,
                    Width = 10,
                    PrintFormat = new OutputFormat(FormatType.F, 8, 2),
                    WriteFormat = new OutputFormat(FormatType.F, 8, 2),
                    MissingValueType = MissingValueType.NoMissingValues,
                    Alignment = Alignment.Centre,

                });

                foreach (var variable in variables)
                {
                    if (measurementLevel.ToLower() == "nominal" && valueLabels != null)
                    {
                        variable.ValueLabels = valueLabels;
                        variable.MeasurementType = MeasurementType.Nominal;
                    }
                }


            }

            // Prepare SPSS file options
            var options = new SpssOptions();

            using (FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
            {
                using (var writer = new SpssWriter(fileStream, variables, options))
                {
                    for (int i = 0; i < maxRows; i++)
                    {
                        var newRecord = writer.CreateRecord();

                        int col = 0;
                        foreach (var entry in adjustedData)
                        {
                            List<double?> values = entry.Value.Item1; // Corrected access
                            newRecord[col] = i < values.Count ? values[i] : null;
                            col++;
                        }

                        writer.WriteRecord(newRecord);
                    }

                    writer.EndFile();
                }
            }


        }

        public static void InsertMultipleVariablesnewa(Dictionary<string, (List<double?>, string, Dictionary<double, string>?)> variablesData)
        {
            // Open Save File Dialog
            string outputFile = GetSaveFilePath();
            if (string.IsNullOrEmpty(outputFile)) return; // Exit if no file selected

            // Find the maximum row count across all variables
            int maxRows = variablesData.Max(entry => entry.Value.Item1.Count);  // Corrected access to Values

            // Normalize all lists to match maxRows by padding with null
            var adjustedData = variablesData.ToDictionary(
                entry => entry.Key,
                entry => (
                    entry.Value.Item1.Concat(Enumerable.Repeat<double?>(null, maxRows - entry.Value.Item1.Count)).ToList(),
                    entry.Value.Item2,
                    entry.Value.Item3
                )
            );

            // Convert input data into SPSS-compatible variables
            var variables = new List<Variable>();


            foreach (var entry in adjustedData)
            {
                var (values, measurementLevel, valueLabels) = entry.Value;

                // Create a new variable for this entry
                var newVariable = new Variable
                {
                    Name = entry.Key,
                    Type = DataType.Numeric,
                    Width = 10,
                    PrintFormat = new OutputFormat(FormatType.F, 8, 2),
                    WriteFormat = new OutputFormat(FormatType.F, 8, 2),
                    MissingValueType = MissingValueType.NoMissingValues,
                    Alignment = Alignment.Centre,
                };

                // Set measurement level and value labels correctly
                if (measurementLevel.ToLower() == "nominal" && valueLabels != null)
                {
                    newVariable.ValueLabels = valueLabels;
                    newVariable.MeasurementType = MeasurementType.Nominal;
                }
                else if (measurementLevel.ToLower() == "scale")
                {
                    newVariable.MeasurementType = MeasurementType.Scale;
                }

                variables.Add(newVariable); // Add variable to list
            }

            // Prepare SPSS file options
            var options = new SpssOptions();

            using (FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
            {
                using (var writer = new SpssWriter(fileStream, variables, options))
                {
                    for (int i = 0; i < maxRows; i++)
                    {
                        var newRecord = writer.CreateRecord();

                        int col = 0;
                        foreach (var entry in adjustedData)
                        {
                            List<double?> values = entry.Value.Item1; // Corrected access
                            newRecord[col] = i < values.Count ? values[i] : null;
                            col++;
                        }

                        writer.WriteRecord(newRecord);
                    }

                    writer.EndFile();
                }
            }


        }

        public static string GetSpssFilePath()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "SPSS Data Files (*.sav)|*.sav";
                openFileDialog.Title = "Select SPSS Data File";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return "";
        }

        public static string GetOutputFilePath()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "SPSS Output Files (*.spo)|*.spo";
                saveFileDialog.Title = "Save SPSS Output File";
                saveFileDialog.DefaultExt = "spo";
                saveFileDialog.AddExtension = true;
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            return "";
        }

        public static void RunSpssSyntax()
        {

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "SPSS Files (*.sav)|*.sav";
                openFileDialog.Title = "Select an SPSS Data File";

                if (openFileDialog.ShowDialog() != DialogResult.OK) return;

                string spssDataFilePath = openFileDialog.FileName;
                string spssExePath = @"C:\Program Files (x86)\IBM\SPSS\Statistics\23\stats.exe";  // Adjust path if needed
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string outputSpoPath = Path.Combine(desktopPath, "SPSS_Output.spo");
                string syntaxFilePath = Path.Combine(Path.GetTempPath(), "temp_syntax.sps");

                // Step 2: Create SPSS syntax to run on the file
                string syntax = $@"
GET FILE='{spssDataFilePath}'.
FREQUENCIES VARIABLES=ALL.
SAVE OUTFILE='{spssDataFilePath}'.
EXECUTE.
";

                File.WriteAllText(syntaxFilePath, syntax); // Save syntax to file

                // Step 3: Run SPSS in hidden batch mode
                Process process = new Process();
                process.StartInfo.FileName = spssExePath;
                process.StartInfo.Arguments = $"/run /syntax=\"{syntaxFilePath}\" /output=\"{outputSpoPath}\"";
                process.StartInfo.Verb = "runas";  // Requests admin permission
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                process.Start();
                process.WaitForExit(); // Wait for SPSS to finish execution

                // Step 4: Confirm output saved
                MessageBox.Show("SPSS syntax executed successfully!\nOutput saved at: " + outputSpoPath);
            }


        }







    }

}
