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



        //public static void InsertMultipleVariables(Dictionary<string, (List<double?> Values, string MeasurementLevel, Dictionary<double, string>? ValueLabels)> variablesData)
        //{
        //    // Open Save File Dialog
        //    string outputFile = GetSaveFilePath();
        //    if (string.IsNullOrEmpty(outputFile)) return; // Exit if no file selected

        //    // Convert input data into SPSS-compatible variables
        //    var variables = new List<Variable>();

        //    foreach (var entry in variablesData)
        //    {
        //        var (values, measurementLevel, valueLabels) = entry.Value;

        //        var variable = new Variable
        //        {
        //            Name = entry.Key, // Variable name from dictionary key
        //            Type = DataType.Numeric,
        //            Width = 10,
        //            PrintFormat = new OutputFormat(FormatType.F, 8, 2),
        //            WriteFormat = new OutputFormat(FormatType.F, 8, 2),
        //            MissingValueType = MissingValueType.NoMissingValues, 
        //            Alignment = Alignment.Centre
        //        };

        //        // Set measurement level (Nominal or Scale)
        //        if (measurementLevel.ToLower() == "nominal" && valueLabels != null)
        //        {
        //            variable.ValueLabels = valueLabels;
        //        }

        //        variables.Add(variable);
        //    }

        //    // Prepare SPSS file options
        //    var options = new SpssOptions();

        //    using (FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
        //    {
        //        using (var writer = new SpssWriter(fileStream, variables, options))
        //        {
        //            // Find the max row count to avoid index out of bounds
        //            int maxRows = variablesData.Max(entry => entry.Value.Values.Count);

        //            // Write records row by row
        //            for (int i = 0; i < maxRows; i++)
        //            {
        //                var newRecord = writer.CreateRecord();
        //                bool rowHasNull = false;

        //                // First pass: Check if any column in this row has null
        //                foreach (var entry in variablesData)
        //                {
        //                    if (i >= entry.Value.Values.Count || entry.Value.Values[i] == null)
        //                    {
        //                        rowHasNull = true;
        //                        break;
        //                    }
        //                }

        //                // Second pass: Assign null to all columns if any column is null
        //                int col = 0;
        //                foreach (var entry in variablesData)
        //                {
        //                    List<double?> values = entry.Value.Values;
        //                    newRecord[col] = rowHasNull ? null : (i < values.Count ? values[i] : null);
        //                    col++;
        //                }

        //                writer.WriteRecord(newRecord);
        //            }

        //            writer.EndFile();
        //        }
        //    }


        //}


    }
}
