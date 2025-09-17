using Aspose.Cells;
using ExcelScore.Classes;
using SpssLib.DataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.StatClasses
{
    public class SpssReaderStat
    {

        public List<SpssVariable> Variables { get; private set; } = new List<SpssVariable>();
        public static string spssFilePath;
        public class SpssParameter
        {
            public string Name { get; set; }
            public string Label { get; set; }
            public string Type { get; set; } // "Nominal" or "Scale"
            public List<string> Values { get; set; } = new();
            public Dictionary<int, string> ValueLabels { get; set; } = new();
        }

        public static List<SpssParameter> ReadSpssParameters(string filePath)
        {
            var result = new List<SpssParameter>();

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            using (SpssReader spssDataset = new SpssReader(fileStream))
            {
                var variables = spssDataset.Variables.ToList();

                // Prepare result for each variable
                foreach (var variable in variables)
                {
                    var param = new SpssParameter
                    {
                        Name = variable.Name,
                        Label = variable.Label,
                        Type = variable.MeasurementType.ToString().IndexOf("Nominal", StringComparison.OrdinalIgnoreCase) >= 0 ? "Nominal" : "Scale",
                        ValueLabels = new Dictionary<int, string>(),
                        Values = new List<string>()
                    };

                    // Add coded labels
                    foreach (var kvp in variable.ValueLabels)
                    {
                        if (kvp.Key is double d && d == (int)d)
                            param.ValueLabels[(int)d] = kvp.Value;
                    }

                    result.Add(param);
                }

                // Read Records
                foreach (var record in spssDataset.Records)
                {
                    for (int i = 0; i < result.Count; i++)
                    {
                        var value = record.GetValue(variables[i]);
                        string strValue = value?.ToString() ?? ".";
                        result[i].Values.Add(strValue);
                    }
                }

                // Post-process Nominal variables: ensure all values are labeled
                foreach (var param in result)
                {
                    if (param.Type == "Nominal")
                    {
                        var distinctInts = param.Values
                            .Where(v => int.TryParse(v, out _))
                            .Select(v => int.Parse(v))
                            .Distinct()
                            .OrderBy(v => v)
                            .ToList();

                        foreach (int val in distinctInts)
                        {
                            if (!param.ValueLabels.ContainsKey(val))
                            {
                                param.ValueLabels[val] = val.ToString(); // Assign raw value as label
                            }
                        }

                        // Optional: sort ValueLabels by key
                        param.ValueLabels = param.ValueLabels
                            .OrderBy(kv => kv.Key)
                            .ToDictionary(kv => kv.Key, kv => kv.Value);
                    }
                }
            }

            return result;
        }

        private void ReadSpssFile_Excel(string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            {
                using (SpssReader spssDataset = new SpssReader(fileStream))
                {
                    Variables.Clear();

                    // Read Variables
                    foreach (var variable in spssDataset.Variables)
                    {
                        Variables.Add(new SpssVariable
                        {
                            Name = variable.Name,
                            Label = variable.Label,
                            ValueLabels = new Dictionary<double, string>(variable.ValueLabels),
                            Measure = variable.MeasurementType.ToString()
                        });
                    }

                    // Read Data Rows
                    foreach (var record in spssDataset.Records)
                    {
                        for (int i = 0; i < Variables.Count; i++)
                        {
                            var value = record.GetValue(spssDataset.Variables.ElementAt(i));
                            Variables[i].Data.Add(value ?? "."); // Handle missing values
                        }
                    }
                }
            }
        }
        public void GetSPSS_Path_Parameters()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "SPSS Data Files (*.sav)|*.sav";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    spssFilePath = openFileDialog.FileName;

                    var spssRaw = ReadSpssParameters(spssFilePath);
                    var statParams = spssRaw.Select(s => new StatParameter
                    {
                        Name = s.Name,
                        Label = s.Label,
                        Type = s.Type,
                        RawValues = new List<string>(s.Values),
                        ValueLabels = new Dictionary<int, string>(s.ValueLabels)
                    }).ToList();

                    
                    FormDataTransfer.Set("SPSS_Parameters", statParams);


                    if (spssFilePath != null)
                    {
                        MessageBox.Show("Spss Imported");
                    }
                }
            }

            


        }

        private void ExportToExcel(string filePath)
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Rawdata";

            // Insert variable names in the first row
            for (int i = 0; i < Variables.Count; i++)
            {
                sheet.Cells[0, i].Value = Variables[i].Name;
            }

            // Insert variable data starting from row 1
            int rowCount = Variables.First().Data.Count;
            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < Variables.Count; col++)
                {
                    sheet.Cells[row + 1, col].Value = Variables[col].Data[row];
                }
            }

            // Add second sheet "ListVariables"
            Worksheet sheet2 = workbook.Worksheets.Add("ListVariables");

            // Insert variable names in the first row
            for (int i = 0; i < Variables.Count; i++)
            {
                sheet2.Cells[0, i].Value = Variables[i].Name;
                sheet2.Cells[1, i].Value = Variables[i].Measure; // Nominal or Scale
            }

            // Insert value labels if they exist
            int maxValueLabelCount = Variables.Max(v => v.ValueLabels.Count);
            for (int row = 0; row < maxValueLabelCount; row++)
            {
                for (int col = 0; col < Variables.Count; col++)
                {
                    if (row < Variables[col].ValueLabels.Count)
                    {
                        var kvp = Variables[col].ValueLabels.ElementAt(row);
                        sheet2.Cells[row + 2, col].Value = $"{kvp.Key} | {kvp.Value}";
                    }
                }
            }

            // Save the file
            workbook.Save(filePath);

            // Extract folder path and copy it to clipboard
            string folderPath = Path.GetDirectoryName(filePath);
            Clipboard.SetText(folderPath);

            MessageBox.Show("Excel file saved successfully! Folder path copied to clipboard.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public void Convert_Spss_to_Excel()
        {
            if(spssFilePath != null)
            {
                ReadSpssFile_Excel(spssFilePath);
                ExportToExcel(Path.Combine(Path.GetDirectoryName(spssFilePath), "DataProg.xlsx"));
            }
            else
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "SPSS Data Files (*.sav)|*.sav";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ReadSpssFile_Excel(openFileDialog.FileName);
                        ExportToExcel(Path.Combine(Path.GetDirectoryName(openFileDialog.FileName), "DataProg.xlsx"));
                    }
                }
            }
        }

        

        
        


        

    }
}
