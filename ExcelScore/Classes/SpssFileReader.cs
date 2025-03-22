using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using SpssLib.DataReader;
using Aspose.Cells;

public class SpssVariable
{
    public string Name { get; set; }
    public string Label { get; set; }
    public Dictionary<double, string> ValueLabels { get; set; } = new Dictionary<double, string>();
    public List<object> Data { get; set; } = new List<object>();
    public string Measure { get; set; }
    public bool HasValueLabels => ValueLabels.Count > 0;
}

public class SpssFileReader
{
    public List<SpssVariable> Variables { get; private set; } = new List<SpssVariable>();
    private string spssFilePath;

    public void LoadSpssFile()
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "SPSS Data Files (*.sav)|*.sav";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                spssFilePath = openFileDialog.FileName;
                ReadSpssFile(spssFilePath);
                //ShowVariablesInMessageBox();
                ExportToExcel(Path.Combine(Path.GetDirectoryName(spssFilePath), "ExportedData.xlsx"));
            }
        }
    }

    private void ReadSpssFile(string filePath)
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

    private void ShowVariablesInMessageBox()
    {
        foreach (var variable in Variables)
        {
            string valueLabelsInfo = variable.HasValueLabels ? "Has Value Labels" : "No Value Labels";
            string valueLabelsDetails = variable.HasValueLabels ? string.Join(", ", variable.ValueLabels.Select(v => $"{v.Key} | {v.Value}")) : "N/A";

            string message = $"Variable: {variable.Name}\nLabel: {variable.Label}\nMeasure: {variable.Measure}\n{valueLabelsInfo}\n\nValue Labels:\n{valueLabelsDetails}\n\nData:\n" +
                             string.Join(", ", variable.Data);
            MessageBox.Show(message, "SPSS Variable Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void ExportToExcel(string filePath)
    {
        Workbook workbook = new Workbook();
        Worksheet rawDataSheet = workbook.Worksheets[0];
        rawDataSheet.Name = "Rawdata";

        // Insert variable names in the first row
        for (int i = 0; i < Variables.Count; i++)
        {
            rawDataSheet.Cells[0, i].Value = Variables[i].Name;
        }

        // Insert variable data starting from row 1
        int rowCount = Variables.First().Data.Count;
        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < Variables.Count; col++)
            {
                rawDataSheet.Cells[row + 1, col].Value = Variables[col].Data[row];
            }
        }

        // Create second sheet "ListVariables"
        Worksheet listVariablesSheet = workbook.Worksheets.Add("ListVariables");

        // Insert variable names in the first row
        for (int i = 0; i < Variables.Count; i++)
        {
            listVariablesSheet.Cells[0, i].Value = Variables[i].Name;
        }

        // Insert measurement type in the second row
        for (int i = 0; i < Variables.Count; i++)
        {
            listVariablesSheet.Cells[1, i].Value = Variables[i].Measure;
        }

        // Insert value labels starting from row 2
        int maxValueLabelCount = Variables.Max(v => v.ValueLabels.Count);
        for (int row = 0; row < maxValueLabelCount; row++)
        {
            for (int col = 0; col < Variables.Count; col++)
            {
                var valueLabels = Variables[col].ValueLabels.ToList();
                if (row < valueLabels.Count)
                {
                    listVariablesSheet.Cells[row + 2, col].Value = $"{valueLabels[row].Key} | {valueLabels[row].Value}";
                }
            }
        }

        // Save the file in the same directory as the SPSS file
        workbook.Save(filePath);
        MessageBox.Show($"Excel file saved successfully at: {filePath}", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
