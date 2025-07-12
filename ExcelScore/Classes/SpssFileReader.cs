using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using SpssLib.DataReader;
using Aspose.Cells;
using System.Diagnostics;


// this is used for Questionnare
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
    public static string spssFilePath;

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
                ExportToExcel(Path.Combine(Path.GetDirectoryName(spssFilePath), "DataProg.xlsx"));
                //RunSpssSyntax(spssFilePath);
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

    public static void RunSpssSyntax(string spssFilePath)
    {
        string folderPath = Path.GetDirectoryName(spssFilePath);
        string syntaxPath = Path.Combine(folderPath, "temp_script.sps");
        string outputPath = Path.Combine(folderPath, "output.spo"); // SPSS output file

        // Create the SPSS syntax dynamically
        File.WriteAllText(syntaxPath, $@"
GET FILE='{spssFilePath}'.
DESCRIPTIVES VARIABLES ALL /STATISTICS=MEAN STDDEV MIN MAX.
SAVE OUTFILE='{outputPath}'.
EXECUTE.
");

        // Run SPSS in batch mode (silent execution)
        ProcessStartInfo processInfo = new ProcessStartInfo
        {
            FileName = @"C:\Program Files (x86)\SPSS\spssspla.exe", // Adjust SPSS path
            Arguments = $"/production {syntaxPath}",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        Process process = Process.Start(processInfo);
        process.WaitForExit(); // Wait until the operation completes

        if (File.Exists(outputPath))
        {
            Clipboard.SetText(folderPath); // Copy folder path to clipboard
            MessageBox.Show("SPSS operation completed! Output file saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("SPSS execution failed. Please check SPSS installation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
