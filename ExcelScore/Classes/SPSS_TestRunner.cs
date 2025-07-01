using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace ExcelScore.Classes
{
    public class SPSS_TestRunner
    {
        public class SPSSFilePaths
        {
            public string SavFilePath;
            public string SyntaxPath;
            public string ResultPath;
            public string SpoPath;
            public string OutputFolder;
        }

        public static SPSSFilePaths PrepareOutputPaths(string savFilePath)
        {
            string baseFolder = Path.GetDirectoryName(savFilePath);
            string outputFolder = Path.Combine(baseFolder, "Program Output");
            Directory.CreateDirectory(outputFolder);

            string syntaxPath = Path.Combine(outputFolder, "auto_crosstab.sps");
            string resultPath = Path.Combine(outputFolder, "result.txt");

            string spoPath = Path.Combine(outputFolder, "output.spo");
            int counter = 1;
            while (File.Exists(spoPath))
                spoPath = Path.Combine(outputFolder, $"output_{counter++}.spo");

            return new SPSSFilePaths
            {
                SavFilePath = savFilePath,
                SyntaxPath = syntaxPath,
                ResultPath = resultPath,
                SpoPath = spoPath,
                OutputFolder = outputFolder
            };
        }

        public static List<CrosstabBlock> RunCrosstabChiSquare(string groupVar, List<string> rowVars)
        {
            string savFile = SpssFileReader.spssFilePath;
            if (string.IsNullOrEmpty(savFile)) return null;

            var paths = PrepareOutputPaths(savFile);
            string tableSyntax = string.Join(" ", rowVars);
            string syntax = $@"
GET FILE='{paths.SavFilePath.Replace(@"\", @"\\")}'.
DATASET NAME DataSet1 WINDOW=ASIS.

OMS
  /SELECT TABLES
  /IF SUBTYPES = ['Crosstabulation', 'Chi-Square Tests']
  /DESTINATION FORMAT = TEXT OUTFILE = '{paths.ResultPath.Replace(@"\", @"\\")}'.

CROSSTABS
  /TABLES={tableSyntax} BY {groupVar}
  /FORMAT=AVALUE TABLES
  /STATISTICS=CHISQ
  /CELLS=COUNT COLUMN
  /COUNT ROUND CELL
  /METHOD=MC CIN(99) SAMPLES(10000).

OMSEND.
";
            File.WriteAllText(paths.SyntaxPath, syntax);

            Type spssType = Type.GetTypeFromProgID("SPSS.Application");
            dynamic spssApp = Activator.CreateInstance(spssType);
            dynamic syntaxDoc = spssApp.OpenSyntaxDoc(paths.SyntaxPath);
            syntaxDoc.Run();

            string outputText = null;
            int waited = 0, maxWaitMs = 7000, intervalMs = 250;

            while (waited < maxWaitMs)
            {
                try
                {
                    bool resultReady = File.Exists(paths.ResultPath);
                    if (resultReady)
                    {
                        dynamic outputDoc = spssApp.GetDesignatedOutputDoc();
                        outputText = File.ReadAllText(paths.ResultPath);

                        if (!string.IsNullOrWhiteSpace(outputText))
                        {
                            try
                            {
                                Thread.Sleep(7000);    // Let it catch up
                                outputDoc.SaveAs(paths.SpoPath);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Failed to save .spo file: " + ex.Message);
                            }

                            break;
                        }
                    }
                }
                catch
                {
                    // Ignore and retry
                }

                Thread.Sleep(intervalMs);
                waited += intervalMs;
            }

            try
            {
                spssApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(spssApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }

            if (!string.IsNullOrWhiteSpace(outputText))
                return ParseMultipleCrosstabBlocks(outputText, groupVar);
            else
                return null;
        }


        public static List<CrosstabBlock> ParseMultipleCrosstabBlocks(string outputText, string groupVar)
        {
            var blocks = new List<CrosstabBlock>();
            var lines = outputText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            CrosstabBlock currentBlock = null;
            List<string> headers = null;
            bool inCrosstab = false, inChi = false;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (line.Contains("* " + groupVar) && line.Trim().EndsWith("Crosstabulation"))
                {
                    string rowVar = line.Split('*')[0].Trim();

                    currentBlock = new CrosstabBlock
                    {
                        RowVariable = rowVar,
                        CrosstabRows = new List<CrosstabRow>(),
                        ChiResults = new List<ChiSquareResult>()
                    };
                    blocks.Add(currentBlock);
                    inCrosstab = true;
                    inChi = false;
                    headers = null;
                    continue;
                }


                if (inCrosstab && line.StartsWith("Chi-Square Tests"))
                {
                    inCrosstab = false;
                    inChi = true;
                    continue;
                }
                if (inCrosstab)
                {
                    // Header line (e.g., A 1.00 2.00 3.00)
                    if (line.StartsWith(groupVar))
                    {
                        headers = Regex.Split(line.Trim(), @"\s+").ToList();
                        headers.Add("Total");
                        currentBlock.Headers = headers;
                    }
                    // Count line (may start with RowVariable or not)
                    else if (line.Contains("Count") && !line.StartsWith("Total"))
                    {
                        var countParts = Regex.Split(line.Trim(), @"\s+");

                        string groupName;
                        List<string> counts;

                        if (countParts[0] == currentBlock.RowVariable && countParts.Count() >= 3)
                        {
                            // Case: B 1.00 Count ...
                            groupName = countParts[1];
                            counts = countParts.Skip(3).ToList();
                        }
                        else
                        {
                            // Case: 1.00 Count ...
                            groupName = countParts[0];
                            counts = countParts.Skip(2).ToList();
                        }

                        // Look ahead for % within A line
                        if (i + 1 < lines.Length)
                        {
                            var nextLine = lines[i + 1].Trim();
                            if (nextLine.Contains("% within"))
                            {
                                var percentParts = Regex.Split(nextLine.Trim(), @"\s+");

                                List<string> percentages;
                                if (percentParts[0] == currentBlock.RowVariable && percentParts.Count() >= 3)
                                    percentages = percentParts.Skip(3).ToList();  // B 1.00 % within ...
                                else
                                    percentages = percentParts.Skip(2).ToList();  // 1.00 % within ...

                                // Initialize the total lists if null
                                currentBlock.TotalRowCounts ??= new List<string>();
                                currentBlock.TotalRowPercentages ??= new List<string>();

                                // Save last item in each row as row total
                                currentBlock.TotalRowCounts.Add(counts.Last());
                                currentBlock.TotalRowPercentages.Add(percentages.Last());

                                // Save main values (excluding the final total column)
                                currentBlock.CrosstabRows.Add(new CrosstabRow
                                {
                                    GroupName = groupName,
                                    Counts = counts.Take(counts.Count - 1).ToList(),
                                    Percentages = percentages.Take(percentages.Count - 1).ToList()
                                });

                                i++; // Skip the % row
                            }
                        }
                    }
                    // Final "Total Count" row (bottom row)
                    else if (line.StartsWith("Total") && line.Contains("Count"))
                    {
                        var countParts = Regex.Split(line.Trim(), @"\s+").Skip(2).ToList(); // Skip "Total Count"
                        currentBlock.TotalColumnCounts = countParts;
                    }
                    else if (line.StartsWith("% within") || (line.StartsWith("Total") && line.Contains("% within")))
                    {
                        var percentParts = Regex.Split(line.Trim(), @"\s+").Skip(2).ToList();
                        currentBlock.TotalColumnPercentages = percentParts;
                    }
                }


                if (inChi && (line.StartsWith("Pearson") || line.StartsWith("Likelih") || line.StartsWith("Fisher") || line.StartsWith("Linear")))
                {
                    var parts = Regex.Split(line.Trim(), @"\s+");
                    currentBlock.ChiResults.Add(new ChiSquareResult
                    {
                        Name = parts[0],
                        Value = parts.ElementAtOrDefault(1)?.Replace("(b)", "").Replace("(c)", ""),
                        df = parts.ElementAtOrDefault(2),
                        AsympSig = parts.ElementAtOrDefault(3),
                        MC_Sig2sided = parts.ElementAtOrDefault(4)?.Replace("(a)", ""),
                        MC_CI_Lower = parts.ElementAtOrDefault(5),
                        MC_CI_Upper = parts.ElementAtOrDefault(6)
                    });
                }
                if (line.Contains("have expected count less than"))
                {
                    var match = Regex.Match(line, @"\((\d+\.\d+%)\)");
                    if (match.Success)
                    {
                        currentBlock.FisherExpectedCountPercentage = match.Groups[1].Value;
                    }
                }
            }

            //Display each block
            //foreach (var block in blocks)
            //{
            //    string crosstab = $"Crosstab for {block.RowVariable} by {groupVar}:\n";
            //    if (block.Headers != null)
            //        crosstab += "     | " + string.Join(" | ", block.Headers) + "\n";

            //    foreach (var row in block.CrosstabRows)
            //    {
            //        crosstab += $"Group = {row.GroupName}\n";
            //        crosstab += "Count     : " + string.Join(" | ", row.Counts) + "\n";
            //        crosstab += "% within : " + string.Join(" | ", row.Percentages) + "\n\n";
            //    }

            //    // ➕ Add total row values (i.e., last column of each data row)
            //    if (block.TotalRowCounts != null && block.TotalRowCounts.Count > 0)
            //    {
            //        crosstab += "Total Row Counts       : " + string.Join(" | ", block.TotalRowCounts) + "\n";
            //        crosstab += "Total Row Percentages  : " + string.Join(" | ", block.TotalRowPercentages) + "\n\n";
            //    }

            //    // ➕ Add total column values (i.e., final totals row in SPSS output)
            //    if (block.TotalColumnCounts != null && block.TotalColumnCounts.Count > 0)
            //    {
            //        crosstab += "Total Column Counts    : " + string.Join(" | ", block.TotalColumnCounts) + "\n";
            //    }

            //    if (block.TotalColumnPercentages != null && block.TotalColumnPercentages.Count > 0)
            //    {
            //        crosstab += "Total Column % within A: " + string.Join(" | ", block.TotalColumnPercentages) + "\n";
            //    }

            //    string chi = $"Chi-Square Tests for {block.RowVariable}:\n";
            //    foreach (var chiTest in block.ChiResults)
            //    {
            //        chi += $"{chiTest.Name}: χ²={chiTest.Value}, df={chiTest.df}, p={chiTest.AsympSig}, MC p={chiTest.MC_Sig2sided}, CI=({chiTest.MC_CI_Lower}, {chiTest.MC_CI_Upper})\n";
            //    }

            //    MessageBox.Show(crosstab);
            //    MessageBox.Show(chi);
            //}

            return blocks;

        }

        public class CrosstabBlock
        {
            public string RowVariable;
            public List<string> Headers;
            public List<CrosstabRow> CrosstabRows;
            public List<ChiSquareResult> ChiResults;

            public List<string> TotalRowCounts;        // Last column (Total) per row
            public List<string> TotalRowPercentages;   // Last column of % within A per row
            public List<string> TotalColumnCounts;     // Bottom row "Total Count"
            public List<string> TotalColumnPercentages;// Bottom row "% within A"
            public string FisherExpectedCountPercentage;
        }

        public class CrosstabRow
        {
            public string GroupName;
            public List<string> Counts;
            public List<string> Percentages;
        }

        public class ChiSquareResult
        {
            public string Name;
            public string Value;
            public string df;
            public string AsympSig;
            public string MC_Sig2sided;
            public string MC_CI_Lower;
            public string MC_CI_Upper;
        }
    }
}
