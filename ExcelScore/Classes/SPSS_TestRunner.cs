using DocumentFormat.OpenXml.Wordprocessing;
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

        public static Dictionary<string, bool> DefaultDescriptiveStats = new Dictionary<string, bool>
{
    { "COUNT", true },
    { "MIN", true },
    { "MAX", true },
    { "MEAN", true },
    { "SEMEAN", true },
    { "STDDEV", true },
    { "MEDIAN", true },
    { "GMEDIAN", true },
    { "SUM", false },
    { "RANGE", false },
    { "FIRST", false },
    { "LAST", false },
    { "VAR", false },
    { "KURT", false },
    { "SEKURT", false },
    { "SKEW", false },
    { "SESKEW", false },
    { "HARMONIC", false },
    { "GEOMETRIC", false },
    { "SPCT", false },
    { "NPCT", false }
};


        public static string ExecuteSpssSyntaxAndGetTextResult(SPSSFilePaths paths, string syntax)
        {
            File.WriteAllText(paths.SyntaxPath, syntax);

            Type spssType = Type.GetTypeFromProgID("SPSS.Application");
            dynamic spssApp = Activator.CreateInstance(spssType);
            dynamic syntaxDoc = spssApp.OpenSyntaxDoc(paths.SyntaxPath);
            syntaxDoc.Run();

            string outputText = null;
            int waited = 0, maxWaitMs = 8000, intervalMs = 250;

            while (waited < maxWaitMs)
            {
                try
                {
                    if (File.Exists(paths.ResultPath))
                    {
                        outputText = File.ReadAllText(paths.ResultPath);

                        if (!string.IsNullOrWhiteSpace(outputText))
                        {
                            Thread.Sleep(10000); // Let it catch up
                            try
                            {
                                dynamic outputDoc = spssApp.GetDesignatedOutputDoc();
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
                catch { }

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

            return outputText;
        }

        public static string RunDescriptiveSyntax(string groupVar, List<string> variables, Dictionary<string, bool> statsEnabled, out string resultText)
        {
            string savFile = SpssFileReader.spssFilePath;
            resultText = null;

            if (string.IsNullOrEmpty(savFile))
                return null;
           // / DESTINATION FORMAT = TEXT OUTFILE = '{paths.ResultPath.Replace(@"\", @"\\")}'.

            var paths = PrepareOutputPaths(savFile);

            var selectedStats = statsEnabled
                .Where(kvp => kvp.Value)
                .Select(kvp => kvp.Key.ToUpper())
                .ToList();

            string cellsPart = string.Join(" ", selectedStats);
            string tablePart = string.Join(" ", variables);

            string syntax = $@"
GET FILE='{paths.SavFilePath.Replace(@"\", @"\\")}'.
DATASET NAME DataSet1 WINDOW=ASIS.

OMS
  /SELECT TABLES
  /IF COMMANDS=['Means']
  /DESTINATION FORMAT = TEXT OUTFILE = '{paths.ResultPath.Replace(@"\", @"\\")}'.

MEANS
  TABLES={tablePart} BY {groupVar}
  /CELLS={cellsPart}.

OMSEND.
";

            resultText = ExecuteSpssSyntaxAndGetTextResult(paths, syntax);
            //MessageBox.Show(resultText);
            return resultText;
        }

        public class DescriptiveResult
        {
            public string VariableName { get; set; }

            // GroupName -> (StatName -> Value)
            public Dictionary<string, Dictionary<string, string>> Stats_Groups { get; set; } = new Dictionary<string, Dictionary<string, string>>();

            public Dictionary<string, string> Stats_Total { get; set; } = new();
        }
        public static List<DescriptiveResult> ParseDescriptiveOutput(string rawText, List<string> variableNames, List<string> groupLabels)
        {
            var results = variableNames.Select(var => new DescriptiveResult
            {
                VariableName = var
            }).ToList();

            var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            string currentGroup = null;

            foreach (var line in lines)
            {
                string trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed))
                    continue;

                // Detect group label
                string matchedGroup = groupLabels.FirstOrDefault(gl => trimmed.StartsWith(gl + " "));
                if (matchedGroup != null)
                {
                    currentGroup = matchedGroup;
                    continue;
                }

                // Skip if group not set yet
                if (string.IsNullOrEmpty(currentGroup))
                    continue;

                // Match full stat name + rest values
                var match = Regex.Match(trimmed, @"^(.+?)(-?\d[\d\.\-Ee ]*)$");
                if (!match.Success) continue;

                string statName = match.Groups[1].Value.Trim();
                string[] statValues = Regex.Split(match.Groups[2].Value.Trim(), @"\s+");

                for (int i = 0; i < variableNames.Count && i < statValues.Length; i++)
                {
                    string varName = variableNames[i];
                    string value = statValues[i];

                    var variableResult = results.First(r => r.VariableName == varName);
                    if (!variableResult.Stats_Groups.ContainsKey(currentGroup))
                        variableResult.Stats_Groups[currentGroup] = new Dictionary<string, string>();

                    variableResult.Stats_Groups[currentGroup][statName] = value;
                }
            }


            foreach (var r in results)
            {
                if (r.Stats_Groups.ContainsKey("Total"))
                {
                    r.Stats_Total = r.Stats_Groups["Total"];
                    r.Stats_Groups.Remove("Total"); // optional: clean from group dict
                }
            }

            


            return results;
        }


        public static string RunPercentilesSyntax(string groupVar, List<string> variables, out string resultText)
        {
            string savFile = SpssFileReader.spssFilePath;
            resultText = null;

            if (string.IsNullOrEmpty(savFile))
                return null;

            var paths = PrepareOutputPaths(savFile);
            string variableList = string.Join(" ", variables);

            string syntax = $@"
GET FILE='{paths.SavFilePath.Replace(@"\", @"\\")}'.
DATASET NAME DataSet1 WINDOW=ASIS.

OMS
  /SELECT TABLES
  /IF SUBTYPES=['Percentiles']
  /DESTINATION FORMAT = TEXT OUTFILE = '{paths.ResultPath.Replace(@"\", @"\\")}'.

EXAMINE
  VARIABLES={variableList} BY {groupVar}
  /PLOT NONE
  /PERCENTILES(25,50,75) HAVERAGE
  /STATISTICS NONE
  /MISSING PAIRWISE
  /TOTAL.

OMSEND.
";

            resultText = ExecuteSpssSyntaxAndGetTextResult(paths, syntax);
            return resultText;
        }

        public class PercentileResult
        {
            public string VariableName { get; set; }

            // For group-specific percentiles
            public Dictionary<string, Dictionary<string, string>> GroupPercentiles { get; set; } = new();

            // For total sample
            public Dictionary<string, string> TotalPercentiles { get; set; } = new();
        }


        public static List<PercentileResult> ParseTukeyTukeyHingesOnly(
    string rawText,
    List<string> variableNames,
    List<string> groupLabels,
    string groupSectionName
)
        {
            var results = variableNames.ToDictionary(v => v, v => new PercentileResult { VariableName = v });
            var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            string currentSection = null; // "Total" or "Groups"
            string currentVariable = null;
            bool inTukey = false;
            bool inPercentileBlock = false;

            foreach (string raw in lines)
            {
                string line = raw.Trim();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Section headers
                if (line.StartsWith("Percentiles"))
                {
                    inPercentileBlock = true;
                    continue;
                }
                if (line.StartsWith("Total Sample"))
                {
                    currentSection = "Total";
                    currentVariable = null;
                    continue;
                }
                if (line.StartsWith(groupSectionName))
                {
                    currentSection = "Groups";
                    currentVariable = null;
                    continue;
                }
                if (!inPercentileBlock)
                    continue;

                // Only process Tukey's Hinges
                if (line.StartsWith("Tukey's Hinges"))
                {
                    inTukey = true;
                    string afterHeader = line.Substring("Tukey's Hinges".Length).Trim();
                    if (!string.IsNullOrEmpty(afterHeader))
                    {
                        // Process the rest of the line as a data line
                        string workingLine = afterHeader;
                        var parts = Regex.Split(workingLine, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
                        if (parts.Count > 0)
                        {
                            if (currentSection == "Total")
                            {
                                if (variableNames.Contains(parts[0]) && parts.Count == 4)
                                {
                                    currentVariable = parts[0];
                                    results[currentVariable].TotalPercentiles = new Dictionary<string, string>
                            {
                                { "25", parts[1] },
                                { "50", parts[2] },
                                { "75", parts[3] }
                            };
                                }
                                else if (parts.Count == 4 && currentVariable != null)
                                {
                                    results[currentVariable].TotalPercentiles = new Dictionary<string, string>
                            {
                                { "25", parts[1] },
                                { "50", parts[2] },
                                { "75", parts[3] }
                            };
                                }
                            }
                            else if (currentSection == "Groups")
                            {
                                if (variableNames.Contains(parts[0]) && groupLabels.Contains(parts[1]) && parts.Count == 5)
                                {
                                    currentVariable = parts[0];
                                    string group = parts[1];
                                    results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
                            {
                                { "25", parts[2] },
                                { "50", parts[3] },
                                { "75", parts[4] }
                            };
                                }
                                else if (groupLabels.Contains(parts[0]) && parts.Count == 4 && currentVariable != null)
                                {
                                    string group = parts[0];
                                    results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
                            {
                                { "25", parts[1] },
                                { "50", parts[2] },
                                { "75", parts[3] }
                            };
                                }
                            }
                        }
                    }
                    continue;
                }
                // If we hit a new method or section, stop Tukey parsing
                if (line.StartsWith("Weighted") || line.StartsWith("Average(Definition 1)") || line.StartsWith("Explore"))
                {
                    inTukey = false;
                    continue;
                }
                if (!inTukey)
                    continue;

                // Remove leading method/indentation
                string workingLine2 = line;
                if (workingLine2.StartsWith("Tukey's Hinges"))
                    workingLine2 = workingLine2.Substring("Tukey's Hinges".Length).Trim();

                // Split by whitespace
                var parts2 = Regex.Split(workingLine2, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
                if (parts2.Count == 0) continue;

                // --- TOTAL SECTION ---
                // [Variable] [P25] [P50] [P75]
                if (currentSection == "Total")
                {
                    if (variableNames.Contains(parts2[0]) && parts2.Count == 4)
                    {
                        currentVariable = parts2[0];
                        results[currentVariable].TotalPercentiles = new Dictionary<string, string>
                {
                    { "25", parts2[1] },
                    { "50", parts2[2] },
                    { "75", parts2[3] }
                };
                    }
                    else if (parts2.Count == 4 && currentVariable != null)
                    {
                        // Variable omitted, use previous
                        results[currentVariable].TotalPercentiles = new Dictionary<string, string>
                {
                    { "25", parts2[1] },
                    { "50", parts2[2] },
                    { "75", parts2[3] }
                };
                    }
                }
                // --- GROUPS SECTION ---
                else if (currentSection == "Groups")
                {
                    // [Variable] [Group] [P25] [P50] [P75]
                    if (variableNames.Contains(parts2[0]) && groupLabels.Contains(parts2[1]) && parts2.Count == 5)
                    {
                        currentVariable = parts2[0];
                        string group = parts2[1];
                        results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
                {
                    { "25", parts2[2] },
                    { "50", parts2[3] },
                    { "75", parts2[4] }
                };
                    }
                    // [Group] [P25] [P50] [P75] (variable omitted)
                    else if (groupLabels.Contains(parts2[0]) && parts2.Count == 4 && currentVariable != null)
                    {
                        string group = parts2[0];
                        results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
                {
                    { "25", parts2[1] },
                    { "50", parts2[2] },
                    { "75", parts2[3] }
                };
                    }
                }
            }

            return results.Values.ToList();
        }




        public static string BuildUnifiedSyntax(string groupVar,
    List<string> nominalVars,
    List<string> scaleVars,
    SPSSFilePaths paths)
        {
            var tablePartNominal = string.Join(" ", nominalVars);
            var tablePartScale = string.Join(" ", scaleVars);
            var cellsPart = string.Join(" ", DefaultDescriptiveStats.Where(kvp => kvp.Value).Select(kvp => kvp.Key.ToUpper()));

            return $@"
GET FILE='{paths.SavFilePath.Replace(@"\", @"\\")}'.
DATASET NAME DataSet1 WINDOW=ASIS.

OMS
  /SELECT TABLES
  /IF SUBTYPES = ['Crosstabulation', 'Chi-Square Tests', 'Means', 'Percentiles']
  /DESTINATION FORMAT = TEXT OUTFILE = '{paths.ResultPath.Replace(@"\", @"\\")}'.

{(nominalVars.Any() ? $@"
CROSSTABS
  /TABLES={tablePartNominal} BY {groupVar}
  /FORMAT=AVALUE TABLES
  /STATISTICS=CHISQ
  /CELLS=COUNT COLUMN
  /COUNT ROUND CELL
  /METHOD=MC CIN(99) SAMPLES(10000)." : "")}

{(scaleVars.Any() ? $@"
MEANS
  TABLES={tablePartScale} BY {groupVar}
  /CELLS={cellsPart}." : "")}

{(scaleVars.Any() ? $@"
EXAMINE
  VARIABLES={tablePartScale} BY {groupVar}
  /PLOT NONE
  /PERCENTILES(25,50,75) HAVERAGE
  /STATISTICS NONE
  /MISSING PAIRWISE
  /TOTAL." : "")}

OMSEND.
";
        }


        public static string RunUnifiedSyntaxAndGetResult(
    string groupVar,
    List<string> nominalVars,
    List<string> scaleVars,
    out string resultText)
        {
            string savFile = SpssFileReader.spssFilePath;
            resultText = null;

            if (string.IsNullOrEmpty(savFile))
                return null;

            var paths = PrepareOutputPaths(savFile);
            var syntax = BuildUnifiedSyntax(groupVar, nominalVars, scaleVars, paths);
            resultText = ExecuteSpssSyntaxAndGetTextResult(paths, syntax);
            return resultText;
        }


        public static List<ParameterAnalysisResult> ParseUnifiedOutput(
    string outputText,
    string groupVar,
    List<(string Name, string Type)> parameters,
    List<string> groupLabels)
        {
            var nominalVars = parameters.Where(p => p.Type == "Nominal").Select(p => p.Name).ToList();
            var scaleVars = parameters.Where(p => p.Type == "Scale").Select(p => p.Name).ToList();

            var allResults = new List<ParameterAnalysisResult>();

            // 1. Parse Crosstab blocks
            var chiBlocks = ParseMultipleCrosstabBlocks(outputText, groupVar);
            foreach (var block in chiBlocks)
            {
                if (nominalVars.Contains(block.RowVariable))
                {
                    allResults.Add(new ParameterAnalysisResult
                    {
                        VariableName = block.RowVariable,
                        Type = "Nominal",
                        GroupVariable = groupVar,
                        ChiSquareBlock = block
                    });
                }
            }

            // 2. Parse Descriptive Stats
            var descResults = ParseDescriptiveOutput(outputText, scaleVars, groupLabels);
            foreach (var desc in descResults)
            {
                var result = allResults.FirstOrDefault(r => r.VariableName == desc.VariableName);
                if (result == null)
                {
                    result = new ParameterAnalysisResult
                    {
                        VariableName = desc.VariableName,
                        Type = "Scale",
                        GroupVariable = groupVar
                    };
                    allResults.Add(result);
                }
                result.Descriptives = desc;
            }

            // 3. Parse Tukey's Hinges
            var tukeyResults = ParseTukeyTukeyHingesOnly(outputText, scaleVars, groupLabels, "Groups");
            foreach (var tukey in tukeyResults)
            {
                var result = allResults.FirstOrDefault(r => r.VariableName == tukey.VariableName);
                if (result == null)
                {
                    result = new ParameterAnalysisResult
                    {
                        VariableName = tukey.VariableName,
                        Type = "Scale",
                        GroupVariable = groupVar
                    };
                    allResults.Add(result);
                }
                result.Percentiles = tukey;
            }

            return allResults;
        }

        public class ParameterAnalysisResult
        {
            public string VariableName { get; set; }
            public string Type { get; set; } // "Nominal" or "Scale"
            public string GroupVariable { get; set; }

            // Nominal-specific
            public CrosstabBlock ChiSquareBlock { get; set; }

            // Scale-specific
            public DescriptiveResult Descriptives { get; set; }
            public PercentileResult Percentiles { get; set; }
        }

        public static List<ParameterAnalysisResult> RunAllParameterAnalyses(
    string groupVar,
    List<(string Name, string Type)> parameters,
    List<string> groupLabels
)
        {
            var nominalVars = parameters.Where(p => p.Type == "Nominal").Select(p => p.Name).ToList();
            var scaleVars = parameters.Where(p => p.Type == "Scale").Select(p => p.Name).ToList();

            var results = new List<ParameterAnalysisResult>();

            // Run Chi-square for nominal
            var chiBlocks = RunCrosstabChiSquare(groupVar, nominalVars);
            if (chiBlocks != null)
            {
                foreach (var block in chiBlocks)
                {
                    results.Add(new ParameterAnalysisResult
                    {
                        VariableName = block.RowVariable,
                        Type = "Nominal",
                        GroupVariable = groupVar,
                        ChiSquareBlock = block
                    });
                }
            }

            // Run Descriptive + Percentile for scale
            if (scaleVars.Any())
            {
                string descText = RunDescriptiveSyntax(groupVar, scaleVars, DefaultDescriptiveStats, out var descResult);
                string tukeyText = RunPercentilesSyntax(groupVar, scaleVars, out var tukeyResult);

                var descriptives = ParseDescriptiveOutput(descText, scaleVars, groupLabels);
                var percentiles = ParseTukeyTukeyHingesOnly(tukeyResult, scaleVars, groupLabels, "Groups");

                foreach (var variable in scaleVars)
                {
                    var desc = descriptives.FirstOrDefault(d => d.VariableName == variable);
                    var perc = percentiles.FirstOrDefault(p => p.VariableName == variable);

                    results.Add(new ParameterAnalysisResult
                    {
                        VariableName = variable,
                        Type = "Scale",
                        GroupVariable = groupVar,
                        Descriptives = desc,
                        Percentiles = perc
                    });
                }
            }

            return results;
        }

        public static List<ParameterAnalysisResult> RunAllFromUnifiedSyntax(
    string groupVar,
    List<(string Name, string Type)> parameters,
    List<string> groupLabels)
        {
            var nominalVars = parameters.Where(p => p.Type == "Nominal").Select(p => p.Name).ToList();
            var scaleVars = parameters.Where(p => p.Type == "Scale").Select(p => p.Name).ToList();

            string outputText = RunUnifiedSyntaxAndGetResult(groupVar, nominalVars, scaleVars, out _);

            if (string.IsNullOrWhiteSpace(outputText)) return new List<ParameterAnalysisResult>();

            return ParseUnifiedOutput(outputText, groupVar, parameters, groupLabels);
        }


    }
}
