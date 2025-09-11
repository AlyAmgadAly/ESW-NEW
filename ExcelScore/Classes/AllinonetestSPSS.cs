using DocumentFormat.OpenXml.Spreadsheet;
using ExcelScore.StatClasses;
using Microsoft.SolverFoundation.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace ExcelScore.Classes
{
    public class SPSSUnifiedRunner
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
            

            if (File.Exists(resultPath))
                File.Delete(resultPath);

            

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

        public static Dictionary<string, bool> DefaultDescriptiveStats = new Dictionary<string, bool>
        {
            { "COUNT", true },
            { "MIN", true },
            { "MAX", true },
            { "MEAN", true },
            { "SEMEAN", true },
            { "STDDEV", true },
            { "MEDIAN", true },
            { "GMEDIAN", true }
        };

        public static string BuildUnifiedSyntax(List<StatTable> tables, SPSSFilePaths paths)
        {
            var syntaxBuilder = new StringBuilder();
            //syntaxBuilder.AppendLine("SET LOCALE = 'ar_EG'.");
            
            //syntaxBuilder.AppendLine("SET UNICODE=ON.");
            // --- Global Header ---
            syntaxBuilder.AppendLine($@"
GET FILE='{paths.SavFilePath.Replace(@"\", @"\\")}'.
DATASET NAME DataSet1 WINDOW=ASIS.

OMS
  /SELECT TABLES
  /DESTINATION FORMAT=TEXT OUTFILE='{paths.ResultPath.Replace(@"\", @"\\")}.'");

            foreach (var table in tables)
            {
                var groupParams = table.GetGroupParameters();
                if (groupParams.Count == 0) continue;

                var percentType = table.PercentType?.ToUpper() ?? "COLUMN";

                var nominalVars = table.GetNonGroupParameters()
                    .Where(p => p.Type == "Nominal")
                    .Select(p => p.Name)
                    .ToList();

                var scaleVars = table.GetNonGroupParameters()
                    .Where(p => p.Type == "Scale")
                    .ToList();

                syntaxBuilder.AppendLine();
                //syntaxBuilder.AppendLine($"* --- Table: {table.TableName} ---.");
                syntaxBuilder.AppendLine($@"TITLE ""{table.TableName}"".");


                // --- CROSSTABS ---
                if (nominalVars.Any())
                {
                    var mainGroup = groupParams.First().Name;
                    syntaxBuilder.AppendLine($@"
CROSSTABS
  /TABLES={string.Join(" ", nominalVars)} BY {mainGroup}
  /FORMAT=AVALUE TABLES
  /STATISTICS=CHISQ
  /CELLS=COUNT {percentType}
  /COUNT ROUND CELL
  /METHOD=MC CIN(99) SAMPLES(10000).");
                }

                // --- MEANS + EXAMINE ---
                if (scaleVars.Any())
                {
                    var tablePartScale = string.Join(" ", scaleVars.Select(p => p.Name));
                    var cellsPart = string.Join(" ", DefaultDescriptiveStats
                        .Where(kvp => kvp.Value)
                        .Select(kvp => kvp.Key.ToUpper()));

                    // CASE 1: One grouping variable
                    if (scaleVars.Any(p => p.GroupedParameterValues?.Count > 0))
                    {
                        var mainGroup = groupParams.First().Name;

                        syntaxBuilder.AppendLine($@"
MEANS
  TABLES={tablePartScale} BY {mainGroup}
  /CELLS={cellsPart}.");

                        syntaxBuilder.AppendLine($@"
EXAMINE
  VARIABLES={tablePartScale} BY {mainGroup}
  /PLOT NONE
  /PERCENTILES(25,50,75) HAVERAGE
  /STATISTICS NONE
  /MISSING PAIRWISE
  /TOTAL.");
                    }
                    // CASE 2: Multiple grouping variables (relation tables)
                    else if (scaleVars.Any(p => p.GroupedParameterValuesRelation?.Count > 0))
                    {
                        var allGroupVars = scaleVars
                            .SelectMany(p => p.GroupedParameterValuesRelation?.Keys ?? Enumerable.Empty<string>())
                            .Distinct();

                        foreach (var groupVar in allGroupVars)
                        {
                            syntaxBuilder.AppendLine($@"* --- Descriptives for GroupVar: {groupVar} ---.");

                            syntaxBuilder.AppendLine($@"
MEANS
  TABLES={tablePartScale} BY {groupVar}
  /CELLS={cellsPart}.");

                            syntaxBuilder.AppendLine($@"
EXAMINE
  VARIABLES={tablePartScale} BY {groupVar}
  /PLOT NONE
  /PERCENTILES(25,50,75) HAVERAGE
  /STATISTICS NONE
  /MISSING PAIRWISE
  /TOTAL.");
                        }
                    }
                }

                // --- STATISTICAL TESTS ---
                foreach (var param in scaleVars)
                {
                    string paramName = param.Name;
                    string normality = param.Normality?.ToUpper();

                    // SINGLE GROUPING VARIABLE
                    if (param.GroupedParameterValues?.Count >= 2)
                    {
                        string groupVar = groupParams.First().Name;

                        var validGroups = param.GroupedParameterValues
                            .Where(g => g.Value.Count > 1)
                            .Select(g => Convert.ToInt32(g.Key))
                            .ToList();

                        if (validGroups.Count < 2) continue;

                        if (validGroups.Count == 2)
                        {
                            int g1 = validGroups[0];
                            int g2 = validGroups[1];

                            if (normality == "NORMAL")
                            {
                                syntaxBuilder.AppendLine($@"
T-TEST
  GROUPS = {groupVar}({g1} {g2})
  /MISSING = ANALYSIS
  /VARIABLES = {paramName}
  /CRITERIA = CI(.95).");
                            }
                            else if (normality == "ABNORMAL")
                            {
                                syntaxBuilder.AppendLine($@"
NPAR TESTS
  /M-W = {paramName} BY {groupVar}({g1} {g2})
  /MISSING ANALYSIS.");
                            }
                        }
                        else
                        {
                            string groupList = string.Join(" ", validGroups);

                            var lowCountGroups = param.GroupedParameterValuesRelation?.ContainsKey(groupVar) == true
                                ? param.GroupedParameterValuesRelation[groupVar]
                                    .Where(g => g.Value.Count <= 1)
                                    .Select(g => Convert.ToInt32(g.Key))
                                    .ToList()
                                : param.GroupedParameterValues
                                    .Where(g => g.Value.Count <= 1)
                                    .Select(g => Convert.ToInt32(g.Key))
                                    .ToList();

                            if (lowCountGroups.Count > 0)
                            {
                                var filterCondition = string.Join("  |  ", validGroups.Select(v => $"{groupVar} = {v}"));
                                var filterLabel = $"{groupVar} = " + string.Join("  |  ", validGroups);

                                syntaxBuilder.AppendLine($@"
USE ALL.
COMPUTE filter_$=({filterCondition}).
VARIABLE LABEL filter_$ '{filterLabel} (FILTER)'.
VALUE LABELS filter_$  0 'Not Selected' 1 'Selected'.
FORMAT filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.");
                            }

                            if (normality == "NORMAL")
                            {
                                syntaxBuilder.AppendLine($@"
ONEWAY
  {paramName} BY {groupVar}
  /MISSING ANALYSIS
  /POSTHOC = {table.PostHoc} ALPHA(.05).");
                            }
                            else if (normality == "ABNORMAL")
                            {
                                syntaxBuilder.AppendLine($@"
NPAR TESTS
  /K-W = {paramName} BY {groupVar}(0 11)
  /MISSING ANALYSIS.");
                            }

                            syntaxBuilder.AppendLine($@"
FILTER OFF.
USE ALL.
EXECUTE.");
                        }

                    }
                    // MULTIPLE GROUPING VARIABLES
                    else if (param.GroupedParameterValuesRelation?.Count >= 1)
                    {
                        foreach (var relationGroup in param.GroupedParameterValuesRelation)
                        {
                            string groupVar = relationGroup.Key;

                            var validGroups = relationGroup.Value
                                .Where(g => g.Value.Count > 1)
                                .Select(g => Convert.ToInt32(g.Key))
                                .ToList();

                            if (validGroups.Count < 2) continue;

                            if (validGroups.Count == 2)
                            {
                                int g1 = validGroups[0];
                                int g2 = validGroups[1];

                                if (normality == "NORMAL")
                                {
                                    syntaxBuilder.AppendLine($@"
T-TEST
  GROUPS = {groupVar}({g1} {g2})
  /MISSING = ANALYSIS
  /VARIABLES = {paramName}
  /CRITERIA = CI(.95).");
                                }
                                else if (normality == "ABNORMAL")
                                {
                                    syntaxBuilder.AppendLine($@"
NPAR TESTS
  /M-W = {paramName} BY {groupVar}({g1} {g2})
  /MISSING ANALYSIS.");
                                }
                            }
                            else
                            {
                                string groupList = string.Join(" ", validGroups);

                                var lowCountGroups = param.GroupedParameterValuesRelation?.ContainsKey(groupVar) == true
                                    ? param.GroupedParameterValuesRelation[groupVar]
                                        .Where(g => g.Value.Count <= 1)
                                        .Select(g => Convert.ToInt32(g.Key))
                                        .ToList()
                                    : param.GroupedParameterValues
                                        .Where(g => g.Value.Count <= 1)
                                        .Select(g => Convert.ToInt32(g.Key))
                                        .ToList();

                                if (lowCountGroups.Count > 0)
                                {
                                    var filterCondition = string.Join("  |  ", validGroups.Select(v => $"{groupVar} = {v}"));
                                    var filterLabel = $"{groupVar} = " + string.Join("  |  ", validGroups);

                                    syntaxBuilder.AppendLine($@"
USE ALL.
COMPUTE filter_$=({filterCondition}).
VARIABLE LABEL filter_$ '{filterLabel} (FILTER)'.
VALUE LABELS filter_$  0 'Not Selected' 1 'Selected'.
FORMAT filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.");
                                }

                                if (normality == "NORMAL")
                                {
                                    syntaxBuilder.AppendLine($@"
ONEWAY
  {paramName} BY {groupVar}
  /MISSING ANALYSIS
  /POSTHOC = {table.PostHoc} ALPHA(.05).");
                                }
                                else if (normality == "ABNORMAL")
                                {
                                    syntaxBuilder.AppendLine($@"
NPAR TESTS
  /K-W = {paramName} BY {groupVar}(0 11)
  /MISSING ANALYSIS.");
                                }
                                syntaxBuilder.AppendLine($@"
FILTER OFF.
USE ALL.
EXECUTE.");
                            }
                        }
                    }
                }
            }

            syntaxBuilder.AppendLine("OMSEND.");
            return syntaxBuilder.ToString();
        }





        public static string BuildUnifiedSyntaxOld(string groupVar,
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

        public static string ExecuteSpssSyntaxAndGetTextResult(SPSSFilePaths paths, string syntax)
        {
            //File.WriteAllText(paths.SyntaxPath, syntax);
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
                            catch { }
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
            catch
            {

            }



            return outputText;
        }

        //public static string RunUnifiedSyntaxAndGetResult(
        //    string groupVar,
        //    List<string> nominalVars,
        //    List<string> scaleVars,
        //    out string resultText)
        //{
        //    string savFile = SpssFileReader.spssFilePath;
        //    resultText = null;
        //    if (string.IsNullOrEmpty(savFile))
        //        return null;
        //    var paths = PrepareOutputPaths(savFile);
        //    //var syntax = BuildUnifiedSyntax(groupVar, nominalVars, scaleVars, paths);
        //    var syntax = BuildUnifiedSyntax(groupVar, nominalVars, scaleVars, paths);
        //    resultText = ExecuteSpssSyntaxAndGetTextResult(paths, syntax);
        //    return resultText;
        //}

        public static string RunUnifiedSyntaxAndGetResult(
            List<StatTable> StatTables,
            out string resultText)
        {
            string savFile = SpssFileReader.spssFilePath;
            resultText = null;
            if (string.IsNullOrEmpty(savFile))
                return null;
            var paths = PrepareOutputPaths(savFile);
            //var syntax = BuildUnifiedSyntax(groupVar, nominalVars, scaleVars, paths);
            var syntax = BuildUnifiedSyntax(StatTables , paths);
            resultText = ExecuteSpssSyntaxAndGetTextResult(paths, syntax);
            return resultText;
        }

        // --- Data Structures ---



        public class CrosstabBlock
        {
            public string RowVariable;
            public List<string> Headers;
            public List<CrosstabRow> CrosstabRows;
            public List<ChiSquareResult> ChiResults;
            public List<string> TotalRowCounts;
            public List<string> TotalRowPercentages;
            public List<string> TotalColumnCounts;
            public List<string> TotalColumnPercentages;
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
        public class DescriptiveResult
        {
            public string VariableName { get; set; }
            public Dictionary<string, Dictionary<string, string>> Stats_Groups { get; set; } = new();
            public Dictionary<string, string> Stats_Total { get; set; } = new();
        }
        public class PercentileResult
        {
            public string VariableName { get; set; }
            public Dictionary<string, Dictionary<string, Dictionary< string,string>>> GroupPercentiles { get; set; } = new();
            public Dictionary<string, string> TotalPercentiles { get; set; } = new();
        }
        public class ParameterAnalysisResult
        {
            public string VariableName { get; set; }
            public string Type { get; set; } // "Nominal" or "Scale"

            public string NormalOrAbnormal { get; set; }
            public string GroupVariable { get; set; }
            public CrosstabBlock ChiSquareBlock { get; set; }
            public DescriptiveResult Descriptives { get; set; }
            public PercentileResult Percentiles { get; set; }

            public List<string> test_Pvalue { get; set; }
        }

        // --- Parsing Methods ---
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
                if ((line.Contains("* " + groupVar) && line.Trim().EndsWith("Crosstabulation")) || (line.Contains("* " + groupVar) && lines[i+2] == "Crosstab"))
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
                    if (line.StartsWith(groupVar))
                    {
                        headers = Regex.Split(line.Trim(), @"\s+").ToList();
                        headers.Add("Total");
                        currentBlock.Headers = headers;
                    }
                    else if (line.Contains("Count") && !line.StartsWith("Total"))
                    {
                        var countParts = Regex.Split(line.Trim(), @"\s+");
                        string groupName;
                        List<string> counts;
                        if (countParts[0] == currentBlock.RowVariable && countParts.Count() >= 3)
                        {
                            groupName = countParts[1];
                            counts = countParts.Skip(3).ToList();
                        }
                        else
                        {
                            groupName = countParts[0];
                            counts = countParts.Skip(2).ToList();
                        }
                        if (i + 1 < lines.Length)
                        {
                            var nextLine = lines[i + 1].Trim();
                            if (nextLine.Contains("% within"))
                            {
                                var percentParts = Regex.Split(nextLine.Trim(), @"\s+");
                                List<string> percentages;
                                if (percentParts[0] == currentBlock.RowVariable && percentParts.Count() >= 3)
                                    percentages = percentParts.Skip(3).ToList();
                                else
                                    percentages = percentParts.Skip(2).ToList();
                                currentBlock.TotalRowCounts ??= new List<string>();
                                currentBlock.TotalRowPercentages ??= new List<string>();
                                currentBlock.TotalRowCounts.Add(counts.Last());
                                currentBlock.TotalRowPercentages.Add(percentages.Last());
                                currentBlock.CrosstabRows.Add(new CrosstabRow
                                {
                                    GroupName = groupName,
                                    Counts = counts.Take(counts.Count - 1).ToList(),
                                    Percentages = percentages.Take(percentages.Count - 1).ToList()
                                });
                                i++;
                            }
                        }
                    }
                    else if (line.StartsWith("Total") && line.Contains("Count"))
                    {
                        var countParts = Regex.Split(line.Trim(), @"\s+").Skip(2).ToList();
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

        public static List<DescriptiveResult> ParseDescriptiveOutput_Multiple(string rawText, List<string> variableNames, List<string> groupLabels)
        {
            var results = variableNames.Select(var => new DescriptiveResult { VariableName = var }).ToList();
            var lines = rawText.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);
            groupLabels.Add("Total");

            // Mapping from your keys to SPSS output labels
            var statKeyToSpssLabel = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
{
    { "COUNT", "N" },
    { "MIN", "Minimum" },
    { "MAX", "Maximum" },
    { "MEAN", "Mean" },
    { "SEMEAN", "Std. Error of Mean" },
    { "STDDEV", "Std. Deviation" },
    { "MEDIAN", "Median" },
    { "GMEDIAN", "Grouped Median" }
};

            // Get only active SPSS labels based on your dictionary
            var activeSpssLabels = DefaultDescriptiveStats
                .Where(kvp => kvp.Value && statKeyToSpssLabel.ContainsKey(kvp.Key))
                .Select(kvp => statKeyToSpssLabel[kvp.Key])
                .ToList();

            // Find the "Report" section
            int reportIdx = Array.FindIndex(lines, l => l.Trim().Equals("Report", StringComparison.OrdinalIgnoreCase));
            if (reportIdx == -1) return results;

            // Find the header line (should contain variable names)
            int headerIdx = reportIdx + 1;
            while (headerIdx < lines.Length && string.IsNullOrWhiteSpace(lines[headerIdx])) headerIdx++;
            if (headerIdx >= lines.Length) return results;

            var headerLine = lines[headerIdx].Trim();
            var headerParts = Regex.Split(headerLine, @"\s+").ToList();
            bool isMultiParam = headerParts.Count > 2 && variableNames.All(v => headerParts.Contains(v));

            if (!isMultiParam) return results; // Skip if unexpected layout

            string currentGroup = null;
            bool totaldone = false;

            for (int i = headerIdx + 1; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrEmpty(line)) continue;
                var parts = Regex.Split(line, @"\s+").ToList();

                // Group start line like "Patient N 35 35"
                if (groupLabels.Any(gl => line.StartsWith(gl + " ")))
                {
                    currentGroup = groupLabels.First(gl => line.StartsWith(gl + " "));


                    if (currentGroup == "Total")
                    {

                        if (parts.Count >= 2)
                        {
                            string statLabel = parts[1];
                            if (activeSpssLabels.Contains(statLabel))
                            {
                                for (int v = 0; v < variableNames.Count; v++)
                                {
                                    string variable = variableNames[v];
                                    string value = parts.Count > v + 2 ? parts[v + 2] : null;
                                    var result = results.First(r => r.VariableName == variable);

                                    if (!result.Stats_Groups.ContainsKey(currentGroup))
                                        result.Stats_Groups[currentGroup] = new Dictionary<string, string>();

                                    if (currentGroup == "Total")
                                    {
                                        if (!result.Stats_Total.ContainsKey(statLabel))
                                            result.Stats_Total[statLabel] = value;
                                    }
                                    else
                                    {
                                        if (!result.Stats_Groups.ContainsKey(currentGroup))
                                            result.Stats_Groups[currentGroup] = new Dictionary<string, string>();

                                        if (!result.Stats_Groups[currentGroup].ContainsKey(statLabel))
                                            result.Stats_Groups[currentGroup][statLabel] = value;
                                    }
                                }
                            }
                        }
                    }
                    else if (parts.Count >= 2)
                    {
                        string statLabel = parts[1];
                        if (activeSpssLabels.Contains(statLabel))
                        {
                            for (int v = 0; v < variableNames.Count; v++)
                            {
                                string variable = variableNames[v];
                                string value = parts.Count > v + 2 ? parts[v + 2] : null;
                                var result = results.First(r => r.VariableName == variable);

                                if (currentGroup == "Total")
                                {
                                    if (!result.Stats_Total.ContainsKey(statLabel))
                                        result.Stats_Total[statLabel] = value;
                                }
                                else
                                {
                                    if (!result.Stats_Groups.ContainsKey(currentGroup))
                                        result.Stats_Groups[currentGroup] = new Dictionary<string, string>();

                                    if (!result.Stats_Groups[currentGroup].ContainsKey(statLabel))
                                        result.Stats_Groups[currentGroup][statLabel] = value;
                                }
                            }
                        }
                    }
                }
                // Stat row within group: "Minimum 6.00 14.00"
                else if (!string.IsNullOrEmpty(currentGroup))
                {
                    // Try to match multi-word stat label from start of line
                    string statLabel = null;
                    for (int p = 0; p < parts.Count; p++)
                    {
                        var candidate = string.Join(" ", parts.Take(p + 1));
                        if (activeSpssLabels.Contains(candidate))
                        {
                            statLabel = candidate;
                            break;
                        }
                    }

                    if (statLabel != null)
                    {
                        int valueStartIndex = statLabel.Split(' ').Length;

                        for (int v = 0; v < variableNames.Count; v++)
                        {
                            string variable = variableNames[v];
                            string value = parts.Count > valueStartIndex + v ? parts[valueStartIndex + v] : null;
                            var result = results.First(r => r.VariableName == variable);

                            


                            if(currentGroup == "Total")
                            {
                                if (!result.Stats_Total.ContainsKey(statLabel))
                                    result.Stats_Total[statLabel] = value;
                            }
                            else
                            {
                                if (!result.Stats_Groups.ContainsKey(currentGroup))
                                    result.Stats_Groups[currentGroup] = new Dictionary<string, string>();

                                if (!result.Stats_Groups[currentGroup].ContainsKey(statLabel))
                                    result.Stats_Groups[currentGroup][statLabel] = value;
                            }
                            
                        }
                    }
                }

                
                
            }

            foreach (var result in results)
            {
                if (result.Stats_Groups.ContainsKey("Total") && (result.Stats_Groups["Total"] == null || result.Stats_Groups["Total"].Count == 0))
                {
                    result.Stats_Groups.Remove("Total");
                }

            }

            return results;
        }


        public static List<DescriptiveResult> ParseDescriptiveOutput_SingleParam(
    string rawText,
    string variableName,
    List<string> groupLabels)
        {
            var results = new List<DescriptiveResult>
    {
        new DescriptiveResult { VariableName = variableName }
    };
            var result = results[0];

            // Define the order of stats as they appear in SPSS output
            var statOrder = new List<(string Key, string Label)>
    {
        ("COUNT", "N"),
        ("MIN", "Minimum"),
        ("MAX", "Maximum"),
        ("MEAN", "Mean"),
        ("SEMEAN", "Std. Error of Mean"),
        ("STDDEV", "Std. Deviation"),
        ("MEDIAN", "Median"),
        ("GMEDIAN", "Grouped Median")
    };

            // Keep only the active ones
            var activeLabels = statOrder
                .Where(s => DefaultDescriptiveStats.ContainsKey(s.Key) && DefaultDescriptiveStats[s.Key])
                .Select(s => s.Label)
                .ToList();

            // Split lines
            var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            // Find "Report"
            int reportIdx = lines.FindIndex(l => l.Trim().Equals("Report", StringComparison.OrdinalIgnoreCase));
            if (reportIdx == -1) return results;

            // Skip 3 header lines (Groups + 2 header wraps)
            int dataStart = reportIdx + 4;

            for (int i = dataStart; i < lines.Count; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                var parts = Regex.Split(line, @"\s+").ToList();
                if (parts.Count < 2) continue;

                string group = parts[0];
                bool isTotal = group.Equals("Total", StringComparison.OrdinalIgnoreCase);

                if (!groupLabels.Contains(group) && !isTotal) continue;

                for (int j = 0; j < activeLabels.Count && j + 1 < parts.Count; j++)
                {
                    string label = activeLabels[j];
                    string value = parts[j + 1];

                    if (isTotal)
                    {
                        result.Stats_Total[label] = value;
                    }
                    else
                    {
                        if (!result.Stats_Groups.ContainsKey(group))
                            result.Stats_Groups[group] = new Dictionary<string, string>();

                        result.Stats_Groups[group][label] = value;
                    }
                }

                if (isTotal)
                    break; // ✅ Done reading after Total row
            }

            return results;
        }

        

        public static List<DescriptiveResult> ParseDescriptiveOutput_Smart(
    string rawText,
    List<string> variableNames,
    List<string> groupLabels)
        {
            if (variableNames == null || variableNames.Count == 0)
                return new List<DescriptiveResult>();

            if (variableNames.Count == 1)
            {
                // Use single-parameter parsing
                return ParseDescriptiveOutput_SingleParam(rawText, variableNames[0], groupLabels);
            }
            else
            {
                // Use multiple-parameter parsing
                return ParseDescriptiveOutput_Multiple(rawText, variableNames, groupLabels);
            }
        }







        //public static List<PercentileResult> ParseTukeyTukeyHingesOnly(
        //    string rawText,
        //    List<string> variableNames,
        //    List<string> groupLabels,
        //    string groupSectionName)
        //{
        //    var results = variableNames.ToDictionary(v => v, v => new PercentileResult { VariableName = v });
        //    var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        //    string currentSection = null;
        //    string currentVariable = null;
        //    bool inTukey = false;
        //    bool inPercentileBlock = false;
        //    foreach (string raw in lines)
        //    {
        //        string line = raw.Trim();
        //        if (string.IsNullOrWhiteSpace(line))
        //            continue;
        //        if (line.StartsWith("Percentiles"))
        //        {
        //            inPercentileBlock = true;
        //            continue;
        //        }
        //        if (line.StartsWith("Total Sample"))
        //        {
        //            currentSection = "Total";
        //            currentVariable = null;
        //            continue;
        //        }
        //        if (line.StartsWith(groupSectionName))
        //        {
        //            currentSection = "Groups";
        //            currentVariable = null;
        //            continue;
        //        }
        //        if (!inPercentileBlock)
        //            continue;
        //        if (line.StartsWith("Tukey's Hinges"))
        //        {
        //            inTukey = true;
        //            string afterHeader = line.Substring("Tukey's Hinges".Length).Trim();
        //            if (!string.IsNullOrEmpty(afterHeader))
        //            {
        //                string workingLine = afterHeader;
        //                var parts = Regex.Split(workingLine, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        //                if (parts.Count > 0)
        //                {
        //                    if (currentSection == "Total")
        //                    {
        //                        if (variableNames.Contains(parts[0]) && parts.Count == 4)
        //                        {
        //                            currentVariable = parts[0];
        //                            results[currentVariable].TotalPercentiles = new Dictionary<string, string>
        //                            {
        //                                { "25", parts[1] },
        //                                { "50", parts[2] },
        //                                { "75", parts[3] }
        //                            };
        //                        }
        //                        else if (parts.Count == 4 && currentVariable != null)
        //                        {
        //                            results[currentVariable].TotalPercentiles = new Dictionary<string, string>
        //                            {
        //                                { "25", parts[1] },
        //                                { "50", parts[2] },
        //                                { "75", parts[3] }
        //                            };
        //                        }
        //                    }
        //                    else if (currentSection == "Groups")
        //                    {
        //                        if (variableNames.Contains(parts[0]) && groupLabels.Contains(parts[1]) && parts.Count == 5)
        //                        {
        //                            currentVariable = parts[0];
        //                            string group = parts[1];
        //                            results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
        //                            {
        //                                { "25", parts[2] },
        //                                { "50", parts[3] },
        //                                { "75", parts[4] }
        //                            };
        //                        }
        //                        else if (groupLabels.Contains(parts[0]) && parts.Count == 4 && currentVariable != null)
        //                        {
        //                            string group = parts[0];
        //                            results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
        //                            {
        //                                { "25", parts[1] },
        //                                { "50", parts[2] },
        //                                { "75", parts[3] }
        //                            };
        //                        }
        //                    }
        //                }
        //            }
        //            continue;
        //        }
        //        if (line.StartsWith("Weighted") || line.StartsWith("Average(Definition 1)") || line.StartsWith("Explore"))
        //        {
        //            inTukey = false;
        //            continue;
        //        }
        //        if (!inTukey)
        //            continue;
        //        string workingLine2 = line;
        //        if (workingLine2.StartsWith("Tukey's Hinges"))
        //            workingLine2 = workingLine2.Substring("Tukey's Hinges".Length).Trim();
        //        var parts2 = Regex.Split(workingLine2, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        //        if (parts2.Count == 0) continue;
        //        if (currentSection == "Total")
        //        {
        //            if (variableNames.Contains(parts2[0]) && parts2.Count == 4)
        //            {
        //                currentVariable = parts2[0];
        //                results[currentVariable].TotalPercentiles = new Dictionary<string, string>
        //                {
        //                    { "25", parts2[1] },
        //                    { "50", parts2[2] },
        //                    { "75", parts2[3] }
        //                };
        //            }
        //            else if (parts2.Count == 4 && currentVariable != null)
        //            {
        //                results[currentVariable].TotalPercentiles = new Dictionary<string, string>
        //                {
        //                    { "25", parts2[1] },
        //                    { "50", parts2[2] },
        //                    { "75", parts2[3] }
        //                };
        //            }
        //        }
        //        else if (currentSection == "Groups")
        //        {
        //            if (variableNames.Contains(parts2[0]) && groupLabels.Contains(parts2[1]) && parts2.Count == 5)
        //            {
        //                currentVariable = parts2[0];
        //                string group = parts2[1];
        //                results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
        //                {
        //                    { "25", parts2[2] },
        //                    { "50", parts2[3] },
        //                    { "75", parts2[4] }
        //                };
        //            }
        //            else if (groupLabels.Contains(parts2[0]) && parts2.Count == 4 && currentVariable != null)
        //            {
        //                string group = parts2[0];
        //                results[currentVariable].GroupPercentiles[group] = new Dictionary<string, string>
        //                {
        //                    { "25", parts2[1] },
        //                    { "50", parts2[2] },
        //                    { "75", parts2[3] }
        //                };
        //            }
        //        }
        //    }
        //    return results.Values.ToList();
        //}

        //public static void ParsePercentiles(List<StatTable> tables, List<string> spssOutputLines)
        //{

        //    for (int i = 0; i < spssOutputLines.Count; i++)
        //    {
        //        string line = spssOutputLines[i];

        //        foreach (var table in tables)
        //        {
        //            if (line == table.TableName)
        //            {
        //                var ScaleParams = table.GetNonGroupParameters()
        //                    .Where(p => string.Equals(p.Type, "Scale", StringComparison.OrdinalIgnoreCase))
        //                    .ToList();

        //                List<string> ScaleParamsNames = ScaleParams.Select(p => p.Name).ToList();


        //                // We entered the block for this table
        //                for (int j = i + 1; j < spssOutputLines.Count;j++)
        //                {
        //                    string innerLine = spssOutputLines[j];
        //                    //string[] parts = innerLine.Split(' ');


        //                    var parts = Regex.Split(innerLine, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        //                    if (parts.Count == 6 && (parts[0] + " " + parts[1]) == "Tukey's Hinges")
        //                    {
                                
        //                        string currentparametername = parts[2];

        //                        var currentparameter = table.GetParameterByName(currentparametername);

        //                        currentparameter.PercentileStats.VariableName = currentparametername;

        //                        currentparameter.PercentileStats.TotalPercentiles["25"] = parts[3];
        //                        currentparameter.PercentileStats.TotalPercentiles["50"] = parts[4];
        //                        currentparameter.PercentileStats.TotalPercentiles["75"] = parts[5];

        //                        //j++;

                                

        //                        for(int k = j+1;k < spssOutputLines.Count;k++)
        //                        {
        //                            string newline = spssOutputLines[k];
        //                            parts = Regex.Split(newline, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        //                            bool found = false;

        //                            if (parts.Count == 4 && ScaleParamsNames.Any(name => string.Equals(name, parts[0], StringComparison.OrdinalIgnoreCase)))
        //                            {
        //                                found = true;
        //                            }
                                    

        //                            if(found)
        //                            {
                                        
        //                                currentparametername = parts[0];

        //                                currentparameter = table.GetParameterByName(currentparametername);

        //                                currentparameter.PercentileStats.TotalPercentiles["25"] = parts[1];
        //                                currentparameter.PercentileStats.TotalPercentiles["50"] = parts[2];
        //                                currentparameter.PercentileStats.TotalPercentiles["75"] = parts[3];
        //                            }

        //                        }



        //                    }
        //                    else if (parts.Count > 6 && (parts[0] + " " + parts[1]) == "Tukey's Hinges")
        //                    {
        //                        //groups section
        //                        //problem here!!!!!!! line has to be  a word with percentile


        //                        //string groupline = spssOutputLines[j - 4];
        //                        //var groupparts = Regex.Split(groupline, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();

        //                        //string GroupParameterName = groupparts.First(p => p != "Percentiles");

        //                        //var CurrentGroup = table.GetParameterByName(GroupParameterName);


        //                        string GroupParameterName = null;
        //                        int newcount = j;
        //                        while (newcount >= 0)
        //                        {
        //                            string groupline = spssOutputLines[newcount];
        //                            var groupparts = Regex.Split(groupline, @"\s+")
        //                                                  .Where(p => !string.IsNullOrWhiteSpace(p))
        //                                                  .ToList();

        //                            if (groupparts.Count == 2 && groupparts.Contains("Percentiles"))
        //                            {
        //                                string candidate = groupparts.First(p => p != "Percentiles");

        //                                var param = table.GetParameterByName(candidate);
        //                                if (param != null && param.IsGroup)
        //                                {
        //                                    GroupParameterName = candidate;
        //                                    break; // found it
        //                                }
        //                            }

        //                            newcount--; // keep going backwards
        //                        }




        //                        string currentparaname = parts[2];
        //                        string GroupLabel = string.Join(" ", parts.Skip(3).Take(parts.Count - 6));

        //                        var currentpara = table.GetParameterByName(currentparaname);


        //                        // Ensure first-level key exists
        //                        if (!currentpara.PercentileStats.GroupPercentiles.ContainsKey(GroupParameterName))
        //                        {
        //                            currentpara.PercentileStats.GroupPercentiles[GroupParameterName]
        //                                = new Dictionary<string, Dictionary<string, string>>();
        //                        }

        //                        // Ensure second-level key exists
        //                        if (!currentpara.PercentileStats.GroupPercentiles[GroupParameterName].ContainsKey(GroupLabel))
        //                        {
        //                            currentpara.PercentileStats.GroupPercentiles[GroupParameterName][GroupLabel]
        //                                = new Dictionary<string, string>();
        //                        }



        //                        currentpara.PercentileStats.GroupPercentiles[GroupParameterName][GroupLabel]["25"] = parts[parts.Count- 3];
        //                        currentpara.PercentileStats.GroupPercentiles[GroupParameterName][GroupLabel]["50"] = parts[parts.Count - 2];
        //                        currentpara.PercentileStats.GroupPercentiles[GroupParameterName][GroupLabel]["75"] = parts[parts.Count - 1];

        //                        for (int k = j + 1; k < spssOutputLines.Count; k++)
        //                        {
        //                            string newline = spssOutputLines[k];
        //                            parts = Regex.Split(newline, @"\s+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();

        //                            if (parts.Count < 4)
        //                            {
        //                                break;
        //                            }

        //                            if (ScaleParamsNames.Any(name => string.Equals(name, parts[0], StringComparison.OrdinalIgnoreCase)))
        //                            {
        //                                currentpara = table.GetParameterByName(parts[0]);
        //                                string newGroupLabel = string.Join(" ", parts.Take(parts.Count - 3));

        //                                if (!currentpara.PercentileStats.GroupPercentiles.ContainsKey(GroupParameterName))
        //                                {
        //                                    currentpara.PercentileStats.GroupPercentiles[GroupParameterName]
        //                                        = new Dictionary<string, Dictionary<string, string>>();
        //                                }

        //                                // Ensure second-level key exists
        //                                if (!currentpara.PercentileStats.GroupPercentiles[GroupParameterName].ContainsKey(newGroupLabel))
        //                                {
        //                                    currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]
        //                                        = new Dictionary<string, string>();
        //                                }


        //                                currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]["25"] = parts[parts.Count - 3];
        //                                currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]["50"] = parts[parts.Count - 2];
        //                                currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]["75"] = parts[parts.Count - 1];

        //                                //New parameter
        //                            }
        //                            else
        //                            {
        //                                //Rest of labels
        //                                string newGroupLabel = string.Join(" ", parts.Take(parts.Count - 3));

        //                                if (!currentpara.PercentileStats.GroupPercentiles.ContainsKey(GroupParameterName))
        //                                {
        //                                    currentpara.PercentileStats.GroupPercentiles[GroupParameterName]
        //                                        = new Dictionary<string, Dictionary<string, string>>();
        //                                }

        //                                // Ensure second-level key exists
        //                                if (!currentpara.PercentileStats.GroupPercentiles[GroupParameterName].ContainsKey(newGroupLabel))
        //                                {
        //                                    currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]
        //                                        = new Dictionary<string, string>();
        //                                }
                                        
        //                                currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]["25"] = parts[parts.Count - 3];
        //                                currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]["50"] = parts[parts.Count - 2];
        //                                currentpara.PercentileStats.GroupPercentiles[GroupParameterName][newGroupLabel]["75"] = parts[parts.Count - 1];


                                        
        //                            }
        //                        }



        //                    }


        //                }
        //            }
        //        }
        //    }


        //}

        public static void ParseUnifiedOutput_Crosstabs(
     string outputText,
     List<StatTable> tables)
        {
            // 1. Gather all unique group variable names across tables
            var groupVars = tables
                .SelectMany(t => t.GetGroupParameters().Select(g => g.Name))
                .Distinct()
                .ToList();

            // 2. Parse and map for each group variable
            foreach (var groupVar in groupVars)
            {
                var blocks = ParseMultipleCrosstabBlocks(outputText, groupVar);

                foreach (var block in blocks)
                {
                    foreach (var table in tables)
                    {
                        // Must have this group variable
                        var groupParam = table.GetGroupParameters()
                            .FirstOrDefault(g => g.Name == groupVar);
                        if (groupParam == null)
                            continue;

                        // Must also have this row variable
                        var param = table.GetNonGroupParameters()
                            .FirstOrDefault(p => p.Name == block.RowVariable);
                        if (param == null)
                            continue;

                        // Assign once only
                        if (param.ChiSquareBlock == null)
                            param.ChiSquareBlock = block;
                    }
                }
            }
        }
        //public static void ParsePercentiles(List<StatTable> tables, string rawText)
        //{
        //    var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        //    StatParameter currentParam = null;
        //    bool inTotalBlock = false;
        //    bool inGroupBlock = false;

        //    foreach (var line in lines)
        //    {
        //        // --- Detect start of a TOTAL block ---
        //        if (line.Contains("Tukey's Hinges"))
        //        {
        //            // Example: "Tukey's Hinges Age     8.0000   11.0000   14.0000"
        //            var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        //            if (parts.Length == 6)
        //            {
        //                string varName = parts[2];

        //                // find parameter in our tables
        //                currentParam = tables
        //                    .SelectMany(t => t.Parameters)
        //                    .FirstOrDefault(p => p.Name == varName && p.Type == "Scale");

        //                if (currentParam == null) continue;

        //                // Reset flags
        //                inTotalBlock = true;
        //                inGroupBlock = false;

        //                // Prepare PercentileResult
        //                var result = new SPSSUnifiedRunner.PercentileResult
        //                {
        //                    VariableName = varName
        //                };

        //                result.TotalPercentiles["25"] = parts[3];
        //                result.TotalPercentiles["50"] = parts[4];
        //                result.TotalPercentiles["75"] = parts[5];

        //                currentParam.PercentileStats["Total"] = result;
        //            }
        //        }

        //        // --- Detect start of a GROUP block ---
        //        else if (line.Contains("Tukey's Hinges") && currentParam != null)
        //        {
        //            // Example: "Age  Patient  Tukey's Hinges   8.0000   10.0000   14.0000"
        //            var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        //            if (parts.Length >= 6)
        //            {
        //                string varName = parts[0];
        //                string groupLabel = parts[1];

        //                // Switch to group mode
        //                inTotalBlock = false;
        //                inGroupBlock = true;

        //                // Ensure group PercentileResult exists
        //                if (!currentParam.PercentileStats.ContainsKey("Groups"))
        //                {
        //                    currentParam.PercentileStats["Groups"] = new SPSSUnifiedRunner.PercentileResult
        //                    {
        //                        VariableName = varName
        //                    };
        //                }

        //                var groupResult = currentParam.PercentileStats["Groups"];

        //                groupResult.GroupPercentiles[groupLabel] = new Dictionary<string, string>
        //        {
        //            { "25", parts[3] },
        //            { "50", parts[4] },
        //            { "75", parts[5] }
        //        };
        //            }
        //        }
        //    }
        //}


    //    public static void ParseUnifiedOutput_Percentiles(
    //string outputText,
    //List<StatTable> tables)
    //    {
    //        foreach (var table in tables)
    //        {
    //            // collect scale parameter names for this table
    //            var scaleParamNames = table.GetNonGroupParameters()
    //                .Where(p => string.Equals(p.Type, "Scale", StringComparison.OrdinalIgnoreCase))
    //                .Select(p => p.Name)
    //                .ToList();

    //            if (!scaleParamNames.Any())
    //                continue;

    //            var groupParams = table.GetGroupParameters();

    //            // If there are no group parameters, parse once without groups
               

    //            // For each group parameter separately
    //            foreach (var g in groupParams)
    //            {
    //                // build ordered distinct labels from raw values
    //                var seen = new HashSet<string>();
    //                var orderedRawDistinct = new List<string>();
    //                foreach (var raw in g.RawValues ?? Enumerable.Empty<string>())
    //                {
    //                    if (string.IsNullOrWhiteSpace(raw)) continue;
    //                    if (raw == ".") continue;
    //                    if (seen.Add(raw))
    //                        orderedRawDistinct.Add(raw);
    //                }

    //                var groupLabels = new List<string>();
    //                foreach (var raw in orderedRawDistinct)
    //                {
    //                    string chosen = null;

    //                    if (double.TryParse(raw, out double d))
    //                    {
    //                        int code = (int)d;
    //                        if (g.ValueLabels != null && g.ValueLabels.TryGetValue(code, out var mappedLabel))
    //                        {
    //                            if (!string.Equals(mappedLabel, code.ToString(), StringComparison.OrdinalIgnoreCase))
    //                                chosen = mappedLabel;
    //                        }

    //                        if (chosen == null)
    //                            chosen = d.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
    //                    }
    //                    else
    //                    {
    //                        chosen = raw;
    //                    }

    //                    if (!groupLabels.Contains(chosen))
    //                        groupLabels.Add(chosen);
    //                }

    //                // call the percentile parser for this group
    //                var percResultsForThisGroup = ParseTukeyHingesOnly_CGPT(outputText, scaleParamNames, groupLabels , g.Name);
                    
    //                foreach (var perc in percResultsForThisGroup)
    //                {
    //                    var param = table.GetParameterByName(perc.VariableName);
    //                    if (param == null) continue;

    //                    if (param.PercentileStats == null)
    //                        param.PercentileStats = new Dictionary<string, SPSSUnifiedRunner.PercentileResult>();

    //                    // store this PercentileResult under the group parameter name
    //                    param.PercentileStats[g.Name] = perc;
    //                }
    //            }
    //        }
    //    }


        public static void ParseUnifiedOutput_Descriptives(
    string outputText,
    List<StatTable> tables)
        {
            foreach (var table in tables)
            {
                // collect scale parameter names for this table
                var scaleParamNames = table.GetNonGroupParameters()
                    .Where(p => string.Equals(p.Type, "Scale", StringComparison.OrdinalIgnoreCase))
                    .Select(p => p.Name)
                    .ToList();

                if (!scaleParamNames.Any())
                    continue;

                var groupParams = table.GetGroupParameters();

                // If there are no group parameters, still call parser once with empty labels
                if (!groupParams.Any())
                {
                    var descResults = ParseDescriptiveOutput_Smart(outputText, scaleParamNames, new List<string>());

                    foreach (var desc in descResults)
                    {
                        var param = table.GetParameterByName(desc.VariableName);
                        if (param == null) continue;

                        // assign into the legacy single-slot
                        param.DescriptiveStats = desc;
                    }

                    continue;
                }

                // For each group parameter separately, parse and assign
                foreach (var g in groupParams)
                {
                    // Build ordered distinct raw values (preserve first appearance order),
                    // excluding missing tokens like "." and empty strings
                    var seen = new HashSet<string>();
                    var orderedRawDistinct = new List<string>();
                    foreach (var raw in g.RawValues ?? Enumerable.Empty<string>())
                    {
                        if (string.IsNullOrWhiteSpace(raw)) continue;
                        if (raw == ".") continue; // treat dot as missing, skip
                        if (seen.Add(raw))
                            orderedRawDistinct.Add(raw);
                    }
                    var groupLabels = new List<string>();
                    foreach (var raw in orderedRawDistinct)
                    {
                        string chosen = null;

                        if (double.TryParse(raw, out double d))
                        {
                            int code = (int)d;
                            if (g.ValueLabels != null && g.ValueLabels.TryGetValue(code, out var mappedLabel))
                            {
                                // Check if mapped label is "real" (different from code)
                                if (!string.Equals(mappedLabel, code.ToString(), StringComparison.OrdinalIgnoreCase))
                                {
                                    chosen = mappedLabel; // use real label (like "Patient")
                                }
                            }

                            // If not labeled or label is just the number → normalize to 2 decimals
                            if (chosen == null)
                                chosen = d.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            // Non-numeric raw → keep as-is
                            chosen = raw;
                        }

                        if (!groupLabels.Contains(chosen))
                            groupLabels.Add(chosen);
                    }


                    // call the existing parser with this group's labels only
                    var descResultsForThisGroup = ParseDescriptiveOutput_Smart(outputText, scaleParamNames, groupLabels);

                    // assign each result back into its StatParameter under the group name
                    foreach (var desc in descResultsForThisGroup)
                    {
                        var param = table.GetParameterByName(desc.VariableName);
                        if (param == null) continue;

                        // ensure the per-group container exists on StatParameter
                        if (param.DescriptiveStatsByGroup == null)
                            param.DescriptiveStatsByGroup = new Dictionary<string, SPSSUnifiedRunner.DescriptiveResult>();

                        // store per group variable (keyed by group parameter name)
                        param.DescriptiveStatsByGroup[g.Name] = desc;

                        // For backward compatibility, if the single DescriptiveStats is empty, set it
                        if (param.DescriptiveStats == null)
                            param.DescriptiveStats = desc;
                    }
                }
            }
        }







        //public static List<ParameterAnalysisResult> ParseUnifiedOutput(
        //    string outputText,
        //    string groupVar,
        //    List<(string Name, string Type)> parameters,
        //    List<string> groupLabels)
        //{
        //    var nominalVars = parameters.Where(p => p.Type == "Nominal").Select(p => p.Name).ToList();
        //    var scaleVars = parameters.Where(p => p.Type == "Scale").Select(p => p.Name).ToList();
        //    var allResults = new List<ParameterAnalysisResult>();
        //    var chiBlocks = ParseMultipleCrosstabBlocks(outputText, groupVar);
        //    foreach (var block in chiBlocks)
        //    {
        //        if (nominalVars.Contains(block.RowVariable))
        //        {
        //            allResults.Add(new ParameterAnalysisResult
        //            {
        //                VariableName = block.RowVariable,
        //                Type = "Nominal",
        //                GroupVariable = groupVar,
        //                ChiSquareBlock = block
        //            });
        //        }
        //    }
        //    var descResults = ParseDescriptiveOutput_Smart(outputText, scaleVars, groupLabels);
        //    foreach (var desc in descResults)
        //    {
        //        var result = allResults.FirstOrDefault(r => r.VariableName == desc.VariableName);
        //        if (result == null)
        //        {
        //            result = new ParameterAnalysisResult
        //            {
        //                VariableName = desc.VariableName,
        //                Type = "Scale",
        //                GroupVariable = groupVar
        //            };
        //            allResults.Add(result);
        //        }
        //        result.Descriptives = desc;
        //    }
        //    var tukeyResults = ParseTukeyTukeyHingesOnly(outputText, scaleVars, groupLabels, "Groups");
        //    foreach (var tukey in tukeyResults)
        //    {
        //        var result = allResults.FirstOrDefault(r => r.VariableName == tukey.VariableName);
        //        if (result == null)
        //        {
        //            result = new ParameterAnalysisResult
        //            {
        //                VariableName = tukey.VariableName,
        //                Type = "Scale",
        //                GroupVariable = groupVar
        //            };
        //            allResults.Add(result);
        //        }
        //        result.Percentiles = tukey;
        //    }
        //    return allResults;
        //}

        //public static List<ParameterAnalysisResult> RunAllFromUnifiedSyntax(
        //    string groupVar,
        //    List<(string Name, string Type)> parameters,
        //    List<string> groupLabels)
        //{
        //    var nominalVars = parameters.Where(p => p.Type == "Nominal").Select(p => p.Name).ToList();
        //    var scaleVars = parameters.Where(p => p.Type == "Scale").Select(p => p.Name).ToList();
        //    string outputText = RunUnifiedSyntaxAndGetResult(groupVar, nominalVars, scaleVars, out _);
        //    if (string.IsNullOrWhiteSpace(outputText)) return new List<ParameterAnalysisResult>();
        //    return ParseUnifiedOutput(outputText, groupVar, parameters, groupLabels);
        //}
        //public static List<ParameterAnalysisResult> RunAllFromUnifiedSyntax(List<StatTable> statTables)
        //{
        //    //var nominalVars = parameters.Where(p => p.Type == "Nominal").Select(p => p.Name).ToList();
        //    //var scaleVars = parameters.Where(p => p.Type == "Scale").Select(p => p.Name).ToList();
        //    string outputText = RunUnifiedSyntaxAndGetResult(statTables, out _);

        //}




        public static void ParsePercentiles(List<StatTable> tables, List<string> spssOutputLines)
        {
            for (int i = 0; i < spssOutputLines.Count; i++)
            {
                string line = spssOutputLines[i];

                foreach (var table in tables)
                {
                    if (!string.Equals(line, table.TableName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var scaleParams = table.GetNonGroupParameters()
                        .Where(p => string.Equals(p.Type, "Scale", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    HashSet<string> scaleParamNames = scaleParams
                        .Select(p => p.Name)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    // We entered the block for this table
                    for (int j = i + 1; j < spssOutputLines.Count; j++)
                    {
                        var parts = SplitParts(spssOutputLines[j]);
                        if (parts.Count < 2) continue;

                        if (parts.Count == 6 && parts[0] == "Tukey's" && parts[1] == "Hinges")
                        {
                            HandleUngroupedPercentiles(parts, spssOutputLines, j, table, scaleParamNames);
                        }
                        else if (parts.Count > 6 && parts[0] == "Tukey's" && parts[1] == "Hinges")
                        {
                            HandleGroupedPercentiles(parts, spssOutputLines, j, table, scaleParamNames);
                        }
                    }
                }
            }



        }

        private static void HandleUngroupedPercentiles(List<string> parts, List<string> lines, int j, StatTable table, HashSet<string> scaleParamNames)
        {
            string paramName = parts[2];
            var param = table.GetParameterByName(paramName);
            if (param == null) return;

            param.PercentileStats.VariableName = paramName;
            param.PercentileStats.TotalPercentiles["25"] = parts[3];
            param.PercentileStats.TotalPercentiles["50"] = parts[4];
            param.PercentileStats.TotalPercentiles["75"] = parts[5];

            // Look for following parameters
            for (int k = j + 1; k < lines.Count; k++)
            {
                var nextParts = SplitParts(lines[k]);
                if (nextParts.Count != 4) break;

                if (!scaleParamNames.Contains(nextParts[0])) break; // stop if new block

                string nextName = nextParts[0];
                var nextParam = table.GetParameterByName(nextName);
                if (nextParam == null) continue;

                nextParam.PercentileStats.TotalPercentiles["25"] = nextParts[1];
                nextParam.PercentileStats.TotalPercentiles["50"] = nextParts[2];
                nextParam.PercentileStats.TotalPercentiles["75"] = nextParts[3];
            }
        }

        private static void HandleGroupedPercentiles(List<string> parts, List<string> lines, int j, StatTable table, HashSet<string> scaleParamNames)
        {
            string groupParamName = FindGroupParameterName(lines, j, table);
            if (groupParamName == null) return;

            string paramName = parts[2];
            var param = table.GetParameterByName(paramName);
            if (param == null) return;

            string groupLabel = string.Join(" ", parts.Skip(3).Take(parts.Count - 6));
            EnsureGroupEntry(param, groupParamName, groupLabel);

            param.PercentileStats.GroupPercentiles[groupParamName][groupLabel]["25"] = parts[parts.Count - 3];
            param.PercentileStats.GroupPercentiles[groupParamName][groupLabel]["50"] = parts[parts.Count - 2];
            param.PercentileStats.GroupPercentiles[groupParamName][groupLabel]["75"] = parts[parts.Count - 1];

            // Parse following lines
            for (int k = j + 1; k < lines.Count; k++)
            {
                var nextParts = SplitParts(lines[k]);
                if (nextParts.Count < 4) break;

                if (scaleParamNames.Contains(nextParts[0]))
                {
                    // ✅ Case 2: new parameter starts
                    paramName = nextParts[0];
                    param = table.GetParameterByName(paramName);
                    if (param == null) continue;

                    string newGroupLabel = string.Join(" ", nextParts.Skip(1).Take(nextParts.Count - 4));
                    EnsureGroupEntry(param, groupParamName, newGroupLabel);

                    param.PercentileStats.GroupPercentiles[groupParamName][newGroupLabel]["25"] = nextParts[nextParts.Count - 3];
                    param.PercentileStats.GroupPercentiles[groupParamName][newGroupLabel]["50"] = nextParts[nextParts.Count - 2];
                    param.PercentileStats.GroupPercentiles[groupParamName][newGroupLabel]["75"] = nextParts[nextParts.Count - 1];
                }
                else
                {
                    // ✅ Case 1: continuation label for same parameter
                    string newGroupLabel = string.Join(" ", nextParts.Take(nextParts.Count - 3));
                    EnsureGroupEntry(param, groupParamName, newGroupLabel);

                    param.PercentileStats.GroupPercentiles[groupParamName][newGroupLabel]["25"] = nextParts[nextParts.Count - 3];
                    param.PercentileStats.GroupPercentiles[groupParamName][newGroupLabel]["50"] = nextParts[nextParts.Count - 2];
                    param.PercentileStats.GroupPercentiles[groupParamName][newGroupLabel]["75"] = nextParts[nextParts.Count - 1];
                }
            }
        }


        private static string FindGroupParameterName(List<string> lines, int startIndex, StatTable table)
        {
            for (int idx = startIndex; idx >= 0; idx--)
            {
                var parts = SplitParts(lines[idx]);
                if (parts.Count == 2 && parts.Contains("Percentiles"))
                {
                    string candidate = parts.First(p => p != "Percentiles");
                    var param = table.GetParameterByName(candidate);
                    if (param?.IsGroup == true) return candidate;
                }
            }
            return null;
        }

        private static void EnsureGroupEntry(StatParameter param, string groupParamName, string groupLabel)
        {
            if (!param.PercentileStats.GroupPercentiles.ContainsKey(groupParamName))
                param.PercentileStats.GroupPercentiles[groupParamName] = new Dictionary<string, Dictionary<string, string>>();

            if (!param.PercentileStats.GroupPercentiles[groupParamName].ContainsKey(groupLabel))
                param.PercentileStats.GroupPercentiles[groupParamName][groupLabel] = new Dictionary<string, string>();
        }

        private static List<string> SplitParts(string line) =>
            line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries).ToList();

    }
}