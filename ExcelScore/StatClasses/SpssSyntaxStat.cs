using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelScore.Base_Class_Extensions;
using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static CenterSpace.NMath.Core.KMeansClustering;
using static SkiaSharp.HarfBuzz.SKShaper;

namespace ExcelScore.StatClasses
{
    public class SpssSyntaxStat
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
        private static string EscapeForSpss(string path)
        {
            if (path.StartsWith(@"\\"))
            {
                // Preserve UNC \\ at start, escape rest
                return @"\\" + path.Substring(2).Replace(@"\", @"\\");
            }
            return path.Replace(@"\", @"\\");
        }
        public static string BuildUnifiedSyntax(List<StatTable> tables, SPSSFilePaths paths)
        {
            var syntaxBuilder = new StringBuilder();

            syntaxBuilder.AppendLine($@"
GET FILE='{EscapeForSpss(paths.SavFilePath)}'.
DATASET NAME DataSet1 WINDOW=ASIS.

OMS
  /SELECT TABLES
  /DESTINATION FORMAT=TEXT OUTFILE='{EscapeForSpss(paths.ResultPath)}'.");

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

        public static string ExecuteSpssSyntaxAndGetTextResult(SPSSFilePaths paths, string syntax)
        {
            //File.WriteAllText(paths.SyntaxPath, syntax);
            //File.WriteAllText(paths.SyntaxPath, syntax);

            using (var writer = new StreamWriter(paths.SyntaxPath, false, Encoding.GetEncoding(1256)))
            {
                writer.Write(syntax);
            }

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
            finally
            {
                ForceKillSPSS();
            }



            return outputText;
        }

        public static void ForceKillSPSS()
        {
            try
            {
                // Find all SPSS processes by name (without .exe)
                var processes = Process.GetProcessesByName("spsswin.exe");

                foreach (var process in processes)
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(); // wait until it's really closed
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error killing SPSS process: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ForceKillSPSS failed: " + ex.Message);
            }
        }

        public static string RunUnifiedSyntaxAndGetResult(
            List<StatTable> StatTables,
            out string resultText)
        {
            string savFile = SpssReaderStat.spssFilePath;
            resultText = null;
            if (string.IsNullOrEmpty(savFile))
                return null;
            var paths = PrepareOutputPaths(savFile);
            //var syntax = BuildUnifiedSyntax(groupVar, nominalVars, scaleVars, paths);
            var syntax = BuildUnifiedSyntax(StatTables, paths);
            resultText = ExecuteSpssSyntaxAndGetTextResult(paths, syntax);
            return resultText;
        }

        public class CrosstabBlockS
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
        public static List<CrosstabBlockS> ParseMultipleCrosstabBlocksS(string outputText, string groupVar)
        {
            var blocks = new List<CrosstabBlockS>();
            var lines = outputText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            CrosstabBlockS currentBlock = null;
            List<string> headers = null;
            bool inCrosstab = false, inChi = false;
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if ((line.Contains("* " + groupVar) && line.Trim().EndsWith("Crosstabulation")) || (line.Contains("* " + groupVar) && lines[i + 2] == "Crosstab"))
                {
                    string rowVar = line.Split('*')[0].Trim();
                    currentBlock = new CrosstabBlockS
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
                var blocks = ParseMultipleCrosstabBlocksS(outputText, groupVar);

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

        public class PercentileResult
        {
            public string VariableName { get; set; }
            public Dictionary<string, Dictionary<string, Dictionary<string, string>>> GroupPercentiles { get; set; } = new();
            public Dictionary<string, string> TotalPercentiles { get; set; } = new();
        }
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
                    if (!(nextParts.Contains("has") && nextParts.Contains("been") && nextParts.Contains("omitted.")))
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
                    return;
                }

                // For each group parameter separately, parse and assign
                foreach (var g in groupParams)
                {
                    var groupLabels = new List<string>();

                    if (g.ValueLabels != null && g.ValueLabels.Count > 0)
                    {
                        foreach (var kvp in g.ValueLabels
                                             .OrderBy(k => k.Key)) // sort by numeric key ascending
                        {
                            var key = kvp.Key;
                            var label = kvp.Value;

                            string chosen;

                            // use label only if it's not identical to the numeric key
                            if (!string.Equals(label, key.ToString(), StringComparison.OrdinalIgnoreCase))
                                chosen = label;
                            else
                                chosen = ((double)key).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);

                            groupLabels.Add(chosen);
                        }
                    }
                    var descResultsForThisGroup = ParseDescriptiveOutput_Smart_New(outputText, scaleParamNames, groupLabels , table , g);

                    // assign each result back into its StatParameter under the group name
                    foreach (var desc in descResultsForThisGroup)
                    {
                        var param = table.GetParameterByName(desc.VariableName);
                        if (param == null) continue;

                        // ensure the per-group container exists on StatParameter
                        if (param.DescriptiveStatsByGroup == null)
                            param.DescriptiveStatsByGroup = new Dictionary<string, SpssSyntaxStat.DescriptiveResult>();

                        // store per group variable (keyed by group parameter name)
                        param.DescriptiveStatsByGroup[g.Name] = desc;

                        // For backward compatibility, if the single DescriptiveStats is empty, set it
                        if (param.DescriptiveStats == null)
                            param.DescriptiveStats = desc;
                    }
                }


            }

           


            //var descResultsForThisGroup = ParseDescriptiveOutput_Smart(outputText, scaleParamNames, groupLabels);

            //// assign each result back into its StatParameter under the group name
            //foreach (var desc in descResultsForThisGroup)
            //{
            //    var param = table.GetParameterByName(desc.VariableName);
            //    if (param == null) continue;

            //    // ensure the per-group container exists on StatParameter
            //    if (param.DescriptiveStatsByGroup == null)
            //        param.DescriptiveStatsByGroup = new Dictionary<string, SpssSyntaxStat.DescriptiveResult>();

            //    // store per group variable (keyed by group parameter name)
            //    param.DescriptiveStatsByGroup[g.Name] = desc;

            //    // For backward compatibility, if the single DescriptiveStats is empty, set it
            //    if (param.DescriptiveStats == null)
            //        param.DescriptiveStats = desc;
            //}
        }
        public class DescriptiveResult
        {
            public string VariableName { get; set; }
            public Dictionary<string, Dictionary<string, string>> Stats_Groups { get; set; } = new();
            public Dictionary<string, string> Stats_Total { get; set; } = new();
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

            foreach (var result in results)
            {
                if (result.Stats_Groups.ContainsKey("Total") && (result.Stats_Groups["Total"] == null || result.Stats_Groups["Total"].Count == 0))
                {
                    result.Stats_Groups.Remove("Total");
                }

            }

            return results;
        }


        public static List<DescriptiveResult> ParseDescriptiveOutput_Smart_New(
    string rawText,
    List<string> variableNames,
    List<string> groupLabels, StatTable stattable ,  StatParameter groupParameter)
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
                return ParseDescriptiveOutput_Multiple_New_Chat(rawText, variableNames, groupLabels , statKeyToSpssLabel , stattable , groupParameter);
            }
        }
        public static Dictionary<string, string> statKeyToSpssLabel = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
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

        //public (bool,int) GroupLabelStart()
        public static List<DescriptiveResult> ParseDescriptiveOutput_Multiple_New(string rawText, List<string> variableNames, List<string> groupLabels , Dictionary<string, string> statKeyToSpssLabel,StatTable statTable , StatParameter groupParameter)
        {


            var results = variableNames.Select(var => new DescriptiveResult { VariableName = var }).ToList();
            var spssOutputLines = rawText.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            

            groupLabels.Add("Total");

            var activeSpssLabels = DefaultDescriptiveStats
               .Where(kvp => kvp.Value && statKeyToSpssLabel.ContainsKey(kvp.Key))
               .Select(kvp => statKeyToSpssLabel[kvp.Key])
               .ToList();



            for (int i = 0; i < spssOutputLines.Count; i++)
            {
                string line = spssOutputLines[i];


                if (!string.Equals(line, statTable.TableName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var scaleParams = statTable.GetNonGroupParameters()
                    .Where(p => string.Equals(p.Type, "Scale", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                HashSet<string> scaleParamNames = scaleParams
                    .Select(p => p.Name)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                // We entered the block for this table
                for (int j = i + 1; j < spssOutputLines.Count; j++)
                {
                    Dictionary<string, bool> Parameter_GroupLabels = new Dictionary<string, bool>(); 
                    var parts = SplitParts(spssOutputLines[j]);
                    if (parts.Count == (1 + scaleParamNames.Count))
                    {
                        if (parts[0] == groupParameter.Name &&
    parts.Skip(1).All(p => scaleParamNames.ContainsExact(p)))
                        {
                            // same result
                            bool TotalFlag = false;
                            string CurrentGroupLabel = "";
                            for (int k = j + 1; k < spssOutputLines.Count; k++)
                            {
                                
                                parts = SplitParts(spssOutputLines[k]);
                                if (parts.Count < 2)
                                {
                                    break;
                                }
                                
                                if (groupLabels.ContainsExact(parts[0]))
                                {
                                    //We started with grouplabel
                                    string statLabel = null;
                                    int valuectr = 0;
                                    for (int p = 1; p < parts.Count; p++)
                                    {
                                        var candidate = string.Join(" ", parts.Skip(1).Take(p));
                                        if (activeSpssLabels.Contains(candidate))
                                        {
                                            statLabel = candidate;
                                            valuectr = p+1;
                                            break;
                                        }
                                    }
                                    if (statLabel == null && parts.Count > 1)
                                    {
                                        statLabel = parts[1];
                                        valuectr = 2;
                                    }

                                    for (int v = 0; v < variableNames.Count; v++)
                                    {
                                        string variable = variableNames[v];
                                        string value = parts.Count > v + 2 ? parts[v + valuectr] : null;
                                        var result = results.First(r => r.VariableName == variable);

                                        if (parts[0] == "Total")
                                        {
                                            TotalFlag = true;
                                            if (!result.Stats_Total.ContainsKey(statLabel))
                                                result.Stats_Total[statLabel] = value;
                                        }
                                        else 
                                        {
                                            CurrentGroupLabel = parts[0];
                                            if (!result.Stats_Groups.ContainsKey(parts[0]))
                                                result.Stats_Groups[parts[0]] = new Dictionary<string, string>();

                                            if (!result.Stats_Groups[parts[0]].ContainsKey(statLabel))
                                                result.Stats_Groups[parts[0]][statLabel] = value;
                                        }
                                    }
                                }
                                else
                                {
                                    string statLabel = null;
                                    int valuectr = 0;
                                    for (int p = 0; p < parts.Count; p++)
                                    {
                                        var candidate = string.Join(" ", parts.Take(p + 1));
                                        if (activeSpssLabels.Contains(candidate))
                                        {
                                            statLabel = candidate;
                                            valuectr = p + 1;
                                            break;
                                        }
                                    }

                                    if (statLabel == null && parts.Count > 1)
                                    {
                                        statLabel = parts[0];
                                        valuectr = 1;
                                    }

                                    for (int v = 0; v < variableNames.Count; v++)
                                    {
                                        string variable = variableNames[v];

                                        if (Parameter_GroupLabels[variable] == false)
                                        {
                                            continue;
                                        }
                                        string value = parts.Count > v + 1 ? parts[v + valuectr] : null;
                                        var result = results.First(r => r.VariableName == variable);

                                        if (TotalFlag)
                                        {
                                            if (!result.Stats_Total.ContainsKey(statLabel))
                                                result.Stats_Total[statLabel] = value;
                                        }
                                        else
                                        {
                                            if (!result.Stats_Groups.ContainsKey(CurrentGroupLabel))
                                                result.Stats_Groups[CurrentGroupLabel] = new Dictionary<string, string>();

                                            if (!result.Stats_Groups[CurrentGroupLabel].ContainsKey(statLabel))
                                                result.Stats_Groups[CurrentGroupLabel][statLabel] = value;
                                        }
                                    }

                                }


                            }
                            break;
                            
                        }
                    }
                    
                    


                }
                
            }
            return results;
        }

        public static bool VerifyParameter_GroupLabel(string paraName , StatTable stattable , StatParameter groupParameter , string CurrentGroupLabel)
        {
            bool Parameter_GroupLabel = false;
            int groupCount = stattable.GetGroupParameters().Count;

            if(groupCount == 1)
            {
                var currentparameter = stattable.GetParameterByName(paraName);
                //var key = myDict.FirstOrDefault(x => x.Value == targetValue).Key;
                
                MessageBox.Show("..");

            }

            return Parameter_GroupLabel;
        }

        public static List<DescriptiveResult> ParseDescriptiveOutput_Multiple_New_Chat(
    string rawText,
    List<string> variableNames,
    List<string> groupLabels,
    Dictionary<string, string> statKeyToSpssLabel,
    StatTable statTable,
    StatParameter groupParameter)
        {
            var results = variableNames.Select(var => new DescriptiveResult { VariableName = var }).ToList();
            var spssOutputLines = rawText.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries).ToList();

            groupLabels.Add("Total");

            var activeSpssLabels = DefaultDescriptiveStats
                .Where(kvp => kvp.Value && statKeyToSpssLabel.ContainsKey(kvp.Key))
                .Select(kvp => statKeyToSpssLabel[kvp.Key])
                .ToList();

            for (int i = 0; i < spssOutputLines.Count; i++)
            {
                string line = spssOutputLines[i];
                if (!string.Equals(line, statTable.TableName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var scaleParams = statTable.GetNonGroupParameters()
                    .Where(p => string.Equals(p.Type, "Scale", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                HashSet<string> scaleParamNames = scaleParams
                    .Select(p => p.Name)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                // We entered the block for this table
                for (int j = i + 1; j < spssOutputLines.Count; j++)
                {
                    // compute max words in any stat label (once per table/block)
                    int maxStatWords = activeSpssLabels
                        .Select(s => s.Split((char[])null, StringSplitOptions.RemoveEmptyEntries).Length)
                        .DefaultIfEmpty(1)
                        .Max();

                    var parts = SplitParts(spssOutputLines[j]);

                    if (parts.Count == (1 + scaleParamNames.Count))
                    {
                        if (parts[0] == groupParameter.Name &&
                            parts.Skip(1).All(p => scaleParamNames.ContainsExact(p)))
                        {
                            bool TotalFlag = false;

                            int k = j + 1;

                            int start_GroupLabel = k;
                            while(k < spssOutputLines.Count)
                            {

                                parts = SplitParts(spssOutputLines[start_GroupLabel]);
                                if (parts.Count < 2)
                                    break;


                                string groupLabel_Space = "";
                                string groupLabel_NoSpace = "";
                                string ProceedLabel = "";

                                for (int tempGroupLabel = start_GroupLabel; tempGroupLabel < start_GroupLabel + activeSpssLabels.Count; tempGroupLabel++)
                                {
                                    string statLabel = null;
                                    int valuectr = 0;
                                    bool found = false;
                                    int start_statlabel = -1;

                                    parts = SplitParts(spssOutputLines[tempGroupLabel]);

                                    for (int s = 0; s < parts.Count && !found; s++)
                                    {
                                        int maxLen = Math.Min(maxStatWords, parts.Count - s);
                                        for (int len = maxLen; len >= 1; len--)
                                        {
                                            var candidate = string.Join(" ", parts.Skip(s).Take(len));
                                            if (activeSpssLabels.Contains(candidate))
                                            {
                                                statLabel = candidate;
                                                valuectr = s + len;
                                                start_statlabel = s;
                                                found = true;
                                                break;
                                            }
                                        }
                                    }

                                    string addition = (start_statlabel > 0 ? string.Join(" ", parts.Take(start_statlabel)) : "").Trim();
                                    if (string.IsNullOrEmpty(addition))
                                        continue;

                                    string candidateWithSpace = (ProceedLabel + " " + addition).Trim();
                                    string candidateNoSpace = (ProceedLabel + addition).Trim();

                                    

                                    foreach (var ExistGroupLabel in groupLabels)
                                    {
                                        if(ExistGroupLabel.Contains(candidateWithSpace))
                                        {
                                            ProceedLabel = candidateWithSpace;
                                        }
                                        else if(ExistGroupLabel.Contains(candidateNoSpace))
                                        {
                                            ProceedLabel = candidateNoSpace;
                                        }

                                    }

                                }
                                

                                for (int tempGroupLabel = start_GroupLabel; tempGroupLabel < start_GroupLabel + activeSpssLabels.Count; tempGroupLabel++)
                                {

                                    string statLabel = null;
                                    int valuectr = 0;
                                    bool found = false;
                                    int start_statlabel = -1;

                                    parts = SplitParts(spssOutputLines[tempGroupLabel]);

                                    for (int s = 0; s < parts.Count && !found; s++)
                                    {
                                        int maxLen = Math.Min(maxStatWords, parts.Count - s);
                                        for (int len = maxLen; len >= 1; len--)
                                        {
                                            var candidate = string.Join(" ", parts.Skip(s).Take(len));
                                            if (activeSpssLabels.Contains(candidate))
                                            {
                                                statLabel = candidate;
                                                valuectr = s + len;
                                                start_statlabel = s;
                                                found = true;
                                                break;
                                            }
                                        }
                                    }


                                    for (int v = 0; v < variableNames.Count; v++)
                                    {
                                        string variable = variableNames[v];

                                        
                                        string value = parts[start_statlabel + v +1];
                                        var result = results.First(r => r.VariableName == variable);

                                        if (ProceedLabel == "Total")
                                        {
                                            if (!result.Stats_Total.ContainsKey(statLabel))
                                                result.Stats_Total[statLabel] = value;
                                        }
                                        else
                                        {
                                            if (!result.Stats_Groups.ContainsKey(ProceedLabel))
                                                result.Stats_Groups[ProceedLabel] = new Dictionary<string, string>();

                                            if (!result.Stats_Groups[ProceedLabel].ContainsKey(statLabel))
                                                result.Stats_Groups[ProceedLabel][statLabel] = value;
                                        }
                                    }
                                }





                                start_GroupLabel += activeSpssLabels.Count;






                            }




                            break; // finished this table
                        }
                    }
                }
            }

            return results;
        }

        public bool Parameter_Group_Label(string parameterName , string groupLabel , StatTable statTable , StatParameter groupParameter)
        {
            bool hasvalues = false;

            int groupParametercount = statTable.GetGroupParameters().Count;

            if(groupParametercount == 1 )
            {


            }



            return hasvalues;
        }
    }
}

//for (int k = j + 1; k < spssOutputLines.Count; k++)
//{
//    parts = SplitParts(spssOutputLines[k]);
//    if (parts.Count < 2)
//        break;


//    string statLabel = null;

//    int valuectr = 0;
//    bool found = false;
//    int start = -1;


//    for (int s = 0; s < parts.Count && !found; s++)
//    {
//        int maxLen = Math.Min(maxStatWords, parts.Count - s);
//        for (int len = maxLen; len >= 1; len--)
//        {
//            var candidate = string.Join(" ", parts.Skip(s).Take(len));
//            if (activeSpssLabels.Contains(candidate))
//            {
//                statLabel = candidate;
//                valuectr = s + len;
//                start = s;
//                found = true;
//                break;
//            }
//        }
//    }

//    string groupLabel = start > 0 ? string.Join(" ", parts.Take(start)).Trim() : "";
//    string groupLabel_Space = start > 0 ? string.Join(" ", parts.Take(start)).Trim() : "";
//    string groupLabel_NoSpace = start > 0 ? string.Join(" ", parts.Take(start)).Trim() : "";

//    bool groupLabel_Space_flag = false;
//    bool groupLabel_NoSpace_flag = false;

//    if (!string.IsNullOrEmpty(groupLabel) && !groupLabels.ContainsExact(groupLabel))
//    {
//        int linePtr = k + 1; 
//        while (linePtr < spssOutputLines.Count)
//        {
//            var nextParts = SplitParts(spssOutputLines[linePtr]); 
//            if (nextParts.Count == 0) break;


//            int nextStart = -1; 
//            for (int s = 0; s < nextParts.Count; s++) 
//            { 
//                int maxLen = Math.Min(maxStatWords, nextParts.Count - s); 
//                for (int len = maxLen; len >= 1; len--) 
//                { 
//                    var candidate = string.Join(" ", nextParts.Skip(s).Take(len)); 
//                    if (activeSpssLabels.Contains(candidate)) 
//                    { 
//                        nextStart = s; 
//                        break; 
//                    } 
//                } 
//                if (nextStart != -1) 
//                    break; 
//            }

//            var addition = nextStart > 0 ? string.Join(" ", nextParts.Take(nextStart)) : string.Join(" ", nextParts);
//            groupLabel = (groupLabel + " " + addition).Trim();


//            if (groupLabels.ContainsExact(groupLabel)) 
//            {
//                break; 
//            } 

//            linePtr++;

//        }
//    }

//    if(groupLabels.Contains(groupLabel))
//    {
//        MessageBox.Show(groupLabel);
//    }
//}
