using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

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
