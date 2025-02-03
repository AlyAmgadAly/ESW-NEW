using Aspose.Cells;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using static ExcelScore.Forms.Waiting.Tool;
using static ExcelScore.Forms.Waiting;

namespace ExcelScore.Forms
{
    public partial class Waiting : Form
    {
        public Waiting()
        {
            InitializeComponent();
        }

        public DataGridView Dgv { get; set; }

        private void Waiting_Load(object sender, EventArgs e)
        {
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(ExcelFunctions.filepath);
            Worksheet Sheet1 = workbook.Worksheets[0];
            MessageBox.Show(Sheet1.Cells[0, 0].Value.ToString());
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
            ChooseFrm chooseFrm = new ChooseFrm();
            chooseFrm.Dgv = Dgv;
            chooseFrm.Show();
        }

        public class Tool
        {
            public int ToolNumber { get; set; }
            public string ToolName { get; set; }
            public List<Subscale> Subscales { get; set; } = new List<Subscale>();
            public List<int> LikertScale { get; set; }
            public List<LevelRange> LevelRanges { get; set; }
            public string LevelDetermination { get; set; }

            public List<double> TotalScores { get; set; } = new List<double>();
            public List<double> AverageScores { get; set; } = new List<double>();
            public List<double> PercentScores { get; set; } = new List<double>();
            public List<int> ComputedLevels { get; set; } = new List<int>();

            public Tool(int toolNumber, string toolName, List<int> likertScale, List<LevelRange> levelRanges, string levelDetermination)
            {
                ToolNumber = toolNumber;
                ToolName = toolName;
                LikertScale = likertScale;
                LevelRanges = levelRanges;
                LevelDetermination = levelDetermination;
            }

            public void AddSubscale(Subscale subscale)
            {
                Subscales.Add(subscale);
            }

            public class Subscale
            {
                public string Name { get; set; }
                public List<Item> Items { get; set; }

                // Lists to store scores for each participant
                public List<double> TotalScores { get; set; } = new List<double>();
                public List<double> AverageScores { get; set; } = new List<double>();
                public List<double> PercentScores { get; set; } = new List<double>();
                public List<int> ComputedLevels { get; set; } = new List<int>();

                public Subscale(string name)
                {
                    Name = name;
                    Items = new List<Item>();
                }

                public void AddItem(Item item)
                {
                    Items.Add(item);
                }

                // Calculate scores for the subscale for a participant
                public void CalculateScores(List<int> responses, List<int> likertScale, List<LevelRange> levelRanges, Tool tool)
                {
                    if (responses.Count != Items.Count)
                        throw new ArgumentException("Responses count must match items count.");

                    // Calculate MinScore and MaxScore dynamically
                    double minScore = Items.Count * likertScale.Min();
                    double maxScore = Items.Count * likertScale.Max();

                    double totalScore = 0;
                    for (int i = 0; i < Items.Count; i++)
                    {
                        int response = responses[i];
                        if (Items[i].IsReverse)
                        {
                            // Reverse score: e.g., 3 becomes 1, 2 becomes 2, etc.
                            response = likertScale.Max() + 1 - response;
                        }
                        totalScore += response;
                    }

                    double averageScore = totalScore / Items.Count;
                    double percentScore = (totalScore - minScore) / (maxScore - minScore) * 100;

                    // Compute level based on tool-specific level ranges
                    int computedLevel = levelRanges
                        .FirstOrDefault(range => tool.IsInRange(percentScore, range.Range))?.Level ?? 0;

                    // Store scores for the participant
                    TotalScores.Add(totalScore);
                    AverageScores.Add(averageScore);
                    PercentScores.Add(percentScore);
                    ComputedLevels.Add(computedLevel);
                }

                public override string ToString()
                {
                    return $"Subscale(Name={Name}, TotalScores={string.Join(", ", TotalScores)}, AverageScores={string.Join(", ", AverageScores)}, PercentScores={string.Join(", ", PercentScores)}, Levels={string.Join(", ", ComputedLevels)})";
                }
            }

            public class Item
            {
                public int Id { get; set; }
                public string Text { get; set; }
                public bool IsReverse { get; set; }
                public List<int> ParticipantResponses { get; set; } = new List<int>();

                public Item(int id, string text, bool isReverse = false)
                {
                    Id = id;
                    Text = text;
                    IsReverse = isReverse;
                }
            }

            public class Participant
            {
                public string Name { get; set; }
                public List<int> Responses { get; set; } = new List<int>();
            }

            public class ExcelReader
            {
                public static List<Tool> ReadTools(Workbook workbook)
                {
                    var tools = new List<Tool>();
                    var toolConfigs = ReadToolConfigs(workbook);
                    var sheet = workbook.Worksheets[0];
                    var headerRow = sheet.Cells.GetRow(0);

                    for (int col = 0; col <= sheet.Cells.MaxDataColumn; col++)
                    {
                        string header = sheet.Cells[0, col].StringValue;

                        if (!header.StartsWith("Q"))
                            continue;

                        var parts = header.Split('.');
                        if (parts.Length < 3)
                            throw new FormatException($"Invalid header format: {header}");

                        int toolNumber = int.Parse(parts[0].Substring(1));
                        string subscaleName = parts[1];
                        int itemNumber = int.Parse(parts[2]);

                        var tool = tools.FirstOrDefault(t => t.ToolNumber == toolNumber);
                        if (tool == null)
                        {
                            tool = new Tool(toolNumber, $"Tool {toolNumber}", new List<int>(), new List<LevelRange>(), "percent");
                            tools.Add(tool);
                        }

                        var subscale = tool.Subscales.FirstOrDefault(s => s.Name == subscaleName);
                        if (subscale == null)
                        {
                            subscale = new Subscale(subscaleName);
                            tool.Subscales.Add(subscale);
                        }

                        subscale.AddItem(new Item(itemNumber, $"Item {itemNumber}"));
                    }

                    foreach (var toolConfig in toolConfigs)
                    {
                        var tool = new Tool(
                            toolConfig.ToolNumber,
                            toolConfig.ToolName,
                            toolConfig.LikertScale,
                            toolConfig.LevelRanges,
                            toolConfig.LevelDetermination
                        );
                        tools.Add(tool);
                    }

                    return tools;
                }

                private static List<ToolConfig> ReadToolConfigs(Workbook workbook)
                {
                    var sheet = workbook.Worksheets["details"];
                    var toolConfigs = new List<ToolConfig>();

                    int row = 0;
                    while (sheet.Cells[row, 0].Value != null)
                    {
                        int toolNumber = sheet.Cells[row, 0].IntValue;
                        string toolName = sheet.Cells[row, 1].StringValue;
                        var toolConfig = new ToolConfig(toolNumber, toolName);

                        row += 2;
                        while (sheet.Cells[row, 0].Value != null && sheet.Cells[row, 0].Value.ToString() != "Level")
                        {
                            toolConfig.LikertScale.Add(sheet.Cells[row, 0].IntValue);
                            toolConfig.LikertLabels.Add(sheet.Cells[row, 1].StringValue);
                            row++;
                        }

                        toolConfig.LevelDetermination = sheet.Cells[row, 1].StringValue;

                        row++;
                        while (sheet.Cells[row, 0].Value != null)
                        {
                            string range = sheet.Cells[row, 1].StringValue.Split('(', ')')[1];
                            string label = sheet.Cells[row, 1].StringValue.Split('(')[0].Trim();
                            toolConfig.LevelRanges.Add(new LevelRange(range, label ,0));
                            row++;
                        }

                        toolConfigs.Add(toolConfig);
                        row++;
                    }

                    return toolConfigs;
                }

                public static List<Participant> ReadParticipants(Workbook workbook)
                {
                    var participants = new List<Participant>();
                    var sheet = workbook.Worksheets[0];

                    for (int row = 1; row < sheet.Cells.MaxDataRow + 1; row++)
                    {
                        var participant = new Participant();
                        var dataRow = sheet.Cells.GetRow(row);

                        participant.Name = dataRow[0].StringValue;

                        for (int col = 1; col <= sheet.Cells.MaxDataColumn; col++)
                        {
                            string header = sheet.Cells[0, col].StringValue;
                            if (!header.StartsWith("Q"))
                                continue;

                            int response = dataRow[col].IntValue;
                            participant.Responses.Add(response);
                        }

                        participants.Add(participant);
                    }

                    return participants;
                }
            }

            public class ToolConfig
            {
                public int ToolNumber { get; set; }
                public string ToolName { get; set; }
                public List<int> LikertScale { get; set; }
                public List<string> LikertLabels { get; set; }
                public string LevelDetermination { get; set; }
                public List<LevelRange> LevelRanges { get; set; }

                public ToolConfig(int toolNumber, string toolName)
                {
                    ToolNumber = toolNumber;
                    ToolName = toolName;
                    LikertScale = new List<int>();
                    LikertLabels = new List<string>();
                    LevelRanges = new List<LevelRange>();
                }
            }

            public class LevelRange
            {
                public string Range { get; set; }
                public string Label { get; set; }
                public int Level { get; set; }  // Added Level property

                public LevelRange(string range, string label, int level)  // Modified constructor to accept level
                {
                    Range = range;
                    Label = label;
                    Level = level;
                }
            }


            private bool IsInRange(double score, string range)
            {
                if (range.StartsWith("<="))
                {
                    double threshold = double.Parse(range.Substring(2));
                    return score <= threshold;
                }
                else if (range.StartsWith("<"))
                {
                    double threshold = double.Parse(range.Substring(1));
                    return score < threshold;
                }
                else if (range.StartsWith(">="))
                {
                    double threshold = double.Parse(range.Substring(2));
                    return score >= threshold;
                }
                else if (range.StartsWith(">"))
                {
                    double threshold = double.Parse(range.Substring(1));
                    return score > threshold;
                }
                else if (range.StartsWith("="))
                {
                    double threshold = double.Parse(range.Substring(1));
                    return score == threshold;
                }
                else if (range.Contains("-"))
                {
                    var parts = range.Split('-');
                    double lowerBound = double.Parse(parts[0]);
                    double upperBound = double.Parse(parts[1]);
                    return score >= lowerBound && score <= upperBound;
                }
                else
                {
                    throw new ArgumentException($"Invalid range format: {range}");
                }
            }

            public void CalculateOverallScores(List<int> responses)
            {
                double totalScore = Subscales.Sum(s => s.TotalScores.Last());
                int totalItems = Subscales.Sum(s => s.Items.Count);

                double minScore = totalItems * LikertScale.Min();
                double maxScore = totalItems * LikertScale.Max();

                double averageScore = totalScore / totalItems;
                double percentScore = (totalScore - minScore) / (maxScore - minScore) * 100;

                double scoreToUse = 0;
                switch (LevelDetermination.ToLower())
                {
                    case "percent":
                        scoreToUse = percentScore;
                        break;
                    case "total":
                        scoreToUse = totalScore;
                        break;
                    case "average":
                        scoreToUse = averageScore;
                        break;
                    default:
                        throw new ArgumentException($"Invalid LevelDetermination: {LevelDetermination}");
                }

                int computedLevel = LevelRanges
                    .FirstOrDefault(range => IsInRange(scoreToUse, range.Range))?.Level ?? 0;

                TotalScores.Add(totalScore);
                AverageScores.Add(averageScore);
                PercentScores.Add(percentScore);
                ComputedLevels.Add(computedLevel);
            }

            public override string ToString()
            {
                return $"Tool(Name={ToolName}, TotalScores={string.Join(", ", TotalScores)}, AverageScores={string.Join(", ", AverageScores)}, PercentScores={string.Join(", ", PercentScores)}, Levels={string.Join(", ", ComputedLevels)})";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(ExcelFunctions.filepath);

                var tools = ExcelReader.ReadTools(workbook);
                var participants = ExcelReader.ReadParticipants(workbook);

                foreach (var tool in tools)
                {
                    foreach (var participant in participants)
                    {
                        foreach (var subscale in tool.Subscales)
                        {
                            subscale.CalculateScores(participant.Responses, tool.LikertScale, tool.LevelRanges , tool);
                        }

                        tool.CalculateOverallScores(participant.Responses);
                    }

                    MessageBox.Show(tool.ToString(), "Tool Results");
                    foreach (var s in tool.Subscales)
                    {
                        MessageBox.Show(s.ToString(), "Subscale Results");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
