using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelScore.Forms.Waiting.Tool;

namespace ExcelScore.Classes
{
    public  class Tool
    {
        public string ToolNumber { get; set; }
        public string ToolName { get; set; }
        public List<Scale> Scales { get; set; } = new List<Scale>();

        public List<Item> ToolItems { get; set; } = new List<Item>();

        public Dictionary<int , string> LikertScale = new Dictionary<int , string>();
        public List<List<LevelRange>> ToolLevels { get; set; } = new List<List<LevelRange>>();
        public List<string> LevelDetermination { get; set; } = new List<string>();

        public List<double> TotalScores { get; set; } = new List<double>();
        public List<double> AverageScores { get; set; } = new List<double>();
        public List<double> PercentScores { get; set; } = new List<double>();
        public List<List<int>> ComputedToolLevels { get; set; } = new List<List<int>>();

        public bool hasscales { get; set; }

        public List<Tool> PeriodsTools { get; set; } = new List<Tool>();

        public void AddScale(Scale scale)
        {
            scale.ParentTool = this;
            Scales.Add(scale);
        }


        public Scale GetScaleFromPeriodTool(string periodToolName, string scaleName)
        {
            Tool periodTool = PeriodsTools.FirstOrDefault(pt => pt.ToolName == periodToolName);
            return periodTool?.Scales.FirstOrDefault(s => s.Scale_Name == scaleName);
        }

        public Tool GetPeriodTool(string toolNumber, string toolName)
        {
            return PeriodsTools.FirstOrDefault(pt => pt.ToolNumber == toolNumber && pt.ToolName == toolName);
        }

        public void UpdateAllPeriodTools()
        {
            foreach (var periodTool in PeriodsTools)
            {
                if (periodTool != null)
                {
                    // Set each scale's ParentTool to the periodTool before transferring
                    foreach (var scale in this.Scales)
                    {
                        scale.ParentTool = periodTool;
                    }

                    // Transfer scales to periodTool and clear from main tool
                    periodTool.Scales = new List<Scale>(this.Scales);
                }

                // Copy other properties
                periodTool.LikertScale = new Dictionary<int, string>(this.LikertScale);
                periodTool.ToolLevels = new List<List<LevelRange>>(this.ToolLevels);
                periodTool.LevelDetermination = new List<string>(this.LevelDetermination);
                periodTool.hasscales = this.hasscales;
            }

            // Clear scales from the main tool after transferring to all PeriodTools
            this.Scales.Clear();
        }





        public void RankScales()
        {
            if (Scales.Count == 0) return;

            // Compute the mean of PercentScores for each scale
            var scaleRanks = Scales
                .Where(scale => scale.PercentScores.Count > 0) // Ignore scales without scores
                .Select(scale => new
                {
                    Scale = scale,
                    MeanPercent = scale.PercentScores.Average()
                })
                .OrderByDescending(x => x.MeanPercent) // Rank from highest to lowest
                .ToList();

            // Assign ranks with proper handling for ties
            int rank = 1;  // The rank to assign
            for (int i = 0; i < scaleRanks.Count; i++)
            {
                if (i > 0 && Math.Abs(scaleRanks[i].MeanPercent - scaleRanks[i - 1].MeanPercent) > 0.0001)
                {
                    // If current MeanPercent is different, update rank to the correct position
                    rank = i + 1;
                }

                scaleRanks[i].Scale.Rank = rank; // Assign rank

                // Debugging output
                //Console.WriteLine($"Scale: {scaleRanks[i].Scale.Scale_Name}, Mean: {scaleRanks[i].MeanPercent}, Rank: {scaleRanks[i].Scale.Rank}");
            }
        }




        public class Scale
        {
            public string Scale_Name { get; set; }

            public string Scale_Full_Name { get; set; }
            public List<Item> Items { get; set; } = new List<Item>();

            public List<Subscale> Subscales { get; set; } = new List<Subscale>();
            // Lists to store scores for each participant
            public List<double> TotalScores { get; set; } = new List<double>();
            public List<double> AverageScores { get; set; } = new List<double>();
            public List<double> PercentScores { get; set; } = new List<double>();
            public List<List<int>> ComputedSubLevels { get; set; } = new List<List<int>>();

            public int Rank { get; set; } = 0;

            public Tool ParentTool { get; set; }

            public void AddSubscale(Subscale subscale)
            {
                subscale.ParentScale = this;
                Subscales.Add(subscale);
            }

            public void AddItem(Item item)
            {
                Items.Add(item);
            }
        }

        public class Subscale
        {
            public string Subscale_Name { get; set; }

            public string Subscale_Full_Name { get; set; }
            public List<Item> Items { get; set; } = new List<Item>();



            // Lists to store scores for each participant
            public List<double> TotalScores { get; set; } = new List<double>();
            public List<double> AverageScores { get; set; } = new List<double>();
            public List<double> PercentScores { get; set; } = new List<double>();
            public List<List<int>> ComputedSubLevels { get; set; } = new List<List<int>>();

            public Scale ParentScale { get; set; }

            public void AddItem(Item item)
            {
                Items.Add(item);
            }
        }

        public class Item
        {
            public string Id { get; set; }
            public string Text { get; set; }
            public bool IsReverse { get; set; }
            public List<double> ParticipantResponses { get; set; } = new List<double>();

            
        }

        public class LevelRange
        {
            public string Range { get; set; }
            public string Label { get; set; }
            
        }


    }
}
