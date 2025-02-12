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

        public Dictionary<int , string> LikertScale = new Dictionary<int , string>();
        public List<List<LevelRange>> ToolLevels { get; set; } = new List<List<LevelRange>>();
        public List<string> LevelDetermination { get; set; } = new List<string>();

        public List<double> TotalScores { get; set; } = new List<double>();
        public List<double> AverageScores { get; set; } = new List<double>();
        public List<double> PercentScores { get; set; } = new List<double>();
        public List<List<int>> ComputedToolLevels { get; set; } = new List<List<int>>();

        

        public void AddScale(Scale scale)
        {
            scale.ParentTool = this;
            Scales.Add(scale);
        }
        

        public class Scale
        {
            public string Scale_Name { get; set; }

            public string Scale_Full_Name { get; set; }
            public List<Item> Items { get; set; }

            public List<Subscale> Subscales { get; set; } = new List<Subscale>();
            // Lists to store scores for each participant
            public List<double> TotalScores { get; set; } = new List<double>();
            public List<double> AverageScores { get; set; } = new List<double>();
            public List<double> PercentScores { get; set; } = new List<double>();
            public List<List<int>> ComputedSubLevels { get; set; } = new List<List<int>>();

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
            public List<Item> Items { get; set; }

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
