using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelScore.Classes
{
    public class NewDomainClass
    {
        public string DomainName { get; set; }

        public int RangeFrom { get; set; }

        public int RangeTo { get; set; }

        public string Level_Percent_Or_TotalScore { get; set; }

        public string Level_Comparison_Choose { get; set; }

        public string Level_RangeFrom { get; set; }

        public string Level_Comparison_Or_RangeTo { get; set; }

        public string Level_Comparison_Value { get; set; }

        public List<string> Questions_Not_Reversed = new List<string>();

        public List<string> Questions_Reversed = new List<string>();    


        public Dictionary<string, int> QuestionsSorted_ReverseValue = new Dictionary<string, int>();

        public List<string> Level_Comparison_Choice = new List<string>();

        public List<int> Level_From_To = new List<int>();

        public List<int> Level_Value = new List<int>();

        public string LikertScore { get; set; }

    }
}
