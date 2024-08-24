using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{

    public class Parameter
    {
        public string Name { get; set; }
        public bool IsGroup { get; set; }

        public bool hasLowerN { get; set; }
        public string NominalOrScale { get; set; }
        public string NormalOrAbnormal { get; set; }
        public Dictionary<double, double> EachGroupCount { get; set; } = new Dictionary<double, double>();
        public List<double> ParameterValues { get; set; } = new List<double>();
        public Dictionary<double, List<double>> GroupedParameterValues { get; set; } = new Dictionary<double, List<double>>();

        public List<string> FPairwise { get; set; } = new List<string>();
        public bool ISFAnovaSig { get; set; }

        public bool IsIndependentCorr { get; set; }
        public Dictionary<int , string> DIC_LablesIfNomainal { get; set; } = new Dictionary<int, string>();
        public List<string> LablesIfNomainal { get; set; } = new List<string>();
        public bool NominalIsYes { get; set; }

        public bool Isfisher { get; set; }
        public Dictionary<(int, int), string> LabelPairwise { get; set; } = new Dictionary<(int, int), string>();

        public Dictionary<double, Dictionary<string, string>> FormattedValues { get; set; } = new Dictionary<double, Dictionary<string, string>>();
        public Parameter()
        {
            GroupedParameterValues = new Dictionary<double, List<double>>();
            ParameterValues = new List<double>();
            LabelPairwise = new Dictionary<(int, int), string>(); // Initialize ParameterValues list in constructor as well
        }
    }
    public class ComparativeTable
    {
        public string TableName { get; set; }
        public List<Parameter> Parameters { get; set; }
        public int GroupColumnIndex { get; set; } // Property to store the group column index
        public bool HasTotalColumn { get; set; }

        public bool HasSelect { get; set; }

        public string CorreType { get; set; }
        public string LetterType { get; set; }
        public List<int> SelectedValues { get; set; }
        public List<string> TestsDone { get; set; }
        public ComparativeTable()
        {
            TestsDone = new List<string>(); 
            Parameters = new List<Parameter>();
            GroupColumnIndex = -1; // Initialize to an invalid value
        }

        public bool hasSig { get; set; }


    }


}
