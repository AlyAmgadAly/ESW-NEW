using MathNet.Numerics.Statistics;
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

        public Dictionary<string ,  Dictionary<double, List<double>>> GroupedParameterValues_Relation { get; set; } = new Dictionary<string, Dictionary<double, List<double>>>();

        public Dictionary<string , Dictionary<double, Dictionary<string, string>>> FormattedValues_Relation { get; set; } = new Dictionary<string, Dictionary<double, Dictionary<string, string>>>();

        public List<string> FPairwise { get; set; } = new List<string>();
        public bool ISFAnovaSig { get; set; }

        public bool IsIndependentCorr { get; set; }
        public Dictionary<int , string> DIC_LablesIfNomainal { get; set; } = new Dictionary<int, string>();
        public List<string> LablesIfNomainal { get; set; } = new List<string>();
        public bool NominalIsYes { get; set; }

        public bool Isfisher { get; set; }

        public bool IsMonteCarlo { get; set; }

        public string Periods_ParameterName { get; set; }

        public bool ParameterNominalSig { get; set; }
        public Dictionary<(int, int), string> LabelPairwise { get; set; } = new Dictionary<(int, int), string>();

        public Dictionary<double, Dictionary<string, string>> FormattedValues { get; set; } = new Dictionary<double, Dictionary<string, string>>();
        public Parameter()
        {
            GroupedParameterValues_Relation = new Dictionary<string, Dictionary<double, List<double>>>();
            GroupedParameterValues = new Dictionary<double, List<double>>();
            ParameterValues = new List<double>();
            LabelPairwise = new Dictionary<(int, int), string>(); // Initialize ParameterValues list in constructor as well
            FormattedValues_Relation = new Dictionary<string, Dictionary<double, Dictionary<string, string>>>();
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

        public string FormatType { get; set; }
        public ComparativeTable()
        {
            TestsDone = new List<string>(); 
            Parameters = new List<Parameter>();
            GroupColumnIndex = -1; // Initialize to an invalid value
        }

        public bool AllNormal()
        {
            return Parameters != null && Parameters
                .Where(p => !p.IsGroup) // Exclude groups
                .All(p => p.NormalOrAbnormal == "Normal");
        }

        public bool AllAbnormal()
        {
            return Parameters != null && Parameters
                .Where(p => !p.IsGroup) // Exclude groups
                .All(p => p.NormalOrAbnormal == "Abnormal");
        }
        public static Parameter GetGroupParamter(ComparativeTable comparativeTable)
        {
            Parameter GroupParameter = null;

            foreach (var Parameter in comparativeTable.Parameters)
            {
                if(Parameter.IsGroup)
                {
                    GroupParameter = Parameter;
                    break;
                }

            }


            return GroupParameter;
        }

        public bool hasSig { get; set; }


    }


}
