using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

namespace ExcelScore.Classes
{
    public class DefinitionUnderTable
    {
        public const string IQRNotBold = "IQR: ";
        public const string IQRBold = "Inter quartile range";

        public const string SDNot = "SD: ";
        public const string SDB = "Standard deviation";


        public const string tstudentTest = "t: Student t-test";
        public const string F_OneWay_Anova = "F: F for One way ANOVA test, Pairwise comparison bet. each 2 groups was done using Post Hoc Test (Tukey)";


        public const string U_Test = "U: Mann Whitney test";
        public const string H_Test = "H: H for Kruskal Wallis test, Pairwise comparison bet. each 2 groups was done using Post Hoc Test (Dunn's for multiple comparisons test)";

        public const string chisqaure_test = "χ²: Chi square test";

        public const string p_stat_Sig = "*: Statistically significant at p ≤ 0.05";

        public string p_comparing(string numberofgroups)
        {
            string p = "p: p value for comparing between the " + numberofgroups + " studied groups";
            return p;
        }








        //public const string String2 = "Value2";

    }
}
