using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{
    public class AnovaTestResult
    {

        public string TestValue { get; set; }
        public string PValue { get; set; }
        public List<string[]> PairwiseComparisons { get; set; }
        
    }
}
