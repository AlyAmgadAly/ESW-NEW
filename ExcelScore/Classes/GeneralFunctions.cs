using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{
    public  class GeneralFunctions
    {


        public bool PvalueHasSig(string pvalue)
        {
            bool IsSig = false;


            if(pvalue == "<0.001")
            {
                IsSig = true;
            }
            else
            {
                double p = double.Parse(pvalue);
                if (p < 0.05)
                {
                    IsSig = true;
                }
                else if (p >= 0.05)
                {
                    IsSig = false;
                }
            }

            return IsSig;


        }

        public static void RemoveInvalidEntries(ref List<List<double>> dataLists)
        {
            if (dataLists == null || dataLists.Count == 0)
                return;

            int minCount = dataLists.Min(list => list.Count);
            HashSet<int> indicesToRemove = new HashSet<int>();

            // Identify indices where any list has -999
            for (int i = 0; i < minCount; i++)
            {
                if (dataLists.Any(list => list[i] == -999))
                {
                    indicesToRemove.Add(i);
                }
            }

            // Remove values at marked indices (iterate backwards to avoid shifting issues)
            foreach (var list in dataLists)
            {
                for (int i = indicesToRemove.Count - 1; i >= 0; i--)
                {
                    int indexToRemove = indicesToRemove.ElementAt(i);
                    if (indexToRemove < list.Count) // Safety check
                    {
                        list.RemoveAt(indexToRemove);
                    }
                }
            }
        }

    }
}
