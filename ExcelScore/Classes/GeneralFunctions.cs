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
    }
}
