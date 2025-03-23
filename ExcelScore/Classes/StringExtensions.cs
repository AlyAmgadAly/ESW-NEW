using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{
    public static class StringExtensions
    {
        public static string RemoveDots(this string input)
        {
            return input.Replace(".", "");
        }
    }
}
