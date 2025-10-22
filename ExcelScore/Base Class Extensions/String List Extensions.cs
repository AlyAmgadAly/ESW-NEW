using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Base_Class_Extensions
{
    public static class String_List_Extensions
    {
        public static bool ContainsExact(this IEnumerable<string> list, string value,
        StringComparison comparison = StringComparison.Ordinal)
        {
            if (list == null || value == null)
                return false;

            return list.Any(x => string.Equals(x, value, comparison));
        }

    }
}
