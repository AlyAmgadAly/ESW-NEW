using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.StatClasses
{
    public class WordTableDesignContext
    {
        public StatTable Table { get; }



        public WordTableDesignContext(StatTable table)
        {
            Table = table ?? throw new ArgumentNullException(nameof(table));
        }
    }
}
