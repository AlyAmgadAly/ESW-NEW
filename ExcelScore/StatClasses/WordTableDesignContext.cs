using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
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

        public WordClass Wordobj { get; }


        public WordTableDesignContext(StatTable table ,WordClass wordobj)
        {
            Table = table ?? throw new ArgumentNullException(nameof(table));
            Wordobj = wordobj ?? throw new ArgumentNullException(nameof(wordobj));
        }
    }
}
