using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.StatClasses
{
    public static class WordTableStatDesign
    {
        

        private static readonly Dictionary<string, Action<WordTableDesignContext>> _designs
        = new(StringComparer.OrdinalIgnoreCase)
    {
        { "Comparative_Default", WordT_Comparative_Default }
    };

        public static void Execute(WordTableDesignContext context)
        {
            string key = $"{context.Table.TableType}_{context.Table.TableDesignType}";

            if (_designs.TryGetValue(key, out var action))
                action(context);
            else
                throw new Exception($"No design found for: {key}");
        }

        // ================= DESIGNS =================

        // TableDesignType = "Default";
        // TableType = "comparative";
        public static void WordT_Comparative_Default(WordTableDesignContext context)
        {
            var table = context.Table;
            MessageBox.Show("testForm");
            // your logic here
        }




    }
}
