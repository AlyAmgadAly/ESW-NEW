using Syncfusion.DocIO.DLS;
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
            var StatTable = context.Table;
            var wordobj = context.Wordobj;
            
            

            IWSection section = wordobj.CreatePortraitSection();

            StatParameter groupParameter = StatTable.GetGroupParameters().FirstOrDefault();

            int groupcount = groupParameter.ValueLabels.Count;

            int ColCount = 3 + groupcount*2 ;




            wordobj.AddComparativeTitle(section,"test",1);
        }

        public static int Get_Comparative_Default_ColCount()
        {


            return 0;
        }




    }
}
