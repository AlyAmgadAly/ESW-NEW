using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{
    public  class TableClasses
    {
        WordClass wordObj = new WordClass();
        public IWTable Items_Descriptive_Nursing_Design(int itemCount , int likertcount , string TableHeaderName , Tool CurrentTool)
        {
            IWSection section = wordObj.CreatePortraitSection();

            int WordTableRows = 2 + itemCount;

            int WordTableColumns = 2 + (CurrentTool.LikertScale.Keys.Count * 2);


            IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

            wordObj.GeneralTableFormat(table);


            wordObj.ApplyMerges(table, likertcount);

            wordObj.ApplyBorders(table, WordTableRows, WordTableColumns);

            wordObj.SetWidths(table, WordTableRows, likertcount);

            wordObj.Add_Header_Items_Nusring(table, likertcount, TableHeaderName, CurrentTool);


            wordObj.AddNo_perc_Center(table, WordTableColumns);

            wordObj.AddQuestionNo(table, itemCount);


            return table;
        }

    }
}
