using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScore.Classes
{
    public class NewNursingExcel
    {
        public static void SetMergedCellValueAndCenterText(Worksheet worksheet, int firstRow, int firstColumn, string value)
        {
            Cell mergedCell = worksheet.Cells[firstRow, firstColumn];
            mergedCell.PutValue(value);

            Style style = mergedCell.GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            mergedCell.SetStyle(style);
        }

        public static void EnsureSheetExists(Workbook workbook, string sheetName, int position)
        {
            
            workbook.Worksheets.RemoveAt(sheetName);
            workbook.Worksheets.Insert(position, SheetType.Worksheet, sheetName);

        }
    }
}
