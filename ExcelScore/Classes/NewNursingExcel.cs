using Aspose.Cells;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Cell = Aspose.Cells.Cell;
using Color = System.Drawing.Color;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;

namespace ExcelScore.Classes
{
    public class NewNursingExcel
    {
        public static Style ScoreStyle(Aspose.Cells.Worksheet ScoreSheet , Color Mycolor , int row , int col)
        {
            Cell mycell = ScoreSheet.Cells[row, col];
            Aspose.Cells.Style style1 = mycell.GetStyle();

            // Set the text alignment to center
            style1.HorizontalAlignment = TextAlignmentType.Center;
            style1.VerticalAlignment = TextAlignmentType.Center;




            // Set the background color of the cell
            style1.ForegroundColor = Mycolor;
            style1.Pattern = BackgroundType.Solid;
            style1.Font.IsBold = true;

            return style1;  
        }

        public static void SetCellValueAndCenterText(Worksheet worksheet, int firstRow, int firstColumn, string value, Color fillColor)
        {
            Cell Mycell = worksheet.Cells[firstRow, firstColumn];
            Mycell.PutValue(value, true);

            Style style = Mycell.GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.Pattern = BackgroundType.Solid;  // Ensure the fill color is applied
            style.Font.IsBold = true;
            style.ForegroundColor = fillColor;  // This is the actual fill color

            Mycell.SetStyle(style);
        }

        public static void SetCellValueAndCenter_int(Worksheet worksheet, int firstRow, int firstColumn, double value)
        {
            Cell Mycell = worksheet.Cells[firstRow, firstColumn];
            Mycell.PutValue(value);

            Style style = Mycell.GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            style.IsFillApplied = false;  

            Mycell.SetStyle(style);
        }


        public static void EnsureSheetExists(Workbook workbook, string sheetName, int position)
        {
            
            workbook.Worksheets.RemoveAt(sheetName);
            workbook.Worksheets.Insert(position, SheetType.Worksheet, sheetName);

        }


    }
}
