using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

using System.Windows.Forms;
using System.Data.Common;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Aspose.Cells;
using System.ComponentModel;
using ExcelScore.Classes;
using System.Diagnostics;

namespace ExcelScore
{
    public class ExcelFunctions
    {

        public static Aspose.Cells.Workbook workbook;
        public static Aspose.Cells.Worksheet worksheet;
        public static Aspose.Cells.Worksheet Sheet2;
        public static string filepath { get; set; }


        //public int ExcelDataCount = Sheet1.Cells.MaxDataRow - 1;


        public string GetNewWorkBook()
        {

            string filepathnew = "";

            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            if (op.ShowDialog() == DialogResult.OK)
            {
                filepathnew = op.FileName;
                
            }

            return filepathnew;
        }


        public Worksheet GetWorksheet()
        {
            return worksheet;
        }

        public Worksheet GetSheet2()
        {
            return Sheet2;
        }

        public Workbook GetWorkbook() 
        {
            return workbook;
        }
        public System.Data.DataTable import()
        {

            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            if (op.ShowDialog() == DialogResult.OK)
            {
                ExcelFunctions.filepath = op.FileName;
                workbook = new Aspose.Cells.Workbook(filepath);

                // Accessing the first worksheet in the Excel file
                worksheet = workbook.Worksheets[0];
                Sheet2 = workbook.Worksheets[1];
                System.Data.DataTable dataTable = new System.Data.DataTable();
                for (int col = 0; col < worksheet.Cells.MaxDataColumn + 1; col++)
                {
                    object columnName = worksheet.Cells[0, col].Value;
                    dataTable.Columns.Add(columnName.ToString().Trim());


                }
                for (int row = 1; row < worksheet.Cells.MaxDataRow + 1; row++)
                {
                    DataRow datarow = dataTable.NewRow();
                    for (int col = 0; col < worksheet.Cells.MaxDataColumn + 1; col++)
                    {
                        object cellvalue = worksheet.Cells[row, col].Value;
                        datarow[col] = cellvalue;

                    }
                    dataTable.Rows.Add(datarow);
                }
               
                LoadExcel.MydataTable = dataTable;
                return dataTable;

            }
            return null;
        }



        public List<string> ReadHeaderColumnsExcel()
        {
            List<string> columns = new List<string>();
            for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
            {
               object cellvalue = worksheet.Cells[0, col].Value;
               string cellvaluestr = cellvalue.ToString();
               columns.Add(cellvaluestr);  
            }
            return columns;
        }

        
        public void ReadLablesIfNominal(ComparativeTable comparativeTable)
        {
            Worksheet Sheet2 = workbook.Worksheets[1];

            for(int column = 0; column <= Sheet2.Cells.MaxDataColumn;column++)
            {
                foreach (var parameter in comparativeTable.Parameters)
                {
                    
                    if (parameter.Name == Sheet2.Cells[0 , column].Value.ToString() && (parameter.NominalOrScale == "Nominal" || parameter.IsGroup))
                    {

                        if (Sheet2.Cells[1, column].Value.ToString() == "Nominal")
                        {
                            for (int row = 2; row <= worksheet.Cells.MaxDataRow; row++)
                            {
                                if (Sheet2.Cells[row, column].Value != null)
                                {
                                    Cell cell = Sheet2.Cells[row, column];
                                    string cellstring = cell.Value.ToString();




                                    var parts = cellstring.Split('|');
                                    if (parts.Length >= 2)
                                    {
                                        var keyPart = parts[0].Trim();
                                        var valuePart = parts[1].Trim();

                                        // Parse the key (assuming it's always an integer at the start of the keyPart)
                                        if (int.TryParse(keyPart, out int key))
                                        {
                                            // Assign to the dictionary
                                            parameter.DIC_LablesIfNomainal[key] = valuePart;
                                            parameter.LablesIfNomainal.Add(cellstring);
                                        }
                                        else
                                        {
                                            // Handle parsing error - invalid key
                                            continue;
                                        }
                                    }



                                }
                            }

                        }
                    }
                    
                }
                
            }

            

        }

        public List<string> GetLevelCount(Worksheet Sheet , string SelectedOverall)
        {
            List<string> level = new List<string>();
            int rowCount = Sheet.Cells.MaxDataRow + 1;

            for(int i=0; i < rowCount; i++)
            {
                string value = Sheet.Cells[i, 0].Value.ToString();
                if (SelectedOverall == value)
                {
                    for (int j = i+1; j < rowCount; j++)
                    {
                        if (Sheet.Cells[j, 0].IsMerged == true)
                        {
                            break;
                        }
                        else if (Sheet.Cells[j, 0].IsMerged == false)
                        {
                            level.Add(Sheet.Cells[j, 2].Value.ToString());
                        }

                    }


                }


            }
            
            return level;

            


        }
        
        



    }
}

