using Aspose.Cells;
using ExcelScore.Classes;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class DescriptiveFrm : Form
    {

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        public List<ComparativeTable> ComparativeTables = new List<ComparativeTable>();

        static ExcelFunctions excelFunctionsobj = new ExcelFunctions();

        Worksheet worksheet = excelFunctionsobj.GetWorksheet();
        public DataGridView Dgv { get; set; }

        List<string> ColumnExcelheaders = excelFunctionsobj.ReadHeaderColumnsExcel();
        public List<ComparativeClass> ComparativeClassesList = new List<ComparativeClass>();

        public string CurrentAction = "Descriptive";
        public DescriptiveFrm()
        {
            InitializeComponent();
        }

        private void panelmove_Paint(object sender, PaintEventArgs e)
        {

        }
        public void AddHeadersToParameter()
        {
            foreach (string columnName in ColumnExcelheaders)
            {
                ComparativeClass comparativeClass = new ComparativeClass();
                comparativeClass.ParameterName = columnName;
                ComparativeClassesList.Add(comparativeClass);
                list_AllParameters.Items.Add(comparativeClass.ParameterName);
            }
        }
        private void DescriptiveFrm_Load(object sender, EventArgs e)
        {
            AddHeadersToParameter();

            lbl_CurrentAction.Text = "Current Action : Descriptive";
        }

        public int CountRows(ComparativeTable table)
        {
            int rowCount = 0;
            bool HasNominal = DetermineHasNominal(table);
            foreach (var parameter in table.Parameters)
            {
                if(HasNominal)
                {
                    if (parameter.NominalOrScale == "Nominal")
                    {
                        int distinctValuesCount = parameter.DIC_LablesIfNomainal.Keys.Count;
                        rowCount += distinctValuesCount + 1;

                    }
                    else if (parameter.NominalOrScale == "Scale")
                    {
                        // For scale parameters, add 4 (assuming you want to count 4 rows per parameter)
                        rowCount += 4;
                    }
                }
                else if (!HasNominal)
                {
                    rowCount++;
                }
                
            }



            return rowCount;
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
            ChooseFrm chooseFrm = new ChooseFrm();
            chooseFrm.Dgv = Dgv;
            chooseFrm.Show();
        }

        private void list_Nominal_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pic_AllParaToNominal_Click(object sender, EventArgs e)
        {
            foreach (object selectedItem in list_AllParameters.SelectedItems)
            {
                list_Nominal.Items.Add(selectedItem.ToString());
            }
        }

        private void pic_AllParaToScale_Click(object sender, EventArgs e)
        {
            foreach (object selectedItem in list_AllParameters.SelectedItems)
            {
                list_Scale.Items.Add(selectedItem.ToString());
            }
        }

        private void pic_RemoveNominalList_Click(object sender, EventArgs e)
        {
            if (list_Nominal.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_Nominal.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_Nominal.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_removeScale_Click(object sender, EventArgs e)
        {
            if (list_Scale.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_Scale.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_Scale.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }
        public void AddTableUI()
        {
            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&
                (list_Nominal.Items.Count > 0 || list_Scale.Items.Count > 0))
            {
                // Check if the table name already exists
                bool tableExists = false;
                foreach (var existingTable in cmb_TableNames.Items)
                {
                    if (string.Equals(existingTable.ToString(), txt_TableName.Text, StringComparison.OrdinalIgnoreCase))
                    {
                        tableExists = true;
                        break;
                    }
                }

                if (!tableExists)
                {
                    cmb_TableNames.Items.Add(txt_TableName.Text);
                    AddTableClass(txt_TableName.Text);
                    MessageBox.Show("Added Table " + txt_TableName.Text);
                }
                else
                {
                    MessageBox.Show("Table name already exists!");
                }
            }
            else
            {
                MessageBox.Show("Please fill in table name and add items to lists!");
            }
        }

        public void AddTableClass(string tableName)
        {
            var comparativeTable = new ComparativeTable
            {
                TableName = tableName,
                Parameters = new List<Parameter>()
            };
            // Add parameters from list_Nominal
            foreach (var item in list_Nominal.Items)
            {
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    NominalOrScale = "Nominal",
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }



            // Add parameters from list_NormalScale
            foreach (var item in list_Scale.Items)
            {
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    NominalOrScale = "Scale",
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }


            ComparativeTables.Add(comparativeTable);




        }
        private void pic_addTable_Click(object sender, EventArgs e)
        {
            AddTableUI();
        }

        private void cmb_TableNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedTableName = cmb_TableNames.SelectedItem.ToString();

            // Find the ComparativeTable instance corresponding to the selected table name
            var selectedTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

            // Clear existing items in list_ViewTableParameters
            list_ViewTableParameters.Items.Clear();

            // Add parameters of the selected table to list_ViewTableParameters
            if (selectedTable != null)
            {
                foreach (var parameter in selectedTable.Parameters)
                {
                    if (parameter.IsGroup)
                    {
                        list_ViewTableParameters.Items.Add("Groups : " + parameter.Name);
                    }
                    else if (!parameter.IsGroup)
                    {
                        list_ViewTableParameters.Items.Add(parameter.Name);
                    }

                }
            }
        }

        private void pic_removeTableSelected_Click(object sender, EventArgs e)
        {
            string selectedTableName = cmb_TableNames.SelectedItem?.ToString();

            if (!string.IsNullOrEmpty(selectedTableName))
            {
                // Remove the selected table name from the ComboBox
                cmb_TableNames.Items.RemoveAt(cmb_TableNames.SelectedIndex);

                // Find and remove the corresponding ComparativeTable instance from the ComparativeTables list
                var tableToRemove = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);
                if (tableToRemove != null)
                {
                    ComparativeTables.Remove(tableToRemove);
                }
            }

            cmb_TableNames.Text = "";
            list_ViewTableParameters.Items.Clear();
        }

        public void DescriptiveBasic()
        {
            GetDataValues();
            
            foreach (ComparativeTable table in ComparativeTables)
            {
                excelFunctionsobj.ReadLablesIfNominal(table);

                foreach (Parameter parameter in table.Parameters)
                {
                    

                    // Get the column index of the parameter
                    int parameterColumnIndex = -1;

                    for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
                    {
                        object cellValue = worksheet.Cells[0, col].Value;
                        string parameterName = cellValue?.ToString();

                        if (!string.IsNullOrEmpty(parameterName) && parameterName == parameter.Name)
                        {
                            parameterColumnIndex = col;
                            break;
                        }
                    }

                    if (parameterColumnIndex != -1)
                    {
                        // Iterate through each row to group the data
                        for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
                        {
                            // Get the parameter value for the current row
                            object parameterCellValue = worksheet.Cells[row, parameterColumnIndex].Value;
                            double parameterValue;

                            if (parameterCellValue.ToString() == ".")
                            {
                                parameter.hasLowerN = true;
                                continue;
                            }
                            if (parameterCellValue != null && double.TryParse(parameterCellValue.ToString(), out parameterValue))
                            {
                                // Add the parameter value to the corresponding group
                                if (!parameter.GroupedParameterValues.ContainsKey(1))
                                {
                                    parameter.GroupedParameterValues[1] = new List<double>();
                                }

                                parameter.GroupedParameterValues[1].Add(parameterValue);
                            }



                        }

                    }
                }
            }
            FormatParameters();
        }

        public void FormatParameters()
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                //MessageBox.Show(table.TableName);
                foreach (Parameter parameter in table.Parameters)
                {
                    

                    foreach (var kvp in parameter.GroupedParameterValues)
                    {
                        double groupValue = kvp.Key;
                        List<double> values = kvp.Value;
                        int totalCount = values.Count;

                        // Initialize the formatted values dictionary for the current group value
                        if (!parameter.FormattedValues.ContainsKey(groupValue))
                        {
                            parameter.FormattedValues[groupValue] = new Dictionary<string, string>();
                        }

                        // Check if the parameter is nominal
                        if (parameter.NominalOrScale == "Nominal")
                        {
                            // Calculate frequency and percentage for each distinct value in the group
                            foreach (var distinctValue in values.Distinct())
                            {
                                int frequency = values.Count(v => v == distinctValue);
                                double percentage = (frequency / (double)totalCount) * 100;


                                //MessageBox.Show(frequency.ToString());
                                // Store frequency and percentage in the FormattedValues dictionary
                                parameter.FormattedValues[groupValue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                parameter.FormattedValues[groupValue][$"Percentage_{distinctValue}"] = $"{percentage:F2}%";
                            }
                        }
                        else if (parameter.NominalOrScale == "Scale")
                        {
                            // Calculate scale statistics
                            double minValue = values.Min();
                            double maxValue = values.Max();
                            double meanValue = values.Average();
                            double stdDevValue = Math.Sqrt(values.Select(x => Math.Pow(x - meanValue, 2)).Sum() / (values.Count - 1));
                            double medianValue;
                            int middleIndex = values.Count / 2;
                            if (values.Count % 2 == 0)
                            {
                                // For even count of elements, take the average of the two middle values
                                double middleValue1 = values.OrderBy(x => x).ElementAt(middleIndex - 1);
                                double middleValue2 = values.OrderBy(x => x).ElementAt(middleIndex);
                                medianValue = (middleValue1 + middleValue2) / 2.0;
                            }
                            else
                            {
                                // For odd count of elements, directly take the middle value
                                medianValue = values.OrderBy(x => x).ElementAt(middleIndex);
                            }
                            double perc25th = CalculateLowerMedian(values);
                            double perc75th = CalculateUpperMedian(values);

                            // Format scale parameter values
                            string formattedMinMax = FormatMinMaxValue(minValue, maxValue);
                            string formattedMeanStd = FormatMeanStdValue(meanValue, stdDevValue);
                            string formattedMedian = FormatSingleValue(medianValue);
                            string formattedIQR = FormatMinMaxValue(perc25th, perc75th);

                            //MessageBox.Show(formattedMinMax);
                            //MessageBox.Show(formattedMinMax);

                            // Store scale statistics in the FormattedValues dictionary
                            parameter.FormattedValues[groupValue]["Min-Max"] = formattedMinMax;
                            parameter.FormattedValues[groupValue]["Mean ± StdDev"] = formattedMeanStd;
                            parameter.FormattedValues[groupValue]["Median"] = formattedMedian + " (" + formattedIQR + ")";
                            //parameter.FormattedValues[groupValue]["IQR"] = formattedIQR;
                        }
                    }
                }
            }
        }
        public double CalculateMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;
            int middle = n / 2;

            if (n % 2 == 0)
            {
                // Even number of elements, average the middle two
                return (values[middle - 1] + values[middle]) / 2.0;
            }
            else
            {
                // Odd number of elements, return the middle one
                return values[middle];
            }
        }

        public double CalculateLowerMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;
            int middle = n / 2;

            return CalculateMedian(values.GetRange(0, middle));
        }

        public double CalculateUpperMedian(List<double> values)
        {
            values.Sort();

            int n = values.Count;


            if (n % 2 == 0)
            {
                int middle = n / 2;
                // Even number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle));
            }
            else
            {
                int middle = (n + 1) / 2;
                // Odd number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle - 1));
            }
        }
        string FormatMinMaxValue(double minValue, double maxValue)
        {
            string minValueString = minValue.ToString("0.00");
            string maxValueString = maxValue.ToString("0.00");

            // Check if one value has three numbers (having three digits after the decimal point)
            bool isMinThreeNumbers = minValueString.Split('.')[0].Length == 3;
            bool isMaxThreeNumbers = maxValueString.Split('.')[0].Length == 3;

            // If one value is three numbers, round it to one decimal place
            if (isMinThreeNumbers)
            {
                minValueString = Math.Round(minValue, 1).ToString("0.0");
            }
            if (isMaxThreeNumbers)
            {
                maxValueString = Math.Round(maxValue, 1).ToString("0.0");
            }

            // Remove trailing zeroes if there are two zeroes after the decimal point
            if (minValueString.EndsWith(".00"))
            {
                minValueString = minValueString.Substring(0, minValueString.Length - 1);
            }
            if (maxValueString.EndsWith(".00"))
            {
                maxValueString = maxValueString.Substring(0, maxValueString.Length - 1);
            }

            return minValueString + " – " + maxValueString;
        }


        string FormatMeanStdValue(double meanValue, double stdDevValue)
        {
            string meanValueString = meanValue.ToString("0.00");
            string stdDevValueString = stdDevValue.ToString("0.00");

            // Check if the values are three numbers (having three digits after the decimal point)
            bool isThreeNumbers = meanValueString.Split('.')[0].Length == 3 && stdDevValueString.Split('.')[0].Length == 3;

            bool isMeanThreeNumbers = meanValueString.Split('.')[0].Length == 3;
            bool isStdThreeNumbers = stdDevValueString.Split('.')[0].Length == 3;
            // If it's three numbers, round to one decimal place
            if (isMeanThreeNumbers)
            {
                meanValueString = Math.Round(meanValue, 1).ToString("0.0");
                
            }
            if (isStdThreeNumbers)
            {
                stdDevValueString = Math.Round(stdDevValue, 1).ToString("0.0");
            }
            

            // Remove trailing zeroes if there are two zeroes after the decimal point
            if (meanValueString.EndsWith(".00"))
            {
                meanValueString = meanValueString.Substring(0, meanValueString.Length - 1);
            }
            if (stdDevValueString.EndsWith(".00"))
            {
                stdDevValueString = stdDevValueString.Substring(0, stdDevValueString.Length - 1);
            }

            return meanValueString + " ± " + stdDevValueString;
        }

        string FormatSingleValue(double value)
        {
            string valueString = value.ToString("0.00");

            // Check if the value is three numbers (having three digits after the decimal point)
            bool isThreeNumbers = valueString.Split('.')[0].Length == 3;

            // If it's three numbers, round to one decimal place
            if (isThreeNumbers)
            {
                valueString = Math.Round(value, 1).ToString("0.0");
            }

            // Remove trailing zeroes if there are two zeroes after the decimal point
            if (valueString.EndsWith(".00"))
            {
                valueString = valueString.Substring(0, valueString.Length - 1);
            }

            return valueString;
        }
        public void GetDataValues()
        {
            // Iterate over each table
            foreach (var table in ComparativeTables)
            {
                // MessageBox.Show("Table " + table.TableName);

                // Iterate over each parameter in the table
                foreach (var parameter in table.Parameters)
                {
                    // Iterate over each column
                    for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
                    {
                        // Get the name of the parameter from the header row
                        object cellValue = worksheet.Cells[0, col].Value;
                        string parameterName = cellValue?.ToString();

                        if (!string.IsNullOrEmpty(parameterName) && parameter.Name == parameterName)
                        {
                            //MessageBox.Show("Parameter " + parameterName);

                            // Iterate over each row in the column
                            for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
                            {
                                // Get the cell data
                                object cellData = worksheet.Cells[row, col].Value;
                                double cellValueDouble;

                                if (cellData.ToString() == ".")
                                {
                                    continue;
                                }
                                // Parse the cell data as double
                                if (cellData != null && double.TryParse(cellData.ToString(), out cellValueDouble))
                                {
                                    // Add the cell value to the parameter's values list
                                    parameter.ParameterValues.Add(cellValueDouble);
                                    //MessageBox.Show(cellValueDouble.ToString());
                                }
                            }
                        }
                    }
                }
            }
        }

        public void printcheck()
        {
            foreach (var ComparativeTable in ComparativeTables)
            {
                foreach (var parameter in ComparativeTable.Parameters)
                {
                    foreach (var values in parameter.GroupedParameterValues.Values)
                    {
                        foreach (var singlevalue in values)
                        {
                            MessageBox.Show(singlevalue.ToString());
                        }
                    }
                }
            }
        }
        WordClass wordObj = new WordClass();


        public bool DetermineHasNominal(ComparativeTable comparativeTable)
        {
            bool HasNominal = false;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.NominalOrScale == "Nominal")
                {
                    HasNominal = true;
                    break;
                }
                
            }
            return HasNominal;
        }
        public void DrawDescriptiveTable()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {

                

                IWSection section = wordObj.CreatePortraitSection();


                wordObj.AddDescriptiveTitle(section, ComparativeTables[tableindex].TableName);

                int Variablerows = CountRows(ComparativeTables[tableindex]);

                //MessageBox.Show(Variablerows.ToString());

                int WordTableColumns = 0;

                bool HasNominal = DetermineHasNominal(ComparativeTables[tableindex]); 

                if (!HasNominal)
                {
                    WordTableColumns = 4;
                }
                else if(HasNominal)
                {
                    WordTableColumns = 3;
                }


                int WordTableRows = 1 + Variablerows;



                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);


                wordObj.GeneralTableFormat(table);


               
                wordObj.ApplyGeneralDescritiveBorders(table, WordTableRows, WordTableColumns);

                wordObj.SetDescriptiveWidths(table, WordTableRows, HasNominal);


                wordObj.Add_GeneralHeaders_Descriptive(table, HasNominal);

                DescriptiveParameterBorders_text(table, WordTableRows, WordTableColumns, ComparativeTables[tableindex] , HasNominal);


                InsertData(table, ComparativeTables[tableindex], HasNominal);

                wordObj.FormatTable(table, 12);





            }
        }

        public void DescriptiveParameterBorders_text(IWTable table, int WordTableRows, int WordTableColumns, ComparativeTable comparativeTable , bool hasnominal)
        {
            int currentrow = 1;
            int count = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {



                if (parameter.NominalOrScale == "Nominal")
                {
                    count = parameter.ParameterValues.Distinct().Count() + 1;
                   
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    count = 4;
                    
                }

                if(hasnominal)
                {
                    if (count > 0)
                    {

                        if (parameter.NominalOrScale == "Scale")
                        {
                            for (int j = 0; j < 4; j++)
                            {
                                table.ApplyHorizontalMerge(currentrow + j, 1, 2);
                                
                            }

                            wordObj.AddPara_NoCenter(table, currentrow, 0, parameter.Name);

                            if (parameter.hasLowerN)
                            {
                                int InsertLowerN = 1;
                                foreach (double group in parameter.GroupedParameterValues.Keys)
                                {
                                    int parameterGroupCount = parameter.GroupedParameterValues[group].Count;
                                    string InsertedN = "(n = " + parameterGroupCount + ")";
                                    wordObj.AddParaCombined(table, currentrow, InsertLowerN, InsertedN, true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);
                                    InsertLowerN = InsertLowerN + 2;
                                }
                            }

                            wordObj.Addpara_NoCenterNoBOLD(table, currentrow + 1, 0, "Min – Max.");
                            wordObj.LeftIntendBeforeText(table, currentrow + 1, 0, 14.17f);
                            

                            wordObj.Addpara_NoCenterNoBOLD(table, currentrow + 2, 0, "Mean ± SD.");
                            wordObj.LeftIntendBeforeText(table, currentrow + 2, 0, 14.17f);

                            wordObj.Addpara_NoCenterNoBOLD(table, currentrow + 3, 0, "Median (IQR)");
                            wordObj.LeftIntendBeforeText(table, currentrow + 3, 0, 14.17f);
                        }
                        else if (parameter.NominalOrScale == "Nominal")
                        {
                            wordObj.AddPara_NoCenter(table, currentrow, 0, parameter.Name);


                            if (parameter.hasLowerN)
                            {
                                int InsertLowerN = 1;
                                foreach (double group in parameter.GroupedParameterValues.Keys)
                                {
                                    table.ApplyHorizontalMerge(currentrow , InsertLowerN, InsertLowerN + 1);
                                    int parameterGroupCount = parameter.GroupedParameterValues[group].Count;
                                    string InsertedN = "(n = " + parameterGroupCount + ")";
                                    wordObj.AddParaCombined(table, currentrow , InsertLowerN, InsertedN, true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);
                                    InsertLowerN = InsertLowerN + 2;
                                }
                            }

                            int row = 1;


                            foreach (var label in parameter.DIC_LablesIfNomainal.Values)
                            {
                                //MessageBox.Show(label.ToString());
                                wordObj.Addpara_NoCenterNoBOLD(table, currentrow + row, 0, label.ToString());
                                wordObj.LeftIntendBeforeText(table, currentrow + row, 0, 14.17f);
                                row++;
                            }


                            //var sortedValues = parameter.ParameterValues.Distinct().OrderBy(value => value);
                            //foreach (var value in sortedValues)
                            //{
                            //    wordObj.Addpara_NoCenterNoBOLD(table, currentrow + row, 0, value.ToString());
                            //    wordObj.LeftIntendBeforeText(table, currentrow + row, 0, 14.17f);
                            //    row++;
                            //}

                        }

                    }



                    if (currentrow + count < WordTableRows - 1)
                    {

                        // Insert bottom border for each parameter
                        if (count > 0)
                        {
                            for (int i = 0; i < WordTableColumns; i++)
                            {
                                table.Rows[currentrow + count].Cells[i].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                                table.Rows[currentrow + count].Cells[i].CellFormat.Borders.Top.LineWidth = 0.5f;
                            }

                            currentrow = currentrow + count;

                        }


                    }
                }


                else if(!hasnominal)
                {
                    wordObj.AddPara_NoCenter(table, currentrow, 0, parameter.Name);

                    if(currentrow < WordTableRows - 1) 
                    {
                        for (int i = 0; i < WordTableColumns; i++)
                        {
                            table.Rows[currentrow].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[currentrow].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }

                        currentrow++;
                    }
                }
                
            }
        }


        

        public void InsertData(IWTable table , ComparativeTable comparativeTable ,bool hasnominal)
        {
            if(hasnominal)
            {
                int startingRow = 2;
                foreach (Parameter parameter in comparativeTable.Parameters)
                {
                    
                    int column = 1;
                    int count = 0;

                    // Skip if the parameter is a group or not nominal
                    if (parameter.IsGroup || parameter.FormattedValues.Count == 0)
                    {
                        continue;
                    }

                    // Determine the count based on parameter type
                    if (parameter.NominalOrScale == "Nominal")
                    {
                        count = parameter.ParameterValues.Distinct().Count() + 1;


                        // Insert frequency and percentage for nominal parameters

                        foreach (var kvp in parameter.FormattedValues)
                        {
                            double groupValue = kvp.Key;
                            Dictionary<string, string> formattedValues = kvp.Value;
                            // Insert frequencies for the current group
                            foreach (var distinctValue in parameter.DIC_LablesIfNomainal.Keys)
                            {

                                string frequency = formattedValues.ContainsKey($"Frequency_{distinctValue}") ? formattedValues[$"Frequency_{distinctValue}"] : "0";
                                wordObj.Addpara_CenterNoBOLD(table, startingRow, column, frequency);
                                MessageBox.Show(frequency);
                                column++;

                                string percentage = formattedValues.ContainsKey($"Percentage_{distinctValue}") ? formattedValues[$"Percentage_{distinctValue}"] : "0.0";
                                percentage = percentage.Replace("%", "");


                                if (percentage.EndsWith("0"))
                                {
                                    //percentage = percentage.Substring(0, percentage.Length - 1);
                                    
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage);
                                }
                                else
                                {
                                    MessageBox.Show(percentage);
                                    //double percentageDouble = Math.Round(double.Parse(percentage), 1);
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage.ToString());
                                }
                                startingRow++;
                                column--;
                            }

                        }

                        startingRow++;



                    }
                    else if (parameter.NominalOrScale == "Scale")
                    {
                        count = 4;

                        // Insert scale statistics for scale parameters
                        foreach (var kvp in parameter.FormattedValues)
                        {
                            double groupValue = kvp.Key;
                            Dictionary<string, string> scaleStats = kvp.Value;

                            foreach (var stat in scaleStats)
                            {
                                wordObj.Addpara_CenterNoBOLD(table, startingRow, column, stat.Value);
                                startingRow++; // Move to the next row
                            }
                        }
                        startingRow++;

                    }

                }
            }
            else if(!hasnominal)
            {
                int startingRow = 1;
                
                foreach (var parameter in comparativeTable.Parameters)
                {
                    int column = 1;
                    foreach (var kvp in parameter.FormattedValues)
                    {
                        double groupValue = kvp.Key;
                        Dictionary<string, string> scaleStats = kvp.Value;

                        foreach (var stat in scaleStats)
                        {
                            wordObj.Addpara_CenterNoBOLD(table, startingRow, column, stat.Value);
                            column++;
                        }
                    }

                    startingRow++;
                }
            }

           
        }
        public void ClearPara()
        {
            foreach (var item in ComparativeTables)
            {
                foreach (var parameter in item.Parameters)
                {
                    parameter.ParameterValues.Clear();
                    parameter.FormattedValues.Clear();
                    parameter.GroupedParameterValues.Clear();
                    parameter.EachGroupCount.Clear();
                    parameter.FPairwise.Clear();

                }
                item.TestsDone.Clear();
            }
        }
        private void btn_Done_Click(object sender, EventArgs e)
        {
            DescriptiveBasic();

            wordObj.InitWord();

            DrawDescriptiveTable();
            ClearPara();
            wordObj.SaveWord();

        }
        public static ComparativeTable DoneSortedComparative { get; set; }
        private void btn_SortTable_Click(object sender, EventArgs e)
        {
            if (cmb_TableNames.SelectedIndex != -1)
            {
                string selectedTableName = cmb_TableNames.SelectedItem.ToString();
                var selectedTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                Sort_Parameters_Frm sort_Parameters_Frm = new Sort_Parameters_Frm();
                sort_Parameters_Frm.SortcomparativeTable = selectedTable;
                sort_Parameters_Frm.TableType = "Descriptive";
                sort_Parameters_Frm.Show();
            }
        }
        public bool CurrentDescriptive = true;
        private void btn_SwicthDescrToComp_Click(object sender, EventArgs e)
        {

            if(CurrentDescriptive)
            {
                lbl_Nominal.Text = "Rows";
                lbl_Scale.Text = "Columns";
                lbl_CurrentAction.Text = "Current Action : Comparative";
                btn_SwicthDescrToComp.Text = "Descriptive";
                CurrentDescriptive = false;
            }
            else if(!CurrentDescriptive)
            {
                lbl_Nominal.Text = "Nominal";
                lbl_Scale.Text = "Scale";
                lbl_CurrentAction.Text = "Current Action : Descriptive";
                btn_SwicthDescrToComp.Text = "Comparative";
                CurrentDescriptive = true;
            }

            




        }

        private void lbl_CurrentAction_Click(object sender, EventArgs e)
        {
            
        }
    }
}
