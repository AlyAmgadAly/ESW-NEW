using Aspose.Cells;
using ExcelScore.Classes;
using MathNet.Numerics;
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
using System.Windows.Markup;
using static System.Windows.Forms.AxHost;

namespace ExcelScore.Forms
{
    public partial class Letters : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public DataGridView Dgv { get; set; }

        public List<ComparativeClass> ComparativeClassesList = new List<ComparativeClass>();
        public List<ComparativeTable> ComparativeTables = new List<ComparativeTable>();

        static ExcelFunctions excelFunctionsobj = new ExcelFunctions();

        List<string> ColumnExcelheaders = excelFunctionsobj.ReadHeaderColumnsExcel();

        Worksheet worksheet = excelFunctionsobj.GetWorksheet();

        Workbook workbook = excelFunctionsobj.GetWorkbook();

        Worksheet Sheet2;
        public Letters()
        {
            InitializeComponent();
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
        private void Letters_Load(object sender, EventArgs e)
        {
            AddHeadersToParameter();
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

        private void pic_AllParaToSelected_Click(object sender, EventArgs e)
        {
            foreach (object selectedItem in list_AllParameters.SelectedItems)
            {
                list_selectedParameters.Items.Add(selectedItem.ToString());
            }
        }

        private void pic_RemoveSelectedPara_Click(object sender, EventArgs e)
        {
            if (list_selectedParameters.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_selectedParameters.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_selectedParameters.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }
        public void AddTableClass(string tableName)
        {
            var comparativeTable = new ComparativeTable
            {
                TableName = tableName,
                Parameters = new List<Parameter>()
            };

            // Add parameters from list_Groups
            foreach (var item in list_Groups.Items)
            {
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    IsGroup = true,
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }

            comparativeTable.LetterType = cmb_letterType.SelectedItem.ToString();





            // Add parameters from list_NormalScale
            foreach (var item in list_selectedParameters.Items)
            {
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    NominalOrScale = "Scale",
                    NormalOrAbnormal = "Normal",
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }

            ComparativeTables.Add(comparativeTable);


        }
        public string AddTableUI()
        {
            string TableName = null;
            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&
                list_Groups.Items.Count > 0 &&
                list_selectedParameters.Items.Count > 0 && cmb_letterType.SelectedIndex != -1)
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
                    TableName = txt_TableName.Text;
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

            return TableName;
        }
        private void pic_addTable_Click(object sender, EventArgs e)
        {
            string TableName = AddTableUI();
            ComparativeBasic(TableName);
        }

        private void pic_removeGroup_Click(object sender, EventArgs e)
        {
            if (list_Groups.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_Groups.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_Groups.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_AllParaToGroups_Click(object sender, EventArgs e)
        {
            foreach (object selectedItem in list_AllParameters.SelectedItems)
            {
                list_Groups.Items.Add(selectedItem.ToString());
            }
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
        public void GetDataValues(string TableName)
        {
            // Iterate over each table
            foreach (var table in ComparativeTables)
            {
                // MessageBox.Show("Table " + table.TableName);
                if (TableName == table.TableName)
                {
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

                                    if (cellData == null)
                                    {
                                        MessageBox.Show("Data is empty at Parameter : " + parameter.Name);
                                        break;
                                    }
                                    else if (cellData.ToString() == ".")
                                    {
                                        continue;
                                    }
                                    object SelectData = worksheet.Cells[row, SelectParameterColIndex].Value;
                                    //MessageBox.Show(SelectParameterColIndex.ToString());

                                    double SelectDataDouble;


                                    if (list_Select.Items.Count == 0)
                                    {
                                        if (cellData != null && double.TryParse(cellData.ToString(), out cellValueDouble))
                                        {
                                            //MessageBox.Show(cellValueDouble.ToString());
                                            // Add the cell value to the parameter's values list
                                            parameter.ParameterValues.Add(cellValueDouble);
                                            //MessageBox.Show(cellValueDouble.ToString());
                                        }
                                    }
                                    else
                                    {
                                        if (SelectData != null && double.TryParse(SelectData.ToString(), out SelectDataDouble))
                                        {
                                            if (SelectedParameterValues_LetterFrm.Contains(SelectDataDouble))
                                            {
                                                if (cellData != null && double.TryParse(cellData.ToString(), out cellValueDouble))
                                                {
                                                    //MessageBox.Show(cellValueDouble.ToString());
                                                    // Add the cell value to the parameter's values list
                                                    //MessageBox.Show("Select value " +SelectData.ToString());
                                                    //MessageBox.Show("Cell value "+cellValueDouble.ToString());

                                                    parameter.ParameterValues.Add(cellValueDouble);
                                                    //MessageBox.Show(cellValueDouble.ToString());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                // Iterate over each parameter in the table
                
            }
        }

        pythonStat pythonStat = new pythonStat();


        public int newFindGroupColumnIndex(ComparativeTable comparativeTable)
        {

            for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
            {
                object cellValue = worksheet.Cells[0, col].Value;
                string parameterName = cellValue?.ToString();

                if (!string.IsNullOrEmpty(parameterName))
                {
                    foreach (var parameter in comparativeTable.Parameters)
                    {
                        if (parameter.Name == parameterName && parameter.IsGroup)
                        {
                            comparativeTable.GroupColumnIndex = col;
                            //MessageBox.Show(col.ToString());// Set the group column index for this table
                            return col;
                        }
                    }
                }
            }

            return -1;

        }
        public void ComparativeBasic(string TableName)
        {
            GetDataValues(TableName);
            
            foreach (ComparativeTable table in ComparativeTables)
            {
                if(TableName == table.TableName)
                {
                    int groupColumnIndex = newFindGroupColumnIndex(table);
                    //MessageBox.Show(groupColumnIndex.ToString());
                    if (groupColumnIndex == -1)
                    {
                        continue;
                    }
                    // Iterate through each parameter of the table
                    foreach (Parameter parameter in table.Parameters)
                    {
                        if (parameter.IsGroup)
                        {
                            // Skip group parameters
                            continue;
                        }

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
                                // Get the group value for the current row
                                object groupCellValue = worksheet.Cells[row, groupColumnIndex].Value;
                                double groupValue;

                                if (groupCellValue != null && double.TryParse(groupCellValue.ToString(), out groupValue))
                                {
                                    // Get the parameter value for the current row
                                    object parameterCellValue = worksheet.Cells[row, parameterColumnIndex].Value;
                                    double parameterValue;


                                    if (parameterCellValue == null)
                                    {

                                        break;
                                    }
                                    if (parameterCellValue.ToString() == ".")
                                    {
                                        parameter.hasLowerN = true;
                                        continue;
                                    }

                                    if (list_Select.Items.Count == 0)
                                    {
                                        if (parameterCellValue != null && double.TryParse(parameterCellValue.ToString(), out parameterValue))
                                        {
                                            // Add the parameter value to the corresponding group
                                            if (!parameter.GroupedParameterValues.ContainsKey(groupValue))
                                            {
                                                parameter.GroupedParameterValues[groupValue] = new List<double>();
                                            }
                                            //MessageBox.Show(groupValue.ToString());
                                            //MessageBox.Show(parameterValue.ToString());
                                            parameter.GroupedParameterValues[groupValue].Add(parameterValue);
                                            //MessageBox.Show(groupValue.ToString());
                                            //MessageBox.Show(parameterValue.ToString());
                                        }
                                    }
                                    else
                                    {
                                        object SelectData = worksheet.Cells[row, SelectParameterColIndex].Value;
                                        //MessageBox.Show(SelectParameterColIndex.ToString());

                                        double SelectDataDouble;
                                        if (SelectData != null && double.TryParse(SelectData.ToString(), out SelectDataDouble))
                                        {
                                            if (SelectedParameterValues_LetterFrm.Contains(SelectDataDouble))
                                            {
                                                if (parameterCellValue != null && double.TryParse(parameterCellValue.ToString(), out parameterValue))
                                                {
                                                    // Add the parameter value to the corresponding group
                                                    if (!parameter.GroupedParameterValues.ContainsKey(groupValue))
                                                    {
                                                        parameter.GroupedParameterValues[groupValue] = new List<double>();
                                                    }

                                                    parameter.GroupedParameterValues[groupValue].Add(parameterValue);
                                                    //MessageBox.Show(groupValue.ToString());
                                                    //MessageBox.Show(parameterValue.ToString());
                                                }
                                            }
                                        }

                                    }

                                }
                            }

                        }
                    }
                }
            }
            FormatParameters(TableName);
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
        public void FormatParameters(string TableName)
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                if(TableName == table.TableName)
                {
                    foreach (Parameter parameter in table.Parameters)
                    {
                        //MessageBox.Show(parameter.Name);
                        if (parameter.IsGroup || parameter.GroupedParameterValues.Count == 0)
                        {
                            continue;
                        }

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
                //MessageBox.Show(table.TableName);
                
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


            bool isMeanThreeNumbers = meanValueString.Split('.')[0].Length == 3;
            bool isStdThreeNumbers = stdDevValueString.Split('.')[0].Length == 3;

            // Check if the values are three numbers (having three digits after the decimal point)
            bool isThreeNumbers = meanValueString.Split('.')[0].Length == 3 && stdDevValueString.Split('.')[0].Length == 3;

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

        WordClass wordObj = new WordClass();
        public int CountGroupValues(ComparativeTable table)
        {
            int groupCount = 0;

            // Find the group parameter
            var groupParameter = table.Parameters.FirstOrDefault(p => p.IsGroup);
            if (groupParameter == null)
            {
                // Handle case where no group parameter is found
                return groupCount;
            }

            // Count the distinct group values
            
            groupCount = groupParameter.ParameterValues.Distinct().Count();
            //MessageBox.Show(groupCount.ToString());

            return groupCount;
        }

        public int GetParametersCount(ComparativeTable comparativeTable)
        {
            int count = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }
                else
                {
                    count++;
                }
            }

            return count;
        }

        public List<string> Groupnames;
        public void GetGroupNames(Worksheet Sheet2)
        {
            for (int row = 1; row < Sheet2.Cells.MaxDataRow + 1; row++)
            {
                object cellvalue = Sheet2.Cells[row, 0].Value;
                Groupnames.Add(cellvalue.ToString());

            }
        }

        public void AddGroupAndVariableNames(IWTable table, int WordTableRows, int WordTableColumns, ComparativeTable comparativeTable)
        {
            int GroupRow = 2;
            for (int i = 0; i < Groupnames.Count; i++)
            {
                wordObj.AddPara_NoCenter(table, GroupRow, 0, Groupnames[i]);
                GroupRow++;
            }

            wordObj.AddParaCombined(table, WordTableRows - 2, 0, "F", true, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
            wordObj.AddPara_Center(table, WordTableRows - 1, 0, "p");


            int VariableColumn = 1;
            foreach (var parameter in comparativeTable.Parameters)
            {

                if (parameter.IsGroup)
                {
                    continue;
                }
                wordObj.AddPara_Center(table, 0, VariableColumn, parameter.Name);
                VariableColumn++;

            }

            for (int i = 1; i < WordTableColumns; i++)
            {
                wordObj.AddPara_Center(table, 1, i, "Mean ± SD.");
            }


        }

        static char GetNextsmallLetter(char currentLetter)
        {
            int asciiValue = (int)currentLetter;
            asciiValue = (asciiValue - 'a' + 1) % 26 + 'a'; // Incrementing ASCII value modulo 26 to wrap around
            return (char)asciiValue;
        }
        public void PrintMessage(object message)
        {
            MessageBox.Show(message.ToString());
        }
        public void HardLetters(IWTable table, Parameter parameter, int InsertColumn)
        {
            int numberofgroups = parameter.GroupedParameterValues.Keys.Distinct().Count();

            
            int pairwiseTotalcount = parameter.FPairwise.Count();
            int pairwisectr = 0;
            
            int Currentgroups = 0;

            Dictionary<int, List<string>> Row_Letter = new Dictionary<int, List<string>>();




            Dictionary<int, double> Parameter_Groups_means = new Dictionary<int, double>();


            //Dictionary<(int, int), string> LabelPairwise = new Dictionary<(int, int), string>();


            // saving Label : pairwise foreach parameter
            while (pairwiseTotalcount > 0)
            {
                for (int j = 1; j <= numberofgroups; j++)
                {
                    for (int i = 2; i <= numberofgroups; i++)
                    {
                        if (i > j)
                        {
                            parameter.LabelPairwise[(j, i)] = parameter.FPairwise[pairwisectr];

                            pairwisectr++;
                            pairwiseTotalcount--;
                        }

                    }



                }

                numberofgroups--;
                Currentgroups++;

            }


            foreach (var kvp in parameter.FormattedValues)
            {
                double groupValue = kvp.Key;
                Dictionary<string, string> scaleStats = kvp.Value;


                foreach (var stat in scaleStats)
                {
                    if (stat.Key == "Mean ± StdDev")
                    {

                        // these get mean which we will append letters to 
                        string MeanStdstr = stat.Value;
                        string[] meanStdArr = MeanStdstr.Split(' ');

                        string mean = meanStdArr[0];
                        Parameter_Groups_means[(int)groupValue] = double.Parse(mean);

                    }
                }

            }

            //sort descending

            var sortedDict = Parameter_Groups_means.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);

            List<int> SortedGroups = new List<int>();

            foreach (var kvp in sortedDict)
            {
                SortedGroups.Add(kvp.Key);
            }

            Dictionary<int, string> GroupLetter = new Dictionary<int, string>();

            int HighestGroup = SortedGroups[0];
            GroupLetter[HighestGroup] = "a";

            List<(int, int)> mycomparisons = new List<(int, int)>();

            int Groupctr = 1;
            int GroupTotalCount = parameter.GroupedParameterValues.Keys.Distinct().Count();
            //MessageBox.Show(GroupTotalCount.ToString());
            for (int i = 0; i < parameter.LabelPairwise.Count; i++)
            {
                int Group = SortedGroups[Groupctr];

                for (int j = Groupctr - 1; j >= 0; j--)
                {
                    if (Groupctr > j)
                    {
                        //string value = parameter.LabelPairwise[(SortedGroups[Groupctr - 1], Group)];

                        int smallerNumber = Math.Min(SortedGroups[j], SortedGroups[Groupctr]);
                        int largerNumber = Math.Max(SortedGroups[j], SortedGroups[Groupctr]);
                        mycomparisons.Add((smallerNumber, largerNumber));
                        i++;
                    }


                }
                Groupctr++;
            }

            
            List<int> BlockedLowerGrps = new List<int>();
            List<int> NotSigGroups = new List<int>();
            int lowergroupCtr = 1;
            bool IsSig = false;
            while (lowergroupCtr < SortedGroups.Count)
            {
                bool flagIsSig = false;
                int higherGroupCtr = lowergroupCtr - 1;

                if (!GroupLetter.ContainsKey(SortedGroups[lowergroupCtr]))
                {
                    // If lowerGroup does not exist, add it with an empty string
                    GroupLetter[SortedGroups[lowergroupCtr]] = "";
                }

                

                while (higherGroupCtr != -1)
                {
                    if (!GroupLetter.ContainsKey(SortedGroups[higherGroupCtr]))
                    {
                        // If higherGroup does not exist, add it with an empty string
                        GroupLetter[SortedGroups[higherGroupCtr]] = "";
                    }
                    int smallerNumber = Math.Min(SortedGroups[higherGroupCtr], SortedGroups[lowergroupCtr]);
                    int largerNumber = Math.Max(SortedGroups[higherGroupCtr], SortedGroups[lowergroupCtr]);
                    string pvaluestr = parameter.LabelPairwise[(smallerNumber, largerNumber)];
                    

                    if (pvaluestr == "<0.001")
                    {
                        IsSig = true;
                    }
                    else
                    {
                        double pdoublevalue = double.Parse(pvaluestr);

                        if (pdoublevalue >= 0.05)
                        {
                            IsSig = false;
                        }
                        else if (pdoublevalue < 0.05)
                        {
                            IsSig = true;
                        }
                    }

                    if(IsSig)
                    {
                        flagIsSig = true;
                        char lastcharacter = GroupLetter[SortedGroups[higherGroupCtr]].LastOrDefault();
                        if (!BlockedLowerGrps.Contains(SortedGroups[lowergroupCtr]))
                        {
                            GroupLetter[SortedGroups[lowergroupCtr]] += GetNextsmallLetter(lastcharacter);
                            BlockedLowerGrps.Add(SortedGroups[lowergroupCtr]);
                            if (NotSigGroups != null)
                            {
                                foreach (int NotSigG in NotSigGroups)
                                {
                                    GroupLetter[NotSigG] += GetNextsmallLetter(lastcharacter);
                                }
                            }
                        }
                        
                    }
                    else if (!IsSig) 
                    {
                        NotSigGroups.Add(SortedGroups[higherGroupCtr]);
                    }
                    


                    --higherGroupCtr;

                    
                }
                
                if (!flagIsSig)
                {
                    char notsigchar = GroupLetter[NotSigGroups[0]].LastOrDefault();
                    GroupLetter[SortedGroups[lowergroupCtr]] += notsigchar;
                }
                NotSigGroups.Clear();
                lowergroupCtr++;

            }








            List<int> keysToModify = new List<int>(GroupLetter.Keys);

            // Iterate over the list of keys and modify the dictionary
            foreach (int key in keysToModify)
            {
                string input = GroupLetter[key];

                // Use a HashSet to remove duplicates while preserving order
                HashSet<char> uniqueChars = new HashSet<char>();
                foreach (char c in input)
                {
                    uniqueChars.Add(c);
                }

                // Construct the string of unique characters
                string uniqueString = string.Join("", uniqueChars);

                // Update the dictionary with the unique string
                GroupLetter[key] = uniqueString;
            }

            //string message = "Dictionary Contents:\n";
            //foreach (var pair in GroupLetter)
            //{
            //    message += $"Key: {pair.Key}, Value: {pair.Value}\n";
            //}

            //// Show the message box
            //MessageBox.Show(message, "Dictionary Contents", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //string message2 = "LabelPairwise dictionary contents:\n\n";
            //foreach (var kvp in parameter.LabelPairwise)
            //{
            //    message2 += $"Key: ({kvp.Key.Item1}, {kvp.Key.Item2}), Value: {kvp.Value}\n";
            //}

            //MessageBox.Show(message2, "LabelPairwise Dictionary", MessageBoxButtons.OK, MessageBoxIcon.Information);


            foreach (var kvp in GroupLetter)
            {
                int row = (kvp.Key)+1;
                string myletters = string.Join("", kvp.Value);
                string Data = table[row, InsertColumn].Paragraphs[0].Text;
                string[] MeanStd = Data.Split(' ');

                WParagraph datapara = table[row, InsertColumn].Paragraphs[0];
                table[row, InsertColumn].Paragraphs[0].Text = "";

                datapara.AppendText(MeanStd[0]);

                WTextRange asteriskt = (WTextRange)datapara.AppendText(myletters);
                asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;


                datapara.AppendText(" ");
                datapara.AppendText(MeanStd[1]);
                datapara.AppendText(" ");
                datapara.AppendText(MeanStd[2]);

            }

            



        }

        public List<string> EasyLetters(IWTable table ,Parameter parameter , int InsertColumn)
        {
            int numberofgroups = parameter.GroupedParameterValues.Keys.Distinct().Count();
            //MessageBox.Show(numberofgroups.ToString());
            List<string> letters = new List<string>();  
            char letter = 'a';
            bool issig = false;
            int pairwiseTotalcount = parameter.FPairwise.Count();
            int pairwisectr = 0;
            int startrow = 3;
            int Currentgroups = 0;


            Dictionary<int , List<string>> Row_Letter = new Dictionary<int , List<string>>();

            while (pairwiseTotalcount > 0)
            {
                startrow = 3 + Currentgroups;
                for (int j = 0; j < numberofgroups-1; j++)
                {
                    
                    if (parameter.FPairwise[pairwisectr] == "<0.001")
                    {
                        issig = true;
                        
                    }
                    else
                    {
                        double p_pairwise = double.Parse(parameter.FPairwise[pairwisectr]);
                        if (p_pairwise >= 0.05)
                        {
                            
                            issig = false;

                        }
                        else if (p_pairwise < 0.05)
                        {
                            
                            issig = true;

                        }
                        
                    }


                    if (issig)
                    {
                        letters.Add(letter.ToString());

                        if (!Row_Letter.ContainsKey(startrow))
                        {
                            Row_Letter[startrow] = new List<string>();
                        }
                        Row_Letter[startrow].Add(letter.ToString());

                    }
                    startrow++;
                    pairwisectr++;
                    pairwiseTotalcount--;
                    

                }
                numberofgroups--;
                Currentgroups++;
                letter = GetNextsmallLetter(letter);

            }

            foreach (var kvp in Row_Letter)
            {
                int row = kvp.Key;
                string myletters = string.Join("", kvp.Value);
                string Data = table[row, InsertColumn].Paragraphs[0].Text;
                string[] MeanStd = Data.Split(' ');

                WParagraph datapara = table[row, InsertColumn].Paragraphs[0];
                table[row, InsertColumn].Paragraphs[0].Text = "";
                
                datapara.AppendText(MeanStd[0]);

                WTextRange asteriskt = (WTextRange)datapara.AppendText(myletters);
                asteriskt.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;


                datapara.AppendText(" ");
                datapara.AppendText(MeanStd[1]);
                datapara.AppendText(" ");
                datapara.AppendText(MeanStd[2]);

            }





            return letters;
            
            

        }


        public void InsertEasyLettersWord(IWTable table , List<string> letters , int insertColumn)
        {

        }

       

        public void InsertMeanStd(IWTable table , ComparativeTable comparativeTable)
        {
            int startColumn = 1;
            foreach (var parameter in comparativeTable.Parameters)
            {
                List<string> Letters = new List<string>();
                int groupcount = parameter.GroupedParameterValues.Keys.Distinct().Count();  
                
                
                if (parameter.IsGroup) 
                {
                    continue;
                }

                int startRow = 2;
                
                if (parameter.NominalOrScale == "Scale")
                {
                    
                    // Insert scale statistics for scale parameters
                    foreach (var kvp in parameter.FormattedValues)
                    {
                        double groupValue = kvp.Key;
                        Dictionary<string, string> scaleStats = kvp.Value;


                        foreach (var stat in scaleStats)
                        {
                            if(stat.Key == "Mean ± StdDev")
                            {

                                // these get mean which we will append letters to 
                                //string first = stat.Value;
                                //string[] tests = first.Split(' ');
                                
                                wordObj.Addpara_CenterNoBOLD(table, startRow, startColumn, stat.Value);
                            }
                        }
                        startRow++;
                    }

                    AnovaTestResult LetteranovaTestResult = new AnovaTestResult();
                    LetteranovaTestResult = pythonStat.ANOVAWithTukeyHSDNewDynamic(parameter);
                    //anovaTestResult = pythonStat.ANOVAWithMultipleComparisons(parameter);



                    AnovaTestResult StatTestResult = new AnovaTestResult();
                    StatTestResult = manualTests.Fanova(parameter);


                    string[] values = { StatTestResult.TestValue, StatTestResult.PValue };

                    if (comparativeTable.LetterType == "Easy")
                    {
                        EasyLetters(table, parameter, startColumn);
                    }
                    else if (comparativeTable.LetterType == "Hard")
                    {
                        HardLetters(table, parameter, startColumn);
                    }
                    
                    

                    wordObj.InsertTest_P_Letters(table, startRow, startColumn, values);

                }

                startColumn++;
            }
            
        }


        ManualTests manualTests = new ManualTests();
        public void DrawLettersTable()
        {

            Sheet2 = workbook.Worksheets[1];


            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                Groupnames = new List<string>();

                IWSection section = wordObj.CreatePortraitSection();


                int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                //MessageBox.Show(numberofgroups.ToString());

                wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                int Variablerows = numberofgroups;

                int WordTableColumns = GetParametersCount(ComparativeTables[tableindex]) + 1;

                int WordTableRows = 4 + Variablerows;

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                wordObj.GeneralLetterTableFormat(table);

               
                wordObj.ApplyGeneralLettersBorders(table, WordTableRows, WordTableColumns, numberofgroups);

                
                wordObj.SetLetterWidths(table, WordTableRows, WordTableColumns);



                GetGroupNames(Sheet2);

                AddGroupAndVariableNames(table , WordTableRows , WordTableColumns , ComparativeTables[tableindex]);


                InsertMeanStd(table, ComparativeTables[tableindex]);




                wordObj.LeftAndRightCellMarginCustom(table , wordObj.SetColumnWidthInCentimeters(0.09f) , wordObj.SetColumnWidthInCentimeters(0.09f));

                wordObj.FormatTable(table, 12);
            }
        }

        public void ClearPara()
        {
            foreach (var item in ComparativeTables)
            {
                foreach (var parameter in item.Parameters)
                {
                    //parameter.ParameterValues.Clear();
                    //parameter.FormattedValues.Clear();
                    //parameter.GroupedParameterValues.Clear();
                    //parameter.EachGroupCount.Clear();
                    parameter.FPairwise.Clear();
                    parameter.LabelPairwise.Clear();
                }
                item.TestsDone.Clear();
            }
        }
        private void btn_Done_Click(object sender, EventArgs e)
        {
            pythonStat.InitPython();


            wordObj.InitWord();

            DrawLettersTable();

            wordObj.SaveWord();

            ClearPara();

        }

        private void btn_ChooseTable_Click(object sender, EventArgs e)
        {

        }
        public int GetSelectParameterCol()
        {
            string SelectParamter = list_Select.Items[0].ToString();
            for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
            {
                int columnindex = -1;
                // Get the name of the parameter from the header row
                object cellValue = worksheet.Cells[0, col].Value;
                string parameterName = cellValue?.ToString();

                if (!string.IsNullOrEmpty(parameterName) && SelectParamter == parameterName)
                {

                    //MessageBox.Show("Parameter " + parameterName);
                    columnindex = col;
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
                            AllParameter_Select_Values.Add(cellValueDouble);
                            //MessageBox.Show(cellValueDouble.ToString());
                        }
                    }
                    //MessageBox.Show(columnindex.ToString());
                    return columnindex;

                }

            }
            return -1;


        }

        public int SelectParameterColIndex { get; set; } = new int();

        public List<double> AllParameter_Select_Values { get; set; } = new List<double>();
        private void pic_AllParaToSelect_Click(object sender, EventArgs e)
        {
            
            foreach (object selectedItem in list_AllParameters.SelectedItems)
            {
                list_Select.Items.Add(selectedItem.ToString());

            }


            SelectParameterColIndex = GetSelectParameterCol();
        }
        public List<double> SelectedParameterValues_LetterFrm { get; set; } = new List<double>();
        private void pic_groups_select_Click(object sender, EventArgs e)
        {
            if (list_Select.Items.Count == 1)
            {



                Select_Groups select_Groupsobj = new Select_Groups(this);
                select_Groupsobj.SelectParaName_SelectGrFrm = list_Select.Items[0].ToString();
                select_Groupsobj.SelectParaValues_SelectGrFrm = AllParameter_Select_Values;


                select_Groupsobj.ShowDialog();
            }
        }

        private void pic_RemoveSelectPara_Click(object sender, EventArgs e)
        {
            if (list_Select.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_Select.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_Select.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }
    }
}
