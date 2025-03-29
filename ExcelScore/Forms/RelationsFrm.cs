using Aspose.Cells;
using ExcelScore.Classes;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ExcelScore.Forms
{
    public partial class RelationsFrm : Form
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

        Worksheet Sheet2 = excelFunctionsobj.GetSheet2();
        public RelationsFrm()
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
                data_allPara.Rows.Add(comparativeClass.ParameterName);
            }

            InsertPicture_Type();
            foreach (DataGridViewColumn column in data_allPara.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        public void InsertPicture_Type()
        {

            for (int column = 0; column <= Sheet2.Cells.MaxDataColumn; column++)
            {
                string Type = "";
                string ParaName = "";
                foreach (string columnName in ColumnExcelheaders)
                {

                    ParaName = columnName;
                    if (columnName == Sheet2.Cells[0, column].Value.ToString())
                    {

                        if (Sheet2.Cells[1, column].Value != null)
                        {

                            Type = Sheet2.Cells[1, column].Value.ToString();

                        }
                    }

                    foreach (DataGridViewRow row in data_allPara.Rows)
                    {
                        if (row.Cells[0].Value != null && row.Cells[0].Value.ToString() == ParaName)
                        {
                            // Insert the image based on the type
                            if (Type == "Nominal")
                            {
                                row.Cells["ColMeasure"].Value = Resource.Final_Nominal_Color; // Replace NominalImage with your actual resource name
                            }
                            else if (Type == "Scale")
                            {
                                row.Cells["ColMeasure"].Value = Resource.FinalScale_Color; // Replace ScaleImage with your actual resource name
                            }
                            break;
                        }
                    }



                }





            }

        }
        private void RelationsFrm_Load(object sender, EventArgs e)
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

        

        

        public string GetDependentType(string ParameterName)
        {
            Worksheet Sheet2 = workbook.Worksheets[1];
            string DependentType = "";
            for (int column = 0; column <= worksheet.Cells.MaxDataColumn; column++)
            {
                if(ParameterName == Sheet2.Cells[0, column].Value.ToString())
                {
                    DependentType = Sheet2.Cells[1, column].Value.ToString();
                }
            }
            return DependentType;

        }

        

        private void list_Dependent_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        private void pic_AllParaToNominal_Click(object sender, EventArgs e)
        {
            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    list_Nominal.Items.Add(item);
                }

            }

            //foreach (object selectedItem in list_AllParameters.SelectedItems)
            //{
            //    list_Nominal.Items.Add(selectedItem.ToString());
            //}
        }

       

  

        private void pic_RemoveNominalList_Click_1(object sender, EventArgs e)
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

        private void pic_AllParaToNormal_Click(object sender, EventArgs e)
        {
            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    list_NormalScale.Items.Add(item);
                }

            }
            //foreach (object selectedItem in list_AllParameters.SelectedItems)
            //{
            //    list_NormalScale.Items.Add(selectedItem.ToString());
            //}
        }

        private void pic_RemoveNormalList_Click_1(object sender, EventArgs e)
        {
            if (list_NormalScale.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_NormalScale.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_NormalScale.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }
        WordDocument document;
        WordClass wordObj = new WordClass();
        private void btn_Done_Click(object sender, EventArgs e)
        {
            ComparativeBasic();
            document = wordObj.InitWord();
            RelationTableLayout_Nominal();
            wordObj.SaveWord();

            ClearPara();


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
                    parameter.LablesIfNomainal.Clear();
                    parameter.DIC_LablesIfNomainal.Clear();

                }
                item.TestsDone.Clear();
            }
        }
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

            return groupCount;
        }
        public int CountRows(ComparativeTable table)
        {
            int rowCount = 0;

            foreach (var parameter in table.Parameters)
            {

                if (parameter.NominalOrScale == "Nominal")
                {
                    //int distinctValuesCount = parameter.ParameterValues.Distinct().Count();
                    int distinctValuesCount = parameter.DIC_LablesIfNomainal.Keys.Count;
                    rowCount += distinctValuesCount + 1;

                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    // For scale parameters, add 4 (assuming you want to count 4 rows per parameter)
                    rowCount += 3;
                }
            }



            return rowCount;
        }
        public void RelationTableLayout_Nominal()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                
                bool TablehasNominal = false;

                IWSection section = wordObj.CreatePortraitSection();

                foreach(var parameter in ComparativeTables[tableindex].Parameters)
                {
                    if(parameter.NominalOrScale == "Nominal")
                    {
                        TablehasNominal = true;
                        break;
                    }
                }


                int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                int Variablerows = CountRows(ComparativeTables[tableindex]);

                int WordTableColumns = 5;
                int WordTableRows = 2 + Variablerows;

                if (TablehasNominal) 
                {
                    WordTableColumns = 7;
                }
                else if(!TablehasNominal)
                {
                    WordTableColumns = 5;   
                }



                if(TablehasNominal)
                {
                    WordTableRows++;
                }

                

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);








                wordObj.GeneralTableFormat(table);
            }
        }

        private void pic_AllParaToAbnormal_Click_1(object sender, EventArgs e)
        {
            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    list_AbnormalScale.Items.Add(item);
                }

            }

            //foreach (object selectedItem in list_AllParameters.SelectedItems)
            //{
            //    list_AbnormalScale.Items.Add(selectedItem.ToString());
            //}
        }

        private void pic_RemoveAbNormalList_Click_1(object sender, EventArgs e)
        {
            if (list_AbnormalScale.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_AbnormalScale.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_AbnormalScale.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }



        private void pic_ifyes_Click_1(object sender, EventArgs e)
        {
            var childforms = Application.OpenForms.OfType<NominalYesOnly>().ToList();
            if (childforms.Count() == 1)
            {
                childforms.FirstOrDefault().Close();
            }
            else
            {
                NominalYesOnly nominalYesOnly = new NominalYesOnly();
                foreach (string nominal in list_Nominal.Items)
                {
                    nominalYesOnly.AllNominal.Add(nominal);
                }

                nominalYesOnly.Show();
            }
        }
        public void AddTableUI()
        {
            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&
               
                (list_Nominal.Items.Count > 0 || list_NormalScale.Items.Count > 0 || list_AbnormalScale.Items.Count > 0))
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
        public static List<string> NominalYes_Compara { get; set; }
        public void AddTableClass(string tableName)
        {
            //bool totalcolumn = check_TotalColumn.Checked;

            var comparativeTable = new ComparativeTable
            {
                TableName = tableName,
                Parameters = new List<Parameter>(),


                //HasTotalColumn = totalcolumn
            };

            // Add parameters from list_Groups
            

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

            foreach (var parameter in comparativeTable.Parameters)
            {
                if (NominalYes_Compara != null)
                {
                    foreach (string parameteryes in NominalYes_Compara)
                    {
                        if (parameteryes == parameter.Name)
                        {
                            parameter.NominalIsYes = true;
                        }
                    }
                }

            }


            // Add parameters from list_NormalScale
            foreach (var item in list_NormalScale.Items)
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

            // Add parameters from list_AbnormalScale
            foreach (var item in list_AbnormalScale.Items)
            {
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    NominalOrScale = "Scale",
                    NormalOrAbnormal = "Abnormal",
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

        private void btn_SortTable_Click(object sender, EventArgs e)
        {
            if (cmb_TableNames.SelectedIndex != -1)
            {
                string selectedTableName = cmb_TableNames.SelectedItem.ToString();
                var selectedTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                Sort_Parameters_Frm sort_Parameters_Frm = new Sort_Parameters_Frm();
                sort_Parameters_Frm.SortcomparativeTable = selectedTable;
                sort_Parameters_Frm.TableType = "Comparative";
                sort_Parameters_Frm.Show();
            }
        }
        public void ComparativeBasic()
        {
            GetDataValues();

            foreach (ComparativeTable table in ComparativeTables)
            {
                excelFunctionsobj.ReadLablesIfNominal(table);
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

                                if (parameterCellValue.ToString() == ".")
                                {
                                    parameter.hasLowerN = true;
                                    continue;
                                }
                                if (parameterCellValue != null && double.TryParse(parameterCellValue.ToString(), out parameterValue))
                                {
                                    // Add the parameter value to the corresponding group
                                    if (!parameter.GroupedParameterValues.ContainsKey(groupValue))
                                    {
                                        parameter.GroupedParameterValues[groupValue] = new List<double>();
                                    }

                                    parameter.GroupedParameterValues[groupValue].Add(parameterValue);
                                }
                            }
                        }

                    }
                }
            }
            FormatParameters();
        }
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

        public void FormatParameters()
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                //MessageBox.Show(table.TableName);
                foreach (Parameter parameter in table.Parameters)
                {
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
                                parameter.FormattedValues[groupValue][$"Percentage_{distinctValue}"] = $"{percentage:F1}%";
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
    }
}
