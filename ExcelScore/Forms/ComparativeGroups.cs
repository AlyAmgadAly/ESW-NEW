using Accord.Math;
using Accord.Statistics.Kernels;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using CenterSpace.NMath.Core;
using ExcelScore.Classes;
using Humanizer;
using MathNet.Numerics.Statistics;
using Microsoft.SolverFoundation.Services;
using Python.Runtime;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskBand;
using Application = System.Windows.Forms.Application;
using Parameter = ExcelScore.Classes.Parameter;

namespace ExcelScore.Forms
{
    public partial class ComparativeGroups : Form
    {

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public DataGridView Dgv { get; set; }

        public List<ComparativeTable> ComparativeTables  = new List<ComparativeTable>();

        static ExcelFunctions excelFunctionsobj = new ExcelFunctions();

        public List<ComparativeClass> ComparativeClassesList = new List<ComparativeClass>();

        List<string> ColumnExcelheaders = excelFunctionsobj.ReadHeaderColumnsExcel();

        Worksheet worksheet = excelFunctionsobj.GetWorksheet();

        Worksheet Sheet2 = excelFunctionsobj.GetSheet2();
        public ComparativeGroups()
        {
            InitializeComponent();
        }

        private ChooseFrm chooseFrmInstance;

        public ComparativeGroups(ChooseFrm chooseFrm)
        {
            InitializeComponent();
            chooseFrmInstance = chooseFrm;
        }

        public void UpdateData(Dictionary<string, string> newData)
        {
            NormalityParaNameList_ComparaGroups = newData;
            InsertNormality();
        }
        public void AddHeadersToParameter()
        {
            foreach (string columnName in ColumnExcelheaders)
            {
                ComparativeClass comparativeClass = new ComparativeClass();
                comparativeClass.ParameterName = columnName;
                ComparativeClassesList.Add(comparativeClass);
                //list_AllParameters.Items.Add(comparativeClass.ParameterName);
                data_allPara.Rows.Add(comparativeClass.ParameterName);
            }

            InsertPicture_Type();
            foreach (DataGridViewColumn column in data_allPara.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            InsertNormality();



        }

        public void InsertNormality()
        {
            if (NormalityParaNameList_ComparaGroups != null)
            {
                foreach (var kvp in NormalityParaNameList_ComparaGroups)
                {
                    string Name = kvp.Key;
                    string Normality = kvp.Value;


                    foreach (DataGridViewRow row in data_allPara.Rows)
                    {
                        if (row.Cells[0].Value != null && row.Cells[0].Value.ToString() == Name)
                        {
                            row.Cells["ColNormality"].Value = Normality;
                        }
                    }

                }
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

        public Dictionary<string, string> NormalityParaNameList_ComparaGroups { get; set; } = new Dictionary<string, string>();
        private void ComparativeGroups_Load(object sender, EventArgs e)
        {
            check_Maha.Visible = false;
            //pythonStat.InitPython();

            //Margins
            //Top 3.5
            //Bottom 3
            //Left 3.25
            //Right 2.75

            //Layout
            //header 2
            //footer 2
            AddHeadersToParameter();

            


        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Hide();
            chooseFrmInstance.Dgv = Dgv;
            chooseFrmInstance.NormalityParaNameList_ChooseFrm = NormalityParaNameList_ComparaGroups;
            chooseFrmInstance.ReturnToChooseFrm();
            

            //this.Close();
            //ChooseFrm chooseFrm = new ChooseFrm();
            //chooseFrm.Dgv = Dgv;
            //chooseFrm.Show();
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

        private void pic_AllParaToScale_Click(object sender, EventArgs e)
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

        private void pic_RemoveNormalList_Click(object sender, EventArgs e)
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

        private void pic_AllParaToAbnormal_Click(object sender, EventArgs e)
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

        private void pic_RemoveAbNormalList_Click(object sender, EventArgs e)
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

        private void pic_AllParaToGroups_Click(object sender, EventArgs e)
        {
            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    list_Groups.Items.Add(item);
                }

            }
            //foreach (object selectedItem in list_AllParameters.SelectedItems)
            //{
            //    list_Groups.Items.Add(selectedItem.ToString());
            //}
        }

        private void pic_RemoveGroupsList_Click(object sender, EventArgs e)
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

        public void AddTableClass(string tableName)
        {
            bool totalcolumn = check_TotalColumn.Checked;
            
            var comparativeTable = new ComparativeTable
            {
                TableName = tableName,
                Parameters = new List<Parameter>() , 


                HasTotalColumn = totalcolumn
            };

            // Add parameters from list_Groups
            foreach (var item in list_Groups.Items)
            {
                //MessageBox.Show(item.ToString());
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    IsGroup = true,
                    GroupedParameterValues = new Dictionary<double, List<double>>() ,
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }

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

            comparativeTable.FormatType = cmb_ChooseTableFormat.Text;

            ComparativeTables.Add(comparativeTable);



            
        }

        public string AddTableUI()
        {
            string TableName = null;
            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&
                list_Groups.Items.Count > 0 && !string.IsNullOrWhiteSpace(cmb_ChooseTableFormat.Text) &&
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
            check_TotalColumn.Checked = false;
            ComparativeBasic(TableName);


        }

        public void GetDataValues(string TableName)
        {
            //MessageBox.Show(SelectParameterColIndex.ToString());
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



                                    if(cellData == null)
                                    {
                                        MessageBox.Show("Data is empty at Parameter : " + parameter.Name);
                                        break;
                                    }
                                    else if (cellData.ToString() == ".")
                                    {
                                        continue;
                                    }



                                    // Parse the cell data as double
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
                                            if (SelectedParameterValues_CompaFrm.Contains(SelectDataDouble))
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

        public void FindGroupColumnIndex()
        {
            foreach (var table in ComparativeTables)
            {
                for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
                {
                    object cellValue = worksheet.Cells[0, col].Value;
                    string parameterName = cellValue?.ToString();

                    if (!string.IsNullOrEmpty(parameterName))
                    {
                        foreach (var parameter in table.Parameters)
                        {
                            if (parameter.Name == parameterName && parameter.IsGroup)
                            {
                                table.GroupColumnIndex = col;
                                MessageBox.Show(col.ToString());// Set the group column index for this table
                                return;
                            }
                        }
                    }
                }
            }
        }

        public void OldGroupData()
        {
            foreach (var table in ComparativeTables)
            {
                //MessageBox.Show(table.TableName);
                int groupColumnIndex = table.GroupColumnIndex; // Retrieve the group column index from the table

                if (groupColumnIndex == -1)
                {
                    // Handle case where no group parameter is found for this table
                    // MessageBox.Show($"Error: No group parameter found for table {table.TableName}");
                    continue;
                }

                // Iterate through each row to group the data
                for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
                {
                    // Get the group value for the current row
                    object groupCellValue = worksheet.Cells[row, groupColumnIndex].Value;
                    double groupValue;

                    if (groupCellValue != null && double.TryParse(groupCellValue.ToString(), out groupValue))
                    {
                        // Iterate through each parameter and add the data to the corresponding group
                        foreach (var parameter in table.Parameters)
                        {
                            if (!parameter.IsGroup)
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
                                    // Get the parameter value for the current row
                                    object parameterCellValue = worksheet.Cells[row, parameterColumnIndex].Value;
                                    double parameterValue;

                                    if (parameterCellValue != null && double.TryParse(parameterCellValue.ToString(), out parameterValue))
                                    {
                                        // Add the parameter value to the corresponding group
                                        if (!parameter.GroupedParameterValues.ContainsKey(groupValue))
                                        {
                                            parameter.GroupedParameterValues[groupValue] = new List<double>();
                                        }

                                        parameter.GroupedParameterValues[groupValue].Add(parameterValue);
                                       // MessageBox.Show("Group "+groupValue.ToString());
                                        //MessageBox.Show(parameterValue.ToString());
                                        
                                    }





                                }
                            }
                        }
                    }
                }
            }
        }


        public void FormatParameters(string TableName)
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                if(TableName == table.TableName)
                {
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


                                    int Totalrowfreq = 0;
                                    foreach (var kvp2 in parameter.GroupedParameterValues)
                                    {
                                        double groupValue2 = kvp2.Key;
                                        List<double> values2 = kvp2.Value;
                                        int frequency2 = values2.Count(v => v == distinctValue);
                                        Totalrowfreq += frequency2;
                                    }
                                    //MessageBox.Show(frequency2.ToString());

                                    if(check_Perc_Row.Checked)
                                    {
                                        int frequency = values.Count(v => v == distinctValue);
                                        double percentage = (frequency / (double)Totalrowfreq) * 100;
                                        parameter.FormattedValues[groupValue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                        parameter.FormattedValues[groupValue][$"Percentage_{distinctValue}"] = $"{percentage:F1}%";
                                    }
                                    else if(!check_Perc_Row.Checked)
                                    {
                                        int frequency = values.Count(v => v == distinctValue);
                                        double percentage = (frequency / (double)totalCount) * 100;
                                        parameter.FormattedValues[groupValue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                        parameter.FormattedValues[groupValue][$"Percentage_{distinctValue}"] = $"{percentage:F1}%";
                                    }

                                    //MessageBox.Show(frequency.ToString());
                                    // Store frequency and percentage in the FormattedValues dictionary
                                    


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
            

            if (n % 2 == 0)
            {
                int middle = n / 2;
                // Even number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(0, middle));
            }
            else
            {
                //The fix was here we get same range from zero to middle but middle index is different

                int middle = (n + 1) / 2;
                
                return CalculateMedian(values.GetRange(0 ,  middle));
            }
            
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
                int middle = (n+1) / 2;
                // Odd number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle-1));
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
                    ClearTable(tableToRemove);
                }
            }

            cmb_TableNames.Text = "";
            list_ViewTableParameters.Items.Clear();

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
                    if(parameter.IsGroup)
                    {
                        list_ViewTableParameters.Items.Add("Groups : " + parameter.Name);
                    }
                    else if(!parameter.IsGroup)
                    {
                        list_ViewTableParameters.Items.Add(parameter.Name);
                    }
                    
                }
            }
        }
        
        


        

        public int FormatParameters(List<ComparativeTable> tables)
        {
            int groupCount = 0;

            foreach (var table in tables)
            {
                foreach (var parameter in table.Parameters)
                {
                    if (parameter.IsGroup) continue;

                    foreach (var kvp in parameter.GroupedParameterValues)
                    {
                        double groupValue = kvp.Key;
                        List<double> values = kvp.Value;

                        // Calculate total count of values in the group
                        int totalCount = values.Count;

                        // Store formatted values in the FormattedValues dictionary
                        if (!parameter.FormattedValues.ContainsKey(groupValue))
                        {
                            parameter.FormattedValues[groupValue] = new Dictionary<string, string>();
                        }

                        // Calculate frequency and percentage for each value in the group
                        if (parameter.NominalOrScale == "Nominal")
                        {
                            // Calculate frequency and percentage for nominal parameters
                            foreach (var distinctValue in values.Distinct())
                            {
                                int frequency = values.Count(v => v == distinctValue);
                                double percentage = (frequency / (double)totalCount) * 100;

                                // Store frequency and percentage separately in the dictionary with distinct keys
                                parameter.FormattedValues[groupValue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                parameter.FormattedValues[groupValue][$"Percentage_{distinctValue}"] = $"{percentage:F2}%";
                            }
                        }


                        else if (parameter.NominalOrScale == "Scale")
                        {
                            // Calculate statistics for scale parameter
                            double minValue = values.Min();
                            double maxValue = values.Max();
                            double meanValue = values.Average();
                            double stdDevValue = Math.Sqrt(values.Select(x => Math.Pow(x - meanValue, 2)).Sum() / (values.Count - 1));
                            double medianValue = values.OrderBy(x => x).ElementAt(values.Count / 2);
                            double perc25th = CalculateTukeysHingesLower(values);
                            double perc75th = CalculateTukeysHingesUpper(values);

                            // Format scale parameter values
                            string formattedMinMax = FormatMinMaxValue(minValue, maxValue);
                            string formattedMeanStd = FormatMeanStdValue(meanValue, stdDevValue);
                            string formattedMedian = FormatSingleValue(medianValue);
                            string formattedIQR = FormatMinMaxValue(perc25th, perc75th);

                            // Store formatted values in the FormattedValues dictionary
                            if (!parameter.FormattedValues.ContainsKey(groupValue))
                            {
                                parameter.FormattedValues[groupValue] = new Dictionary<string, string>();
                            }

                            parameter.FormattedValues[groupValue]["Min-Max"] = formattedMinMax;
                            parameter.FormattedValues[groupValue]["Mean ± StdDev"] = formattedMeanStd;
                            parameter.FormattedValues[groupValue]["Median"] = formattedMedian;
                            parameter.FormattedValues[groupValue]["IQR"] = formattedIQR;
                        }
                    }
                }

                foreach (var parameter in table.Parameters)
                {
                    if (parameter.IsGroup) continue;
                    groupCount += parameter.GroupedParameterValues.Count;
                }
            }

            return groupCount;
        }


        public double CalculateTukeysHingesLower(List<double> values)
        {
            values.Sort();
            int n = values.Count;
            int lowerHingeIndex = n / 4;
            double lowerHinge = (n % 2 == 0)
            ? (values[lowerHingeIndex - 1] + values[lowerHingeIndex]) / 2.0
            : values[lowerHingeIndex];



            return (lowerHinge + values[lowerHingeIndex - 1]) / 2.0;
        }

        public double CalculateTukeysHingesUpper(List<double> values)
        {
            values.Sort();
            int n = values.Count;

            int upperHingeIndex = 3 * n / 4;
            double upperHinge = (n % 2 == 0)
            ? (values[upperHingeIndex - 1] + values[upperHingeIndex]) / 2.0
            : values[upperHingeIndex];



            return (upperHinge + values[upperHingeIndex + 1]) / 2.0;
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
            if(isStdThreeNumbers)
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
                    // For nominal parameters, add the count of distinct values
                    rowCount += distinctValuesCount+1;
                    
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                   
                    // For scale parameters, add 4 (assuming you want to count 4 rows per parameter)
                    rowCount += 4;
                }
            }


            //MessageBox.Show(rowCount.ToString());

            

            return rowCount;
        }

        public List<double> GetEachGroupCOUNT(ComparativeTable table)
        {
            var groupParameter = table.Parameters.FirstOrDefault(p => p.IsGroup);

            List<double> groupColumn = groupParameter.ParameterValues;

            List<double> EachGroupCount = new List<double>();



            foreach (int group in groupColumn)
            {
                
                if (!groupParameter.EachGroupCount.ContainsKey(group))
                {
                    
                    groupParameter.EachGroupCount[group] = 1;
                }
                else
                {
                    
                    groupParameter.EachGroupCount[group]++;
                }
            }

            var sortedKeys = groupParameter.EachGroupCount.Keys.OrderBy(key => key).ToList();
            foreach (int group in sortedKeys)
            {
                EachGroupCount.Add(groupParameter.EachGroupCount[group]);
                //MessageBox.Show(groupParameter.EachGroupCount[group].ToString());
            }

            return EachGroupCount;
        }


        public void InsertNeachGroup(IWTable table , List<double> EachGroupCount , ComparativeTable comparativeTable)
        {
            //int numberofgroups = EachGroupCount.Count;
            //int GroupCountStr = 1;
            int column = 1;

            //foreach (double GroupCount in EachGroupCount)
            //{
            //    wordObj.AddPara_Center(table, 0, column, "Group " + GroupCountStr + "\n(n = " + GroupCount + ")");
            //    column = column + 2;
            //    GroupCountStr++;
            //}

            Parameter Groupparameter = null;
            foreach (var parameter in comparativeTable.Parameters)
            {
                
                if(parameter.IsGroup)
                {
                    Groupparameter = parameter;
                    break;
                }    
            }

            int groupctr = 0;
            

            foreach (var label in Groupparameter.DIC_LablesIfNomainal.Values)
            {
                //MessageBox.Show(label);
                wordObj.AddPara_Center(table, 0, column, label + "\n(n = " + EachGroupCount[groupctr] + ")");

                column = column + 2;
                groupctr++; 
            }

            


        }

        DefinitionUnderTable definitionUnderTableObj = new DefinitionUnderTable(); 

   

        public void InsertPairwiseF(IWTable table ,ComparativeTable comparativeTable , int WordTableColumns)
        {
            
            int startrow = 1;
            int count = 0;

            Parameter lastParameter = null;
            foreach (var parameter in comparativeTable.Parameters)
            {
               // MessageBox.Show(parameter.Name);
                lastParameter = parameter;
            }

            foreach (var parameter in comparativeTable.Parameters)
            {
                
                if (parameter.IsGroup)
                {
                    continue;
                }
                if(parameter.NominalOrScale =="Nominal")
                {
                    //count = parameter.ParameterValues.Distinct().Count() + 2;
                    count = parameter.DIC_LablesIfNomainal.Keys.Count() + 2;
                }
                else if(parameter.NominalOrScale == "Scale")
                {
                    count = 4;
                }


                if(parameter.ISFAnovaSig)
                {
                    
                    WTableRow row;
                     
                    int insert = startrow + 4;



                    row = table.AddRow();
                    table.Rows.Insert(insert, row);

  
                    int rowindex = table.Rows.IndexOf(row);



                    wordObj.AddParaCombined(table, rowindex, 0, "Sig. bet. grps.", true, true , Syncfusion.Drawing.Color.Empty , Syncfusion.Drawing.Color.Black);
                    table.ApplyHorizontalMerge(rowindex, 1, WordTableColumns-3);


                    for (int i = 0; i < WordTableColumns; i++)
                    {
                        if (parameter == lastParameter)
                        {
                            
                            table.Rows[rowindex].Cells[i].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[rowindex].Cells[i].CellFormat.Borders.Top.LineWidth = 0.5f;
                        }
                        else if(!(parameter == lastParameter))
                        {
                            //MessageBox.Show(rowindex.ToString());
                            table.Rows[rowindex].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[rowindex].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }
                    }

                    WParagraph paragraph = (WParagraph)table[rowindex, 1].AddParagraph();
                    for (int i = 0; i < parameter.FPairwise.Count; i++)
                    {
                        string groupnumber = "";
                        
                        bool issig = false;
                        try
                        {
                            double checkforP = double.Parse(parameter.FPairwise[i]);
                            if(checkforP < 0.05)
                            {
                                issig = true;   
                                //SubSuperScriptText
                            }
                        }
                        catch(FormatException)
                        {
                            issig = true;
                            //wordObj.AddParaCombined(table, rowindex, 0, "Sig. bet. grps.", false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }
                        groupnumber = (i+1).ToString();
                        

                        //wordObj.AddParaCombined(table, rowindex, 1, "p" , false , true , Syncfusion.Drawing.Color.Empty , Syncfusion.Drawing.Color.Black);
                        //wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, groupnumber, "Sub");

                        wordObj.InsertPairwiseF(table, rowindex, 1, "p" , paragraph , false , true , Syncfusion.Drawing.Color.Empty , Syncfusion.Drawing.Color.Black);
                        wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty , groupnumber , "Sub");
                        if(!parameter.FPairwise[i].Contains("<"))
                        {
                            wordObj.InsertPairwiseF(table, rowindex, 1, "=", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }
                        
                        wordObj.InsertPairwiseF(table, rowindex, 1, parameter.FPairwise[i], paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                        if (issig)
                        {
                            wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, "*", "Super");
                        }

                        if (i < parameter.FPairwise.Count-1)
                        {
                            wordObj.InsertPairwiseF(table, rowindex, 1, ",", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }

                    }

                    WTableRow other;
                    other = table.AddRow(false);
                    table.Rows.Insert(insert + 1, other);


                    table.ApplyHorizontalMerge(insert + 1, 1, WordTableColumns - 3);

                    string Column0text = table[insert, 0].Paragraphs[0].Text;
                    wordObj.AddParaCombined(table, insert+1, 0, Column0text, true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                    WParagraph NEWpara = (WParagraph)table[insert + 1, 1].AddParagraph();
                    for (int i = 0; i < parameter.FPairwise.Count; i++)
                    {
                        string groupnumber = "";

                        bool issig = false;
                        try
                        {
                            double checkforP = double.Parse(parameter.FPairwise[i]);
                            if (checkforP < 0.05)
                            {
                                issig = true;
                                //SubSuperScriptText
                            }
                        }
                        catch (FormatException)
                        {
                            issig = true;
                            //wordObj.AddParaCombined(table, rowindex, 0, "Sig. bet. grps.", false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }
                        groupnumber = (i + 1).ToString();

                        wordObj.InsertPairwiseF(table, insert + 1, 1, "p", NEWpara, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        wordObj.SubSuperScriptText(table, insert + 1, 1, Syncfusion.Drawing.Color.Empty, groupnumber, "Sub");
                        if (!parameter.FPairwise[i].Contains("<"))
                        {
                            wordObj.InsertPairwiseF(table, insert + 1, 1, "=", NEWpara, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }

                        wordObj.InsertPairwiseF(table, insert + 1, 1, parameter.FPairwise[i], NEWpara, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                        if (issig)
                        {
                            wordObj.SubSuperScriptText(table, insert + 1, 1, Syncfusion.Drawing.Color.Empty, "*", "Super");
                        }

                        if (i < parameter.FPairwise.Count - 1)
                        {
                            wordObj.InsertPairwiseF(table, insert + 1, 1, ",", NEWpara, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }

                    }

                    table.Rows[insert + 1].Cells[0].CellFormat.Borders.Right.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[insert + 1].Cells[0].CellFormat.Borders.Right.LineWidth = 1.5f;

                    table.Rows[insert + 1].Cells[WordTableColumns-2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[insert + 1].Cells[WordTableColumns - 2].CellFormat.Borders.Left.LineWidth = 1.5f;

                    for (int i = 0; i < WordTableColumns; i++)
                    {
                        if (parameter == lastParameter)
                        {

                            table.Rows[insert + 1].Cells[i].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[insert + 1].Cells[i].CellFormat.Borders.Top.LineWidth = 0.5f;
                        }
                        else if (!(parameter == lastParameter))
                        {
                            //MessageBox.Show(rowindex.ToString());
                            table.Rows[insert + 1].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[insert + 1].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }
                    }


                    table.Rows[insert + 1].Cells[WordTableColumns - 2].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;
                    table.Rows[insert + 1].Cells[WordTableColumns - 1].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;


                    table.Rows.RemoveAt(insert);


                    startrow = startrow + 1;
                }

                startrow = startrow + count;    
            }




        }

        public void ComparativeBasic(string TableName)
        {
            GetDataValues(TableName);
            
            foreach (ComparativeTable table in ComparativeTables)
            {
                if(TableName == table.TableName)
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

                                            //log(groupValue);
                                            //log(parameterValue);
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
                                            if (SelectedParameterValues_CompaFrm.Contains(SelectDataDouble))
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


        WordDocument document;
        public void AddTotalColumn(IWTable table , ComparativeTable comparativeTable)
        {
            bool tableHasNominal = false;
            foreach (WTableRow row in table.Rows)
            {
                // Add a cell to each row at the specified column index
                WTableCell newCell = new WTableCell(document);

                row.Cells.Insert(1, newCell);

            }

            foreach (WTableRow row in table.Rows)
            {
                // Add a cell to each row at the specified column index
                WTableCell newCell = new WTableCell(document);

                row.Cells.Insert(1, newCell);

            }

            //Merge Total

            table.ApplyHorizontalMerge(0, 1, 2);
            table.ApplyHorizontalMerge(1, 1, 2);

            //table.ApplyVerticalMerge(1, 0, 1);



            //boders

            table.Rows[0].Cells[1].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
            table.Rows[0].Cells[1].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            int TotalN = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(!parameter.hasLowerN)
                {
                    TotalN = parameter.ParameterValues.Count;
                    break;
                }
            }

            wordObj.AddPara_Center(table, 0, 1 , "Total\n(n = "+ TotalN + ")");


            int startrow = 2;

            foreach (var parameter in comparativeTable.Parameters)
            {
                var values = parameter.ParameterValues;
                int count = 0;
                if (parameter.IsGroup)
                {

                    continue;
                }
                if (parameter.NominalOrScale == "Nominal")
                {
                    tableHasNominal = true;
                    count = parameter.DIC_LablesIfNomainal.Keys.Count+1;
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    if (parameter.ISFAnovaSig)
                    {
                        count = 5;
                    }
                    else
                    {
                        count = 4;
                    }

                }




                //border of each parameter
                // except the first one 
                if(startrow!= 2)
                {
                    table.Rows[startrow].Cells[1].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[startrow].Cells[1].CellFormat.Borders.Top.LineWidth = 0.5f;

                    table.Rows[startrow].Cells[2].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[startrow].Cells[2].CellFormat.Borders.Top.LineWidth = 0.5f;
                }


                int totalCount = parameter.ParameterValues.Count;

                if (parameter.NominalOrScale == "Nominal")
                {
                    // Calculate frequency and percentage for each distinct value in the group
                    foreach (var distinctValue in parameter.ParameterValues.Distinct())
                    {
                        int frequency = parameter.ParameterValues.Count(v => v == distinctValue);
                        double percentage = (frequency / (double)totalCount) * 100;

                        if (parameter.DIC_LablesIfNomainal.ContainsKey((int)distinctValue))
                        {
                            if (startrow != 2)
                            {
                                wordObj.Addpara_CenterNoBOLD(table, startrow + (int)distinctValue , 1, frequency.ToString());
                                wordObj.Addpara_CenterNoBOLD(table, startrow + (int)distinctValue , 2, percentage.ToString("0.0"));
                            }
                            else if (startrow == 2)
                            {
                                wordObj.Addpara_CenterNoBOLD(table, startrow + (int)distinctValue+1, 1, frequency.ToString());
                                wordObj.Addpara_CenterNoBOLD(table, startrow + (int)distinctValue+1, 2, percentage.ToString("0.0"));
                            }


                        }
                    }
                }
                else if (parameter.NominalOrScale == "Scale")
                {


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


                    //merge first


                    //for (int i = 0; i < 4; i++)
                    //{
                    //    table.ApplyHorizontalMerge(startrow + i, 1, 2);
                    //}

                    



                    if(tableHasNominal)
                    {
                        table.Rows[1].Cells[1].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[1].Cells[1].CellFormat.Borders.Bottom.LineWidth = 1.5f;
                        table.ApplyVerticalMerge(1, 0, 1);
                        table.ApplyHorizontalMerge(startrow , 1, 2);
                        table.ApplyHorizontalMerge(startrow + 1, 1, 2);
                        table.ApplyHorizontalMerge(startrow + 2, 1, 2);
                        table.ApplyHorizontalMerge(startrow + 3, 1, 2);


                        wordObj.Addpara_CenterNoBOLD(table, startrow +1, 1, formattedMinMax);
                        wordObj.Addpara_CenterNoBOLD(table, startrow + 2, 1, formattedMeanStd);
                        wordObj.Addpara_CenterNoBOLD(table, startrow + 3, 1, formattedMedian + " (" + formattedIQR + ")");



                        if (parameter.NormalOrAbnormal == "Normal")
                        {
                            wordObj.HighlightCellContent(table, startrow + 3, 2);

                        }
                        else
                        {
                            wordObj.HighlightCellContent(table, startrow + 2, 2);

                        }
                    }
                    else if(!tableHasNominal)
                    {
                        table.ApplyHorizontalMerge(startrow, 1, 2);
                        table.ApplyHorizontalMerge(startrow + 1, 1, 2);
                        table.ApplyHorizontalMerge(startrow + 2, 1, 2);
                        //table.ApplyHorizontalMerge(startrow + 3, 1, 2);


                        wordObj.Addpara_CenterNoBOLD(table, startrow + 0, 1, formattedMinMax);
                        wordObj.Addpara_CenterNoBOLD(table, startrow + 1, 1, formattedMeanStd);
                        wordObj.Addpara_CenterNoBOLD(table, startrow + 2, 1, formattedMedian + " (" + formattedIQR + ")");
                    }

                   
                    

                    // Calculate scale statistics



                }



                startrow = startrow + count;


            }




            
            


        }
        public void log(object message)
        {
            MessageBox.Show(message.ToString());
        }

        public int CountPeriodsUp_Groups_Columns(ComparativeTable comparativeTable)
        {
            int ColumnCtr = 0;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.NominalOrScale == "Scale")
                {
                    ColumnCtr++;
                }

            }

            return ColumnCtr;
        }
        public void ComparativeTablePeriodsUp_Groups_Layout()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Periods Groups")
                {
                    bool TableHasSigI = false;

                    IWSection section = wordObj.CreatePortraitSection();




                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = (5 * numberofgroups) +3;

                    //log(Variablerows);


                    int WordTableColumns = 3 + CountPeriodsUp_Groups_Columns(ComparativeTables[tableindex]);

                    int WordTableRows = Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.GeneralTableFormat(table);

                    wordObj.ApplyGeneralPeriodsUp_Groups_ComparativeMerges(table, WordTableColumns , WordTableRows, numberofgroups, ComparativeTables[tableindex]);


                    wordObj.ApplyGeneral_PeriodsUp_Groups_ComparativeBorders(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);

                    wordObj.SetComparative_PeriodsUp_GroupsWidths(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);


                    wordObj.SetComparative_PeriodsUp_Groups_Headers(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);


                    InsertData_scale_PeriodsUp_Groups(table, ComparativeTables[tableindex]);

                    GetGroupsTest_PeriodsUp_groups(table , ComparativeTables[tableindex] , WordTableRows);


                    GetPeriodsTest_PeriodsUp_groups(table, ComparativeTables[tableindex] , WordTableColumns);


                    wordObj.LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    wordObj.FormatTable(table, 12);
                }
            }
        }
        public void GetPeriodsTest_PeriodsUp_groups(IWTable table, ComparativeTable comparativeTable, int WordTableColumns)
        {
            var scaleParameters = comparativeTable.Parameters
                                                  .Where(p => p.NominalOrScale == "Scale")
                                                  .ToList();

            if (scaleParameters.Count == 2)
            {

               


                int row = 3;
                var firstParameterValues = new List<double>();
                var secondParameterValues = new List<double>();

                int count = 0;
                int tempctr = 0;
                foreach (var kvp in scaleParameters[0].GroupedParameterValues)
                {
                    var key = kvp.Key;
                    var values = kvp.Value;
                    count = values.Count;

                    
                    for (int i = 0; i < count; i++)
                    {
                        firstParameterValues.Add(scaleParameters[0].ParameterValues[i]);
                        secondParameterValues.Add(scaleParameters[1].ParameterValues[i]);


                    }

                    var firstParameterArray = firstParameterValues.ToArray();
                    var secondParameterArray = secondParameterValues.ToArray();

                    string[] result = manual.Tpaired(firstParameterArray, secondParameterArray);
                    wordObj.InsertTest_P(table, row, WordTableColumns - 2, result);
                    row += 5;


                    firstParameterValues.Clear();
                    secondParameterValues.Clear();
                    firstParameterArray.Clear();
                    secondParameterArray.Clear();



                }




            }
        }

        public void GetGroupsTest_PeriodsUp_groups(IWTable table ,ComparativeTable comparativeTable , int WordTableRows)
        {

            int col = 1;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.IsGroup)
                {
                    continue;
                }

                if(parameter.NominalOrScale == "Scale")
                {

                    
                    if(parameter.NormalOrAbnormal == "Normal")
                    {
                        List<double> group1Values = new List<double>();
                        List<double> group2Values = new List<double>();

                        SplitGroupedParameterValues(parameter, out group1Values, out group2Values);

                        string[] values = manual.StudentT_Unpaired(group1Values, group2Values);



                        wordObj.AddParaCombined(table, WordTableRows-1, col, values[0], false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);


                        WParagraph testparaHighlight = (WParagraph)table[WordTableRows-1, col].Paragraphs[0];
                        

                        WTextRange PText = new WTextRange(testparaHighlight.Document);
                        PText.Text = " (" + values[1] + ")";

                        testparaHighlight.ChildEntities.Insert(1, PText);
                        
                        


                        col++;
                    }







                }
            }
            
        }

        public void InsertData_scale_PeriodsUp_Groups(IWTable table , ComparativeTable comparativeTable)
        {
            int col = 1;
            int row = 3;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }
                if (parameter.NominalOrScale == "Scale")
                {

                    var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();



                    foreach (var groupValue in sortedKeys)
                    {

                        Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];


                        foreach (var stat in formattedValues)
                        {
                            wordObj.Addpara_CenterNoBOLD(table, row, col, stat.Value);
                            row++;

                        }
                        row = row + 2;



                    }

                }

                col++;
                row = 3;
            }
        }
        private void btn_Done_Click(object sender, EventArgs e)
        {
            //ComparativeBasic();
            

            pythonStat.InitPython();

            document = wordObj.InitWord();

            
                
            ComparativeTableGroups_Layout();
            ComparativeTablePeriodsUp_Groups_Layout();



            wordObj.SaveWord();



            ClearPara();

            

        }
        WordClass wordObj = new WordClass();
        

        public void ClearTable(ComparativeTable comparativeTable)
        {
            foreach (var parameter in comparativeTable.Parameters)
            {
                parameter.ParameterValues.Clear();
                parameter.FormattedValues.Clear();
                parameter.GroupedParameterValues.Clear();
                parameter.EachGroupCount.Clear();
                parameter.FPairwise.Clear();
                parameter.LablesIfNomainal.Clear();
                parameter.DIC_LablesIfNomainal.Clear();

            }
            comparativeTable.TestsDone.Clear();
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
                    parameter.EachGroupCount.Clear();
                    //parameter.FPairwise.Clear();
                    //parameter.LablesIfNomainal.Clear();
                    //parameter.DIC_LablesIfNomainal.Clear();

                }
                //item.TestsDone.Clear();
            }
        }




        public void newComparativeParamaeterBorders(IWTable table, int WordTableRows, int WordTableColumns, int tableindex, int numberofgroups)
        {
            int currentRow = 1;

            foreach (var parameter in ComparativeTables[tableindex].Parameters)
            {
                int count = 0;
                if (parameter.NominalOrScale == "Nominal")
                {
                    // Get the count of distinct elements for nominal parameters
                    count = parameter.ParameterValues.Distinct().Count() + 1;
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    // For scale parameters, the count is 4
                    count = 4;
                }

                if (count > 0)
                {
                    // Check if the current row and count are within table bounds
                    if (currentRow + count < WordTableRows)
                    {
                        // Apply vertical merges
                        table.ApplyVerticalMerge(WordTableColumns - 1, currentRow + 2, currentRow + count);
                        table.ApplyVerticalMerge(WordTableColumns - 2, currentRow + 2, currentRow + count);
                    }

                    // Insert data for nominal parameters
                    if (parameter.NominalOrScale == "Nominal")
                    {
                        // Insert the parameter name
                        wordObj.AddPara_NoCenter(table, currentRow + 1, 0, parameter.Name);

                        // Start inserting data from row 2 (index 1)
                        int row = 2;

                        // Iterate over each distinct value of the nominal parameter
                        foreach (var value in parameter.ParameterValues.Distinct())
                        {
                            // Insert the value into the table
                            wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, 0, value.ToString());
                            wordObj.LeftIntendBeforeText(table, currentRow + row, 0, 14.17f);

                            // Iterate over each group and insert the frequency and percentage values
                            int col = 1; // Start from column 1
                            foreach (var groupValue in parameter.FormattedValues.Keys)
                            {
                                // Debugging message to check if the value exists in the FormattedValues dictionary
                                MessageBox.Show($"Checking for value: {value}, Group Value: {groupValue}, Keys (FormattedValues): {string.Join(", ", parameter.FormattedValues[groupValue].Keys)}");

                                // Check if the value exists in the FormattedValues dictionary for the current groupValue
                                if (parameter.FormattedValues[groupValue].ContainsKey(value.ToString()))
                                {
                                    // Value exists in the group, retrieve frequency and percentage
                                    string frequency = parameter.FormattedValues[groupValue][value.ToString()];
                                    string percentage = parameter.FormattedValues[groupValue][value.ToString() + " Percentage"];

                                    // Insert frequency and percentage into the table
                                    wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col, frequency);
                                    wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col + 1, percentage);
                                }
                                else
                                {
                                    // Value doesn't exist in the group, insert 0 for both frequency and percentage
                                    wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col, "0");
                                    wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col + 1, "0%");
                                }

                                col += 2; // Move to the next group columns
                            }

                            row++; // Move to the next row for the next distinct value
                        }
                    }

                    // Update currentRow
                    currentRow += count;
                }
            }


        }
        


        public void ComparativeParamaeterBorders(IWTable table, int WordTableRows, int WordTableColumns , int tableindex , int numberofgroups)
        {
            int currentRow = 1;

            foreach (var parameter in ComparativeTables[tableindex].Parameters)
            {
               // MessageBox.Show(parameter.Name);
                int count = 0;
                if (parameter.NominalOrScale == "Nominal")
                {

                    // Get the count of distinct elements for nominal parameters
                    //count = parameter.ParameterValues.Distinct().Count()+1;
                    count = parameter.DIC_LablesIfNomainal.Keys.Count + 1;
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    // For scale parameters, the count is 4
                    count = 4;
                }

                if (count > 0)
                {
                    if (currentRow + count < WordTableRows)
                    {

                        table.ApplyVerticalMerge(WordTableColumns - 1, currentRow + 2, currentRow + count);
                        table.ApplyVerticalMerge(WordTableColumns - 2, currentRow + 2, currentRow + count);
                    }

                    if(parameter.NominalOrScale == "Scale")
                    {
                        for (int j = 0; j < 4; j++)
                        {
                            for (int i = 1; i <= numberofgroups * 2; i = i+2)
                            {
                                table.ApplyHorizontalMerge(currentRow + j+1, i, i + 1);
                            }
                        }

                        wordObj.AddPara_NoCenter(table,currentRow+1 , 0 , parameter.Name);

                        if(parameter.hasLowerN)
                        {
                            int InsertLowerN = 1;
                            foreach (double group in parameter.GroupedParameterValues.Keys)
                            {
                                int parameterGroupCount = parameter.GroupedParameterValues[group].Count;
                                string InsertedN = "(n = " + parameterGroupCount + ")";
                                wordObj.AddParaCombined(table, currentRow + 1, InsertLowerN, InsertedN, true, true,  Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);
                                InsertLowerN = InsertLowerN + 2;
                            }
                        }

                        wordObj.Addpara_NoCenterNoBOLD(table, currentRow + 2, 0, "Min – Max.");
                        wordObj.LeftIntendBeforeText(table, currentRow + 2, 0, 14.17f);

                        wordObj.Addpara_NoCenterNoBOLD(table, currentRow + 3, 0, "Mean ± SD.");
                        wordObj.LeftIntendBeforeText(table, currentRow + 3, 0, 14.17f);

                        wordObj.Addpara_NoCenterNoBOLD(table, currentRow + 4, 0, "Median (IQR)");
                        wordObj.LeftIntendBeforeText(table, currentRow + 4, 0, 14.17f);
                    }
                    else if(parameter.NominalOrScale == "Nominal")
                    {


                        wordObj.AddPara_NoCenter(table, currentRow + 1, 0, parameter.Name);


                        if (parameter.hasLowerN)
                        {
                            int InsertLowerN = 1;
                            foreach (double group in parameter.GroupedParameterValues.Keys)
                            {
                                table.ApplyHorizontalMerge(currentRow + 1 , InsertLowerN , InsertLowerN+1);
                                int parameterGroupCount = parameter.GroupedParameterValues[group].Count;
                                string InsertedN = "(n = " + parameterGroupCount + ")";
                                wordObj.AddParaCombined(table, currentRow + 1, InsertLowerN, InsertedN, true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);
                                InsertLowerN = InsertLowerN + 2;
                            }
                        }

                        int row = 2;

                        foreach (var label in parameter.DIC_LablesIfNomainal.Values)
                        {
                            wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, 0, label.ToString());
                            wordObj.LeftIntendBeforeText(table, currentRow + row, 0, 14.17f);
                            row++;
                        }

                        //var sortedValues = parameter.ParameterValues.Distinct().OrderBy(value => value);
                        //foreach (var value in sortedValues)
                        //{
                        //    wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, 0, value.ToString());
                        //    wordObj.LeftIntendBeforeText(table, currentRow + row, 0, 14.17f);
                        //    row++; 
                        //}

                    }

                }


                if (currentRow + count < WordTableRows-1)
                {
                    
                    // Insert bottom border for each parameter
                    if (count > 0)
                    {
                        for (int i = 0; i < WordTableColumns; i++)
                        {
                            table.Rows[currentRow + count].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[currentRow + count].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }

                        currentRow += count;
                        
                    }
                    
                    
                }
            }
            
        }

        public void InsertDataIntoWord(IWTable table, int WordTableRows, int WordTableColumns, int tableindex, int numberofgroups)
        {
            int currentRow = 1;
            foreach (var parameter in ComparativeTables[tableindex].Parameters)
            {
                if (parameter.NominalOrScale == "Nominal")
                {
                    // Insert the parameter name
                    wordObj.AddPara_NoCenter(table, currentRow + 1, 0, parameter.Name);

                    // Determine the number of distinct values for the nominal parameter
                    int distinctCount = parameter.ParameterValues.Distinct().Count();

                    // Start inserting data from row 2 (index 1)
                    int row = 2;

                    // Iterate over each distinct value of the nominal parameter
                    foreach (var value in parameter.ParameterValues.Distinct())
                    {
                        // Insert the value into the table
                        wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, 0, value.ToString());
                        wordObj.LeftIntendBeforeText(table, currentRow + row, 0, 14.17f);

                        // Iterate over each group and insert the frequency and percentage values
                        int col = 2; // Start from column 2
                        foreach (var groupValue in parameter.FormattedValues.Keys)
                        {
                            if (parameter.FormattedValues[groupValue].ContainsKey(value.ToString()))
                            {
                                string frequency = parameter.FormattedValues[groupValue][value.ToString()];
                                string percentage = parameter.FormattedValues[groupValue][value.ToString() + " Percentage"];

                                // Insert frequency and percentage into the table
                                wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col, frequency);
                                wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col + 1, percentage);
                            }
                            else
                            {
                                // Insert 0 for missing values
                                wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col, "0");
                                wordObj.Addpara_NoCenterNoBOLD(table, currentRow + row, col + 1, "0%");
                            }

                            col += 2; // Move to the next group columns
                        }

                        row++; // Move to the next row for the next distinct value
                    }
                }

            }
        }

        public string ParseData(string input)
        {
            // Define regular expressions to match different formats
            Regex rangeRegex = new Regex(@"([\d.]+) – ([\d.]+)");
            Regex meanStdDevRegex = new Regex(@"([\d.]+) ± ([\d.]+)");
            Regex medianQuartileRegex = new Regex(@"([\d.]+) \(([\d.]+) – ([\d.]+)\)");

            // Match the input against the regular expressions
            Match rangeMatch = rangeRegex.Match(input);
            Match meanStdDevMatch = meanStdDevRegex.Match(input);
            Match medianQuartileMatch = medianQuartileRegex.Match(input);

            // Process based on the matched format
            if (rangeMatch.Success)
            {
                double min = double.Parse(rangeMatch.Groups[1].Value);
                double max = double.Parse(rangeMatch.Groups[2].Value);
                return $"{FormatNumber(min)} – {FormatNumber(max)}";
            }
            else if (meanStdDevMatch.Success)
            {
                double mean = double.Parse(meanStdDevMatch.Groups[1].Value);
                double stdDev = double.Parse(meanStdDevMatch.Groups[2].Value);
                return $"{FormatNumber(mean)} ± {FormatNumber(stdDev)}";
            }
            else if (medianQuartileMatch.Success)
            {
                double median = double.Parse(medianQuartileMatch.Groups[1].Value);
                double lowerQuartile = double.Parse(medianQuartileMatch.Groups[2].Value);
                double upperQuartile = double.Parse(medianQuartileMatch.Groups[3].Value);
                return $"{FormatNumber(median)} ({FormatNumber(lowerQuartile)} – {FormatNumber(upperQuartile)})";
            }
            else
            {
                // No match found, return original input
                return input;
            }
        }

        string FormatNumber(double number)
        {
            // Check if the number is in the format "x.y"
            if (Math.Abs(number - Math.Round(number)) < 0.001)
            {
                return $"{number:F1}";
            }
            // Check if the number is in the format "x.y - z"
            else if (number.ToString().Contains(" - "))
            {
                string[] parts = number.ToString().Split(new string[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
                return $"{double.Parse(parts[0]):F1} - {double.Parse(parts[1]):F1}";
            }
            // Check if the number is in the format "x.y(z - w)"
            else if (number.ToString().Contains("(") && number.ToString().Contains(")"))
            {
                int startIndex = number.ToString().IndexOf("(");
                int endIndex = number.ToString().IndexOf(")");
                string mainPart = number.ToString().Substring(0, startIndex).Trim();
                string rangePart = number.ToString().Substring(startIndex + 1, endIndex - startIndex - 1).Trim();
                string[] rangeParts = rangePart.Split(new string[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
                return $"{double.Parse(mainPart):F0}({double.Parse(rangeParts[0]):F1} - {double.Parse(rangeParts[1]):F1})";
            }
            // Default case: use "F1" format specifier
            else
            {
                return $"{number:F1}";
            }
        }
        

        public void SplitGroupedParameterValues(Parameter parameter, out List<double> group1Values, out List<double> group2Values)
        {
            group1Values = new List<double>();
            group2Values = new List<double>();

            int count = 0;
            foreach (var kvp in parameter.GroupedParameterValues)
            {
                List<double> currentValues = kvp.Value; // Get the list of values associated with the current key

                // Determine which list to add the values to based on the index
                if (count % 2 == 0)
                {
                    group1Values.AddRange(currentValues);
                }
                else
                {
                    group2Values.AddRange(currentValues);
                }

                count++;
            }
        }
        

        pythonStat pythonStat = new pythonStat();
        ManualTests manual = new ManualTests();
        public void PerformTest( IWTable table, int numberofgroups ,  Parameter parameter , int addrow , int addColumn , int WordTableColumns , bool issame , ComparativeTable comparativeTable)
        {
            bool TableHasSigI;


            if (parameter.IsGroup)
            {
                return;
            }
            if (parameter.NominalOrScale == "Nominal")
            {
                    // chi

            }
            if (parameter.NominalOrScale == "Scale")
            {

                if (numberofgroups == 2)
                {
                    if (parameter.NormalOrAbnormal == "Normal")
                    {
                        comparativeTable.TestsDone.Add("tstudent");
                        List<double> group1Values = new List<double>();
                        List<double> group2Values = new List<double>();

                        SplitGroupedParameterValues(parameter, out group1Values, out group2Values);

                        string[] values = manual.StudentT_Unpaired(group1Values, group2Values);
                        //string[] values = pythonStat.PerformTTest(group1Values, group2Values);

                        

                        TableHasSigI = TableHasSig(values);

                        if (TableHasSigI)
                        {
                            comparativeTable.hasSig = true;
                        }

                        if (issame)
                        {
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }
                        else if (!issame)
                        {
                            values[0] = "t=\n" + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }


                        wordObj.HighlightCellContent(table, addrow + 2, WordTableColumns);
                        //wordObj.InsertTest_P(table, addrow, addColumn+1 , values[1]);
                    }
                    else if (parameter.NormalOrAbnormal == "Abnormal")
                    {
                        comparativeTable.TestsDone.Add("U");
                        List<double> group1Values = new List<double>();
                        List<double> group2Values = new List<double>();

                        SplitGroupedParameterValues(parameter, out group1Values, out group2Values);
                        //pythonStat.InitPython();
                        //string[] values = pythonStat.MannWhitneyUTest(group1Values, group2Values);
                        string[] values = manual.UTest(group1Values, group2Values);

                        MessageBox.Show(values[1]);

                        TableHasSigI = TableHasSig(values);

                        if (TableHasSigI)
                        {
                            comparativeTable.hasSig = true;
                        }
                        if (issame)
                        {
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }
                        else if (!issame)
                        {
                            values[0] = "U=\n" + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }



                        wordObj.HighlightCellContent(table, addrow + 1, WordTableColumns);
                        //wordObj.InsertTest_P(table, addrow, addColumn+1, values[1]);


                        //U test
                    }
                }

                else if (numberofgroups > 2)
                {

                    if (parameter.NormalOrAbnormal == "Normal")
                    {
                        comparativeTable.TestsDone.Add("FAnova");
                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        //anovaTestResult = pythonStat.ANOVAWithTukeyHSDNewDynamic(parameter);

                        anovaTestResult = manual.Fanova(parameter);

                        string[] values = { anovaTestResult.TestValue, anovaTestResult.PValue };



                        TableHasSigI = TableHasSig(values);

                        if (TableHasSigI)
                        {
                            comparativeTable.hasSig = true;
                        }
                        if (issame)
                        {
                            wordObj.InsertTest_P(table, addrow, addColumn, values);

                        }
                        else if (!issame)
                        {
                            values[0] = "F=\n" + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);

                        }



                        wordObj.HighlightCellContent(table, addrow + 2, WordTableColumns);

                        // F anova
                    }
                    else if (parameter.NormalOrAbnormal == "Abnormal")
                    {
                        //H kriskual
                        //
                        comparativeTable.TestsDone.Add("H");
                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        //anovaTestResult = pythonStat.KruskalWallisNoDunnDynamic(parameter);

                        anovaTestResult = manual.Kruskal_H(parameter);
                        string[] values = { anovaTestResult.TestValue, anovaTestResult.PValue };


                        TableHasSigI = TableHasSig(values);

                        if (TableHasSigI)
                        {
                            comparativeTable.hasSig = true;
                        }
                        if (issame)
                        {
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }
                        else if (!issame)
                        {
                            values[0] = "H=\n" + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }
                        wordObj.HighlightCellContent(table, addrow + 1, WordTableColumns);
                        // wordObj.Addpara_CenterNoBOLD(table, addrow, addColumn +1, anovaTestResult.PValue);

                    }

                }

            }




        }
        
        public (string, string , bool) InsertSeperateTest_TestOfSig(ComparativeTable comparativeTable , int numberofgroups)
        {
            bool allsame = true;
            string type = "";
            string NominalOrScale = "";

            // Counters for each condition
            int nominalCount = 0;
            int normalAndScaleCount = 0;
            int abnormalAndScaleCount = 0;
            int totalParameterCount = 0;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup) { continue; }
                totalParameterCount++; // Increment total count only for non-group parameters

                if (parameter.NominalOrScale == "Nominal")
                {
                    NominalOrScale = "Nominal";

                    nominalCount++;
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    if (parameter.NormalOrAbnormal == "Normal")
                    {
                        normalAndScaleCount++;
                    }
                    else if (parameter.NormalOrAbnormal == "Abnormal")
                    {
                        abnormalAndScaleCount++;
                    }
                }
            }

            // Check conditions
            if (nominalCount == totalParameterCount)
            {
                type = "χ²";
            }
            else if (normalAndScaleCount == totalParameterCount)
            {
                
                if (numberofgroups == 2)
                {
                    type = "t";
                }
                else if (numberofgroups > 2)
                {
                    type = "F";
                }
            }
            else if (abnormalAndScaleCount == totalParameterCount)
            {
                
                if (numberofgroups == 2)
                {
                    type = "U";
                }
                else if (numberofgroups > 2)
                {
                    type = "H";
                }
            }
            else
            {
                type = "Test of Sig.";
                allsame = false;
            }

            return  (type , NominalOrScale , allsame);
        }

        public bool TableHasSig(string[] values)
        {
            bool Significant = false;
            if (values[1] == "<0.001")
            {
                Significant = true; 
            }
            else
            {
                double p = double.Parse(values[1]);
                if(p < 0.05)
                {
                    Significant = true;
                }
                else if(p >= 0.05)
                {
                    Significant = false;
                }


            }


            return Significant;
        }


        public void ComparativeTableGroups_Layout()
        {
            for(int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Default")
                {
                    bool TableHasSigI = false;

                    IWSection section = wordObj.CreatePortraitSection();




                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = CountRows(ComparativeTables[tableindex]);


                    //MessageBox.Show(Variablerows.ToString());

                    //MessageBox.Show(Variablerows.ToString());

                    int WordTableColumns = 3 + (numberofgroups * 2);

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);


                    wordObj.GeneralTableFormat(table);

                    wordObj.ApplyGeneralComparativeMerges(table, WordTableColumns, numberofgroups);

                    wordObj.ApplyGeneralComparativeBorders(table, WordTableRows, WordTableColumns, numberofgroups);




                    wordObj.SetComparativeWidths(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);


                    wordObj.Add_GeneralHeaders_Comparative_Center(table, WordTableRows, WordTableColumns, numberofgroups);

                    ComparativeParamaeterBorders(table, WordTableRows, WordTableColumns, tableindex, numberofgroups);

                    int startingRow;
                    int newRowCount = 3;

                    (string testtype, string NominalOrScale, bool issame) = InsertSeperateTest_TestOfSig(ComparativeTables[tableindex], numberofgroups);

                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, testtype);



                    foreach (Parameter parameter in ComparativeTables[tableindex].Parameters)
                    {
                        startingRow = newRowCount;
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
                            // count = parameter.ParameterValues.Distinct().Count() + 1;
                            count = parameter.DIC_LablesIfNomainal.Keys.Count() + 1;
                            ComparativeTables[tableindex].TestsDone.Add("Chi");

                            //string[] values = pythonStat.PerformChiSquareTest(parameter);
                            string[] values = manual.getchi(parameter);

                            TableHasSigI = TableHasSig(values);

                            if (TableHasSigI)
                            {
                                ComparativeTables[tableindex].hasSig = true;
                            }

                            if (issame)
                            {
                                wordObj.InsertTest_P(table, startingRow, WordTableColumns - 2, values);
                            }
                            else if (!issame)
                            {
                                values[0] = "χ²=\n" + values[0];
                                wordObj.InsertTest_P(table, startingRow, WordTableColumns - 2, values);
                            }

                            //MessageBox.Show(parameter.Isfisher.ToString());
                            if (parameter.Isfisher)
                            {
                                WParagraph testparaHighlight = (WParagraph)table[startingRow, WordTableColumns - 1].Paragraphs[0];
                                WTextRange FEtext = new WTextRange(testparaHighlight.Document);
                                FEtext.Text = "FE";
                                FEtext.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

                                WTextRange PText = new WTextRange(testparaHighlight.Document);
                                PText.Text = "p=";

                                testparaHighlight.ChildEntities.Insert(0, PText);
                                testparaHighlight.ChildEntities.Insert(0, FEtext);

                            }



                            //wordObj.Addpara_CenterNoBOLD(table, startingRow, WordTableColumns - 1, values[1]);

                            // Insert frequency and percentage for nominal parameters

                            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();

                            //foreach (var kvp in parameter.FormattedValues)
                            foreach (var groupValue in sortedKeys)
                            {
                                //double groupValue = kvp.Key;
                                // MessageBox.Show(groupValue.ToString());
                                Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];
                                //MessageBox.Show(groupValue.ToString());
                                // Insert frequencies for the current group
                                foreach (var distinctValue in parameter.DIC_LablesIfNomainal.Keys)
                                {

                                    string frequency = formattedValues.ContainsKey($"Frequency_{distinctValue}") ? formattedValues[$"Frequency_{distinctValue}"] : "0";
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, frequency);
                                    //startingRow++; // Move to the next row
                                    //MessageBox.Show(frequency.ToString());
                                    column++;

                                    string percentage = formattedValues.ContainsKey($"Percentage_{distinctValue}") ? formattedValues[$"Percentage_{distinctValue}"] : "0.0";
                                    percentage = percentage.Replace("%", "");


                                    if (percentage.EndsWith("0"))
                                    {
                                        //percentage = percentage.Substring(0, percentage.Length - 1);
                                        //MessageBox.Show(percentage);
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage);
                                    }

                                    else
                                    {

                                        //double percentageDouble = Math.Round(double.Parse(percentage), 1);
                                        //MessageBox.Show(percentageDouble.ToString());
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage.ToString());
                                    }
                                    //wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage);
                                    startingRow++; // Move to the next row
                                    column--;
                                }


                                // Reset startingRow and increment column for the next group
                                startingRow = newRowCount;
                                column = column + 2;
                            }




                        }

                        else if (parameter.NominalOrScale == "Scale")
                        {

                            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();
                            count = 4; // Count for scale parameters

                            // Insert scale statistics for scale parameters
                            foreach (var groupValue in sortedKeys)
                            {
                                //double groupValue = kvp.Key;
                                Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];
                                //Dictionary<string, string> scaleStats = kvp.Value;

                                foreach (var stat in formattedValues)
                                {
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, stat.Value);
                                    startingRow++; // Move to the next row
                                }

                                // Reset startingRow and increment column for the next group
                                startingRow = newRowCount;

                                column = column + 2;
                            }
                            PerformTest(table, numberofgroups, parameter, startingRow, WordTableColumns - 2, WordTableColumns, issame, ComparativeTables[tableindex]);
                        }

                        // Increment newRowCount for the next parameter
                        newRowCount += count;
                    }

                    if (!(NominalOrScale == "Nominal"))
                    {
                        table.Rows.RemoveAt(1);
                        for (int j = 0; j < WordTableColumns; j++)
                        {
                            table.Rows[0].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[0].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
                        }
                    }



                    List<double> eachgroupcount = GetEachGroupCOUNT(ComparativeTables[tableindex]);

                    //InsertNeachGroup(table, eachgroupcount);
                    InsertNeachGroup(table, eachgroupcount, ComparativeTables[tableindex]);

                    InsertPairwiseF(table, ComparativeTables[tableindex], WordTableColumns);




                    SetMarginBasedOngroups(table, numberofgroups);


                    string groupText = numberofgroups.ToWords();

                    InsertYesonly(table, ComparativeTables[tableindex], WordTableColumns);


                    if (ComparativeTables[tableindex].HasTotalColumn)
                    {
                        AddTotalColumn(table, ComparativeTables[tableindex]);
                    }

                    wordObj.FormatTable(table, 12);

                    ModifyGroupSpaces(table, ComparativeTables[tableindex], numberofgroups);

                    //MessageBox.Show(issame.ToString());
                    if (!issame)
                    {
                        ModifyTestSpaces(table, ComparativeTables[tableindex], WordTableColumns);
                    }


                    InsertDefUnderTableNew(table, section, ComparativeTables[tableindex], groupText);


                    //Designed for maha till now
                    CustomMarginFormat(table);
                }

               



            }

        }

        public void CustomMarginFormat(IWTable table)
        {
            //Designed for maha till now
            if (check_Maha.Checked) 
            {
                wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.09f), wordObj.SetColumnWidthInCentimeters(0.05f));
            }
            
        }
        public void SetMarginBasedOngroups(IWTable table , int numberofgroups)
        {
            if (numberofgroups == 2)
            {
                wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.09f), wordObj.SetColumnWidthInCentimeters(0.09f));
            }
            else if(numberofgroups > 2)
            {
                wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.09f), wordObj.SetColumnWidthInCentimeters(0.09f));
            }
        }
        public void ModifyGroupSpaces(IWTable table , ComparativeTable comparativeTable , int numberofgroups)
        {
            if(!comparativeTable.HasTotalColumn)
            {
                for (int i = 1; i < numberofgroups * 2; i = i + 2)
                {
                    WParagraph groupparagraph = table[0, i].Paragraphs[0];
                    groupparagraph.ParagraphFormat.BeforeSpacing = 0;
                    groupparagraph.ParagraphFormat.AfterSpacing = 0;
                }
            }
            else if(comparativeTable.HasTotalColumn) 
            {
                for (int i = 3; i < numberofgroups * 2; i = i + 2)
                {
                    WParagraph groupparagraph = table[0, i].Paragraphs[0];
                    groupparagraph.ParagraphFormat.BeforeSpacing = 0;
                    groupparagraph.ParagraphFormat.AfterSpacing = 0;
                }
            }
            
           
        }
        public void ModifyTestSpaces(IWTable table , ComparativeTable comparativeTable , int WordTableColumns)
        {
            
            int currentrow = 2;
            int count = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.IsGroup)
                {
                    continue;
                }

                if(parameter.NominalOrScale == "Nominal")
                {
                    //count = parameter.ParameterValues.Distinct().Count() + 1;
                    if(parameter.NominalIsYes)
                    {
                        count = 1;
                    }
                    else
                    {
                        count = parameter.DIC_LablesIfNomainal.Keys.Count() + 1;
                    }
                    
                }
                else if(parameter.NominalOrScale == "Scale")
                {
                    if(parameter.ISFAnovaSig)
                    {
                        count = 5;
                    }
                    else
                    {
                        count = 4;
                    }
                    
                }
                
                if (!comparativeTable.HasTotalColumn) 
                {
                    if (parameter.NominalIsYes)
                    {
                        table[currentrow, WordTableColumns - 2].AddParagraph().AppendText(Text);
                        WParagraph paragraph = table[currentrow, WordTableColumns-2].Paragraphs[0];
                        paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                        table[currentrow, WordTableColumns-2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        paragraph.ParagraphFormat.BeforeSpacing = 0;
                        paragraph.ParagraphFormat.AfterSpacing = 0;
                    }
                    else
                    {
                        WParagraph paragraph = table[currentrow , WordTableColumns - 2].Paragraphs[0];
                        paragraph.ParagraphFormat.BeforeSpacing = 0;
                        paragraph.ParagraphFormat.AfterSpacing = 0;

                        WParagraph paragraph2 = table[currentrow+1, WordTableColumns - 2].Paragraphs[0];
                        paragraph2.ParagraphFormat.BeforeSpacing = 0;
                        paragraph2.ParagraphFormat.AfterSpacing = 0;
                    }
                    
                }

                else if(comparativeTable.HasTotalColumn)
                {
                    WParagraph paragraph = table[currentrow + 1, WordTableColumns].Paragraphs[0];
                    paragraph.ParagraphFormat.BeforeSpacing = 0;
                    paragraph.ParagraphFormat.AfterSpacing = 0;
                }
                
                
                



                currentrow = currentrow + count;
            }
        }

        public void InsertYesonly(IWTable table,ComparativeTable comparativeTable , int WordTableColumns)
        {
            int startrow = 2;
            int count = 0;

            Parameter Last = null;
            foreach (var parameter in comparativeTable.Parameters)
            {
                Last = parameter;
            }


            //int ParameterOnly = 1;
            foreach (var parameter in comparativeTable.Parameters) 
            {
                if(parameter.IsGroup)
                {
                    continue;
                }
                if (parameter.NominalOrScale == "Nominal")
                {
                    if(parameter.NominalIsYes) 
                    {
                        count = 1;
                        //MessageBox.Show(parameter.Name);
                    }
                    else
                    {
                        //count = parameter.ParameterValues.Distinct().Count() + 1;
                        count = parameter.DIC_LablesIfNomainal.Keys.Count() + 1;
                    }
                    
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    if(parameter.ISFAnovaSig)
                    {
                        count = 5;
                    }
                    else
                    {
                        count = 4;
                    }
                    
                }



                if(parameter.NominalOrScale == "Nominal")
                {
                    if(parameter.NominalIsYes)
                    {

                        for (int i = 1; i < WordTableColumns - 2 ; i++)
                        {
                            //table[startrow, i].Paragraphs[0].Text = table[startrow + 2, i].Paragraphs[0].Text;
                            string Text = table[startrow + 2, i].Paragraphs[0].Text;
                            wordObj.Addpara_CenterNoBOLD(table, startrow, i, Text);
                        }

                        string Test = table[startrow + 1, WordTableColumns - 2].Paragraphs[0].Text;
                        //MessageBox.Show(Test);
                        IWParagraph paragraph =  table[startrow, WordTableColumns - 2].AddParagraph();
                        WTextRange textRange = (WTextRange)paragraph.AppendText(Test);
                        textRange.CharacterFormat.HighlightColor = Syncfusion.Drawing.Color.Yellow;
                        paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                        




                        //table[startrow, WordTableColumns - 1].Paragraphs[0].Text = table[startrow + 1, WordTableColumns - 1].Paragraphs[0].Text;
                        string p = table[startrow + 1, WordTableColumns - 1].Paragraphs[0].Text;
                        wordObj.Addpara_CenterNoBOLD(table, startrow, WordTableColumns - 1, p);


                        for (int i = 0; i < WordTableColumns; i++)
                        {
                            table.Rows[startrow].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[startrow].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                        }

                        table.Rows.RemoveAt(startrow+1);
                        table.Rows.RemoveAt(startrow+1);

                        if (Last == parameter)
                        {
                            for (int i = 0; i < WordTableColumns; i++)
                            {
                                table.Rows[startrow].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.ThickThinMediumGap;
                                table.Rows[startrow].Cells[i].CellFormat.Borders.Bottom.LineWidth = 2.25f;
                            }
                        }

                    }


                }


                startrow = startrow + count;

            }


            
        }


        public void InsertDefUnderTableNew(IWTable table,IWSection section , ComparativeTable comparativeTable , string grouptext)
        {
            bool isScaleExist = false;
            

            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.NominalOrScale == "Scale")
                {
                    isScaleExist = true;
                    break;
                }
            }
            if(isScaleExist) 
            {
                IWParagraph firstParagraph = section.AddParagraph();
                WParagraphFormat paragraphFormat = firstParagraph.ParagraphFormat;


                paragraphFormat.BeforeSpacing = 10;
                paragraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
                paragraphFormat.FirstLineIndent = wordObj.SetColumnWidthInCentimeters(-0.5f);
                paragraphFormat.LeftIndent = wordObj.SetColumnWidthInCentimeters(0.5f);

                WTextRange IQR = (WTextRange)firstParagraph.AppendText(DefinitionUnderTable.IQRNotBold);
                WTextRange InterQuar = (WTextRange)firstParagraph.AppendText(DefinitionUnderTable.IQRBold+"\t\t");
                WTextRange SD = (WTextRange)firstParagraph.AppendText(DefinitionUnderTable.SDNot);
                WTextRange StandardDev = (WTextRange)firstParagraph.AppendText(DefinitionUnderTable.SDB);
                
                
                StandardDev.CharacterFormat.Bold = true;
                InterQuar.CharacterFormat.Bold = true;


                foreach (ParagraphItem item in firstParagraph.ChildEntities)
                {
                    if (item is WTextRange)
                    {

                        WTextRange text = item as WTextRange;
                        //Modifies the character format of the text
                        //text.CharacterFormat.Bold = true;
                        text.CharacterFormat.FontName = "Times New Roman";
                        text.CharacterFormat.FontSize = 10;

                    }
                }
            }

            IWParagraph TestsParag = section.AddParagraph();
            WParagraphFormat TestsparagraphFormat = TestsParag.ParagraphFormat;

            if(!isScaleExist)
            {
                TestsparagraphFormat.BeforeSpacing = 10;
            }

            TestsparagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
            TestsparagraphFormat.FirstLineIndent = wordObj.SetColumnWidthInCentimeters(-0.5f);
            TestsparagraphFormat.LeftIndent = wordObj.SetColumnWidthInCentimeters(0.5f);

            
            foreach (string Test in comparativeTable.TestsDone.Distinct())
            {
                if(Test == "Chi")
                {
                    WTextRange Chi = (WTextRange)TestsParag.AppendText(DefinitionUnderTable.chisqaure_test + "\n");
                    Chi.CharacterFormat.Bold = true;
                    
                }
                else if(Test == "tstudent")
                {
                    WTextRange tstudent = (WTextRange)TestsParag.AppendText(DefinitionUnderTable.tstudentTest+"\n");
                    tstudent.CharacterFormat.Bold = true;
                    
                }
                else if(Test == "FAnova")
                {
                    WTextRange Fanova = (WTextRange)TestsParag.AppendText(DefinitionUnderTable.F_OneWay_Anova + "\n");
                    Fanova.CharacterFormat.Bold = true;
                    
                }
                else if (Test == "U")
                {
                    WTextRange U = (WTextRange)TestsParag.AppendText(DefinitionUnderTable.U_Test + "\n");
                    U.CharacterFormat.Bold = true;
                    
                }
                else if (Test == "H")
                {
                    WTextRange H = (WTextRange)TestsParag.AppendText(DefinitionUnderTable.H_Test + "\n");
                    H.CharacterFormat.Bold = true;
                    
                }

                
            }

            string p_comparing = definitionUnderTableObj.p_comparing(grouptext);
            WTextRange p_comparingText = (WTextRange)TestsParag.AppendText(p_comparing);

            if(comparativeTable.hasSig)
            {
                WTextRange p_statSig = (WTextRange)TestsParag.AppendText("\n"+DefinitionUnderTable.p_stat_Sig);
            }
            


            foreach (ParagraphItem item in TestsParag.ChildEntities)
            {
                if (item is WTextRange)
                {

                    WTextRange text = item as WTextRange;
                    text.CharacterFormat.FontName = "Times New Roman";
                    text.CharacterFormat.FontSize = 10;

                }
            }

            TestsParag.ParagraphFormat.PageBreakAfter = true;
        }


        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        
        public static List<string> NominalYes_Compara { get; set; }
        private void pic_ifyes_Click(object sender, EventArgs e)
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

        private void label4_Click(object sender, EventArgs e)
        {

        }

        public static ComparativeTable DoneSortedComparative { get; set; }
        private void button1_Click(object sender, EventArgs e)
        {
            
            if (cmb_TableNames.SelectedIndex != -1 )
            {
                string selectedTableName = cmb_TableNames.SelectedItem.ToString();
                var selectedTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                Sort_Parameters_Frm sort_Parameters_Frm = new Sort_Parameters_Frm();
                sort_Parameters_Frm.SortcomparativeTable = selectedTable;
                sort_Parameters_Frm.TableType = "Comparative";
                sort_Parameters_Frm.Show();
            }
            
        }
        public static ComparativeTable DoneSelectionComparaTable { get; set; }

        public Parameter ParameterSelectName_Compara { get; set; }

        public List<double> SelectedParameterValues_CompaFrm { get; set; } = new List<double>();
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

        private void check_TotalColumn_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void check_Maha_CheckedChanged(object sender, EventArgs e)
        {

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
        public List<double> AllParameter_Select_Values { get; set; } = new List<double>();

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

        public int SelectParameterColIndex { get;set; } = new int();    
        private void pic_AllParaToSelect_Click(object sender, EventArgs e)
        {

            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    list_Select.Items.Add(item);
                }

            }
            //foreach (object selectedItem in list_AllParameters.SelectedItems)
            //{
            //    list_Select.Items.Add(selectedItem.ToString());

            //}


            SelectParameterColIndex = GetSelectParameterCol();

            //MessageBox.Show(SelectParameterColIndex.ToString());
            



        }

        

        

        private void btn_Update_Click(object sender, EventArgs e)
        {
            string filepath = ExcelFunctions.filepath;
            Workbook workbook = new Aspose.Cells.Workbook(filepath);

            // Accessing the first worksheet in the Excel file
            worksheet = workbook.Worksheets[0];

            MessageBox.Show("File Updated");
        }

        private void cmb_ChooseTableFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cmb_ChooseTableFormat.Text == "Default")
            {

                lbl_Nominal.Visible = true;
                list_Nominal.Visible = true;
                pic_RemoveNominalList.Visible = true;
                pic_AllParaToNominal.Visible = true;
                pic_ifyes.Visible = true;
            }
            else if(cmb_ChooseTableFormat.Text == "Periods Groups")
            {
                

                list_Nominal.Visible = false;
                pic_RemoveNominalList.Visible = false;
                pic_AllParaToNominal.Visible=false;
                pic_ifyes.Visible=false;
                lbl_Nominal.Visible = false;



            }
        }

        private void pic_TableFormat_Click(object sender, EventArgs e)
        {
            if(cmb_ChooseTableFormat.SelectedIndex != -1) 
            {
                if(cmb_ChooseTableFormat.Text == "Default")
                {
                    ChooseTable chooseTable = new ChooseTable();
                    chooseTable.Table_ChooseTable = Resource.Default_Comparative;
                    chooseTable.ShowDialog();
                }
            }
        }
            
        private void txt_ParaName_TextChanged(object sender, EventArgs e)
        {

            string filterText = txt_ParaName.Text.ToLower(); // Convert to lowercase for case-insensitive comparison

            // Iterate through all rows in the DataGridView
            foreach (DataGridViewRow row in data_allPara.Rows)
            {
                // Skip the new row placeholder if it's visible
                if (row.IsNewRow) continue;

                // Get the value of the cell in the first column
                string cellValue = row.Cells[0].Value.ToString().ToLower(); // Convert to lowercase for case-insensitive comparison

                // Check if the cell value contains the filter text
                bool shouldShow = cellValue.Contains(filterText);

                // Set the row's visibility
                row.Visible = shouldShow;
            }
        }

        private void data_allPara_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
