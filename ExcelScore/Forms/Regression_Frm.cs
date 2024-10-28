using Aspose.Cells;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
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
using Parameter = ExcelScore.Classes.Parameter;
using Worksheet = Aspose.Cells.Worksheet;

namespace ExcelScore.Forms
{
    public partial class Regression_Frm : Form
    {

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public DataGridView Dgv { get; set; }

        public List<ComparativeTable> ComparativeTables = new List<ComparativeTable>();

        static ExcelFunctions excelFunctionsobj = new ExcelFunctions();

        public List<ComparativeClass> ComparativeClassesList = new List<ComparativeClass>();

        List<string> ColumnExcelheaders = excelFunctionsobj.ReadHeaderColumnsExcel();

        Worksheet worksheet = excelFunctionsobj.GetWorksheet();

        Worksheet Sheet2 = excelFunctionsobj.GetSheet2();
        public Regression_Frm()
        {
            InitializeComponent();

        }

        public void InsertPicture_Type_new()
        {
            // Step 1: Create a dictionary for quick row lookup
            var paraNameToRowMap = new Dictionary<string, DataGridViewRow>();

            // Fill the dictionary with the paraName and corresponding row
            foreach (DataGridViewRow row in data_allPara.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    string paraName = row.Cells[0].Value.ToString();
                    if (!paraNameToRowMap.ContainsKey(paraName))
                    {
                        paraNameToRowMap[paraName] = row;
                    }
                }
            }

            // Step 2: Process the Excel sheet and update the DataGridView rows
            for (int column = 0; column <= Sheet2.Cells.MaxDataColumn; column++)
            {
                string type = "";
                string paraName = "";

                foreach (string columnName in ColumnExcelheaders)
                {
                    paraName = columnName;
                    if (columnName == Sheet2.Cells[0, column].Value.ToString())
                    {
                        if (Sheet2.Cells[1, column].Value != null)
                        {
                            type = Sheet2.Cells[1, column].Value.ToString();
                        }

                        // Update the corresponding DataGridView row if exists
                        if (paraNameToRowMap.TryGetValue(paraName, out var targetRow))
                        {
                            // Insert the image based on the type
                            if (type == "Nominal")
                            {
                                targetRow.Cells["ColMeasure"].Value = Resource.Final_Nominal_Color;
                            }
                            else if (type == "Scale")
                            {
                                targetRow.Cells["ColMeasure"].Value = Resource.FinalScale_Color;
                            }
                        }

                        // No need to continue with other column names once a match is found
                        break;
                    }
                }
            }
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

            InsertPicture_Type_new();
            foreach (DataGridViewColumn column in data_allPara.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }


        }
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
        private void Regression_Frm_Load(object sender, EventArgs e)
        {
            AddHeadersToParameter();
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
            ChooseFrm chooseFrm = new ChooseFrm();
            chooseFrm.Dgv = Dgv;
            chooseFrm.Show();
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
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
        public static ComparativeTable DoneSortedComparative { get; set; }
        private void btn_SortTable_Click(object sender, EventArgs e)
        {
            if (cmb_TableNames.SelectedIndex != -1)
            {
                string selectedTableName = cmb_TableNames.SelectedItem.ToString();
                var selectedTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                Sort_Parameters_Frm sort_Parameters_Frm = new Sort_Parameters_Frm();
                sort_Parameters_Frm.SortcomparativeTable = selectedTable;
                sort_Parameters_Frm.TableType = "Regression";
                sort_Parameters_Frm.Show();
            }
        }

        private void btn_Update_Click(object sender, EventArgs e)
        {
            string filepath = ExcelFunctions.filepath;
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);

            // Accessing the first worksheet in the Excel file
            worksheet = workbook.Worksheets[0];
            Sheet2 = workbook.Worksheets[1];

            MessageBox.Show("File Updated");
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
                    ClearTable(tableToRemove);
                }
            }

            cmb_TableNames.Text = "";
            list_ViewTableParameters.Items.Clear();
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
                    list_Dependent.Items.Add(item);
                }

            }
        }

        private void pic_RemoveNominalList_Click(object sender, EventArgs e)
        {
            if (list_Dependent.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_Dependent.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_Dependent.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_ClearNominalList_Click(object sender, EventArgs e)
        {
            var selectedItemsNominal = new List<object>();
            foreach (var selectedItemNominal in list_Dependent.Items)
            {
                selectedItemsNominal.Add(selectedItemNominal);
            }


            foreach (var selectedItemNominal in selectedItemsNominal)
            {
                list_Dependent.Items.Remove(selectedItemNominal);
            }
        }

        private void pic_AllLists_Click(object sender, EventArgs e)
        {

            var selectedItemsNominal = new List<object>();
            foreach (var selectedItemNominal in list_Dependent.Items)
            {
                selectedItemsNominal.Add(selectedItemNominal);
            }


            foreach (var selectedItemNominal in selectedItemsNominal)
            {
                list_Dependent.Items.Remove(selectedItemNominal);
            }




            var selectedItemsNormal = new List<object>();
            foreach (var selectedItemNormal in list_NotSeperatedScale.Items)
            {
                selectedItemsNormal.Add(selectedItemNormal);
            }


            foreach (var selectedItemNormal in selectedItemsNormal)
            {
                list_NotSeperatedScale.Items.Remove(selectedItemNormal);
            }





            var selectedItemsNomianlNotSeperated = new List<object>();
            foreach (var selectedItemAbnormal in list_NotSeperatedNominal.Items)
            {
                selectedItemsNomianlNotSeperated.Add(selectedItemAbnormal);
            }


            foreach (var selectedItemAbnormal in selectedItemsNomianlNotSeperated)
            {
                list_NotSeperatedNominal.Items.Remove(selectedItemAbnormal);
            }




            var selectedItemsAbnormal = new List<object>();
            foreach (var selectedItemAbnormal in list_Seperated.Items)
            {
                selectedItemsAbnormal.Add(selectedItemAbnormal);
            }


            foreach (var selectedItemAbnormal in selectedItemsAbnormal)
            {
                list_Seperated.Items.Remove(selectedItemAbnormal);
            }
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
                    list_NotSeperatedScale.Items.Add(item);
                }

            }
        }

        private void pic_RemoveNormalList_Click(object sender, EventArgs e)
        {
            if (list_NotSeperatedScale.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_NotSeperatedScale.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_NotSeperatedScale.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_clearNormalList_Click(object sender, EventArgs e)
        {
            var selectedItemsNormal = new List<object>();
            foreach (var selectedItemNormal in list_NotSeperatedScale.Items)
            {
                selectedItemsNormal.Add(selectedItemNormal);
            }


            foreach (var selectedItemNormal in selectedItemsNormal)
            {
                list_NotSeperatedScale.Items.Remove(selectedItemNormal);
            }
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
                    list_Seperated.Items.Add(item);
                }

            }
        }

        private void pic_RemoveAbNormalList_Click(object sender, EventArgs e)
        {
            if (list_Seperated.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_Seperated.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_Seperated.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_ClearAbnormalList_Click(object sender, EventArgs e)
        {
            var selectedItemsAbnormal = new List<object>();
            foreach (var selectedItemAbnormal in list_Seperated.Items)
            {
                selectedItemsAbnormal.Add(selectedItemAbnormal);
            }


            foreach (var selectedItemAbnormal in selectedItemsAbnormal)
            {
                list_Seperated.Items.Remove(selectedItemAbnormal);
            }
        }
        public string AddTableUI()
        {
            string TableName = null;
            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&
                 !string.IsNullOrWhiteSpace(cmb_ChooseTableFormat.Text) && list_Dependent.Items.Count > 0 &&
                (list_NotSeperatedScale.Items.Count > 0 || list_Seperated.Items.Count > 0 || list_NotSeperatedNominal.Items.Count >0))
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

        public void AddTableClass(string tableName)
        {

            var comparativeTable = new ComparativeTable
            {
                TableName = tableName,
                Parameters = new List<Parameter>(),


            };



            // Add parameters from list_Nominal
            foreach (var item in list_Dependent.Items)
            {
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    NominalOrScale = "Dependent",
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }




            // Add parameters from list_NormalScale
            foreach (var item in list_NotSeperatedScale.Items)
            {
                var parameter = new Parameter
                {

                    Name = item.ToString(),
                    NominalOrScale = "Scale",
                    NormalOrAbnormal = "Not Seperated",
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };


                comparativeTable.Parameters.Add(parameter);
            }

            foreach (var item in list_NotSeperatedNominal.Items)
            {
                var parameter = new Parameter
                {

                    Name = item.ToString(),
                    NominalOrScale = "Nominal",
                    NormalOrAbnormal = "Not Seperated",
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };


                comparativeTable.Parameters.Add(parameter);
            }




            // Add parameters from list_AbnormalScale
            foreach (var item in list_Seperated.Items)
            {
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    NominalOrScale = "Nominal",
                    NormalOrAbnormal = "Seperated",
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }



            comparativeTable.FormatType = cmb_ChooseTableFormat.Text;

            ComparativeTables.Add(comparativeTable);




        }
        public void CheckFullEmptyParameters(string TableName)
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                if (TableName == table.TableName)
                {
                    foreach (var parameter in table.Parameters)
                    {
                        if (parameter.ParameterValues.Count == 0)
                        {
                            MessageBox.Show("Parameter : " + parameter.Name + " is Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                }
            }
        }

        public void CheckForOthers_inNominal(string TableName)
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                if (TableName == table.TableName)
                {
                    foreach (var parameter in table.Parameters)
                    {
                        List<double> valuesErrored = new List<double>();
                        if (parameter.NominalOrScale == "Nominal")
                        {
                            List<double> keys = new List<double>();
                            foreach (var kvp in parameter.DIC_LablesIfNomainal)
                            {
                                double key = kvp.Key;
                                keys.Add(key);
                            }

                            foreach (var value in parameter.ParameterValues)
                            {
                                if (!keys.Contains(value))
                                {
                                    valuesErrored.Add(value);
                                    
                                }

                            }

                            if(valuesErrored.Count > 7 )
                            {
                                MessageBox.Show("Did you enter a Scale Parameter in Nominal ? : " + parameter.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                foreach (var Errorvalue in valuesErrored)
                                {
                                    MessageBox.Show("Value " + Errorvalue + " doesn't exist at Parameter : " + parameter.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                
                            }
                        }
                    }
                }
            }
        }

        public void GetDataValues(string TableName)
        {
            foreach (var table in ComparativeTables)
            {

                if (TableName == table.TableName)
                {
                    foreach (var parameter in table.Parameters)
                    {
                        for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
                        {
                            object cellValue = worksheet.Cells[0, col].Value;
                            string parameterName = cellValue?.ToString();

                            if (!string.IsNullOrEmpty(parameterName) && parameter.Name == parameterName)
                            {

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
                                   


                                    // Parse the cell data as double
                                    object SelectData = worksheet.Cells[row, SelectParameterColIndex].Value;

                                    double SelectDataDouble;


                                    if (list_Select.Items.Count == 0)
                                    {

                                        if (cellData.ToString() == ".")
                                        {
                                            parameter.hasLowerN = true;
                                            parameter.ParameterValues.Add(-999);
                                            continue;
                                        }


                                        if (cellData != null && double.TryParse(cellData.ToString(), out cellValueDouble))
                                        {
                                            parameter.ParameterValues.Add(cellValueDouble);

                                        }
                                    }
                                    else
                                    {
                                        if (SelectData != null && double.TryParse(SelectData.ToString(), out SelectDataDouble))
                                        {
                                            if (SelectedParameterValues_CompaFrm.Contains(SelectDataDouble))
                                            {
                                                if (cellData.ToString() == ".")
                                                {
                                                    parameter.hasLowerN = true;
                                                    parameter.ParameterValues.Add(-999);
                                                    continue;
                                                }


                                                if (cellData != null && double.TryParse(cellData.ToString(), out cellValueDouble))
                                                {

                                                    parameter.ParameterValues.Add(cellValueDouble);
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
        }
        public int SelectParameterColIndex { get; set; } = new int();
        private void pic_addTable_Click(object sender, EventArgs e)
        {
            string TableName = AddTableUI();


            GetDataValues(TableName);

            foreach (ComparativeTable table in ComparativeTables)
            {

                if (TableName == table.TableName)
                {
                    
                    excelFunctionsobj.ReadLablesIfNominal(table);
                }
            }




            CheckFullEmptyParameters(TableName);
            CheckForOthers_inNominal(TableName);
        }

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


            SelectParameterColIndex = GetSelectParameterCol();
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

        public List<double> AllParameter_Select_Values { get; set; } = new List<double>();

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
        WordClass wordObj = new WordClass();

        public int CountRows(ComparativeTable comparativeTable)
        {
            int TotalRows = 0;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.NominalOrScale == "Dependent")
                {
                    continue;
                }

                else if(parameter.NormalOrAbnormal == "Not Seperated")
                {
                    TotalRows++;
                }
                else if(parameter.NormalOrAbnormal == "Seperated")
                {
                    
                    int count = parameter.DIC_LablesIfNomainal.Keys.Count+1;
                    TotalRows += count;
                }

               


            }





            return TotalRows;
        }
        ManualTests manual = new ManualTests();
        ColorClass colorClass = new ColorClass();
        Syncfusion.Drawing.Color lighterOrange;

        public void LinearRegression_Layout()
        {
            lighterOrange = colorClass.LightOrange();

            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Linear")
                {

                    IWSection section = wordObj.CreatePortraitSection();

                    wordObj.AddRgressionTitle(section, ComparativeTables[tableindex].TableName, "linear");

                    int Variablerows = CountRows(ComparativeTables[tableindex]);


                    int WordTableColumns = 5;

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);


                    wordObj.GeneralTableFormat(table);


                    //Merges
                    wordObj.ApplyRegression_OuterMerges(table, WordTableRows, WordTableColumns);

                    //Borders
                    wordObj.ApplyRegression_OuterBorders(table, WordTableRows, WordTableColumns);

                    //Widths
                    wordObj.ApplyRegression_Widths(table, WordTableRows, WordTableColumns);


                    //Outer Headers
                    wordObj.Apply_Linear_Regression_OuterHeaders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, "B");



                    wordObj.InsertRegression_InnerHeader_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);


                    Parameter DependentParameter = GetDependentParameter(ComparativeTables[tableindex]);



                    int StartingRow = 2;
                    foreach (var parameter in ComparativeTables[tableindex].Parameters)
                    {
                        List<double> ResultIndependent = new List<double>();
                        List<double> ResultDependent = new List<double>();
                        if (parameter.NominalOrScale == "Dependent")
                        {
                            continue;
                        }

                        else if (parameter.NominalOrScale == "Scale")
                        {
                            (ResultIndependent, ResultDependent) = ReturnTrueParameterValues(parameter, DependentParameter);

                            List<double> RegressionResult = new List<double>();
                            RegressionResult = manual.LinearRegressionn(ResultIndependent.ToArray(), ResultDependent.ToArray());


                            InsertDataRegression(table, RegressionResult, StartingRow);


                            StartingRow++;


                        }
                        //NominalOrScale = "Nominal",
                        //NormalOrAbnormal = "Not Seperated",
                        else if (parameter.NominalOrScale == "Nominal")
                        {

                            if (parameter.NormalOrAbnormal == "Not Seperated")
                            {
                                (ResultIndependent, ResultDependent) = ReturnTrueParameterValues(parameter, DependentParameter);

                                List<double> RegressionResult = new List<double>();
                                RegressionResult = manual.LinearRegressionn(ResultIndependent.ToArray(), ResultDependent.ToArray());


                                InsertDataRegression(table, RegressionResult, StartingRow);


                                StartingRow++;
                            }
                            else if (parameter.NormalOrAbnormal == "Seperated")
                            {
                                StartingRow++;
                                (ResultIndependent, ResultDependent) = ReturnTrueParameterValues(parameter, DependentParameter);

                                var SortedIndependentKeys = ResultIndependent.Distinct().OrderBy(key => key).ToList();

                                var SortedGroupDIC = parameter.DIC_LablesIfNomainal.Keys.OrderBy(key => key).ToList();

                                foreach (var item in SortedGroupDIC)
                                {
                                    if (SortedIndependentKeys.Contains(item))
                                    {
                                        List<double> ResultIndependent_Seperated = new List<double>();
                                        foreach (var Independentvalue in ResultIndependent)
                                        {
                                            if (Independentvalue == item)
                                            {
                                                ResultIndependent_Seperated.Add(1);
                                            }
                                            else
                                            {
                                                ResultIndependent_Seperated.Add(0);
                                            }
                                        }
                                        List<double> RegressionResult = new List<double>();
                                        RegressionResult = manual.LinearRegressionn(ResultIndependent_Seperated.ToArray(), ResultDependent.ToArray());

                                        InsertDataRegression(table, RegressionResult, StartingRow);
                                        StartingRow++;
                                    }
                                    else if (!SortedIndependentKeys.Contains(item))
                                    {
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 1, "–");
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, "–");
                                        StartingRow++;
                                    }
                                }

                            }

                        }
                    }


                    wordObj.LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    wordObj.FormatTableCustom(table, 9.5f, 0, 0);




                }
            }
        }
        public void LogisticRegression_Layout()
        {
            lighterOrange = colorClass.LightOrange();

            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Logistic")
                {

                    IWSection section = wordObj.CreatePortraitSection();

                    wordObj.AddRgressionTitle(section, ComparativeTables[tableindex].TableName , "logistic");

                    int Variablerows = CountRows(ComparativeTables[tableindex]);


                    int WordTableColumns = 5;

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);


                    wordObj.GeneralTableFormat(table);


                    //Merges
                    wordObj.ApplyRegression_OuterMerges(table, WordTableRows, WordTableColumns);

                    //Borders
                    wordObj.ApplyRegression_OuterBorders(table, WordTableRows, WordTableColumns);

                    //Widths
                    wordObj.ApplyRegression_Widths(table, WordTableRows, WordTableColumns);


                    //Outer Headers
                    wordObj.Apply_Linear_Regression_OuterHeaders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns , "OR");



                    wordObj.InsertRegression_InnerHeader_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);


                    Parameter DependentParameter = GetDependentParameter(ComparativeTables[tableindex]);



                    int StartingRow = 2;
                    foreach (var parameter in ComparativeTables[tableindex].Parameters)
                    {
                        List<double> ResultIndependent = new List<double>();
                        List<double> ResultDependent = new List<double>();
                        if (parameter.NominalOrScale == "Dependent")
                        {
                            continue;
                        }

                        else if(parameter.NominalOrScale == "Scale")
                        {
                            (ResultIndependent, ResultDependent) = ReturnTrueParameterValues(parameter, DependentParameter);

                            List<double> RegressionResult = new List<double>();
                            RegressionResult = manual.LogisticRegression(ResultIndependent.ToArray(), ResultDependent.ToArray());
                            

                            InsertDataRegression(table, RegressionResult, StartingRow);


                            StartingRow++;


                        }
                        //NominalOrScale = "Nominal",
                    //NormalOrAbnormal = "Not Seperated",
                        else if(parameter.NominalOrScale == "Nominal")
                        {

                            if(parameter.NormalOrAbnormal == "Not Seperated")
                            {
                                (ResultIndependent, ResultDependent) = ReturnTrueParameterValues(parameter, DependentParameter);

                                List<double> RegressionResult = new List<double>();
                                RegressionResult = manual.LogisticRegression(ResultIndependent.ToArray(), ResultDependent.ToArray());


                                InsertDataRegression(table, RegressionResult, StartingRow);


                                StartingRow++;
                            }
                            else if (parameter.NormalOrAbnormal == "Seperated")
                            {
                                StartingRow++;
                                (ResultIndependent, ResultDependent) = ReturnTrueParameterValues(parameter, DependentParameter);

                                var SortedIndependentKeys = ResultIndependent.Distinct().OrderBy(key => key).ToList();

                                var SortedGroupDIC = parameter.DIC_LablesIfNomainal.Keys.OrderBy(key => key).ToList();

                                foreach (var item in SortedGroupDIC)
                                {
                                    if(SortedIndependentKeys.Contains(item))
                                    {
                                        List<double> ResultIndependent_Seperated = new List<double>();
                                        foreach (var Independentvalue in ResultIndependent)
                                        {
                                            if(Independentvalue == item)
                                            {
                                                ResultIndependent_Seperated.Add(1);
                                            }
                                            else
                                            {
                                                ResultIndependent_Seperated.Add(0);
                                            }
                                        }
                                        List<double> RegressionResult = new List<double>();
                                        RegressionResult = manual.LogisticRegression(ResultIndependent_Seperated.ToArray(), ResultDependent.ToArray());

                                        InsertDataRegression(table, RegressionResult, StartingRow);
                                        StartingRow++;
                                    }
                                    else if(!SortedIndependentKeys.Contains(item))
                                    {
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 1, "–");
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, "–");
                                        StartingRow++;
                                    }
                                }

                            }

                        }
                    }


                    wordObj.LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    wordObj.FormatTableCustom(table, 9.5f, 0, 0);
                    



                }
            }
        }


        public void InsertDataRegression(IWTable table ,List<double> RegressionResult , int StartingRow)
        {
            string pValueString = RegressionResult[3] < 0.001 ? "<0.001" : RegressionResult[3].ToString("0.000");
            string B = RegressionResult[0].ToString("0.000");
            string CI_L = RegressionResult[1].ToString("0.000");
            string CI_U = RegressionResult[2].ToString("0.000");



            wordObj.Addpara_CenterNoBOLD(table, StartingRow, 1, pValueString);

            string B_Ci = $"{B} ({CI_L} – {CI_U})";
            wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, B_Ci);

            bool PSig = generalFunctions.PvalueHasSig(pValueString);

            if (PSig)
            {
                wordObj.SubSuperScriptText(table, StartingRow, 1, Syncfusion.Drawing.Color.Empty, "*", "Super");
                table.Rows[StartingRow].Cells[1].CellFormat.BackColor = lighterOrange;
                table.Rows[StartingRow].Cells[2].CellFormat.BackColor = lighterOrange;
            }
        }

        GeneralFunctions generalFunctions = new GeneralFunctions();
        public Parameter GetDependentParameter(ComparativeTable comparativeTable)
        {
            Parameter DependentParameter = null;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.NominalOrScale == "Dependent")
                {
                    DependentParameter = parameter;
                    break;
                }
            }

            return DependentParameter;  


        }

        public (List<double>, List<double>) ReturnTrueParameterValues(Parameter Independent , Parameter Dependent)
        {
            List<double> TrueIndependentParaValues = new List<double>();
            List<double> TrueDependentParaValues = new List<double>();
            for (int i = 0;i< Independent.ParameterValues.Count;i++)
            {
                double IndependentValue = Independent.ParameterValues[i];
                double DependentValue = Dependent.ParameterValues[i];

                if((IndependentValue != -999) && (DependentValue!= -999))
                {
                    TrueIndependentParaValues.Add(IndependentValue);
                    TrueDependentParaValues.Add(DependentValue);
                }



            }

            return (TrueIndependentParaValues, TrueDependentParaValues);
        }


        WordDocument document;
        private void btn_Done_Click(object sender, EventArgs e)
        {
            document = wordObj.InitWord();

            LinearRegression_Layout();
            LogisticRegression_Layout();

            string filepath = wordObj.SaveWord();

            wordObj.removeHeader(filepath);

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    list_NotSeperatedNominal.Items.Add(item);
                }

            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (list_NotSeperatedNominal.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_NotSeperatedNominal.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_NotSeperatedNominal.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_clearNotSeperatedNominal_Click(object sender, EventArgs e)
        {
            var selectedItemsNormal = new List<object>();
            foreach (var selectedItemNormal in list_NotSeperatedNominal.Items)
            {
                selectedItemsNormal.Add(selectedItemNormal);
            }


            foreach (var selectedItemNormal in selectedItemsNormal)
            {
                list_NotSeperatedNominal.Items.Remove(selectedItemNormal);
            }
        }
    }
}
