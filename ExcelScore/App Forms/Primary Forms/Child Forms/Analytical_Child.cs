using ExcelScore.App_UI;
using ExcelScore.Classes;
using ExcelScore.Forms;
using ExcelScore.FormsDesigns;
using ExcelScore.Properties;
using ExcelScore.StatClasses;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    public partial class Analytical_Child : Form
    {
        public Analytical_Child()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }

        public List<StatParameter> SpssParameters;
        List<StatTable> StatTables = new List<StatTable>();

        private void Analytical_Child_Load(object sender, EventArgs e)
        {
            SpssParameters = FormDataTransfer.Get<List<StatParameter>>("SPSS_Parameters");
            PopDataGrid_AllSpssParam();
        }

        private void btn_ChooseFormat_Click(object sender, EventArgs e)
        {
            FormManager.ShowStandaloneForm<ChooseAnalyticFormat_Child>();
        }

        private void list_Nominal_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        public void PopDataGrid_AllSpssParam()
        {
            //Load Parameters from SPSS
            data_allPara.Rows.Clear();
            foreach (var spssparam in SpssParameters)
            {
                data_allPara.Rows.Add(spssparam.Name);
            }

            //Make Sorting Off 
            foreach (DataGridViewColumn column in data_allPara.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            //Scroll Wheel and disabled scroll bars
            data_allPara.MouseWheel += myDataGridView_MouseWheel;

            //Making Col Resizable
            data_allPara.AllowUserToResizeColumns = false; 
            data_allPara.Columns[0].Resizable = DataGridViewTriState.True;

            //Insert Picture Type
            InsertPicture_Type_new();
        }


        private void myDataGridView_MouseWheel(object sender, MouseEventArgs e)
        {
            int currentIndex = data_allPara.FirstDisplayedScrollingRowIndex;
            int scrollLines = SystemInformation.MouseWheelScrollLines;

            if (e.Delta > 0) // Scroll up
            {
                int newIndex = Math.Max(0, currentIndex - scrollLines);
                data_allPara.FirstDisplayedScrollingRowIndex = newIndex;
            }
            else if (e.Delta < 0) // Scroll down
            {
                int newIndex = Math.Min(data_allPara.RowCount - 1, currentIndex + scrollLines);
                data_allPara.FirstDisplayedScrollingRowIndex = newIndex;
            }
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
            foreach(var spssparam in SpssParameters)
            {
                
                string paraName = spssparam.Name;

                if (paraNameToRowMap.TryGetValue(paraName, out var targetRow))
                {
                    // Insert the image based on the type
                    if (spssparam.Type == "Nominal")
                    {
                        targetRow.Cells["ColMeasure"].Value = Resource.Final_Nominal_Color;
                    }
                    else if (spssparam.Type == "Scale")
                    {
                        targetRow.Cells["ColMeasure"].Value = Resource.FinalScale_Color;
                    }
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

        //Adding To Lists from datagrid
        public List<string> OrderedParameters = new List<string>();
        public void AddTo_Lists_From_DataGrid(ListBox AddedListbox)
        {
            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    AddedListbox.Items.Add(item);
                    OrderedParameters.Add(item);
                }

            }
        }

        //Removing Single Items From DataGrid
        public void Remove_Single_List(ListBox RemovedListbox)
        {
            if (RemovedListbox.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in RemovedListbox.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    RemovedListbox.Items.Remove(selectedItem);
                }


                foreach (var item in selectedItems)
                {
                    OrderedParameters.Remove(item.ToString());
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }
        public void Clear_List(ListBox ListCleared)
        {
            var selectedItems = new List<object>();
            foreach (var selectedItem in ListCleared.Items)
            {
                selectedItems.Add(selectedItem);
            }


            foreach (var selectedItem in selectedItems)
            {
                ListCleared.Items.Remove(selectedItem);
            }

            foreach (var item in selectedItems)
            {
                OrderedParameters.Remove(item.ToString());
            }
        }
        private void btn_NameLabel_Click(object sender, EventArgs e)
        {
            
        }

        private void pic_previous_Click(object sender, EventArgs e)
        {
            SelectionMainForm selectionMainForm = new SelectionMainForm();
            selectionMainForm.Show();
            this.Hide();
        }

        private SelectFrm_Child selectFrm_Child;
        private void btn_Select_Click(object sender, EventArgs e)
        {
            if (selectFrm_Child == null || selectFrm_Child.IsDisposed)
            {
                selectFrm_Child = new SelectFrm_Child();
            }

            selectFrm_Child.Show();
            
        }

        private void pic_AllParaToNominal_Click(object sender, EventArgs e)
        {
            AddTo_Lists_From_DataGrid(list_Nominal);
        }

        private void pic_AllParaToNormal_Click(object sender, EventArgs e)
        {
            AddTo_Lists_From_DataGrid(list_NormalScale);
        }

        private void pic_AllParaToAbnormal_Click(object sender, EventArgs e)
        {
            AddTo_Lists_From_DataGrid(list_AbnormalScale);
        }

        private void pic_AllParaToGroups_Click(object sender, EventArgs e)
        {
            //We didn't use add to lists because it adds to ordered parameters
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
        }

        private void pic_RemoveNominalList_Click(object sender, EventArgs e)
        {
            Remove_Single_List(list_Nominal);
        }

        private void pic_RemoveNormalList_Click(object sender, EventArgs e)
        {
            Remove_Single_List(list_NormalScale);
        }

        private void pic_RemoveAbnormalList_Click(object sender, EventArgs e)
        {
            Remove_Single_List(list_AbnormalScale);
        }

        private void pic_RemoveGroupsList_Click(object sender, EventArgs e)
        {
            Remove_Single_List(list_Groups);
        }

        private void pic_ClearNominalList_Click(object sender, EventArgs e)
        {
            Clear_List(list_Nominal);
        }

        private void pic_ClearNormalList_Click(object sender, EventArgs e)
        {
            Clear_List(list_NormalScale);
        }

        private void pic_ClearAbnormalList_Click(object sender, EventArgs e)
        {
            Clear_List(list_AbnormalScale);
        }

        private void pic_ClearGroupsList_Click(object sender, EventArgs e)
        {
            Clear_List(list_Groups);
        }
        public void AddTableClassNew(string tableName, string tableType)
        {
            StatClasses.StatTable table;

            // --- Dynamically create the appropriate table type ---
            switch (tableType.ToLower())
            {
                case "comparative":
                    table = new ComparativeStatTable();
                    break;
                case "descriptive":
                    table = new DescriptiveStatTable();
                    break;
                case "relation":
                    table = new RelationStatTable();
                    break;
                default:
                    MessageBox.Show($"Unknown table type: {tableType}");
                    return;
            }

            table.TableName = tableName;
            table.TableDesignType = FormDataTransfer.Get<string>("TableDesignType");

            // --- Add Group Parameters ---
            foreach (var item in list_Groups.Items)
            {
                table.Parameters.Add(new StatParameter
                {
                    Name = item.ToString(),
                    IsGroup = true,
                    Type = "Nominal",
                    RawValues = new List<string>()
                });
            }
            

            // --- Add Main Parameters ---
            foreach (var paramName in OrderedParameters)
            {
                if (table.Parameters.Any(p => p.Name == paramName))
                    continue;

                StatParameter parameter = null;

                if (list_Nominal.Items.Contains(paramName))
                {
                    parameter = new StatParameter
                    {
                        Name = paramName,
                        Type = "Nominal",
                        RawValues = new List<string>()
                    };

                }
                else if (list_NormalScale.Items.Contains(paramName))
                {
                    parameter = new StatParameter
                    {
                        Name = paramName,
                        Type = "Scale",
                        Normality = "Normal",
                        RawValues = new List<string>()
                    };
                }
                else if (list_AbnormalScale.Items.Contains(paramName))
                {
                    parameter = new StatParameter
                    {
                        Name = paramName,
                        Type = "Scale",
                        Normality = "Abnormal",
                        RawValues = new List<string>()
                    };
                }

                if (parameter != null)
                    table.Parameters.Add(parameter);
            }

            // --- Add the new table to your master list ---
            StatTables.Add(table);
        }
        public string AddTableUINew()
        {
            string tableName = null;

            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&
                list_Groups.Items.Count > 0 &&
                (list_Nominal.Items.Count > 0 || list_NormalScale.Items.Count > 0 || list_AbnormalScale.Items.Count > 0))
            {
                // Check if the table name already exists
                bool tableExists = cmb_TableNames.Items
                    .Cast<object>()
                    .Any(existing => string.Equals(existing.ToString(), txt_TableName.Text, StringComparison.OrdinalIgnoreCase));

                if (!tableExists)
                {
                    // Get table type from transfer storage
                    var tableType = FormDataTransfer.Get<string>("TableType");

                    if (string.IsNullOrEmpty(tableType))
                    {
                        MessageBox.Show("Table type was not set. Cannot add table.");
                        return null;
                    }


                    // Create and add table
                    AddTableClassNew(txt_TableName.Text, tableType);
                    cmb_TableNames.Items.Add(txt_TableName.Text);
                    tableName = txt_TableName.Text;
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

            return tableName;
        }
        public void StatBasic(string tableName)
        {
            // Step 1: Find the table (any type)
            var table = StatTables.FirstOrDefault(t => t.TableName == tableName);
            if (table == null)
            {
                return;
            }

            // Step 2: Load SPSS parameters from memory
            var spssParams = SpssParameters;
            if (spssParams == null || spssParams.Count == 0)
            {
                return;
            }

            // Step 3: Match parameters and assign raw values, labels, etc.
            foreach (var param in table.Parameters)
            {
                var match = spssParams.FirstOrDefault(p => p.Name == param.Name);
                if (match != null)
                {
                    param.RawValues = new List<string>(match.RawValues);
                    param.ValueLabels = new Dictionary<int, string>(match.ValueLabels);
                    param.Type = match.Type;
                    //param.Normality = match.Normality;
                }
            }

            // Step 4: Group the data inside the table
            table.AssignGroupedValuesToAll();

            MessageBox.Show($"Table '{tableName}' filled successfully.");
        }
        private void pic_AddStatTable_Click(object sender, EventArgs e)
        {
            var designType = FormDataTransfer.Get<string>("TableDesignType");
            if (string.IsNullOrEmpty(designType))
            {
                MessageBox.Show("Please Choose Design Type!");
                return;
            }
            string tableName = AddTableUINew();

            if (!string.IsNullOrEmpty(tableName))
            {
                StatBasic(tableName); // Only call if table creation succeeded

                //Assign Select Statement
                var table = StatTables.FirstOrDefault(t => t.TableName == tableName);
                string SelectComm = FormDataTransfer.Get<string>("IfSelectCommand");

                table.SelectStatement = SelectComm;

            }
            
        }
        


        private void btn_Done_Click(object sender, EventArgs e)
        {
            string outputText = SpssSyntaxStat.RunUnifiedSyntaxAndGetResult(StatTables, out _);
            SpssSyntaxStat.ParseUnifiedOutput_Crosstabs(outputText, StatTables);

            GeneralStatFunctions.ParseUnifiedOutput_Descriptives(StatTables);


            foreach (var item in StatTables)
            {
                var context = new WordTableDesignContext(item);
                WordTableStatDesign.Execute(context);
            }
            

            MessageBox.Show(".");


        }

        

        private void pic_removeTableSelected_Click(object sender, EventArgs e)
        {
            string selectedTableName = cmb_TableNames.SelectedItem?.ToString();

            if (!string.IsNullOrEmpty(selectedTableName))
            {
                // Remove the selected table name from the ComboBox
                cmb_TableNames.Items.RemoveAt(cmb_TableNames.SelectedIndex);

                // Find and remove the corresponding ComparativeTable instance from the ComparativeTables list
                var tableToRemove = StatTables.FirstOrDefault(table => table.TableName == selectedTableName);
                if (tableToRemove != null)
                {
                    StatTables.Remove(tableToRemove);
                }
            }

            cmb_TableNames.Text = "";
            list_ViewTableParameters.Items.Clear();
        }

        private void cmb_TableNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedTableName = cmb_TableNames.SelectedItem.ToString();

            // Find the ComparativeTable instance corresponding to the selected table name
            var selectedTable = StatTables.FirstOrDefault(table => table.TableName == selectedTableName);

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
                    else
                    {
                        list_ViewTableParameters.Items.Add(parameter.Name);
                    }

                }
            }

        }

        private void btn_SortTable_Click(object sender, EventArgs e)
        {
            if (cmb_TableNames.SelectedIndex != -1)
            {
                string selectedTableName = cmb_TableNames.SelectedItem.ToString();
                var selectedTable = StatTables.FirstOrDefault(table => table.TableName == selectedTableName);
                FormDataTransfer.Set("SortStatTable", selectedTable);

                SortFrm_Child sortFrm_Child = new SortFrm_Child();
                sortFrm_Child.Show();
            }

            

        }

        private void pic_ClearAllLists_Click(object sender, EventArgs e)
        {
            var selectedItemsNominal = new List<object>();
            foreach (var selectedItemNominal in list_Nominal.Items)
            {
                selectedItemsNominal.Add(selectedItemNominal);
            }


            foreach (var selectedItemNominal in selectedItemsNominal)
            {
                list_Nominal.Items.Remove(selectedItemNominal);
            }




            var selectedItemsNormal = new List<object>();
            foreach (var selectedItemNormal in list_NormalScale.Items)
            {
                selectedItemsNormal.Add(selectedItemNormal);
            }


            foreach (var selectedItemNormal in selectedItemsNormal)
            {
                list_NormalScale.Items.Remove(selectedItemNormal);
            }





            var selectedItemsAbnormal = new List<object>();
            foreach (var selectedItemAbnormal in list_AbnormalScale.Items)
            {
                selectedItemsAbnormal.Add(selectedItemAbnormal);
            }


            foreach (var selectedItemAbnormal in selectedItemsAbnormal)
            {
                list_AbnormalScale.Items.Remove(selectedItemAbnormal);
            }



            var itemsToRemove = new List<object>();
            itemsToRemove.AddRange(selectedItemsNominal);
            itemsToRemove.AddRange(selectedItemsNormal);
            itemsToRemove.AddRange(selectedItemsAbnormal);


            foreach (var item in itemsToRemove)
            {
                OrderedParameters.Remove(item.ToString());
            }
        }
        
        private void btn_Load_Click(object sender, EventArgs e)
        {
            if (cmb_TableNames.SelectedIndex != -1)
            {
                App_UI.Custom_UI_Functions.ClearListBox(list_Nominal);
                App_UI.Custom_UI_Functions.ClearListBox(list_NormalScale);
                App_UI.Custom_UI_Functions.ClearListBox(list_AbnormalScale);
                App_UI.Custom_UI_Functions.ClearListBox(list_Groups);




                string selectedTableName = cmb_TableNames.SelectedItem.ToString();

                var comparativeTable = StatTables.FirstOrDefault(table => table.TableName == selectedTableName);

                foreach (var parameter in comparativeTable.Parameters)
                {
                    if (parameter.IsGroup)
                    {
                        list_Groups.Items.Add(parameter.Name);
                    }
                    else if (parameter.Type == "Nominal")
                    {
                        list_Nominal.Items.Add(parameter.Name);
                        OrderedParameters.Add(parameter.Name);
                    }
                    else if (parameter.Type == "Scale")
                    {
                        if (parameter.Normality == "Normal")
                        {
                            list_NormalScale.Items.Add(parameter.Name);
                            OrderedParameters.Add(parameter.Name);
                        }
                        else if (parameter.Normality == "Abnormal")
                        {
                            list_AbnormalScale.Items.Add(parameter.Name);
                            OrderedParameters.Add(parameter.Name);
                        }
                    }
                }


                MessageBox.Show("Lists Updated", "Update Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public StatParameter GetSpssParamByName(string parameterName)
        {
            return SpssParameters.FirstOrDefault(p => p.Name == parameterName);
        }


        private void btn_Update_Click(object sender, EventArgs e)
        {
            var spssRaw = SpssReaderStat.ReadSpssParameters(SpssReaderStat.spssFilePath);
            var statParams = spssRaw.Select(s => new StatParameter
            {
                Name = s.Name,
                Label = s.Label,
                Type = s.Type,
                RawValues = new List<string>(s.Values),
                ValueLabels = new Dictionary<int, string>(s.ValueLabels)
            }).ToList();


            FormDataTransfer.Set("SPSS_Parameters", statParams);

            SpssParameters = statParams;

            PopDataGrid_AllSpssParam();
        }
        private ParameterInfoFrm_Child parameterfrm_obj;
        private void data_allPara_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Check if the clicked column is the button column
                if (data_allPara.Columns[e.ColumnIndex] is DataGridViewButtonColumn)
                {
                    // Get the row where the button was clicked
                    DataGridViewRow clickedRow = data_allPara.Rows[e.RowIndex];

                    // Retrieve parameter name
                    string paraName = clickedRow.Cells[0].Value.ToString();

                    // If the form is not open, create it
                    if (parameterfrm_obj == null || parameterfrm_obj.IsDisposed)
                    {
                        parameterfrm_obj = new ParameterInfoFrm_Child();
                        parameterfrm_obj.Show();
                    }

                    StatParameter viewedParameter = GetSpssParamByName(paraName);

                    // Update the existing form with the new data
                    parameterfrm_obj.UpdateParameter(viewedParameter);
                }
            }
        }

        private void groupbx_Nominal_Enter(object sender, EventArgs e)
        {

        }
    }
}
