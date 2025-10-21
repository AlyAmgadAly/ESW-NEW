using Accord.Math;
using Accord.Statistics.Distributions.Univariate;
using Accord.Statistics.Kernels;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using CenterSpace.NMath.Core;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
using Humanizer;
using MathNet.Numerics.LinearAlgebra.Factorization;
using MathNet.Numerics.Statistics;
using Microsoft.SolverFoundation.Services;
using Python.Runtime;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using System.Xml.Linq;
using static Humanizer.On;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskBand;
using Application = System.Windows.Forms.Application;
using Parameter = ExcelScore.Classes.Parameter;
using Worksheet = Aspose.Cells.Worksheet;
using System.IO;
using Accord.Statistics.Testing;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Syncfusion.Pdf.Tables;
using ExcelScore.FormsDesigns;
using DocumentFormat.OpenXml.Presentation;
using ExcelScore.StatClasses;
using ClosedXML.Excel;
using System.Security.Cryptography;
using System.Diagnostics;

namespace ExcelScore.Forms
{
    public partial class ComparativeGroups : Form
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

        public static Worksheet worksheet = excelFunctionsobj.GetWorksheet();

        public static Worksheet Sheet2 = excelFunctionsobj.GetSheet2();
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
            data_allPara.Rows.Clear();

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
        //to retrieve the % parts in default (% row , col , total)
        public Dictionary<string, bool> CheckedDataExtraGeneral = new Dictionary<string, bool>();

        public List<string> OrderedParameters = new List<string>();
        public Dictionary<string, string> NormalityParaNameList_ComparaGroups { get; set; } = new Dictionary<string, string>();
        private void ComparativeGroups_Load(object sender, EventArgs e)
        {
            lbl_PeriodsCount.Visible = false;
            txt_PeriodCount.Visible = false;


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
                    OrderedParameters.Add(item);
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


                foreach (var item in selectedItems)
                {
                    OrderedParameters.Remove(item.ToString());
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
                    OrderedParameters.Add(item);
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

                foreach (var item in selectedItems)
                {
                    OrderedParameters.Remove(item.ToString());
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
                    OrderedParameters.Add(item);
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

                foreach (var item in selectedItems)
                {
                    OrderedParameters.Remove(item.ToString());
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
            table.TableDesignType = cmb_ChooseTableFormat.Text;

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

            // --- Add SubGroup Parameters ---
            foreach (var item in list_SubGroups.Items)
            {
                table.Parameters.Add(new StatParameter
                {
                    Name = item.ToString(),
                    IsSubGroup = true,
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
        List<StatTable> StatTables = new List<StatTable>();
        public void AddTableClass(string tableName)
        {
            bool totalcolumn = Total_Column_Comparative;

            var comparativeTable = new ComparativeTable
            {
                TableName = tableName,
                Parameters = new List<Parameter>(),
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
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }


            foreach (var item in list_SubGroups.Items)
            {
                //MessageBox.Show(item.ToString());
                var parameter = new Parameter
                {
                    Name = item.ToString(),
                    IsSubGroup = true,
                    GroupedParameterValues = new Dictionary<double, List<double>>(),
                    FormattedValues = new Dictionary<double, Dictionary<string, string>>() // Initialize FormattedValues dictionary
                };
                comparativeTable.Parameters.Add(parameter);
            }
            //IsSubGroup

            foreach (var paramName in OrderedParameters)
            {
                Parameter parameter = null;

                
                if (list_Nominal.Items.Contains(paramName))
                {
                    parameter = new Parameter
                    {
                        Name = paramName,
                        NominalOrScale = "Nominal",
                        GroupedParameterValues = new Dictionary<double, List<double>>(),
                        FormattedValues = new Dictionary<double, Dictionary<string, string>>()
                    };

                    if (NominalYes_Compara != null && NominalYes_Compara.Contains(paramName))
                        parameter.NominalIsYes = true;
                }
                else if (list_NormalScale.Items.Contains(paramName))
                {
                    parameter = new Parameter
                    {
                        Name = paramName,
                        NominalOrScale = "Scale",
                        NormalOrAbnormal = "Normal",
                        GroupedParameterValues = new Dictionary<double, List<double>>(),
                        FormattedValues = new Dictionary<double, Dictionary<string, string>>()
                    };
                }
                else if (list_AbnormalScale.Items.Contains(paramName))
                {
                    parameter = new Parameter
                    {
                        Name = paramName,
                        NominalOrScale = "Scale",
                        NormalOrAbnormal = "Abnormal",
                        GroupedParameterValues = new Dictionary<double, List<double>>(),
                        FormattedValues = new Dictionary<double, Dictionary<string, string>>()
                    };
                }

                if (parameter != null)
                {
                    comparativeTable.Parameters.Add(parameter);
                }
            }

            comparativeTable.FormatType = cmb_ChooseTableFormat.Text;

            ComparativeTables.Add(comparativeTable);




        }
        public string AddTableUINew()
        {
            string tableName = null;

            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&
                list_Groups.Items.Count > 0 &&
                !string.IsNullOrWhiteSpace(cmb_ChooseTableFormat.Text) &&
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

                    // Save the design type for consistency
                    FormDataTransfer.Set("TableDesignType", cmb_ChooseTableFormat.Text);

                    // Create and add table
                    AddTableClassNew(txt_TableName.Text, tableType);
                    cmb_TableNames.Items.Add(txt_TableName.Text);

                    //MessageBox.Show("Added Table " + txt_TableName.Text);
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
                    AddTableClassNew(txt_TableName.Text , FormDataTransfer.Get<string>("TableType"));
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

                            if (valuesErrored.Count > 7)
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

        public List<int> newFindGroups_ColumnIndexes_Relation(ComparativeTable comparativeTable)
        {
            List<int> Group_Indexes = new List<int>();
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
                    {
                        object cellValue = worksheet.Cells[0, col].Value;
                        string parameterName = cellValue?.ToString();
                        if (!string.IsNullOrEmpty(parameterName))
                        {
                            if (parameter.Name == parameterName && parameter.IsGroup)
                            {
                                Group_Indexes.Add(col);
                                break;
                            }
                        }

                    }
                }
            }

            return Group_Indexes;

        }
        public void ComparativeBasic_Relation(string TableName)
        {
            GetDataValues(TableName);

            foreach (ComparativeTable table in ComparativeTables)
            {
                if (TableName == table.TableName)
                {
                    excelFunctionsobj.ReadLablesIfNominal(table);
                    List<int> GroupColumn_Indexes = newFindGroups_ColumnIndexes_Relation(table);
                    foreach (Parameter parameter in table.Parameters)
                    {

                        if (parameter.IsGroup)
                        {
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
                            foreach (int GroupColumn_index in GroupColumn_Indexes)
                            {
                                for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
                                {
                                    object groupCellValue = worksheet.Cells[row, GroupColumn_index].Value;
                                    double groupValue;

                                    object groupcellName = worksheet.Cells[0, GroupColumn_index].Value;
                                    string GroupName = groupcellName.ToString();


                                    if (groupCellValue != null && double.TryParse(groupCellValue.ToString(), out groupValue))
                                    {
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

                                                if (!parameter.GroupedParameterValues_Relation.ContainsKey(GroupName))
                                                {
                                                    parameter.GroupedParameterValues_Relation[GroupName] = new Dictionary<double, List<double>>();
                                                }

                                                if (!parameter.GroupedParameterValues_Relation[GroupName].ContainsKey(groupValue))
                                                {
                                                    parameter.GroupedParameterValues_Relation[GroupName][groupValue] = new List<double>();
                                                }
                                                parameter.GroupedParameterValues_Relation[GroupName][groupValue].Add(parameterValue);




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

                                                        if (!parameter.GroupedParameterValues_Relation.ContainsKey(GroupName))
                                                        {
                                                            parameter.GroupedParameterValues_Relation[GroupName] = new Dictionary<double, List<double>>();
                                                        }

                                                        if (!parameter.GroupedParameterValues_Relation[GroupName].ContainsKey(groupValue))
                                                        {
                                                            parameter.GroupedParameterValues_Relation[GroupName][groupValue] = new List<double>();
                                                        }
                                                        parameter.GroupedParameterValues_Relation[GroupName][groupValue].Add(parameterValue);




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

            FormatParameters_Relation(TableName);


        }

        public void FormatParameters_Relation(string TableName)
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                if (TableName == table.TableName)
                {
                    foreach (Parameter parameter in table.Parameters)
                    {
                        if (parameter.IsGroup || parameter.GroupedParameterValues_Relation.Count == 0)
                        {
                            continue;
                        }

                        foreach (var kvp in parameter.GroupedParameterValues_Relation)
                        {
                            string GroupName = kvp.Key;
                            Dictionary<double, List<double>> Value = kvp.Value;


                            if (!parameter.FormattedValues_Relation.ContainsKey(GroupName))
                            {
                                parameter.FormattedValues_Relation[GroupName] = new Dictionary<double, Dictionary<string, string>>();
                            }

                            foreach (var kvp1 in parameter.GroupedParameterValues_Relation[GroupName])
                            {

                                double groupvalue = kvp1.Key;
                                List<double> values = kvp1.Value;
                                int totalCount = values.Count;


                                if (!parameter.FormattedValues_Relation[GroupName].ContainsKey(groupvalue))
                                {
                                    parameter.FormattedValues_Relation[GroupName][groupvalue] = new Dictionary<string, string>();
                                }

                                if (parameter.NominalOrScale == "Nominal")
                                {
                                    // Calculate frequency and percentage for each distinct value in the group
                                    foreach (var distinctValue in values.Distinct())
                                    {


                                        int Totalrowfreq = 0;
                                        foreach (var kvp2 in parameter.GroupedParameterValues_Relation[GroupName])
                                        {
                                            double groupValue2 = kvp2.Key;
                                            List<double> values2 = kvp2.Value;
                                            int frequency2 = values2.Count(v => v == distinctValue);
                                            Totalrowfreq += frequency2;
                                        }

                                        if (Row_Percent_Comparative)
                                        {
                                            int frequency = values.Count(v => v == distinctValue);
                                            double percentage = (frequency / (double)Totalrowfreq) * 100;
                                            parameter.FormattedValues_Relation[GroupName][groupvalue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                            parameter.FormattedValues_Relation[GroupName][groupvalue][$"Percentage_{distinctValue}"] = $"{percentage:F1}%";
                                        }
                                        else if (!Row_Percent_Comparative)
                                        {
                                            int frequency = values.Count(v => v == distinctValue);
                                            double percentage = (frequency / (double)totalCount) * 100;
                                            parameter.FormattedValues_Relation[GroupName][groupvalue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                            parameter.FormattedValues_Relation[GroupName][groupvalue][$"Percentage_{distinctValue}"] = $"{percentage:F1}%";
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
                                    parameter.FormattedValues_Relation[GroupName][groupvalue]["Min-Max"] = formattedMinMax;
                                    parameter.FormattedValues_Relation[GroupName][groupvalue]["Mean ± StdDev"] = formattedMeanStd;
                                    parameter.FormattedValues_Relation[GroupName][groupvalue]["Median"] = formattedMedian + " (" + formattedIQR + ")";
                                    //parameter.FormattedValues[groupValue]["IQR"] = formattedIQR;
                                }


                            }

                        }
                    }
                }
                //MessageBox.Show(table.TableName);

            }
        }

        public static bool Row_Percent_Comparative { get; set; }
        public static bool Maha_Comparative { get; set; }
        public static bool Total_Column_Comparative { get; set; }
        private void pic_addTable_Click(object sender, EventArgs e)
        {

            string TableName = AddTableUI();


            ///check_TotalColumn.Checked = false;




            if (cmb_ChooseTableFormat.Text == "Relation" || cmb_ChooseTableFormat.Text == "Relation Scale Pathology" || cmb_ChooseTableFormat.Text == "Relation IQR" || cmb_ChooseTableFormat.Text == "Relation Median No IQR")
            {
                ComparativeBasic_Relation(TableName);
            }
            else
            {
                ComparativeBasic(TableName);
            }



            CheckFullEmptyParameters(TableName);
            CheckForOthers_inNominal(TableName);

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



                                    if (cellData == null)
                                    {
                                        MessageBox.Show("Data is empty at Parameter : " + parameter.Name);
                                        break;
                                    }
                                    else if (cellData.ToString() == ".")
                                    {
                                        parameter.hasLowerN = true;
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






        public void FormatParameters(string TableName)
        {
            foreach (ComparativeTable table in ComparativeTables)
            {
                if (TableName == table.TableName)
                {
                    CheckedDataExtraGeneral = FormDataTransfer.Get<Dictionary<string, bool>>("nodeCheckedStatusExtra");
                    Parameter GroupParameter = ComparativeTable.GetGroupParamter(table);
                    double GrandTotalCount = GroupParameter.ParameterValues.Count;
                    foreach (Parameter parameter in table.Parameters)
                    {
                        //|| parameter.GroupedParameterValues.Count == 0
                        if (parameter.IsGroup)
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
                                bool RowPercent = CheckedDataExtraGeneral["RowPercentage"];
                                bool ColumnPercent = CheckedDataExtraGeneral["ColumnPercentage"];
                                bool TotalPercent = CheckedDataExtraGeneral["TotalPercentage"];
                                
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

                                    if (RowPercent)
                                    {
                                        int frequency = values.Count(v => v == distinctValue);
                                        double percentage = (frequency / (double)Totalrowfreq) * 100;
                                        parameter.FormattedValues[groupValue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                        parameter.FormattedValues[groupValue][$"Percentage_{distinctValue}"] = $"{percentage:F1}%";
                                    }
                                    else if (ColumnPercent)
                                    {
                                        int frequency = values.Count(v => v == distinctValue);
                                        double percentage = (frequency / (double)totalCount) * 100;
                                        parameter.FormattedValues[groupValue][$"Frequency_{distinctValue}"] = frequency.ToString();
                                        parameter.FormattedValues[groupValue][$"Percentage_{distinctValue}"] = $"{percentage:F1}%";
                                    }
                                    else if(TotalPercent)
                                    {
                                        int frequency = values.Count(v => v == distinctValue);
                                        double percentage = (frequency / (double)GrandTotalCount) * 100;
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

            if (n == 1)
            {
                return values[0];
            }
            else if (n % 2 == 0)
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

                return CalculateMedian(values.GetRange(0, middle));
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
            else if (n == 1)
            {
                return values[0];
            }
            else if (n == 3)
            {
                return (values[1] + values[2]) / 2.0;
            }

            else
            {
                int middle = (n + 1) / 2;
                // Odd number of elements, calculate median of upper half excluding the median
                return CalculateMedian(values.GetRange(middle, n - middle - 1));

            }
        }

        public void RemoveTable()
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




            string tabletype = selectedTable.FormatType;

            cmb_UpdateTableType.Text = tabletype;







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

            string minValueString = "";
            string maxValueString = "";



            minValueString = minValue.ToString("0.00");
            maxValueString = maxValue.ToString("0.00");


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

        public int CountRows_NoIQR(ComparativeTable table)
        {
            int rowCount = 0;

            foreach (var parameter in table.Parameters)
            {

                if (parameter.NominalOrScale == "Nominal")
                {
                    int distinctValuesCount = parameter.DIC_LablesIfNomainal.Keys.Count;
                    rowCount += distinctValuesCount + 1;

                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    if (parameter.ISFAnovaSig)
                    {
                        rowCount += 4;
                    }
                    else
                    {
                        rowCount += 3;
                    }

                }
            }



            return rowCount;
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
                    rowCount += distinctValuesCount + 1;

                }
                else if (parameter.NominalOrScale == "Scale")
                {
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

        public void InsertNeachGroup_New(IWTable table, List<double> EachGroupCount, ComparativeTable comparativeTable)
        {
            int col = 1;

            Parameter Groupparameter = null;
            foreach (var parameter in comparativeTable.Parameters)
            {

                if (parameter.IsGroup)
                {
                    Groupparameter = parameter;
                    break;
                }
            }

            foreach (var kvp in Groupparameter.DIC_LablesIfNomainal)
            {
                int key = kvp.Key;

                string label = kvp.Value;
                int count = Groupparameter.ParameterValues.Count(x => x == key);

                if (count != 0)
                {
                    wordObj.AddPara_Center(table, 0, col, label + Convert.ToChar(11) + "(n = " + count + ")");
                    col = col + 2;
                }

            }
        }
        public void InsertNeachGroup(IWTable table, List<double> EachGroupCount, ComparativeTable comparativeTable)
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

                if (parameter.IsGroup)
                {
                    Groupparameter = parameter;
                    break;
                }
            }

            int groupctr = 0;


            foreach (var label in Groupparameter.DIC_LablesIfNomainal.Values)
            {
                //MessageBox.Show(label);
                wordObj.AddPara_Center(table, 0, column, label + Convert.ToChar(11) + "(n = " + EachGroupCount[groupctr] + ")");

                column = column + 2;
                groupctr++;
            }




        }

        DefinitionUnderTable definitionUnderTableObj = new DefinitionUnderTable();



        public void InsertPairwiseF(IWTable table, ComparativeTable comparativeTable, int WordTableColumns)
        {
            bool Hasnominal = false;
            foreach (var parameter in comparativeTable.Parameters)
            {
                // MessageBox.Show(parameter.Name);
                if (parameter.NominalOrScale == "Nominal")
                {
                    Hasnominal = true;
                }
            }
            int startrow = 0;
            if (Hasnominal)
            {
                startrow = 2;
            }
            else
            {
                startrow = 1;
            }




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
                if (parameter.NominalOrScale == "Nominal")
                {
                    //count = parameter.ParameterValues.Distinct().Count() + 2;
                    count = parameter.DIC_LablesIfNomainal.Keys.Count() + 1;
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    count = 4;
                }


                if (parameter.ISFAnovaSig)
                {

                    WTableRow row;

                    int insert = startrow + 4;



                    row = table.AddRow();
                    table.Rows.Insert(insert, row);


                    int rowindex = table.Rows.IndexOf(row);



                    wordObj.AddParaCombined(table, rowindex, 0, "Sig. bet. grps.", true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                    table.ApplyHorizontalMerge(rowindex, 1, WordTableColumns - 3);


                    for (int i = 0; i < WordTableColumns; i++)
                    {
                        if (parameter == lastParameter)
                        {

                            table.Rows[rowindex].Cells[i].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                            table.Rows[rowindex].Cells[i].CellFormat.Borders.Top.LineWidth = 0.5f;
                        }
                        else if (!(parameter == lastParameter))
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


                        //wordObj.AddParaCombined(table, rowindex, 1, "p" , false , true , Syncfusion.Drawing.Color.Empty , Syncfusion.Drawing.Color.Black);
                        //wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, groupnumber, "Sub");

                        wordObj.InsertPairwiseF(table, rowindex, 1, "p", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, groupnumber, "Sub");
                        if (!parameter.FPairwise[i].Contains("<"))
                        {
                            wordObj.InsertPairwiseF(table, rowindex, 1, "=", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }

                        wordObj.InsertPairwiseF(table, rowindex, 1, parameter.FPairwise[i], paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                        if (issig)
                        {
                            wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, "*", "Super");
                        }

                        if (i < parameter.FPairwise.Count - 1)
                        {
                            wordObj.InsertPairwiseF(table, rowindex, 1, ",", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                        }

                    }

                    WTableRow other;
                    other = table.AddRow(false);
                    table.Rows.Insert(insert + 1, other);


                    table.ApplyHorizontalMerge(insert + 1, 1, WordTableColumns - 3);

                    string Column0text = table[insert, 0].Paragraphs[0].Text;
                    wordObj.AddParaCombined(table, insert + 1, 0, Column0text, true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

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

                    table.Rows[insert + 1].Cells[WordTableColumns - 2].CellFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
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
        public void StatBasic(string tableName)
        {
            // Step 1: Find the table (any type)
            var table = StatTables.FirstOrDefault(t => t.TableName == tableName);
            if (table == null)
            {
                return;
            }

            // Step 2: Load SPSS parameters from memory
            var spssParams = FormDataTransfer.Get<List<StatParameter>>("SPSS_Parameters");
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


        public void ComparativeBasic(string TableName)
        {
            GetDataValues(TableName);

            foreach (ComparativeTable table in ComparativeTables)
            {
                if (TableName == table.TableName)
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

                                            parameter.GroupedParameterValues[groupValue].Add(parameterValue);

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
        public void AddTotalColumn(IWTable table, ComparativeTable comparativeTable)
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


            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {

                    continue;
                }
                if (parameter.NominalOrScale == "Nominal")
                {
                    tableHasNominal = true;
                }

            }
            //Merge Total

            table.ApplyHorizontalMerge(0, 1, 2);
            //table.ApplyHorizontalMerge(1, 1, 2);

            //table.ApplyVerticalMerge(1, 0, 1);


            if (tableHasNominal)
            {
                wordObj.AddPara_Center(table, 1, 1, "No.");
                wordObj.AddPara_Center(table, 1, 2, "%");
            }
            //boders
            if (tableHasNominal)
            {
                table.Rows[0].Cells[1].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[1].CellFormat.Borders.Bottom.LineWidth = 0.5f;

                table.Rows[0].Cells[0].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[0].CellFormat.Borders.Bottom.LineWidth = 0.5f;


                table.Rows[1].Cells[1].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[1].CellFormat.Borders.Bottom.LineWidth = 1.5f;

                table.Rows[1].Cells[2].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[1].Cells[2].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }
            else if (!tableHasNominal)
            {
                table.Rows[0].Cells[1].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[1].CellFormat.Borders.Bottom.LineWidth = 1.5f;

                table.Rows[0].Cells[0].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                table.Rows[0].Cells[0].CellFormat.Borders.Bottom.LineWidth = 1.5f;
            }



            int TotalN = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (!parameter.hasLowerN)
                {
                    TotalN = parameter.ParameterValues.Count;
                    break;
                }
            }

            wordObj.AddPara_Center(table, 0, 1, "Total\n(n = " + TotalN + ")");


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
                    //tableHasNominal = true;
                    count = parameter.DIC_LablesIfNomainal.Keys.Count + 1;
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
                if (startrow != 2)
                {
                    if (tableHasNominal)
                    {
                        table.Rows[startrow].Cells[1].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[startrow].Cells[1].CellFormat.Borders.Top.LineWidth = 0.5f;

                        table.Rows[startrow].Cells[2].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[startrow].Cells[2].CellFormat.Borders.Top.LineWidth = 0.5f;
                    }
                    else if (!tableHasNominal)
                    {
                        table.Rows[startrow - 1].Cells[1].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[startrow - 1].Cells[1].CellFormat.Borders.Top.LineWidth = 0.5f;

                        table.Rows[startrow - 1].Cells[2].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[startrow - 1].Cells[2].CellFormat.Borders.Top.LineWidth = 0.5f;
                    }

                }


                int totalCount = parameter.ParameterValues.Count;

                //if (parameter.NominalOrScale == "Nominal")
                //{
                //    // Calculate frequency and percentage for each distinct value in the group
                //    int incrow = 1;
                //    foreach (var distinctValue in parameter.ParameterValues.Distinct())
                //    {
                //        int frequency = parameter.ParameterValues.Count(v => v == distinctValue);
                //        double percentage = (frequency / (double)totalCount) * 100;

                //        if (parameter.DIC_LablesIfNomainal.ContainsKey((int)distinctValue))
                //        {
                //            if (startrow != 2)
                //            {
                //                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 1, frequency.ToString());
                //                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 2, percentage.ToString("0.0"));
                //                incrow++;
                //            }
                //            else if (startrow == 2)
                //            {
                //                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 1, frequency.ToString());
                //                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 2, percentage.ToString("0.0"));
                //                incrow++;
                //            }




                //        }
                //    }
                //}
                if (parameter.NominalOrScale == "Nominal")
                {
                    // Calculate frequency and percentage for each distinct value in the group
                    int incrow = 1;
                    foreach (var distinctValue in parameter.DIC_LablesIfNomainal.Keys)
                    {
                        int frequency = parameter.ParameterValues.Count(v => v == distinctValue);
                        double percentage = (frequency / (double)totalCount) * 100;

                        if (parameter.DIC_LablesIfNomainal.ContainsKey((int)distinctValue))
                        {
                            if (startrow != 2)
                            {
                                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 1, frequency.ToString());
                                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 2, percentage.ToString("0.0"));
                                incrow++;
                            }
                            else if (startrow == 2)
                            {
                                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 1, frequency.ToString());
                                wordObj.Addpara_CenterNoBOLD(table, startrow + incrow, 2, percentage.ToString("0.0"));
                                incrow++;
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





                    if (tableHasNominal)
                    {
                        table.Rows[1].Cells[1].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[1].Cells[1].CellFormat.Borders.Bottom.LineWidth = 1.5f;

                        table.Rows[1].Cells[2].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[1].Cells[2].CellFormat.Borders.Bottom.LineWidth = 1.5f;




                        table.ApplyVerticalMerge(1, 0, 1);
                        table.ApplyHorizontalMerge(startrow, 1, 2);
                        table.ApplyHorizontalMerge(startrow + 1, 1, 2);
                        table.ApplyHorizontalMerge(startrow + 2, 1, 2);
                        table.ApplyHorizontalMerge(startrow + 3, 1, 2);


                        wordObj.Addpara_CenterNoBOLD(table, startrow + 1, 1, formattedMinMax);
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
                    else if (!tableHasNominal)
                    {
                        table.ApplyHorizontalMerge(startrow - 1, 1, 2);

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
                if (parameter.NominalOrScale == "Scale")
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
                   

                    IWSection section = wordObj.CreatePortraitSection();




                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = (5 * numberofgroups) + 3;

                    //log(Variablerows);


                    int WordTableColumns = 3 + CountPeriodsUp_Groups_Columns(ComparativeTables[tableindex]);

                    int WordTableRows = Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.GeneralTableFormat(table);

                    wordObj.ApplyGeneralPeriodsUp_Groups_ComparativeMerges(table, WordTableColumns, WordTableRows, numberofgroups, ComparativeTables[tableindex]);


                    wordObj.ApplyGeneral_PeriodsUp_Groups_ComparativeBorders(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);

                    wordObj.SetComparative_PeriodsUp_GroupsWidths(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);


                    wordObj.SetComparative_PeriodsUp_Groups_Headers(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);


                    InsertData_scale_PeriodsUp_Groups(table, ComparativeTables[tableindex]);

                    GetGroupsTest_PeriodsUp_groups(table, ComparativeTables[tableindex], WordTableRows);


                    GetPeriodsTest_PeriodsUp_groups(table, ComparativeTables[tableindex], WordTableColumns);


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
                    count = values.Count + tempctr;


                    for (int i = tempctr; i < count; i++)
                    {
                        firstParameterValues.Add(scaleParameters[0].ParameterValues[i]);
                        secondParameterValues.Add(scaleParameters[1].ParameterValues[i]);

                    }
                    tempctr = count;

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

        public void GetGroupsTest_PeriodsUp_groups(IWTable table, ComparativeTable comparativeTable, int WordTableRows)
        {

            int col = 1;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }

                if (parameter.NominalOrScale == "Scale")
                {


                    if (parameter.NormalOrAbnormal == "Normal")
                    {
                        List<double> group1Values = new List<double>();
                        List<double> group2Values = new List<double>();

                        SplitGroupedParameterValues(parameter, out group1Values, out group2Values);



                        string[] values = manual.StudentT_Unpaired(group1Values, group2Values);



                        wordObj.AddParaCombined(table, WordTableRows - 1, col, values[0], false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);


                        WParagraph testparaHighlight = (WParagraph)table[WordTableRows - 1, col].Paragraphs[0];


                        WTextRange PText = new WTextRange(testparaHighlight.Document);
                        PText.Text = " (" + values[1] + ")";

                        testparaHighlight.ChildEntities.Insert(1, PText);




                        col++;
                    }


                    else if (parameter.NormalOrAbnormal == "Abnormal")
                    {
                        List<double> group1Values = new List<double>();
                        List<double> group2Values = new List<double>();

                        SplitGroupedParameterValues(parameter, out group1Values, out group2Values);

                        string[] values = manual.UTest(group1Values, group2Values);

                        wordObj.AddParaCombined(table, WordTableRows - 1, col, values[0], false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);


                        WParagraph testparaHighlight = (WParagraph)table[WordTableRows - 1, col].Paragraphs[0];


                        WTextRange PText = new WTextRange(testparaHighlight.Document);
                        PText.Text = " (" + values[1] + ")";

                        testparaHighlight.ChildEntities.Insert(1, PText);




                        col++;
                    }







                }
            }

        }

        public void InsertData_scale_PeriodsUp_Groups(IWTable table, ComparativeTable comparativeTable)
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

        public void ComparativeTableGroups_Layout_NoIQR()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Default No IQR")
                {
                    

                    IWSection section = wordObj.CreatePortraitSection();




                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);


                    int Variablerows = CountRows_NoIQR(ComparativeTables[tableindex]);
                    int WordTableRows = 2 + Variablerows;

                    if (!ComparativeTables[tableindex].HasTotalColumn)
                    {
                        int WordTableColumns = 3 + (numberofgroups * 2);
                    }
                    else if (ComparativeTables[tableindex].HasTotalColumn)
                    {
                        int WordTableColumns = 3 + (numberofgroups * 2) + 1;
                    }





                }
            }
        }

        public void PaperComparative_Layout()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Paper")
                {
                    bool TableHasSigI = false;

                    IWSection section = wordObj.CreatePortraitSection();




                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = CountRows(ComparativeTables[tableindex]);



                    int WordTableColumns = 3 + numberofgroups;

                    int WordTableRows = 1 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.PaperGeneralFormat(table);

                    wordObj.PaperBordersGeneral(table, WordTableColumns);
                    wordObj.PaperComparative_widths(table, WordTableRows, numberofgroups, WordTableColumns);

                    wordObj.AddHeaderPaper(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);


                    (string testtype, string NominalOrScale, bool issame) = InsertSeperateTest_TestOfSig(ComparativeTables[tableindex], numberofgroups);

                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, testtype);



                    int startingRow;
                    int newRowCount = 2;

                    foreach (Parameter parameter in ComparativeTables[tableindex].Parameters)
                    {
                        startingRow = newRowCount;
                        int column = 1;
                        int count = 0;

                        if (parameter.IsGroup || parameter.FormattedValues.Count == 0)
                        {
                            continue;
                        }

                        // Determine the count based on parameter type
                        if (parameter.NominalOrScale == "Nominal")
                        {
                            count = parameter.DIC_LablesIfNomainal.Keys.Count() + 1;
                            ComparativeTables[tableindex].TestsDone.Add("Chi");
                            pythonStat.InitPython();
                            string[] values = pythonStat.PerformChiSquareTest(parameter);
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


                            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();

                            foreach (var groupValue in sortedKeys)
                            {
                                Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];
                                foreach (var distinctValue in parameter.DIC_LablesIfNomainal.Keys)
                                {

                                    string frequency = formattedValues.ContainsKey($"Frequency_{distinctValue}") ? formattedValues[$"Frequency_{distinctValue}"] : "0";
                                    //wordObj.Addpara_CenterNoBOLD(table, startingRow, column, frequency);

                                    string percentage = formattedValues.ContainsKey($"Percentage_{distinctValue}") ? formattedValues[$"Percentage_{distinctValue}"] : "0.0";
                                    percentage = percentage.Replace("%", "");


                                    if (percentage.EndsWith("0"))
                                    {
                                        string InsertedString = frequency + " (" + percentage + "%)";
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, InsertedString);
                                    }

                                    else
                                    {
                                        string InsertedString = frequency + " (" + percentage + "%)";
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, InsertedString);
                                    }
                                    startingRow++; // Move to the next row
                                }
                                // Reset startingRow and increment column for the next group
                                startingRow = newRowCount;
                                column++;
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

                                column++;
                            }

                            PerformTest(table, numberofgroups, parameter, startingRow, WordTableColumns - 2, WordTableColumns, issame, ComparativeTables[tableindex]);

                            if (parameter.NormalOrAbnormal == "Normal")
                            {
                                wordObj.HighlightCellContent_Paper(table, startingRow + 2, WordTableColumns);
                            }
                            else if (parameter.NormalOrAbnormal == "Abnormal")
                            {
                                wordObj.HighlightCellContent_Paper(table, startingRow + 1, WordTableColumns);
                            }

                        }

                        // Increment newRowCount for the next parameter
                        newRowCount += count;
                    }

                    //Remove No and %
                    try
                    {
                        InsertPairwiseF(table, ComparativeTables[tableindex], WordTableColumns);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Remove Nominal Parameters!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }


                    wordObj.FormatTable(table, 12);
                }
            }
        }

        public int Relation_DependentNumber_Rows(ComparativeTable comparativeTable)
        {
            int TotalRows = 0;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    int count = parameter.DIC_LablesIfNomainal.Keys.Count + 1;
                    TotalRows += count;
                }


            }


            return TotalRows;
        }

        public int Relation_DependentNumber_Rows_Pathology(ComparativeTable comparativeTable)
        {
            int TotalRows = 0;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    int count = parameter.DIC_LablesIfNomainal.Keys.Count;
                    TotalRows += count;
                }


            }


            return TotalRows;
        }

        public void Periods_Layout()
        {

        }

        public void Relation_Layout_DependentNumber_Pathology()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Relation Scale Pathology")
                {
                    IWSection section = wordObj.CreatePortraitSection();



                    wordObj.AddRelationTitle(section, ComparativeTables[tableindex].TableName);



                    int Variablerows = Relation_DependentNumber_Rows_Pathology(ComparativeTables[tableindex]);


                    int WordTableColumns = 7;

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);


                    //Merges
                    wordObj.ApplyRelation_OuterMerges_Pathology(table, WordTableRows, WordTableColumns);

                    //Borders
                    wordObj.ApplyRelation_OuterBorders_Pathology(table, WordTableRows, WordTableColumns);

                    //Widths
                    wordObj.ApplyRelation_Widths_Pathology(table, WordTableRows, WordTableColumns);


                    //Outer Headers
                    wordObj.ApplyRelation_OuterHeaders_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    wordObj.InsertRelation_InnerHeader_Merges_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);

                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test of Sig.");


                    foreach (var parameter in ComparativeTables[tableindex].Parameters)
                    {

                        if (parameter.IsGroup || parameter.GroupedParameterValues_Relation.Count == 0)
                        {
                            continue;
                        }
                        int StartingRow = 2;
                        foreach (var GroupNameKvp in parameter.FormattedValues_Relation)
                        {
                            string GroupName = GroupNameKvp.Key;


                            Parameter groupParameter = ComparativeTables[tableindex].Parameters.SingleOrDefault(p => p.Name == GroupName);
                            PerformTest_Relation(ComparativeTables[tableindex], table, StartingRow, WordTableColumns - 2, parameter, groupParameter);


                            var sortedGroups = parameter.FormattedValues_Relation[GroupName].Keys.OrderBy(key => key).ToList();

                            var SortedGroupDIC = groupParameter.DIC_LablesIfNomainal.Keys.OrderBy(key => key).ToList();

                            foreach (var GroupValue in SortedGroupDIC)
                            {
                                if (sortedGroups.Contains((int)GroupValue))
                                {
                                    if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count > 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];

                                        //string MinMax

                                        string MeanSD = "";
                                        string Median = "";
                                        string MinMax = "";
                                        string FormattedMedianMinMax = "";

                                        foreach (var stat in formattedValues)
                                        {
                                            if (stat.Key == "Mean ± StdDev")
                                            {
                                                MeanSD = stat.Value;
                                            }
                                            if (stat.Key == "Median")
                                            {

                                                string[] valuessplit = stat.Value.Split(' ');
                                                Median = valuessplit[0];
                                            }
                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                            }

                                        }

                                        FormattedMedianMinMax = Median + " (" + MinMax + ")";
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, MeanSD);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 4, FormattedMedianMinMax);

                                        StartingRow++;
                                    }




                                    else if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count == 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];
                                        string MinMax = "";
                                        string OnlyValue = "";
                                        foreach (var stat in formattedValues)
                                        {

                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                                string[] values = MinMax.Split(' ');
                                                OnlyValue = values[0];
                                            }

                                        }
                                        table.ApplyHorizontalMerge(StartingRow, 3, 4);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, OnlyValue);
                                        wordObj.SubSuperScriptText(table, StartingRow, 3, Syncfusion.Drawing.Color.White, "#", "Super");
                                        StartingRow++;

                                    }





                                }
                                else
                                {
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 4, "–");
                                    StartingRow++;
                                }



                            }




                        }




                    }






                    wordObj.LeftAndRightCellMarginCustom(table, 0, 0);
                    wordObj.FormatTableCustom(table, 10.5f, 0, 0);

                }
            }
        }
        public void Relation_Layout_DependentNumber_IQR()
        {

            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Relation IQR")
                {
                    

                    IWSection section = wordObj.CreatePortraitSection();



                    wordObj.AddRelationTitle(section, ComparativeTables[tableindex].TableName);



                    int Variablerows = Relation_DependentNumber_Rows(ComparativeTables[tableindex]);


                    int WordTableColumns = 7;

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);


                    //Merges
                    wordObj.ApplyRelation_IQR_OuterMerges(table, WordTableRows, WordTableColumns);

                    //Borders
                    wordObj.ApplyRelation_IQR_OuterBorders(table, WordTableRows, WordTableColumns);

                    //Widths
                    wordObj.ApplyRelation_IQR_Widths(table, WordTableRows, WordTableColumns);


                    //Outer Headers
                    wordObj.ApplyRelation_IQR_OuterHeaders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);


                    //
                    wordObj.InsertRelation_InnerHeader_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test of Sig.");


                    foreach (var parameter in ComparativeTables[tableindex].Parameters)
                    {

                        if (parameter.IsGroup || parameter.GroupedParameterValues_Relation.Count == 0)
                        {
                            continue;
                        }
                        int StartingRow = 2;
                        foreach (var GroupNameKvp in parameter.FormattedValues_Relation)
                        {
                            StartingRow++;
                            string GroupName = GroupNameKvp.Key;
                            Parameter groupParameter = ComparativeTables[tableindex].Parameters.SingleOrDefault(p => p.Name == GroupName);
                            PerformTest_Relation(ComparativeTables[tableindex], table, StartingRow, WordTableColumns - 2, parameter, groupParameter);


                            var sortedGroups = parameter.FormattedValues_Relation[GroupName].Keys.OrderBy(key => key).ToList();

                            var SortedGroupDIC = groupParameter.DIC_LablesIfNomainal.Keys.OrderBy(key => key).ToList();

                            foreach (var GroupValue in SortedGroupDIC)
                            {
                                if (sortedGroups.Contains((int)GroupValue))
                                {
                                    if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count > 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];

                                        //string MinMax

                                        string MeanSD = "";
                                        string MedianIQR = "";
                                        string MinMax = "";
                                        //string FormattedMedianMinMax = "";

                                        foreach (var stat in formattedValues)
                                        {
                                            if (stat.Key == "Mean ± StdDev")
                                            {
                                                MeanSD = stat.Value;
                                            }
                                            if (stat.Key == "Median")
                                            {
                                                MedianIQR = stat.Value;
                                            }
                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                            }

                                        }

                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, MinMax);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, MeanSD);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 4, MedianIQR);

                                        StartingRow++;
                                    }



                                    else if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count == 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];
                                        string MinMax = "";
                                        string OnlyValue = "";
                                        foreach (var stat in formattedValues)
                                        {

                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                                string[] values = MinMax.Split(' ');
                                                OnlyValue = values[0];
                                            }

                                        }
                                        table.ApplyHorizontalMerge(StartingRow, 2, 4);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, OnlyValue);
                                        wordObj.SubSuperScriptText(table, StartingRow, 2, Syncfusion.Drawing.Color.White, "#", "Super");
                                        StartingRow++;

                                    }

                                }

                                else
                                {
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 4, "–");
                                    StartingRow++;
                                }

                            }




                        }




                    }







                    //wordObj.FormatTableCustom(table, 11, 0, 0);
                    wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.09f), wordObj.SetColumnWidthInCentimeters(0.09f));

                }
            }
        }
        public void Relation_Layout_DependentNumber_Median_NoIQR()
        {

            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Relation Median No IQR")
                {
                   

                    IWSection section = wordObj.CreatePortraitSection();



                    wordObj.AddRelationTitle(section, ComparativeTables[tableindex].TableName);



                    int Variablerows = Relation_DependentNumber_Rows(ComparativeTables[tableindex]);


                    int WordTableColumns = 7;

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);


                    //Merges
                    wordObj.ApplyRelation_IQR_OuterMerges(table, WordTableRows, WordTableColumns);

                    //Borders
                    wordObj.ApplyRelation_IQR_OuterBorders(table, WordTableRows, WordTableColumns);

                    //Widths
                    wordObj.ApplyRelation_IQR_Widths(table, WordTableRows, WordTableColumns);


                    //Outer Headers
                    wordObj.ApplyRelation_Median_no_IQR_OuterHeaders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);


                    //
                    wordObj.InsertRelation_InnerHeader_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test of Sig.");


                    foreach (var parameter in ComparativeTables[tableindex].Parameters)
                    {

                        if (parameter.IsGroup || parameter.GroupedParameterValues_Relation.Count == 0)
                        {
                            continue;
                        }
                        int StartingRow = 2;
                        foreach (var GroupNameKvp in parameter.FormattedValues_Relation)
                        {
                            StartingRow++;
                            string GroupName = GroupNameKvp.Key;
                            Parameter groupParameter = ComparativeTables[tableindex].Parameters.SingleOrDefault(p => p.Name == GroupName);
                            PerformTest_Relation(ComparativeTables[tableindex], table, StartingRow, WordTableColumns - 2, parameter, groupParameter);


                            var sortedGroups = parameter.FormattedValues_Relation[GroupName].Keys.OrderBy(key => key).ToList();

                            var SortedGroupDIC = groupParameter.DIC_LablesIfNomainal.Keys.OrderBy(key => key).ToList();

                            foreach (var GroupValue in SortedGroupDIC)
                            {
                                if (sortedGroups.Contains((int)GroupValue))
                                {
                                    if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count > 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];

                                        //string MinMax

                                        string MeanSD = "";
                                        string MedianIQR = "";
                                        string MinMax = "";
                                        //string FormattedMedianMinMax = "";

                                        foreach (var stat in formattedValues)
                                        {
                                            if (stat.Key == "Mean ± StdDev")
                                            {
                                                MeanSD = stat.Value;
                                            }
                                            if (stat.Key == "Median")
                                            {
                                                string[] parts = stat.Value.Split(' ');
                                                MedianIQR = parts[0];
                                            }
                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                            }

                                        }

                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, MinMax);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, MeanSD);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 4, MedianIQR);

                                        StartingRow++;
                                    }



                                    else if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count == 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];
                                        string MinMax = "";
                                        string OnlyValue = "";
                                        foreach (var stat in formattedValues)
                                        {

                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                                string[] values = MinMax.Split(' ');
                                                OnlyValue = values[0];
                                            }

                                        }
                                        table.ApplyHorizontalMerge(StartingRow, 2, 4);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, OnlyValue);
                                        wordObj.SubSuperScriptText(table, StartingRow, 2, Syncfusion.Drawing.Color.White, "#", "Super");
                                        StartingRow++;

                                    }

                                }

                                else
                                {
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 4, "–");
                                    StartingRow++;
                                }

                            }




                        }




                    }







                    //wordObj.FormatTableCustom(table, 11, 0, 0);
                    wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.09f), wordObj.SetColumnWidthInCentimeters(0.09f));

                }
            }
        }
        public void Relation_Layout_DependentNumber()
        {

            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Relation")
                {
                    

                    IWSection section = wordObj.CreatePortraitSection();



                    wordObj.AddRelationTitle(section, ComparativeTables[tableindex].TableName);



                    int Variablerows = Relation_DependentNumber_Rows(ComparativeTables[tableindex]);


                    int WordTableColumns = 6;

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);


                    //Merges
                    wordObj.ApplyRelation_OuterMerges(table, WordTableRows, WordTableColumns);

                    //Borders
                    wordObj.ApplyRelation_OuterBorders(table, WordTableRows, WordTableColumns);

                    //Widths
                    wordObj.ApplyRelation_Widths(table, WordTableRows, WordTableColumns);


                    //Outer Headers
                    wordObj.ApplyRelation_OuterHeaders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);


                    //
                    wordObj.InsertRelation_InnerHeader_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test of Sig.");


                    foreach (var parameter in ComparativeTables[tableindex].Parameters)
                    {

                        if (parameter.IsGroup || parameter.GroupedParameterValues_Relation.Count == 0)
                        {
                            continue;
                        }
                        int StartingRow = 2;
                        foreach (var GroupNameKvp in parameter.FormattedValues_Relation)
                        {
                            StartingRow++;
                            string GroupName = GroupNameKvp.Key;
                            Parameter groupParameter = ComparativeTables[tableindex].Parameters.SingleOrDefault(p => p.Name == GroupName);
                            PerformTest_Relation(ComparativeTables[tableindex], table, StartingRow, WordTableColumns - 2, parameter, groupParameter);


                            var sortedGroups = parameter.FormattedValues_Relation[GroupName].Keys.OrderBy(key => key).ToList();

                            var SortedGroupDIC = groupParameter.DIC_LablesIfNomainal.Keys.OrderBy(key => key).ToList();

                            foreach (var GroupValue in SortedGroupDIC)
                            {
                                if (sortedGroups.Contains((int)GroupValue))
                                {
                                    if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count > 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];

                                        //string MinMax

                                        string MeanSD = "";
                                        string Median = "";
                                        string MinMax = "";
                                        string FormattedMedianMinMax = "";

                                        foreach (var stat in formattedValues)
                                        {
                                            if (stat.Key == "Mean ± StdDev")
                                            {
                                                MeanSD = stat.Value;
                                            }
                                            if (stat.Key == "Median")
                                            {

                                                string[] valuessplit = stat.Value.Split(' ');
                                                Median = valuessplit[0];
                                            }
                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                            }

                                        }

                                        FormattedMedianMinMax = Median + " (" + MinMax + ")";
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, MeanSD);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, FormattedMedianMinMax);

                                        StartingRow++;
                                    }



                                    else if (parameter.GroupedParameterValues_Relation[GroupName][GroupValue].Count == 1)
                                    {
                                        Dictionary<string, string> formattedValues = parameter.FormattedValues_Relation[GroupName][GroupValue];
                                        string MinMax = "";
                                        string OnlyValue = "";
                                        foreach (var stat in formattedValues)
                                        {

                                            if (stat.Key == "Min-Max")
                                            {
                                                MinMax = stat.Value;
                                                string[] values = MinMax.Split(' ');
                                                OnlyValue = values[0];
                                            }

                                        }
                                        table.ApplyHorizontalMerge(StartingRow, 2, 3);
                                        wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, OnlyValue);
                                        wordObj.SubSuperScriptText(table, StartingRow, 2, Syncfusion.Drawing.Color.White, "#", "Super");
                                        StartingRow++;

                                    }

                                }

                                else
                                {
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 2, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, StartingRow, 3, "–");
                                    StartingRow++;
                                }

                            }




                        }




                    }







                    //wordObj.FormatTableCustom(table, 11, 0, 0);
                    wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.09f), wordObj.SetColumnWidthInCentimeters(0.09f));

                }
            }
        }

        public void PerformTest_Relation(ComparativeTable comparativeTable, IWTable table, int addrow, int addColumn, Parameter parameter, Parameter GroupParameter)
        {
            pythonStat.InitPython();
            int numberofgroups = parameter.FormattedValues_Relation[GroupParameter.Name].Keys.Count;

            var SortedGroupsValues = parameter.GroupedParameterValues_Relation[GroupParameter.Name].Keys;
            int ValueOne = 0;
            foreach (var key in SortedGroupsValues)
            {
                if (parameter.GroupedParameterValues_Relation[GroupParameter.Name][key].Count == 1)
                {
                    ValueOne++;
                }
            }

            if (numberofgroups - ValueOne <= 2)
            {
                List<double> group1Values = new List<double>();
                List<double> group2Values = new List<double>();


                int count = 0;
                foreach (var kvp in parameter.GroupedParameterValues_Relation[GroupParameter.Name])
                {
                    List<double> currentValues = kvp.Value;

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

                if (parameter.NormalOrAbnormal == "Normal")
                {
                    comparativeTable.TestsDone.Add("tstudent");


                    string[] values = manual.StudentT_Unpaired(group1Values, group2Values);
                    values[0] = "t=" + Convert.ToChar(11) + values[0];
                    wordObj.InsertTest_P(table, addrow, addColumn, values);

                }

                else if (parameter.NormalOrAbnormal == "Abnormal")
                {
                    comparativeTable.TestsDone.Add("U");

                    string[] values = manual.UTest(group1Values, group2Values);

                    string[] Threezerosarray = values[0].Split('.');
                    string threezeros = "000";

                    try
                    {
                        if (Threezerosarray[1] == threezeros)
                        {
                            double doubleValue = double.Parse(values[0]); // Assuming values[0] is a string containing a double
                            values[0] = doubleValue.ToString("0.0");
                        }
                    }
                    catch (Exception)
                    {

                    }
                    finally
                    {
                        values[0] = "U=" + Convert.ToChar(11) + values[0];
                        wordObj.InsertTest_P(table, addrow, addColumn, values);
                    }



                }

            }

            if (numberofgroups - ValueOne > 2)
            {
                if (parameter.NormalOrAbnormal == "Normal")
                {
                    comparativeTable.TestsDone.Add("FAnova");

                    AnovaTestResult anovaTest = new AnovaTestResult();

                    anovaTest = manual.Fanova_Relation(parameter, GroupParameter);
                    string[] values = { anovaTest.TestValue, anovaTest.PValue };





                    values[0] = "F=" + Convert.ToChar(11) + values[0];
                    wordObj.InsertTest_P(table, addrow, addColumn, values);

                }
                else if (parameter.NormalOrAbnormal == "Abnormal")
                {

                    comparativeTable.TestsDone.Add("H");


                    AnovaTestResult anovaTestResult = new AnovaTestResult();

                    anovaTestResult = manual.Kruskal_H_Relation(parameter, GroupParameter);
                    string[] values = { anovaTestResult.TestValue, anovaTestResult.PValue };

                    values[0] = "H=" + Convert.ToChar(11) + values[0];
                    wordObj.InsertTest_P(table, addrow, addColumn, values);


                }

            }


        }
        public bool DetermineHasScale(ComparativeTable comparativeTable)
        {
            bool HasScale = false;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.NominalOrScale == "Scale")
                {
                    HasScale = true;
                    break;
                }

            }
            return HasScale;
        }
        public bool DetermineHasNominal(ComparativeTable comparativeTable)
        {
            bool HasNominal = false;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.NominalOrScale == "Nominal")
                {
                    HasNominal = true;
                    break;
                }

            }
            return HasNominal;
        }
        public int Count_Rows_Descriptive(ComparativeTable table)
        {

            int rowCount = 0;
            bool HasNominal = DetermineHasNominal(table);
            foreach (var parameter in table.Parameters)
            {
                if (HasNominal)
                {
                    if (parameter.NominalOrScale == "Nominal")
                    {
                        int distinctValuesCount = parameter.DIC_LablesIfNomainal.Keys.Count;
                        rowCount += distinctValuesCount + 1;

                    }
                    else if (parameter.NominalOrScale == "Scale")
                    {
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

        public void TestPic(string filePath)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // Get path to desktop
            using (FileStream docStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Load an existing Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
                {
                    // Create a new instance of DocIORenderer
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        // Convert an entire Word document to images
                        Stream[] imageStreams = wordDocument.RenderAsImages();
                        for (int i = 0; i < imageStreams.Length; i++)
                        {
                            string outputPath = Path.Combine(desktopPath, "WordToImage_" + i + ".jpeg");
                            using (FileStream fileStreamOutput = File.Create(outputPath))
                            {
                                imageStreams[i].CopyTo(fileStreamOutput);
                            }
                        }
                    }
                }
            }
        }
        public void DrawDescriptiveTable()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {

                if (ComparativeTables[tableindex].FormatType == "Descriptive")
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
                    else if (HasNominal)
                    {
                        WordTableColumns = 3;
                    }


                    int WordTableRows = 1 + Variablerows;



                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);


                    wordObj.GeneralTableFormat(table);



                    wordObj.ApplyGeneralDescritiveBorders(table, WordTableRows, WordTableColumns);

                    wordObj.SetDescriptiveWidths(table, WordTableRows, HasNominal);


                    wordObj.Add_GeneralHeaders_Descriptive(table, HasNominal);

                    DescriptiveParameterBorders_text(table, WordTableRows, WordTableColumns, ComparativeTables[tableindex], HasNominal);


                    InsertData_Descriptive(table, ComparativeTables[tableindex], HasNominal);

                    //pythonStat.InitPython();

                    //var contingencyTable = new List<List<int>>();
                    //foreach (var par in ComparativeTables[tableindex].Parameters)
                    //{
                    //    if(par.IsGroup)
                    //    {
                    //        continue;
                    //    }

                    //    var allValues = par.GroupedParameterValues.Values.SelectMany(x => x).Distinct().ToList();

                    //    // Prepare contingency table

                    //    foreach (var value in allValues)
                    //    {
                    //        var categoryCounts = new List<int>();
                    //        foreach (var groupData in par.GroupedParameterValues.Values)
                    //        {
                    //            categoryCounts.Add(groupData.Count(x => x == value));
                    //        }
                    //        contingencyTable.Add(categoryCounts);
                    //    }
                    //}

                    //pythonStat.McNemarTest(contingencyTable);

                    wordObj.FormatTable(table, 12);
                }







            }
        }

        public void DescriptiveParameterBorders_text(IWTable table, int WordTableRows, int WordTableColumns, ComparativeTable comparativeTable, bool hasnominal)
        {
            int currentrow = 1;
            int count = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {



                if (parameter.NominalOrScale == "Nominal")
                {
                    count = parameter.DIC_LablesIfNomainal.Keys.Count + 1;

                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    count = 4;

                }

                if (hasnominal)
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
                                    table.ApplyHorizontalMerge(currentrow, InsertLowerN, InsertLowerN + 1);
                                    int parameterGroupCount = parameter.GroupedParameterValues[group].Count;
                                    string InsertedN = "(n = " + parameterGroupCount + ")";
                                    wordObj.AddParaCombined(table, currentrow, InsertLowerN, InsertedN, true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);
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


                else if (!hasnominal)
                {
                    wordObj.AddPara_NoCenter(table, currentrow, 0, parameter.Name);

                    if (parameter.hasLowerN)
                    {
                        wordObj.AddParaCombined(table, currentrow, 0, "(n = " + parameter.ParameterValues.Count + ")", true, false, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);
                    }

                    if (currentrow < WordTableRows - 1)
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




        public void InsertData_Descriptive(IWTable table, ComparativeTable comparativeTable, bool hasnominal)
        {
            if (hasnominal)
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
                                //MessageBox.Show(frequency);
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
                                    //MessageBox.Show(percentage);
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
            else if (!hasnominal)
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

        public int Count_Rows_Columns_Descriptive_Periods(ComparativeTable comparativeTable, int numberofperiods)
        {
            int rowcount = 0;
            bool hasNominal = false;


            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }

                if (parameter.NominalOrScale == "Nominal")
                {
                    rowcount += parameter.DIC_LablesIfNomainal.Keys.Count + 1;
                    hasNominal = true;
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    rowcount += 4;
                }

            }

            rowcount = rowcount / numberofperiods;

            if (hasNominal)
            {
                rowcount++;
            }











            return rowcount;
        }

        public bool Tablehasnominal_fn(ComparativeTable comparativeTable)
        {

            bool hasNominal = false;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }
                if (parameter.NominalOrScale == "Nominal")
                {
                    hasNominal = true;
                }
            }

            return hasNominal;
        }
        public void DescriptivePeriodsNoTest()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Descriptive Periods No Test")
                {
                    IWSection section = wordObj.CreatePortraitSection();

                    

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, 1);

                    int PeriodCount = 0;

                    if (!string.IsNullOrWhiteSpace(txt_PeriodCount.Text) && int.TryParse(txt_PeriodCount.Text, out PeriodCount))
                    {
                        PeriodCount = int.Parse(txt_PeriodCount.Text);
                    }
                    else
                    {
                        MessageBox.Show("Please enter a valid period count");
                    }

                    int Variablerows = Count_Rows_Columns_Descriptive_Periods(ComparativeTables[tableindex], PeriodCount);






                    int WordTableColumns = 1 + (PeriodCount * 2);

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);

                    //Merges
                    wordObj.Apply_Descriptive_periodsNoTest_OuterMerges(table, WordTableRows, WordTableColumns, PeriodCount);

                    ///Stopped here (below is false)

                    //Borders
                    //wordObj.ApplyRelation_OuterBorders_Pathology(table, WordTableRows, WordTableColumns);
                    //Apply_Descriptive_OuterBorders_PeriodsNoTest()

                    ////Widths
                    //wordObj.ApplyRelation_Widths_Pathology(table, WordTableRows, WordTableColumns);


                    ////Outer Headers
                    //wordObj.ApplyRelation_OuterHeaders_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    //wordObj.InsertRelation_InnerHeader_Merges_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);

                    //wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test of Sig.");









                    wordObj.LeftAndRightCellMarginCustom(table, 0, 0);
                    wordObj.FormatTableCustom(table, 12f, 0, 0);

                }
            }
        }

        public void DescriptivePeriodsTest()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Descriptive Periods Test")
                {
                    IWSection section = wordObj.CreatePortraitSection();

                    

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, 1);

                    int PeriodCount = 0;

                    if (!string.IsNullOrWhiteSpace(txt_PeriodCount.Text) && int.TryParse(txt_PeriodCount.Text, out PeriodCount))
                    {
                        PeriodCount = int.Parse(txt_PeriodCount.Text);
                    }
                    else
                    {
                        MessageBox.Show("Please enter a valid period count");
                    }

                    int Variablerows = Count_Rows_Columns_Descriptive_Periods(ComparativeTables[tableindex], PeriodCount);




                    // 1 for names 2 for tests

                    int WordTableColumns = 1 + (PeriodCount * 2) + 2;

                    int WordTableRows = 2 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);

                    //Merges
                    wordObj.Apply_Descriptive_periodsNoTest_OuterMerges(table, WordTableRows, WordTableColumns, PeriodCount);

                    ///Stopped here (below is false)

                    //Borders
                    //wordObj.ApplyRelation_OuterBorders_Pathology(table, WordTableRows, WordTableColumns);
                    //Apply_Descriptive_OuterBorders_PeriodsNoTest()

                    ////Widths
                    //wordObj.ApplyRelation_Widths_Pathology(table, WordTableRows, WordTableColumns);


                    ////Outer Headers
                    //wordObj.ApplyRelation_OuterHeaders_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    //wordObj.InsertRelation_InnerHeader_Merges_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);

                    //wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test of Sig.");









                    wordObj.LeftAndRightCellMarginCustom(table, 0, 0);
                    wordObj.FormatTableCustom(table, 12f, 0, 0);

                }
            }
        }
        public int Groups_Side_Periods_Up_Rows(ComparativeTable comparativeTable, int parametercount)
        {
            int rows = 0;

            Parameter CurrentParameterNotGroup = null;

            Parameter GroupParameter = null;

            foreach (var CurrentParameter in comparativeTable.Parameters)
            {
                if (CurrentParameter.IsGroup)
                {
                    GroupParameter = CurrentParameter;
                    continue;
                }
                CurrentParameterNotGroup = CurrentParameter;
            }



            for (int i = 0; i < parametercount / 3; i++)
            {
                foreach (var GroupKeyValue in GroupParameter.DIC_LablesIfNomainal.Keys)
                {
                    rows += 5;
                }
                rows++;
            }




            return rows;
        }
        public int Groups_Side_Periods_Up_threePeriods_ParameterCount(ComparativeTable comparativeTable)
        {
            int ParaCount = 0;

            foreach (var CurrentParameter in comparativeTable.Parameters)
            {
                if (CurrentParameter.IsGroup)
                    continue;

                ParaCount++;
            }

            return ParaCount;

        }

        public void Groups_Side_Periods_Up_threePeriods_periodstest(IWTable table, ComparativeTable comparativeTable, List<Parameter> CurrentParameters, int insertrow, Parameter GroupParameter, int GroupColIndex, int WordTableColumns, ref List<int> RowsToremoveSigPeriods)
        {
            List<Dictionary<int, List<double>>> CorrectDataParameters = new List<Dictionary<int, List<double>>>();

            string NormalOrAbnormal = CurrentParameters[0].NormalOrAbnormal;

            foreach (var CurrentParameter in CurrentParameters)
            {
                Dictionary<int, List<double>> GroupedCorrectData = ExcelFunctions.GetDataWithMissingValues_withGroup(worksheet, CurrentParameter.Name, GroupParameter.Name, GroupColIndex);
                CorrectDataParameters.Add(GroupedCorrectData);
            }

            HashSet<int> uniqueGroupKeys = new HashSet<int>();
            foreach (var dict in CorrectDataParameters)
            {
                foreach (var key in dict.Keys)
                {
                    uniqueGroupKeys.Add(key);
                }
            }




            int loopCounter = 0;

            foreach (int groupKey in uniqueGroupKeys)
            {
                List<List<double>> valuesLists = new List<List<double>>();



                for (int i = 0; i < CurrentParameters.Count; i++)
                {
                    if (CorrectDataParameters[i].TryGetValue(groupKey, out List<double> values))
                    {
                        valuesLists.Add(new List<double>(values)); // Clone the list before cleaning
                    }
                    else
                    {
                        valuesLists.Add(new List<double>()); // Maintain list count alignment
                    }
                }

                GeneralFunctions.RemoveInvalidEntries(ref valuesLists);

                List<Parameter> parametersForGroup = new List<Parameter>();

                for (int i = 0; i < CurrentParameters.Count; i++)
                {
                    if (valuesLists[i].Count > 0) // Ensure there's still valid data
                    {
                        Parameter newParameter = new Parameter
                        {
                            Name = CurrentParameters[i].Name, // Copy name from original list
                            ParameterValues = valuesLists[i]  // Assign cleaned values
                        };

                        parametersForGroup.Add(newParameter);
                    }
                }



                if (NormalOrAbnormal == "Normal")
                {
                    AnovaTestResult testresult = pythonStat.RepeatedMeasuresAnovaBoth(parametersForGroup);

                    bool pvalueSig = generalFunctions.PvalueHasSig(testresult.PValue);



                    int rowNumber = (loopCounter % 2 == 0) ? (insertrow - 9) : (insertrow - 4);
                    int rowNumber_Sig_Periods = (loopCounter % 2 == 0) ? (insertrow - 6) : (insertrow - 1);

                    if (comparativeTable.AllNormal())
                    {
                        wordObj.AddParaCombined(table, rowNumber, WordTableColumns - 2, testresult.TestValue, false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                    }
                    else
                    {
                        wordObj.AddParaCombined(table, rowNumber, WordTableColumns - 2, "F=" + Convert.ToChar(11) + testresult.TestValue, false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                    }


                    wordObj.Addpara_CenterNoBOLD(table, rowNumber, WordTableColumns - 1, testresult.PValue);


                    if (!pvalueSig)
                    {
                        RowsToremoveSigPeriods.Add(rowNumber_Sig_Periods);
                    }
                    else
                    {
                        wordObj.SubSuperScriptText(table, rowNumber, WordTableColumns - 2, Syncfusion.Drawing.Color.Yellow, "*", "Super");
                        wordObj.SubSuperScriptText(table, rowNumber, WordTableColumns - 1, Syncfusion.Drawing.Color.Empty, "*", "Super");

                        int pnumber = 1;
                        int totalComparisons = testresult.PairwiseComparisons.Count;
                        StringBuilder pValuesText = new StringBuilder();

                        foreach (var Values in testresult.PairwiseComparisons)
                        {
                            if (Values.Length > 2) // Ensure index 2 exists
                            {
                                string pvalueString = Values[2];

                                bool psig = generalFunctions.PvalueHasSig(pvalueString);

                                // Try parsing the p-value; if it's a valid number, format it
                                string formattedPValue;
                                if (double.TryParse(pvalueString, out double parsedValue))
                                {
                                    formattedPValue = parsedValue.ToString("0.000"); // Format to 3 decimal places
                                }
                                else
                                {
                                    formattedPValue = pvalueString; // Keep "<0.001" as is
                                }

                                // **Insert "p" first**
                                wordObj.AddParaCombined_New(table, rowNumber_Sig_Periods, 3, "p", false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                                // **Now insert subscripted number**
                                wordObj.SubSuperScriptText(table, rowNumber_Sig_Periods, 3, Syncfusion.Drawing.Color.Transparent, pnumber.ToString(), "Sub");

                                // **Now insert the p-value after the subscript**
                                string finalText = (formattedPValue == "<0.001") ? "<0.001" : $"={formattedPValue}";
                                wordObj.AddParaCombined_New(table, rowNumber_Sig_Periods, 3, finalText, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                                if (psig)
                                {
                                    wordObj.SubSuperScriptText(table, rowNumber_Sig_Periods, 3, Syncfusion.Drawing.Color.Empty, "*", "Super");
                                }
                                // Add comma if it's NOT the last value
                                if (pnumber < totalComparisons)
                                {
                                    wordObj.AddParaCombined_New(table, rowNumber_Sig_Periods, 3, ",", false, true, Syncfusion.Drawing.Color.Transparent, Syncfusion.Drawing.Color.Black);
                                }


                                pnumber++;
                            }
                        }

                        //wordObj.AddPara_Center(table, rowNumber_Sig_Periods, WordTableColumns-2, "");
                        table.Rows[rowNumber_Sig_Periods].Cells[WordTableColumns - 2].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;

                        //wordObj.AddPara_Center(table, rowNumber_Sig_Periods, WordTableColumns - 1, "");
                        table.Rows[rowNumber_Sig_Periods].Cells[WordTableColumns - 1].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;



                    }



                    loopCounter++;

                    //wordObj.Addpara_CenterNoBOLD(table , )

                }
                else if (NormalOrAbnormal == "Abnormal")
                {
                    List<List<double>> Datalist = new List<List<double>>();



                    foreach (var parameter in parametersForGroup)
                    {
                        Datalist.Add(parameter.ParameterValues);
                    }

                    GeneralFunctions.RemoveInvalidEntries(ref Datalist);



                    AnovaTestResult testresult = pythonStat.PerformFriedmanWithDunnTest(Datalist);

                    bool pvalueSig = generalFunctions.PvalueHasSig(testresult.PValue);



                    int rowNumber = (loopCounter % 2 == 0) ? (insertrow - 9) : (insertrow - 4);
                    int rowNumber_Sig_Periods = (loopCounter % 2 == 0) ? (insertrow - 6) : (insertrow - 1);

                    if (comparativeTable.AllAbnormal())
                    {
                        wordObj.AddParaCombined(table, rowNumber, WordTableColumns - 2, testresult.TestValue, false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                    }
                    else
                    {
                        wordObj.AddParaCombined(table, rowNumber, WordTableColumns - 2, "Fr=" + Convert.ToChar(11) + testresult.TestValue, false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                    }


                    wordObj.Addpara_CenterNoBOLD(table, rowNumber, WordTableColumns - 1, testresult.PValue);


                    if (!pvalueSig)
                    {
                        RowsToremoveSigPeriods.Add(rowNumber_Sig_Periods);
                    }
                    else
                    {
                        wordObj.SubSuperScriptText(table, rowNumber, WordTableColumns - 2, Syncfusion.Drawing.Color.Yellow, "*", "Super");
                        wordObj.SubSuperScriptText(table, rowNumber, WordTableColumns - 1, Syncfusion.Drawing.Color.Empty, "*", "Super");

                        int pnumber = 1;
                        int totalComparisons = testresult.PairwiseComparisons.Count;
                        StringBuilder pValuesText = new StringBuilder();

                        foreach (var Values in testresult.PairwiseComparisons)
                        {
                            if (Values.Length > 1) // Ensure index 2 exists
                            {
                                string pvalueString = Values[1];

                                bool psig = generalFunctions.PvalueHasSig(pvalueString);

                                // Try parsing the p-value; if it's a valid number, format it
                                string formattedPValue;
                                if (double.TryParse(pvalueString, out double parsedValue))
                                {
                                    formattedPValue = parsedValue.ToString("0.000"); // Format to 3 decimal places
                                }
                                else
                                {
                                    formattedPValue = pvalueString; // Keep "<0.001" as is
                                }

                                // **Insert "p" first**
                                wordObj.AddParaCombined_New(table, rowNumber_Sig_Periods, 3, "p", false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                                // **Now insert subscripted number**
                                wordObj.SubSuperScriptText(table, rowNumber_Sig_Periods, 3, Syncfusion.Drawing.Color.Transparent, pnumber.ToString(), "Sub");

                                // **Now insert the p-value after the subscript**
                                string finalText = (formattedPValue == "<0.001") ? "<0.001" : $"={formattedPValue}";
                                wordObj.AddParaCombined_New(table, rowNumber_Sig_Periods, 3, finalText, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);


                                if (psig)
                                {
                                    wordObj.SubSuperScriptText(table, rowNumber_Sig_Periods, 3, Syncfusion.Drawing.Color.Empty, "*", "Super");
                                }
                                // Add comma if it's NOT the last value
                                if (pnumber < totalComparisons)
                                {
                                    wordObj.AddParaCombined_New(table, rowNumber_Sig_Periods, 3, ",", false, true, Syncfusion.Drawing.Color.Transparent, Syncfusion.Drawing.Color.Black);
                                }


                                pnumber++;
                            }
                        }

                        wordObj.AddPara_Center(table, rowNumber_Sig_Periods, WordTableColumns - 2, "");
                        table.Rows[rowNumber_Sig_Periods].Cells[WordTableColumns - 2].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;

                        wordObj.AddPara_Center(table, rowNumber_Sig_Periods, WordTableColumns - 1, "");
                        table.Rows[rowNumber_Sig_Periods].Cells[WordTableColumns - 1].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;


                    }







                    loopCounter++;
                }




            }

        }
        public void Groups_Side_Periods_Up_threePeriods()
        {
            pythonStat.InitPython();
            string filepath = ExcelFunctions.filepath;
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);

            worksheet = workbook.Worksheets[0];
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Groups Side Periods Up")
                {
                    List<int> RowsToremoveSigPeriods = new List<int>();

                    

                    int ParameterCount = Groups_Side_Periods_Up_threePeriods_ParameterCount(ComparativeTables[tableindex]);

                    int GroupColIndex = newFindGroupColumnIndex(ComparativeTables[tableindex]);

                    IWSection section = wordObj.CreatePortraitSection();

                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = Groups_Side_Periods_Up_Rows(ComparativeTables[tableindex], ParameterCount);

                    int WordTableRows = 1 + Variablerows;

                    int WordTableColumns = 8;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.GeneralTableFormat(table);


                    wordObj.Groups_Side_Periods_Up_threePeriods_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups, ParameterCount);

                    wordObj.Groups_Side_Periods_Up_threePeriods_Borders(table, WordTableRows, WordTableColumns);

                    wordObj.Groups_Side_Periods_Up_threePeriods_Widths(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);


                    wordObj.Groups_Side_Periods_Up_threePeriods_Headers(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, ParameterCount);


                    //the code applies data when group parameter is first and 3 periods (Note)***

                    //Insert Data
                    int parameterCounter = 1;
                    int row = 1;
                    int tempctr = 0;


                    Parameter GroupPara = ComparativeTable.GetGroupParamter(ComparativeTables[tableindex]);
                    var SortedGroupKeysAll = GroupPara.ParameterValues.Distinct().OrderBy(key => key).ToList();

                    //BIG parameter (3 parameters)
                    for (int paractr = 0; paractr < ParameterCount / 3; paractr++)
                    {
                        int col = 3;
                        //small parameters DAY 1 , etc
                        List<Parameter> periodsparameter = new List<Parameter>();
                        for (int SubParactr = 0; SubParactr < 3; SubParactr++)
                        {
                            periodsparameter.Add(ComparativeTables[tableindex].Parameters[parameterCounter]);
                            row = 1 + tempctr;
                            var sortedKeys = ComparativeTables[tableindex].Parameters[parameterCounter].FormattedValues.Keys.OrderBy(key => key).ToList();

                            foreach (int groupValue in SortedGroupKeysAll)
                            {
                                if (sortedKeys.Contains(groupValue))
                                {

                                    //Add n
                                    int n = ComparativeTables[tableindex].Parameters[parameterCounter].GroupedParameterValues[groupValue].Count;

                                    int groupn = GroupPara.GetValueCount(groupValue);

                                    if (n == groupn)
                                    {
                                        wordObj.AddPara_Center(table, row, col, "(n = " + n + ")");
                                    }
                                    else
                                    {
                                        wordObj.AddParaCombined_New(table, row, col, "(n = " + n + ")", true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red); ;
                                    }




                                    Dictionary<string, string> formattedValues = ComparativeTables[tableindex].Parameters[parameterCounter].FormattedValues[groupValue];

                                    int currentRow = row + 1;

                                    foreach (var stat in formattedValues)
                                    {
                                        if (stat.Key == "Mean ± StdDev")
                                        {
                                            if (ComparativeTables[tableindex].Parameters[parameterCounter].NormalOrAbnormal == "Abnormal")
                                            {
                                                wordObj.AddParaCombined(table, currentRow, col, stat.Value, false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                                            }
                                            else
                                            {
                                                wordObj.Addpara_CenterNoBOLD(table, currentRow, col, stat.Value);
                                            }
                                        }
                                        else if (stat.Key == "Median")
                                        {
                                            if (ComparativeTables[tableindex].Parameters[parameterCounter].NormalOrAbnormal == "Normal")
                                            {
                                                wordObj.AddParaCombined(table, currentRow, col, stat.Value, false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                                            }
                                            else
                                            {
                                                wordObj.Addpara_CenterNoBOLD(table, currentRow, col, stat.Value);
                                            }
                                        }
                                        else
                                        {
                                            wordObj.Addpara_CenterNoBOLD(table, currentRow, col, stat.Value);
                                        }

                                        currentRow++;
                                    }


                                }
                                else
                                {
                                    wordObj.AddParaCombined_New(table, row, col, "(n = " + 0 + ")", true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Red);

                                    wordObj.Addpara_CenterNoBOLD(table, row + 1, col, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, row + 2, col, "–");
                                    wordObj.Addpara_CenterNoBOLD(table, row + 3, col, "–");
                                }
                                row = row + 5;


                            }

                            PerformTest_Groups_Side(table, ComparativeTables[tableindex], ComparativeTables[tableindex].Parameters[parameterCounter], row, col, WordTableColumns);


                            col++;
                            parameterCounter++;
                        }



                        Groups_Side_Periods_Up_threePeriods_periodstest(table, ComparativeTables[tableindex], periodsparameter, row, GroupPara, GroupColIndex, WordTableColumns, ref RowsToremoveSigPeriods);
                        tempctr = tempctr + 11;
                    }

                    //table.Rows.RemoveAt(5);
                    //table.Rows.RemoveAt(10);



                    wordObj.LeftAndRightCellMarginCustom(table, 0.09f, 0.09f);
                    wordObj.FormatTableCustom(table, 10, 0, 0);

                    int rowindexSubtract = 0;
                    foreach (var rowindex in RowsToremoveSigPeriods)
                    {
                        table.Rows.RemoveAt(rowindex - rowindexSubtract);
                        rowindexSubtract++;
                    }
                }
            }

        }
        GeneralFunctions generalFunctions = new GeneralFunctions();
        public void PerformTest_Groups_Side(IWTable table, ComparativeTable comparativeTable, Parameter CurrenParameter, int InsertRow, int InsertColumn, int WordTableColumns)
        {
            //int GroupsCountPara = CurrenParameter.GroupedParameterValues.Keys.Count;

            int numberofgroupsTrue = CurrenParameter.FormattedValues.Keys.Count;
            var SortedGroupsValues = CurrenParameter.GroupedParameterValues.Keys;

            int ValueOne = 0;
            foreach (var key in SortedGroupsValues)
            {
                if (CurrenParameter.GroupedParameterValues[key].Count == 1)
                {
                    ValueOne++;
                }
            }

            if (numberofgroupsTrue - ValueOne <= 2)
            {
                if (CurrenParameter.NormalOrAbnormal == "Normal")
                {
                    List<double> group1Values = new List<double>();
                    List<double> group2Values = new List<double>();

                    SplitGroupedParameterValues(CurrenParameter, out group1Values, out group2Values);

                    string[] values = manual.StudentT_Unpaired(group1Values, group2Values);
                    bool pvalueSig = generalFunctions.PvalueHasSig(values[1]);

                    wordObj.AddParaCombined_New(table, InsertRow, InsertColumn, values[0], false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);

                    if (pvalueSig)
                    {
                        wordObj.SubSuperScriptText(table, InsertRow, InsertColumn, Syncfusion.Drawing.Color.Yellow, "*", "Super");
                    }

                    wordObj.AddParaCombined_New(table, InsertRow, InsertColumn, " (" + values[1], false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                    if (pvalueSig)
                    {
                        wordObj.SubSuperScriptText(table, InsertRow, InsertColumn, Syncfusion.Drawing.Color.Empty, "*", "Super");
                    }

                    wordObj.AddParaCombined_New(table, InsertRow, InsertColumn, ")", false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                    // wordObj.AddPara_Center(table, InsertRow, WordTableColumns - 2, "");
                    table.Rows[InsertRow].Cells[WordTableColumns - 2].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;

                    //wordObj.AddPara_Center(table, InsertRow, WordTableColumns - 1, "");
                    table.Rows[InsertRow].Cells[WordTableColumns - 1].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;
                }
                else if (CurrenParameter.NormalOrAbnormal == "Abnormal")
                {
                    List<double> group1Values = new List<double>();
                    List<double> group2Values = new List<double>();

                    SplitGroupedParameterValues(CurrenParameter, out group1Values, out group2Values);

                    string[] values = manual.UTest(group1Values, group2Values);

                    bool pvalueSig = generalFunctions.PvalueHasSig(values[1]);

                    wordObj.AddParaCombined_New(table, InsertRow, InsertColumn, values[0], false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);

                    if (pvalueSig)
                    {
                        wordObj.SubSuperScriptText(table, InsertRow, InsertColumn, Syncfusion.Drawing.Color.Yellow, "*", "Super");
                    }

                    wordObj.AddParaCombined_New(table, InsertRow, InsertColumn, " (" + values[1], false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                    if (pvalueSig)
                    {
                        wordObj.SubSuperScriptText(table, InsertRow, InsertColumn, Syncfusion.Drawing.Color.Empty, "*", "Super");
                    }

                    wordObj.AddParaCombined_New(table, InsertRow, InsertColumn, ")", false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                    //wordObj.AddParaCombined(table, InsertRow, WordTableColumns - 2, "", false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                    table.Rows[InsertRow].Cells[WordTableColumns - 2].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;

                    //ordObj.AddParaCombined(table, InsertRow, WordTableColumns - 1, "" , false , true , Syncfusion.Drawing.Color.Empty , Syncfusion.Drawing.Color.Black);
                    table.Rows[InsertRow].Cells[WordTableColumns - 1].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;
                }
            }


        }

        public int  letter_table_group_Subgroups_Rows(ComparativeTable comparativeTable)
        {
            int rows = 0;

            Parameter SubGroup = null;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.IsSubGroup)
                {
                    SubGroup = parameter;

                    break;
                }
            }


            foreach (var SubGroupValue in SubGroup.DIC_LablesIfNomainal.Keys)
            {
                rows = rows + 2;
            }



            rows = rows + 3;

            return rows;


        }
        public int letter_table_group_Subgroups_Cols(ComparativeTable comparativeTable)
        {
            int columns = 0;


            Parameter Group = null;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    Group = parameter;

                    break;
                }
            }


            foreach (var GroupValue in Group.DIC_LablesIfNomainal.Keys)
            {
                columns++;
            }


            columns = columns + 3;

            return columns;
        }

        
        public void letter_table_group_Subgroups()
        {

            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Letter Groups SubGroups")
                {
                    IWSection section = wordObj.CreatePortraitSection();

                    wordObj.AddRelationTitle(section, ComparativeTables[tableindex].TableName);


                    int WordTableRows = letter_table_group_Subgroups_Rows(ComparativeTables[tableindex]);


                    int WordTableColumns = letter_table_group_Subgroups_Cols(ComparativeTables[tableindex]);

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    //wordObj.GeneralTableFormat(table);

                    wordObj.GeneralLetterTableFormat(table);
                    Parameter Groupparameter = ComparativeTable.GetGroupParamter(ComparativeTables[tableindex]);

                    //Merges
                    wordObj.letter_table_group_Subgroups_merges(table, WordTableRows, WordTableColumns , Groupparameter);

                    //Borders
                    wordObj.letter_table_group_Subgroups_borders(table, WordTableRows, WordTableColumns);


                    //Widths
                    wordObj.letter_table_group_Subgroups_widths(table, WordTableRows, WordTableColumns);


                    //Outer Headers
                    wordObj.letter_table_group_Subgroups_header(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);


                }
            }
          


            //        //Outer Headers
            //        wordObj.ApplyRelation_IQR_OuterHeaders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);


            //        //
            //        wordObj.InsertRelation_InnerHeader_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



            //        wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, "Test of Sig.");


        }

        public static void PrintPercentiles(List<StatTable> tables)
        {
            foreach (var table in tables)
            {
                string output = $"Table: {table.TableName}\n";

                foreach (var param in table.Parameters)
                {
                    output += $"\nParameter: {param.Name}";

                    // --- Total percentiles ---
                    if (param.PercentileStats?.TotalPercentiles?.Count > 0)
                    {
                        output += "\n  Total Percentiles:";
                        foreach (var kv in param.PercentileStats.TotalPercentiles)
                        {
                            output += $"\n    {kv.Key}% = {kv.Value}";
                        }
                    }

                    // --- Group percentiles ---
                    if (param.PercentileStats?.GroupPercentiles?.Count > 0)
                    {
                        output += "\n  Group Percentiles:";
                        foreach (var groupKvp in param.PercentileStats.GroupPercentiles)
                        {
                            string groupParam = groupKvp.Key;
                            foreach (var labelKvp in groupKvp.Value)
                            {
                                string label = labelKvp.Key;
                                output += $"\n    {groupParam} = {label}";

                                foreach (var pKvp in labelKvp.Value)
                                {
                                    output += $"\n      {pKvp.Key}% = {pKvp.Value}";
                                }
                            }
                        }
                    }

                    output += "\n";
                }

                MessageBox.Show(output, "Parsed Percentiles");
            }
        }

        public static void ForceKillSPSS()
        {
            try
            {
                // Find all SPSS processes by name (without .exe)
                var processes = Process.GetProcessesByName("spsswin.exe");

                foreach (var process in processes)
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(); // wait until it's really closed
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error killing SPSS process: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ForceKillSPSS failed: " + ex.Message);
            }
        }

        private void btn_Done_Click(object sender, EventArgs e)
        {
            //ComparativeBasic();
            string outputText = SPSSUnifiedRunner.RunUnifiedSyntaxAndGetResult(StatTables, out _);

            SPSSUnifiedRunner.ParseUnifiedOutput_Crosstabs(outputText, StatTables);

            //SPSSUnifiedRunner.ParseUnifiedOutput_Descriptives(outputText, StatTables);

            var lines = outputText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            SPSSUnifiedRunner.ParsePercentiles(StatTables, lines);

            //PrintPercentiles(StatTables);
            //ForceKillSPSS();

            //List<string> groupLabels = new List<string>() { "Patient" , "Control" , "3.00" , "4.00" };
            //List<string> vars = new  List<string>() { "Age" };

            //var result = SPSSUnifiedRunner.ParseTukeyTukeyHingesOnly(outputText, vars, groupLabels, "Groups");




            //pythonStat.InitPython();

            document = wordObj.InitWord();

            NewComparativeGroups_Fn();

            letter_table_group_Subgroups();

            Anesthesia_Periods();

            ComparativeTableGroups_2_periods();

            //ComparativeTableGroups_Layout();

            

            ComparativeTablePeriodsUp_Groups_Layout();

            Groups_Side_Periods_Up_threePeriods();

            PaperComparative_Layout();



            Pathology_Layout_NoIQR();

            Pathology_Layout();


            Relation_Layout_DependentNumber_IQR();

            Relation_Layout_DependentNumber_Median_NoIQR();

            Relation_Layout_DependentNumber();


            Relation_Layout_DependentNumber_Pathology();

            DrawDescriptiveTable();

            DescriptivePeriodsNoTest();

            DescriptivePeriodsTest();

            string filepath = wordObj.SaveWord();

            wordObj.removeHeader(filepath);

            ClearPara();






        }

        WordClass wordObj = new WordClass();

        //anesthesia

        public void Anesthesia_Periods()
        {
            pythonStat.InitPython();
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Anesthesia Periods")
                {


                    Anesthesia_Periods_EachGroup(ComparativeTables[tableindex]);


                }
            }
        }
        public int Anesthesia_Periods_EachGroup_rows(Parameter CheckParameterNormality)
        {
            int rows = 0;

            if(CheckParameterNormality.NormalOrAbnormal == "Normal")
            {
                rows = 7;
            }
            else if(CheckParameterNormality.NormalOrAbnormal == "Abnormal")
            {
                rows = 9;
            }



            return rows;


        }

        public void Periods_test_Group_MissingValues(IWTable table ,ComparativeTable comparativeTable , Parameter GroupParameter , int GroupColIndex , int CurrentGroupKey , int WordTableRows, int WordTableColumns)
        {
            AnovaTestResult anovaTestResult = null;
            List<Dictionary<int, List<double>>> CorrectDataParameters = new List<Dictionary<int, List<double>>>();
            List<Parameter> CurrentParameters = new List<Parameter>();

            foreach (var Parameter in comparativeTable.Parameters)
            {
                if (Parameter.IsGroup)
                    continue;

                CurrentParameters.Add(Parameter);

            }

            string NormalOrAbnormal = CurrentParameters[0].NormalOrAbnormal;

            foreach (var CurrentParameter in CurrentParameters)
            {
                Dictionary<int, List<double>> GroupedCorrectData = ExcelFunctions.GetDataWithMissingValues_withGroup(worksheet, CurrentParameter.Name, GroupParameter.Name, GroupColIndex);
                CorrectDataParameters.Add(GroupedCorrectData);
            }

            HashSet<int> uniqueGroupKeys = new HashSet<int>();
            foreach (var dict in CorrectDataParameters)
            {
                foreach (var key in dict.Keys)
                {
                    uniqueGroupKeys.Add(key);
                }
            }





            List<List<double>> valuesLists = new List<List<double>>();

            Dictionary<string, bool> ParaName_999Missing = new Dictionary<string, bool>();

            for (int i = 0; i < CurrentParameters.Count; i++)
            {
                if (CorrectDataParameters[i].TryGetValue(CurrentGroupKey, out List<double> values))
                {
                    if(values.Contains(-999))
                    {
                        ParaName_999Missing.Add(CurrentParameters[i].Name, true);
                    }
                    else
                    {
                        ParaName_999Missing.Add(CurrentParameters[i].Name, false);
                    }

                    valuesLists.Add(new List<double>(values));

                }
            }




            

            List<Parameter> parametersForGroup = new List<Parameter>();

            for (int i = 0; i < CurrentParameters.Count; i++)
            {
                if ((valuesLists[i].Count > 0) && ParaName_999Missing[CurrentParameters[i].Name] == false) // Ensure there's still valid data
                {
                    Parameter newParameter = new Parameter
                    {
                        Name = CurrentParameters[i].Name, // Copy name from original list
                        ParameterValues = valuesLists[i]  // Assign cleaned values
                    };

                    parametersForGroup.Add(newParameter);
                }
            }


            if (NormalOrAbnormal == "Normal")
            {
                anovaTestResult = pythonStat.RepeatedMeasuresAnovaBoth(parametersForGroup);
            }
            else if (NormalOrAbnormal == "Abnormal")
            {
                List<List<double>> Datalist = new List<List<double>>();



                foreach (var parameter in parametersForGroup)
                {
                    Datalist.Add(parameter.ParameterValues);
                }

                anovaTestResult = pythonStat.PerformFriedmanWithDunnTest(Datalist);




            }



            if (anovaTestResult.PairwiseComparisons.Count > 0)
            {
                if (NormalOrAbnormal == "Normal")
                {
                    
                    int pairwisectr = 0;
                    int PairwiseColumn = 2;
                    for (int i = 1; i < CurrentParameters.Count; i++)
                    {
                        if (ParaName_999Missing[CurrentParameters[i].Name] == false)
                        {
                            wordObj.AddPara_Center(table, WordTableRows - 1, PairwiseColumn, anovaTestResult.PairwiseComparisons[pairwisectr][2]);
                            bool Psig = generalFunctions.PvalueHasSig(anovaTestResult.PairwiseComparisons[pairwisectr][2]);
                            if(Psig)
                            {
                                wordObj.SubSuperScriptText(table, WordTableRows - 1, PairwiseColumn, Syncfusion.Drawing.Color.Empty, "*", "Super");
                            }
                            
                            pairwisectr++;
                            PairwiseColumn++;
                        }
                        else
                        {
                            List<List<double>> tpairedLists = new List<List<double>>
                                {
                                new List<double>(valuesLists[0]), // Create a new list for period1
                                new List<double>(valuesLists[i])  // Create a new list for period2
                                };

                            GeneralFunctions.RemoveInvalidEntries(ref tpairedLists);

                            List<double> period1 = tpairedLists[0];
                            List<double> period2 = tpairedLists[1];

                            string[] result = pythonStat.TpairedTest(period1, period2);
                            wordObj.AddPara_Center(table, WordTableRows - 1, PairwiseColumn, result[1]);

                            bool Psig = generalFunctions.PvalueHasSig(result[1]);
                            if (Psig)
                            {
                                wordObj.SubSuperScriptText(table, WordTableRows - 1, PairwiseColumn, Syncfusion.Drawing.Color.Empty, "*", "Super");
                            }

                            PairwiseColumn++;
                        }
                    }


                       
                }
                else if (NormalOrAbnormal == "Abnormal")
                {
                    int pairwisectr = 0;
                    int PairwiseColumn = 2;
                    for (int i = 1; i < CurrentParameters.Count; i++)
                    {
                        if (ParaName_999Missing[CurrentParameters[i].Name] == false)
                        {
                            wordObj.AddPara_Center(table, WordTableRows - 1, PairwiseColumn, anovaTestResult.PairwiseComparisons[pairwisectr][1]);
                            bool Psig = generalFunctions.PvalueHasSig(anovaTestResult.PairwiseComparisons[pairwisectr][1]);
                            if (Psig)
                            {
                                wordObj.SubSuperScriptText(table, WordTableRows - 1, PairwiseColumn, Syncfusion.Drawing.Color.Empty, "*", "Super");
                            }

                            pairwisectr++;
                            PairwiseColumn++;
                        }
                        else
                        {
                            List<List<double>> tpairedLists = new List<List<double>>
                                {
                                new List<double>(valuesLists[0]), // Create a new list for period1
                                new List<double>(valuesLists[i])  // Create a new list for period2
                                };

                            GeneralFunctions.RemoveInvalidEntries(ref tpairedLists);

                            List<double> period1 = tpairedLists[0];
                            List<double> period2 = tpairedLists[1];

                            string[] result = manual.Zpaired(period1.ToArray(), period2.ToArray());
                            wordObj.AddPara_Center(table, WordTableRows - 1, PairwiseColumn, result[1]);

                            bool Psig = generalFunctions.PvalueHasSig(result[1]);
                            if (Psig)
                            {
                                wordObj.SubSuperScriptText(table, WordTableRows - 1, PairwiseColumn, Syncfusion.Drawing.Color.Empty, "*", "Super");
                            }

                            PairwiseColumn++;
                        }
                    }

                }



            }











        }
        public void Anesthesia_Periods_EachGroup(ComparativeTable comparativeTable)
        {
            int GroupColIndex = newFindGroupColumnIndex(comparativeTable);

            Parameter GroupParam = null;
            Parameter CurrentParameter = null;
            int parametercount = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    GroupParam = parameter;
                }
                else
                {
                    CurrentParameter = parameter;
                    parametercount++;
                }

            }

            string[] ParameterExtractName = CurrentParameter.Name.Split('.');
            string parameterName = ParameterExtractName[0]; 

            foreach (var KVP in GroupParam.DIC_LablesIfNomainal)
            {
                int fakeorder = 1;
                IWSection section = wordObj.CreatePortraitSection();


                int group_number_value = KVP.Key;
                string group_name_value = KVP.Value;


                wordObj.AddTitle_Dynamic(section, "Change in "+ parameterName + " in " + group_name_value + " Group", ref fakeorder );

                int Current_GroupNumber_Count = GroupParam.GetValueCount(group_number_value);

                int Variablerows = Anesthesia_Periods_EachGroup_rows(CurrentParameter);


                int WordTableRows = Current_GroupNumber_Count + Variablerows;

                int WordTableColumns = 1 + parametercount;


                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                wordObj.GeneralTableFormat(table);

                ////Merges
                wordObj.Anesthesia_Periods_EachGroup_Merges(table,  WordTableColumns);

                ////Borders
                wordObj.Anesthesia_Periods_EachGroup_Borders(table, WordTableRows, WordTableColumns , Current_GroupNumber_Count);

                ////Widths
                wordObj.Anesthesia_Periods_EachGroup_Widths(table, WordTableRows, WordTableColumns , parametercount);

                ////Outer Headers
                wordObj.Anesthesia_Periods_EachGroup_Headers(table, WordTableRows, WordTableColumns , parameterName , comparativeTable , CurrentParameter , Current_GroupNumber_Count , parametercount);

                int insert_col = 1;
                foreach (var CurrentParameter_Insert in comparativeTable.Parameters)
                {
                    if (CurrentParameter_Insert.IsGroup)
                        continue;

                    int insert_row = 2;

                    Dictionary<int, List<double>> GroupedCorrectData = ExcelFunctions.GetDataWithMissingValues_withGroup(worksheet, CurrentParameter_Insert.Name, GroupParam.Name, GroupColIndex);


                    foreach (var GroupCorrected in GroupedCorrectData)
                    {
                        int GroupCorrected_key = GroupCorrected.Key;
                        if(group_number_value == GroupCorrected_key)
                        {
                            foreach (var item in GroupCorrected.Value)
                            {
                                if(item != -999)
                                {
                                    wordObj.Addpara_CenterNoBOLD(table, insert_row, insert_col, item.ToString());
                                }
                                else
                                {
                                    wordObj.Addpara_CenterNoBOLD(table, insert_row, insert_col, "-");
                                }
                                
                                insert_row++;
                            }

                        }
                    }


                    if (CurrentParameter_Insert.FormattedValues.ContainsKey(group_number_value))
                    {
                        foreach (var ParameterFormattedValues in CurrentParameter_Insert.FormattedValues[group_number_value])
                        {
                            string key = ParameterFormattedValues.Key;
                            string value = ParameterFormattedValues.Value;

                            switch (key)
                            {
                                case "Min-Max":
                                case "Mean ± StdDev":
                                    string[] splitValues = value.Split(' ');
                                    string firstValue = splitValues[0];
                                    string lastValue = splitValues[splitValues.Length - 1];

                                    wordObj.Addpara_CenterNoBOLD(table, insert_row++, insert_col, firstValue);
                                    wordObj.Addpara_CenterNoBOLD(table, insert_row++, insert_col, lastValue);
                                    break;

                                case "Median":
                                    if (CurrentParameter_Insert.NormalOrAbnormal == "Abnormal")
                                    {
                                        string[] splitMedian = value.Split(new string[] { " (" }, StringSplitOptions.None);
                                        if (splitMedian.Length == 2)  // Ensure proper format
                                        {
                                            string median = splitMedian[0];
                                            string IQR = splitMedian[1].TrimEnd(')');
                                            IQR = IQR.Replace(" ", "");

                                            wordObj.Addpara_CenterNoBOLD(table, insert_row++, insert_col, median);
                                            wordObj.Addpara_CenterNoBOLD(table, insert_row++, insert_col, IQR);
                                        }
                                    }
                                    break;
                            }
                        }
                    }


                    insert_col++;

                }


                
               Periods_test_Group_MissingValues(table,comparativeTable, GroupParam, GroupColIndex, group_number_value , WordTableRows , WordTableColumns);

               





                if (parametercount <= 6)
                {
                    wordObj.LeftAndRightCellMarginCustom(table, 0.01f, 0.01f);
                    wordObj.FormatTableCustom(table, 10.5f, 2, 1);
                }
                else if (parametercount > 6 && parametercount < 10)
                {
                    wordObj.LeftAndRightCellMarginCustom(table, 0.01f, 0.01f);
                    wordObj.FormatTableCustom(table, 10f , 2 , 1);
                }
                else if(parametercount > 10)
                {
                    wordObj.LeftAndRightCellMarginCustom(table, 0.05f, 0.05f);
                    wordObj.FormatTableCustom(table, 9.5f, 2, 1);
                }
                

               


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
                    parameter.FPairwise.Clear();
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
        
        public void InsertGroupTitles_NewComparativeGroups_Fn(IWTable table,bool HasNominal , ComparativeTable comparativeTable)
        {
            int startcol = 1;
            int totalcount = 0;

            if(Total_Column_Comparative && HasNominal)
            {
                startcol = 3;
            }
            else if(Total_Column_Comparative && !HasNominal)
            {
                startcol = 2;
            }

            Parameter groupparameter = ComparativeTable.GetGroupParamter(comparativeTable);
            Dictionary<int, int> ValuewithCounts = groupparameter.GetValueCounts_AllIncludingUnknowns();
            foreach (var kvp in groupparameter.DIC_LablesIfNomainal)
            {
                string GroupName = kvp.Value;
                int CurrentGroupCount = ValuewithCounts[kvp.Key];
                string GroupNameCount = GroupName + Convert.ToChar(11) + "(n = " + CurrentGroupCount+ ")";

                
                wordObj.AddPara_Center(table, 0, startcol, GroupNameCount);
                if(HasNominal)
                {
                    startcol = startcol + 2;
                }
                else
                {
                    startcol++;
                }

                totalcount += CurrentGroupCount;

            }


            if(Total_Column_Comparative)
            {
                string TotalCount = "Total" + Convert.ToChar(11) + "(n = " + totalcount + ")";
                wordObj.AddPara_Center(table, 0, 1, TotalCount);
            }


        }
        public void InsertData_NewComparativeGroups_Fn(ComparativeTable comparativeTable , int startingrow)
        {
            int insertcol = 1;

            var groupparameter = ComparativeTable.GetGroupParamter(comparativeTable);
            if (Total_Column_Comparative)
            {

            }
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.NominalOrScale == "Nominal")
                {
                    var results = SPSS_TestRunner.RunCrosstabChiSquare(groupparameter.Name, new List<string> { parameter.Name });
                    foreach (var result in results)
                    {
                        foreach (var Row in result.CrosstabRows)
                        {

                        }
                        MessageBox.Show(result.FisherExpectedCountPercentage);
                    }
                }
            }

        }

        //public void ComparativeParamaeterBorders_NewComparativeGroups_Fn(IWTable table, int WordTableRows, int WordTableColumns, ComparativeTable comparativeTable, int numberofgroups , bool HasScale , bool HasNominal , List<string> CheckedScaleDataNeeded , int PairwiseCount)
        //{
        //    int startingrow = 0;
        //    if (HasNominal)
        //    {
        //        startingrow = 1; 
        //    }
        //    foreach (var parameter in comparativeTable.Parameters)
        //    {

        //        if (parameter.IsGroup)
        //            continue;

        //        //getting count for nominal and scale
        //        int count = 0;

        //        if (parameter.NominalOrScale == "Nominal")
        //        {
        //            Dictionary<int , int> Correctcount = parameter.GetValueCounts_AllIncludingUnknowns();
        //            count = Correctcount.Keys.Count+1;
        //        }
        //        else if (parameter.NominalOrScale == "Scale")
        //        {
        //            count = CheckedScaleDataNeeded.Count + 1 + PairwiseCount;
        //        }

        //        //Merge for all parameters test
        //        if(count > 0)
        //        {

        //            //Merges of test

        //            if (parameter.NominalOrScale == "Nominal")
        //            {
        //                if (startingrow + count < WordTableRows)
        //                {
        //                    table.ApplyVerticalMerge(WordTableColumns - 1, startingrow + 2, startingrow + count);
        //                    table.ApplyVerticalMerge(WordTableColumns - 2, startingrow + 2, startingrow + count);
        //                }

        //            }
        //            else if (parameter.NominalOrScale == "Scale")
        //            {
        //                if (startingrow + count < WordTableRows)
        //                {
        //                    table.ApplyVerticalMerge(WordTableColumns - 1, startingrow + 2, startingrow + count-PairwiseCount);
        //                    table.ApplyVerticalMerge(WordTableColumns - 2, startingrow + 2, startingrow + count- PairwiseCount);
        //                }
        //            }

        //            if (parameter.NominalOrScale == "Scale")
        //            {
        //                if(HasNominal)
        //                {
        //                    for (int j = 0; j <= CheckedScaleDataNeeded.Count; j++)
        //                    {
        //                        for (int i = 1; i < WordTableColumns - 2; i = i + 2)
        //                        {
        //                            table.ApplyHorizontalMerge(startingrow + j + 1, i, i + 1);

        //                        }

        //                    }
        //                }


        //                int pairwiseStartRow = startingrow + CheckedScaleDataNeeded.Count + 2;

        //                for (int i = 0; i < PairwiseCount; i++)
        //                {
        //                    int currentRow = pairwiseStartRow + i;


        //                    // Only merge on the last row of the pairwise block
        //                    if (i == PairwiseCount - 1)
        //                    {
        //                        table.ApplyHorizontalMerge(currentRow, 1, WordTableColumns-3);
        //                    }

        //                    if (HasNominal && !(i == PairwiseCount - 1))
        //                    {
        //                        for (int colctrpair = 1; colctrpair < WordTableColumns-2; colctrpair = colctrpair+2)
        //                        {
        //                            table.ApplyHorizontalMerge(currentRow, colctrpair, colctrpair +1);
        //                        }
        //                    }

        //                    for (int toppairwiseborder = 0; toppairwiseborder < WordTableColumns; toppairwiseborder++)
        //                    {
        //                        table.Rows[currentRow].Cells[toppairwiseborder].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
        //                        table.Rows[currentRow].Cells[toppairwiseborder].CellFormat.Borders.Top.LineWidth = 0.5f;
        //                    }

        //                }

        //                wordObj.AddPara_NoCenter(table, startingrow + 1, 0, parameter.Name);

        //            }

        //        }


        //        if (startingrow + count < WordTableRows - 1)
        //        {

        //            // Insert bottom border for each parameter
        //            if (count > 0)
        //            {
        //                for (int i = 0; i < WordTableColumns; i++)
        //                {
        //                    table.Rows[startingrow + count].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
        //                    table.Rows[startingrow + count].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
        //                }

        //                startingrow += count;

        //            }


        //        }


        //    }


        //}


        //string[] keysToCheck = { "NumberOfCasesFirst", "MinMaxFirst", "MeanSDFirst", "MedianIQRFirst", "MedianMinMaxSecond" };

        public void ComparativeParamaeterBorders_NewComparativeGroups_Fn(
    IWTable table, int WordTableRows, int WordTableColumns,
    ComparativeTable comparativeTable, int numberofgroups,
    bool HasScale, bool HasNominal, List<string> CheckedScaleDataNeeded, int PairwiseCount, Dictionary<string, bool> CheckedExtraData)
        {
            InsertGroupTitles_NewComparativeGroups_Fn(table,HasNominal , comparativeTable);
            int startingRow = HasNominal ? 1 : 0;

            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup) continue;

                int count = GetRowCount(parameter, CheckedScaleDataNeeded, PairwiseCount);

                if (count <= 0) continue;

                ApplyVerticalMerges(table, WordTableColumns, WordTableRows, parameter, startingRow, count, PairwiseCount);
                InsertTitles(table, parameter, CheckedScaleDataNeeded, startingRow , CheckedExtraData , CheckedScaleDataNeeded.Count , PairwiseCount);

                if (parameter.NominalOrScale == "Scale")
                {
                    if (HasNominal)
                    {
                        ApplyScaleHorizontalMerges(table, startingRow, CheckedScaleDataNeeded.Count, WordTableColumns);
                    }

                    ApplyPairwiseSection(table, startingRow, WordTableColumns, CheckedScaleDataNeeded.Count, PairwiseCount, HasNominal);
                    wordObj.AddPara_NoCenter(table, startingRow + 1, 0, parameter.Name);
                }
                else if (parameter.NominalOrScale == "Nominal")
                {
                    wordObj.AddPara_NoCenter(table, startingRow + 1, 0, parameter.Name);
                }

                // Add bottom border after the parameter block if not the last
                if (startingRow + count < WordTableRows - 1)
                {
                    ApplyBottomBorder(table, startingRow + count, WordTableColumns);
                }

                startingRow += count;
            }
        }
        public Dictionary<string, string> DIC_ParameterDataNames = new Dictionary<string, string>
        {
                { "NumberOfCasesFirst", "N" },
                { "MinMaxFirst", "Min. – Max." },
                { "MeanSDFirst", "Mean ± SD." },
                { "MedianIQRFirst", "Median (IQR)" },
                { "MedianMinMaxSecond", "Median (Min. – Max.)" }
        };
        public void InsertTitles(IWTable table,Parameter parameter , List<string> CheckedParameterNames, int startrow, Dictionary<string,bool> CheckedExtraData , int scaleCount , int pairwiseCount)
        {
            int pairwiseStartRow = startrow + scaleCount + 2;

            if (parameter.NominalOrScale == "Scale")
            {
                foreach (var kvp in DIC_ParameterDataNames)
                {
                    if (CheckedParameterNames.Contains(kvp.Key))
                    {
                        string DataInsertValue = kvp.Value;
                        wordObj.Addpara_NoCenterNoBOLD(table, startrow + 2, 0, DataInsertValue);
                        wordObj.LeftIntendBeforeText(table, startrow + 2, 0, 14.17f);
                        startrow++;
                    }
                }



                for (int i = 0; i < pairwiseCount; i++)
                {
                    int currentRow = pairwiseStartRow + i;

                    if (i == pairwiseCount - 1)
                    {
                        wordObj.AddPara_Center(table, currentRow, 0, "Sig. bet. groups");
                    }
                    else
                    {
                        if(i == 0)
                        {
                            if (CheckedExtraData["PcontrolFirst"] || CheckedExtraData["PcontrolLast"])
                            {

                                wordObj.AddPara_Center(table, currentRow, 0, "p");
                                wordObj.SubSuperScriptText(table, currentRow, 0, Syncfusion.Drawing.Color.Empty, "Control", "Sub");
                                
                            }
                            else
                            {
                                wordObj.AddPara_Center(table, currentRow, 0, "p");
                                wordObj.SubSuperScriptText(table, currentRow, 0, Syncfusion.Drawing.Color.Empty, "0", "Sub");
                                
                            }
                        }
                        else
                        {
                            wordObj.AddPara_Center(table, currentRow, 0, "p");
                            wordObj.SubSuperScriptText(table, currentRow, 0, Syncfusion.Drawing.Color.Empty, i.ToString(), "Sub");
                            
                        }
                        
                    }
                }
            }
            else if (parameter.NominalOrScale == "Nominal")
            {
                foreach (var label in parameter.DIC_LablesIfNomainal.Values)
                {
                    wordObj.Addpara_NoCenterNoBOLD(table, startrow+2 , 0, label.ToString());
                    wordObj.LeftIntendBeforeText(table, startrow+2 , 0, 14.17f);
                    startrow++;
                }
            }
        }
        //string[] keysToCheck = { "NumberOfCasesFirst", "MinMaxFirst", "MeanSDFirst", "MedianIQRFirst", "MedianMinMaxSecond" };
        private int GetRowCount(Parameter parameter, List<string> checkedScaleItems, int pairwiseCount)
        {
            if (parameter.NominalOrScale == "Nominal")
            {
                var countDict = parameter.GetValueCounts_AllIncludingUnknowns();
                return countDict.Keys.Count + 1;
            }
            else if (parameter.NominalOrScale == "Scale")
            {
                return checkedScaleItems.Count + 1 + pairwiseCount;
            }

            return 0;
        }

        private void ApplyVerticalMerges(IWTable table, int cols, int rows, Parameter param, int startRow, int count, int pairwiseCount)
        {
            if (startRow + count >= rows) return;

            int endRow = param.NominalOrScale == "Scale" ? startRow + count - pairwiseCount : startRow + count;

            table.ApplyVerticalMerge(cols - 1, startRow + 2, endRow);
            table.ApplyVerticalMerge(cols - 2, startRow + 2, endRow);
        }

        private void ApplyScaleHorizontalMerges(IWTable table, int startRow, int scaleCount, int cols)
        {
            for (int j = 0; j <= scaleCount; j++)
            {
                for (int i = 1; i < cols - 2; i += 2)
                {
                    table.ApplyHorizontalMerge(startRow + j + 1, i, i + 1);
                }
            }
        }

        private void ApplyPairwiseSection(IWTable table, int startRow, int cols, int scaleCount, int pairwiseCount, bool hasNominal)
        {
            int pairwiseStartRow = startRow + scaleCount + 2;

            for (int i = 0; i < pairwiseCount; i++)
            {
                int currentRow = pairwiseStartRow + i;

                // Top border for each pairwise row
                for (int c = 0; c < cols; c++)
                {
                    var cell = table.Rows[currentRow].Cells[c];
                    cell.CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    cell.CellFormat.Borders.Top.LineWidth = 0.5f;
                }

                if (i == pairwiseCount - 1)
                {
                    // Last row gets full width merge
                    table.ApplyHorizontalMerge(currentRow, 1, cols - 3);
                }
                else if (hasNominal)
                {
                    // Mid rows get pairwise 2-column merges
                    for (int col = 1; col < cols - 2; col += 2)
                    {
                        table.ApplyHorizontalMerge(currentRow, col, col + 1);
                    }
                }

                // Optional: Insert text here if needed
                // wordObj.AddPara_Center(table, currentRow, 0, "Pairwise result...");
            }
        }

        private void ApplyBottomBorder(IWTable table, int rowIndex, int cols)
        {
            for (int i = 0; i < cols; i++)
            {
                var cell = table.Rows[rowIndex].Cells[i];
                cell.CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                cell.CellFormat.Borders.Bottom.LineWidth = 0.5f;
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
                            var sortedKeys = parameter.GroupedParameterValues.Keys.OrderBy(key => key).ToList();
                            foreach (double group in sortedKeys)
                            {
                                int parameterGroupCount = parameter.GroupedParameterValues[group].Count;
                                //MessageBox.Show(parameterGroupCount.ToString());
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



            foreach (var GroupParameter in comparativeTable.Parameters)
            {

            }



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
                int numberofgroupsTrue = parameter.FormattedValues.Keys.Count;
                var SortedGroupsValues = parameter.GroupedParameterValues.Keys;

                int ValueOne = 0;
                foreach (var key in SortedGroupsValues)
                {
                    if (parameter.GroupedParameterValues[key].Count == 1)
                    {
                        ValueOne++;
                    }
                }


                if (numberofgroupsTrue - ValueOne <= 2)
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
                            values[0] = "t=" + Convert.ToChar(11) + values[0];
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
                        pythonStat.InitPython();

                        string[] values = manual.UTest(group1Values, group2Values); 



                        //MessageBox.Show(values[1]);

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
                            values[0] = "U="+ Convert.ToChar(11) + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }



                        wordObj.HighlightCellContent(table, addrow + 1, WordTableColumns);
                        //wordObj.InsertTest_P(table, addrow, addColumn+1, values[1]);


                        //U test



                       // List<Parameter> TestParameters = new List<Parameter>();
                       // List<List<double>> dataParaValues = new List<List<double>>();
                       // List<string> ParameterLabels = new List<string>();

                       // foreach (var testparameter in comparativeTable.Parameters)
                       // {
                       //     if (testparameter.IsGroup)
                       //     {
                       //         continue;
                       //     }
                       //     dataParaValues.Add(testparameter.ParameterValues);
                       //     ParameterLabels.Add(testparameter.Name);
                       //      //TestParameters.Add(testparameter);
                       // }

                       //pythonStat.PerformFriedmanWithDunnTest(dataParaValues);
                         //manual.Zpaired(dataParaValues[0].ToArray() , dataParaValues[1].ToArray());


                    }
                }

                else if (numberofgroupsTrue - ValueOne > 2)
                {

                    if (parameter.NormalOrAbnormal == "Normal")
                    {
                        pythonStat.InitPython();
                        comparativeTable.TestsDone.Add("FAnova");
                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        
                        //anovaTestResult = pythonStat.newANOVAWithTukeyHSDNewDynamic(parameter);

                        anovaTestResult = pythonStat.ANOVAWithTukeyHSDNewDynamic(parameter);
                        
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
                            values[0] = "F="+ Convert.ToChar(11) + values[0];
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

                        pythonStat.InitPython();

                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        anovaTestResult = pythonStat.KruskalWallisWithDunnDynamic(parameter);

                       // anovaTestResult = manual.Kruskal_H(parameter);
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
                            values[0] = "H="+ Convert.ToChar(11) + values[0];
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

        public int Count_Groups_2_Periods_rows(ComparativeTable comparativeTable)
        {
            int variablerows = 0;
            int parametercount = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.IsGroup)
                {
                    continue;
                }
                if(parameter.NominalOrScale == "Scale")
                {
                    parametercount++;
                }
            }
            int periodTestrows = parametercount / 2;
            variablerows = 1 + (parametercount * 4) + periodTestrows;


            return variablerows;
        }
        public void ComparativeTableGroups_2_periods()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Groups 2 periods")
                {
                    IWSection section = wordObj.CreatePortraitSection();
                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int WordTableRows = Count_Groups_2_Periods_rows(ComparativeTables[tableindex]);

                    int WordTableColumns = 4 + (numberofgroups);

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.GeneralTableFormat(table);

                    wordObj.ApplyGeneral_Groups_2Periods_Merges(table, ComparativeTables[tableindex] , WordTableColumns , WordTableRows);
                    wordObj.ApplyGeneral_Groups_2Periods_Borders(table, WordTableRows, WordTableColumns, ComparativeTables[tableindex]);

                    wordObj.SetComparative_GroupsUp_2Periods_Widths(table  , WordTableRows , WordTableColumns , numberofgroups);


                    wordObj.AddHeaderComparative_GroupsUp_2Periods(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex]);


                    InsertScale_GroupsUp_2Periods(table, ComparativeTables[tableindex]);

                    GetPeriodsTests_GroupsUp_2Periods(table, ComparativeTables[tableindex]);

                    GetGroupsTests_GroupsUp_2Periods(table, ComparativeTables[tableindex] , numberofgroups , WordTableColumns);


                    wordObj.FormatTable(table, 11);
                }
            }
        }

        public void GetGroupsTests_GroupsUp_2Periods(IWTable table, ComparativeTable comparativeTable , int numberofgroups , int WordTableColumns)
        {
            int row = 2;
            int flag = 0;
            string[] values = new string[2];
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.IsGroup)
                {
                    continue;
                }

                if (numberofgroups == 2)
                {
                    List<double> group1Values = new List<double>();
                    List<double> group2Values = new List<double>();

                    SplitGroupedParameterValues(parameter, out group1Values, out group2Values);

                    if(parameter.NormalOrAbnormal == "Normal")
                    {
                        values = manual.StudentT_Unpaired(group1Values, group2Values);
                    }
                    else if(parameter.NormalOrAbnormal == "Abnormal")
                    {
                        values = manual.UTest(group1Values, group2Values);
                    }    


                    wordObj.InsertTest_P(table, row, WordTableColumns-2, values);

                }

                if(flag == 0)
                {
                    row = row + 4;
                    flag = 1;
                }
                else if(flag == 1)
                {
                    row = row + 5;
                    flag = 0;
                }

                
            }
        }

        public bool CheckSig(string p)
        {
            bool isSig = false;
            if(p == "<0.001")
            {
                isSig = true;
            }
            else
            {
                double pvalue;
                if(double.TryParse(p , out pvalue))
                {
                    if(pvalue < 0.05)
                    {
                        isSig = true;
                    }
                    else if(pvalue >= 0.05)
                    {
                        isSig = false;
                    }
                }
            }


            return isSig;
        }
        public void GetPeriodsTests_GroupsUp_2Periods(IWTable table, ComparativeTable comparativeTable)
        {
            pythonStat.InitPython();
            int row = 9;
            for(int i = 0;i<comparativeTable.Parameters.Count;)
            {
                Parameter currentparameter = comparativeTable.Parameters[i];    
                if(currentparameter.IsGroup)
                {
                    i++;
                    continue;
                    
                }

                else if(!currentparameter.IsGroup)
                {
                    Parameter nextparameter = comparativeTable.Parameters[i+1];


                    var firstParameterValues = new List<double>();
                    var secondParameterValues = new List<double>();


                    int col = 2;
                    int count = 0;
                    int tempctr = 0;

                    if(currentparameter.NormalOrAbnormal == "Normal")
                    {
                        wordObj.AddParaCombined(table, row, 1, "t", false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                        wordObj.SubSuperScriptText(table, row, 1, Syncfusion.Drawing.Color.Yellow, "0", "Sub");
                        
                    }
                    else if(currentparameter.NormalOrAbnormal == "Abnormal")
                    {
                        wordObj.AddParaCombined(table, row, 1, "Z", false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                    }
                   
                    WParagraph testparaHighlightname = (WParagraph)table[row, 1].Paragraphs[0];
                    WTextRange PTextname = new WTextRange(testparaHighlightname.Document);
                    PTextname.Text = " (" + "p";
                    testparaHighlightname.ChildEntities.Insert(1, PTextname);
                    
                    wordObj.SubSuperScriptText(table, row, 1, Syncfusion.Drawing.Color.White, "0", "Sub");

                    WTextRange PTextname_Parenthses = new WTextRange(testparaHighlightname.Document);
                    PTextname_Parenthses.Text = ")";
                    testparaHighlightname.ChildEntities.Insert(3, PTextname_Parenthses);


                    foreach (var kvp in currentparameter.GroupedParameterValues)
                    {
                        string[] result = new string[2];
                        var key = kvp.Key;
                        var values = kvp.Value;
                        count = values.Count + tempctr;


                        for (int j = tempctr; j < count; j++)
                        {
                            firstParameterValues.Add(currentparameter.ParameterValues[j]);
                            secondParameterValues.Add(nextparameter.ParameterValues[j]);

                        }
                        tempctr = count;

                        var firstParameterArray = firstParameterValues.ToArray();
                        var secondParameterArray = secondParameterValues.ToArray();


                        //wordObj.AddPara_Center()
                        



                        if(currentparameter.NormalOrAbnormal == "Normal")
                        {
                            result = pythonStat.TpairedTest(firstParameterArray.ToList(), secondParameterArray.ToList());
                            //result = manual.Tpaired(firstParameterArray, secondParameterArray);
                        }
                        else if(currentparameter.NormalOrAbnormal == "Abnormal")
                        {
                            result = manual.Zpaired(firstParameterArray, secondParameterArray);
                        }



                        //string[] result = pythonStat.TpairedTest(firstParameterArray.ToList(), secondParameterArray.ToList());

                        bool issig = CheckSig(result[1]);

                        if(!issig)
                        {
                            wordObj.AddParaCombined(table, row, col, result[0], false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                            WParagraph testparaHighlight = (WParagraph)table[row, col].Paragraphs[0];
                            WTextRange PText = new WTextRange(testparaHighlight.Document);
                            PText.Text = " (" + result[1] + ")";
                            testparaHighlight.ChildEntities.Insert(1, PText);
                        }
                        else if(issig)
                        {
                            wordObj.AddParaCombined(table, row, col, result[0], false, true, Syncfusion.Drawing.Color.Yellow, Syncfusion.Drawing.Color.Black);
                            wordObj.SubSuperScriptText(table, row, col, Syncfusion.Drawing.Color.Yellow, "*", "Super");

                            WParagraph testparaHighlight = (WParagraph)table[row, col].Paragraphs[0];
                            WTextRange PText = new WTextRange(testparaHighlight.Document);
                            PText.Text = " (" + result[1];
                            testparaHighlight.ChildEntities.Insert(2, PText);

                            wordObj.SubSuperScriptText(table, row, col, Syncfusion.Drawing.Color.White, "*", "Super");

                            WTextRange PTextvalue_Parenthses = new WTextRange(testparaHighlight.Document);
                            PTextvalue_Parenthses.Text = ")";
                            testparaHighlight.ChildEntities.Insert(4, PTextvalue_Parenthses);


                        }


                        col++;
                        firstParameterValues.Clear();
                        secondParameterValues.Clear();
                        firstParameterArray.Clear();
                        secondParameterArray.Clear();



                    }

                    row += 9;
                    i += 2;




                }





            }


        }
        public void InsertScale_GroupsUp_2Periods(IWTable table, ComparativeTable comparativeTable)
        {
            int row = 1;
            int flag = 0;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if (parameter.IsGroup)
                {
                    continue;
                }
                if (parameter.NominalOrScale == "Scale")
                {

                    var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();


                    int col = 2;
                    foreach (var groupValue in sortedKeys)
                    {

                        Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];

                        int temprow = row+1;
                        foreach (var stat in formattedValues)
                        {
                            wordObj.Addpara_CenterNoBOLD(table, temprow, col, stat.Value);
                            temprow++;

                        }
                        col++;



                    }
                    if (flag == 0)
                    {
                        row = row + 4;
                        flag = 1;
                    }
                    else if (flag == 1)
                    {
                        row = row + 5;
                        flag = 0;
                    }

                }
            }
            
        }
        public int CountPathologyrows_IQR(ComparativeTable comparativeTable)
        {
            int rowCount = 0;

            foreach (var parameter in comparativeTable.Parameters)
            {

                if (parameter.NominalOrScale == "Nominal")
                {
                    int distinctValuesCount = parameter.DIC_LablesIfNomainal.Keys.Count;
                    rowCount += distinctValuesCount;

                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    rowCount += 3;
                }
            }

            return rowCount;
        }


        public int CountPathologyrows(ComparativeTable comparativeTable)
        {
            int rowCount = 0;

            foreach (var parameter in comparativeTable.Parameters)
            {

                if (parameter.NominalOrScale == "Nominal")
                {
                    int distinctValuesCount = parameter.DIC_LablesIfNomainal.Keys.Count;
                    rowCount += distinctValuesCount;

                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    rowCount += 2;
                }
            }

            return rowCount;
        }
        public void Pathology_Layout_NoIQR()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Pathology Relation")
                {
                    bool TableHasSigI = false;

                    IWSection section = wordObj.CreatePortraitSection();

                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = CountPathologyrows(ComparativeTables[tableindex]);


                    int WordTableColumns = 5 + (numberofgroups * 2);

                    int WordTableRows = 3 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);

                    wordObj.ApplyPathology_Outer_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);

                    wordObj.ApplyPathology_Outer_Borders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);

                    wordObj.ApplyPathology_Widths(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);


                    wordObj.Apply_AllHeaders_Inner_Merges_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);

                    (string testtype, string NominalOrScale, bool issame) = InsertSeperateTest_TestOfSig(ComparativeTables[tableindex], numberofgroups);

                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, testtype);


                    int startingRow;
                    int newRowCount = 3;

                    foreach (Parameter parameter in ComparativeTables[tableindex].Parameters)
                    {
                        startingRow = newRowCount;
                        int column = 3;
                        int count = 0;

                        if (parameter.IsGroup || parameter.FormattedValues.Count == 0)
                        {
                            continue;
                        }

                        // Determine the count based on parameter type
                        if (parameter.NominalOrScale == "Nominal")
                        {
                            count = parameter.DIC_LablesIfNomainal.Keys.Count();
                            ComparativeTables[tableindex].TestsDone.Add("Chi");
                            pythonStat.InitPython();
                            string[] values = pythonStat.PerformChiSquareTest(parameter);
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
                                values[0] = "χ²="+ Convert.ToChar(11) + values[0];
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


                            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();

                            foreach (var groupValue in sortedKeys)
                            {
                                Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];
                                foreach (var distinctValue in parameter.DIC_LablesIfNomainal.Keys)
                                {

                                    string frequency = formattedValues.ContainsKey($"Frequency_{distinctValue}") ? formattedValues[$"Frequency_{distinctValue}"] : "0";
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, frequency);
                                    column++;

                                    string percentage = formattedValues.ContainsKey($"Percentage_{distinctValue}") ? formattedValues[$"Percentage_{distinctValue}"] : "0.0";
                                    percentage = percentage.Replace("%", "");


                                    if (percentage.EndsWith("0"))
                                    {
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage);
                                    }

                                    else
                                    {
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage.ToString());
                                    }
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
                            count = 2; // Count for scale parameters

                            // Insert scale statistics for scale parameters
                            foreach (var groupValue in sortedKeys)
                            {
                                Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];

                                if (parameter.GroupedParameterValues[groupValue].Count == 1)
                                {
                                    double only1 = parameter.GroupedParameterValues[groupValue][0];
                                    string only1str = only1.ToString("0.0");
                                    table.ApplyVerticalMerge(column, startingRow, startingRow + 1);
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, only1str);
                                    wordObj.SubSuperScriptText(table, startingRow, column, Syncfusion.Drawing.Color.White, "#", "Super");
                                }
                                else
                                {
                                    string MeanSD = "";
                                    string Median = "";
                                    string MinMax = "";
                                    string FormattedMedianMinMax = "";
                                    bool isTwonumbers = false;
                                    foreach (var stat in formattedValues)
                                    {
                                        if (stat.Key == "Mean ± StdDev")
                                        {
                                            MeanSD = stat.Value;
                                        }
                                        if (stat.Key == "Median")
                                        {

                                            string[] valuessplit = stat.Value.Split(' ');
                                            Median = valuessplit[0];
                                        }
                                        if (stat.Key == "Min-Max")
                                        {
                                            MinMax = stat.Value;
                                            string[] MinMaxsplit = MinMax.Split(' ');
                                            string Max = MinMaxsplit[2];
                                            isTwonumbers = Max.Split('.')[0].Length >= 2;
                                        }

                                    }
                                    if(isTwonumbers)
                                    {
                                        FormattedMedianMinMax = Median + Convert.ToChar(11) + " (" + MinMax + ")";
                                    }
                                    else
                                    {
                                        FormattedMedianMinMax = Median + " (" + MinMax + ")";
                                    }
                                    
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, MeanSD);
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow + 1, column, FormattedMedianMinMax);
                                }

 
                                startingRow = newRowCount;

                                column = column + 2;
                            }
                            PerformTest_Pathology(table, numberofgroups, parameter, startingRow, WordTableColumns - 2, WordTableColumns, issame, ComparativeTables[tableindex]);
                        }


                        newRowCount += count;

                    }

                    RemoveIf_AllScale_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.01f), wordObj.SetColumnWidthInCentimeters(0.01f));
                    wordObj.FormatTableCustom(table, 10, 0, 0);
                }
            }

        }


        public void Pathology_Layout()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Pathology Comparative")
                {
                    bool TableHasSigI = false;

                    IWSection section = wordObj.CreatePortraitSection();

                    int numberofgroups = CountGroupValues(ComparativeTables[tableindex]);

                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = CountPathologyrows_IQR(ComparativeTables[tableindex]);


                    int WordTableColumns = 5 + (numberofgroups * 2);

                    int WordTableRows = 3 + Variablerows;

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);

                    wordObj.ApplyPathology_Outer_Merges(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);

                    wordObj.ApplyPathology_Outer_Borders(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);

                    wordObj.ApplyPathology_Widths(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);


                    wordObj.Apply_AllHeaders_Inner_Merges_Pathology_IQR(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns, numberofgroups);

                    (string testtype, string NominalOrScale, bool issame) = InsertSeperateTest_TestOfSig(ComparativeTables[tableindex], numberofgroups);

                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, testtype);


                    int startingRow;
                    int newRowCount = 3;

                    foreach (Parameter parameter in ComparativeTables[tableindex].Parameters)
                    {
                        startingRow = newRowCount;
                        int column = 3;
                        int count = 0;

                        if (parameter.IsGroup || parameter.FormattedValues.Count == 0)
                        {
                            continue;
                        }

                        // Determine the count based on parameter type
                        if (parameter.NominalOrScale == "Nominal")
                        {
                            count = parameter.DIC_LablesIfNomainal.Keys.Count();
                            ComparativeTables[tableindex].TestsDone.Add("Chi");
                            pythonStat.InitPython();
                            string[] values = pythonStat.PerformChiSquareTest(parameter);
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
                                values[0] = "χ²=" + Convert.ToChar(11) + values[0];
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


                            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();

                            foreach (var groupValue in sortedKeys)
                            {
                                Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];
                                foreach (var distinctValue in parameter.DIC_LablesIfNomainal.Keys)
                                {

                                    string frequency = formattedValues.ContainsKey($"Frequency_{distinctValue}") ? formattedValues[$"Frequency_{distinctValue}"] : "0";
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, frequency);
                                    column++;

                                    string percentage = formattedValues.ContainsKey($"Percentage_{distinctValue}") ? formattedValues[$"Percentage_{distinctValue}"] : "0.0";
                                    percentage = percentage.Replace("%", "");


                                    if (percentage.EndsWith("0"))
                                    {
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage);
                                    }

                                    else
                                    {
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage.ToString());
                                    }
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
                            count = 3; // Count for scale parameters

                            // Insert scale statistics for scale parameters
                            foreach (var groupValue in sortedKeys)
                            {
                                Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];

                                if (parameter.GroupedParameterValues[groupValue].Count == 1)
                                {
                                    double only1 = parameter.GroupedParameterValues[groupValue][0];
                                    string only1str = only1.ToString("0.0");
                                    table.ApplyVerticalMerge(column, startingRow, startingRow + 2);
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, only1str);
                                    wordObj.SubSuperScriptText(table, startingRow, column, Syncfusion.Drawing.Color.White, "#", "Super");
                                }
                                else
                                {
                                    string MeanSD = "";
                                    
                                    string MinMax = "";
                                    string FormattedMedianIQR = "";
                                    bool isTwonumbers = false;
                                    foreach (var stat in formattedValues)
                                    {


                                        if (stat.Key == "Min-Max")
                                        {
                                            MinMax = stat.Value;
                                            string[] MinMaxsplit = MinMax.Split(' ');
                                            string Max = MinMaxsplit[2];
                                            isTwonumbers = Max.Split('.')[0].Length >= 2;
                                        }
                                        if (stat.Key == "Mean ± StdDev")
                                        {
                                            MeanSD = stat.Value;
                                        }
                                        if (stat.Key == "Median")
                                        {
                                            FormattedMedianIQR = stat.Value;
                                        }
                                        

                                    }
                                    


                                    wordObj.Addpara_CenterNoBOLD(table, startingRow, column, MinMax);
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow+1, column, MeanSD);
                                    wordObj.Addpara_CenterNoBOLD(table, startingRow + 2, column, FormattedMedianIQR);
                                }


                                startingRow = newRowCount;

                                column = column + 2;
                            }
                            PerformTest_Pathology_IQR(table, numberofgroups, parameter, startingRow, WordTableColumns - 2, WordTableColumns, issame, ComparativeTables[tableindex]);
                        }


                        newRowCount += count;

                    }

                    RemoveIf_AllScale_Pathology(table, ComparativeTables[tableindex], WordTableRows, WordTableColumns);



                    wordObj.LeftAndRightCellMarginCustom(table, wordObj.SetColumnWidthInCentimeters(0.01f), wordObj.SetColumnWidthInCentimeters(0.01f));
                    wordObj.FormatTableCustom(table, 10, 0, 0);
                }
            }

        }
        public void RemoveIf_AllScale_Pathology(IWTable table, ComparativeTable comparativeTable , int WordTableRows , int WordTableColumns)
        {
            bool TableHasNominal = false;
            foreach (var parameter in comparativeTable.Parameters)
            {
                if(parameter.NominalOrScale == "Nominal")
                {
                    TableHasNominal = true;
                    break;
                }
            }
            if(!TableHasNominal)
            {
                table.Rows.RemoveAt(2);
                for (int j = 0; j < WordTableColumns; j++)
                {
                    table.Rows[1].Cells[j].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                    table.Rows[1].Cells[j].CellFormat.Borders.Bottom.LineWidth = 1.5f;
                }
            }
            
        }
        public void PerformTest_Pathology(IWTable table, int numberofgroups, Parameter parameter, int addrow, int addColumn, int WordTableColumns, bool issame, ComparativeTable comparativeTable)
        {
            pythonStat.InitPython();
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
                            values[0] = "t="+ Convert.ToChar(11) + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }


                        wordObj.HighlightCellContent(table, addrow +1, WordTableColumns);
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

                        //MessageBox.Show(values[1]);

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
                            values[0] = "U="+ Convert.ToChar(11) + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }



                        wordObj.HighlightCellContent(table, addrow , WordTableColumns);
                        //wordObj.InsertTest_P(table, addrow, addColumn+1, values[1]);


                        //U test
                    }
                }

                else if (numberofgroups > 2)
                {

                    if (parameter.NormalOrAbnormal == "Normal")
                    {
                        pythonStat.InitPython();
                        comparativeTable.TestsDone.Add("FAnova");
                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        anovaTestResult = pythonStat.ANOVAWithTukeyHSDNewDynamic(parameter);

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
                            values[0] = "F="+ Convert.ToChar(11) + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);

                        }



                        wordObj.HighlightCellContent(table, addrow + 1, WordTableColumns);

                        // F anova
                    }
                    else if (parameter.NormalOrAbnormal == "Abnormal")
                    {
                        //H kriskual
                        //
                        comparativeTable.TestsDone.Add("H");

                        pythonStat.InitPython();

                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        anovaTestResult = pythonStat.KruskalWallisWithDunnDynamic(parameter);

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
                            values[0] = "H="+ Convert.ToChar(11) + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }
                        wordObj.HighlightCellContent(table, addrow , WordTableColumns);
                        // wordObj.Addpara_CenterNoBOLD(table, addrow, addColumn +1, anovaTestResult.PValue);

                    }

                }

            }
        }



        public void PerformTest_Pathology_IQR(IWTable table, int numberofgroups, Parameter parameter, int addrow, int addColumn, int WordTableColumns, bool issame, ComparativeTable comparativeTable)
        {
            pythonStat.InitPython();
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
                            values[0] = "t=" + Convert.ToChar(11) + values[0];
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

                        //MessageBox.Show(values[1]);

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
                            values[0] = "U=" + Convert.ToChar(11) + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }



                        wordObj.HighlightCellContent(table, addrow+1, WordTableColumns);
                        //wordObj.InsertTest_P(table, addrow, addColumn+1, values[1]);


                        //U test
                    }
                }

                else if (numberofgroups > 2)
                {

                    if (parameter.NormalOrAbnormal == "Normal")
                    {
                        pythonStat.InitPython();
                        comparativeTable.TestsDone.Add("FAnova");
                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        anovaTestResult = pythonStat.ANOVAWithTukeyHSDNewDynamic(parameter);

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
                            values[0] = "F=" + Convert.ToChar(11) + values[0];
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

                        pythonStat.InitPython();

                        AnovaTestResult anovaTestResult = new AnovaTestResult();
                        anovaTestResult = pythonStat.KruskalWallisWithDunnDynamic(parameter);

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
                            values[0] = "H=" + Convert.ToChar(11) + values[0];
                            wordObj.InsertTest_P(table, addrow, addColumn, values);
                        }
                        wordObj.HighlightCellContent(table, addrow+1, WordTableColumns);
                        // wordObj.Addpara_CenterNoBOLD(table, addrow, addColumn +1, anovaTestResult.PValue);

                    }

                }

            }
        }
        public int CountRows_NewComparativeGroups_Fn(ComparativeTable comparativeTable, List<string> CheckedPrimaryNeeded)
        {
            int rowCount = 0;

            bool hasNominal = false;
            foreach (var parameter in comparativeTable.Parameters)
            {

                if (parameter.NominalOrScale == "Nominal")
                {
                    Dictionary<int, int> parameterDistinctCount = parameter.GetValueCounts_AllIncludingUnknowns();
                    
                    rowCount += parameterDistinctCount.Keys.Count + 1;
                    hasNominal = true;
                }
                else if (parameter.NominalOrScale == "Scale")
                {
                    rowCount += CheckedPrimaryNeeded.Count + 1;
                }
            }

            if(!hasNominal)
            {
                rowCount--;
            }
            //MessageBox.Show(rowCount.ToString());



            return rowCount;


        }
        public int CountGroupValues_NewComparativeGroups_Fn(ComparativeTable comparativeTable)
        {
            int groupCount = 0;

            // Find the group parameter
            var groupParameter = comparativeTable.Parameters.FirstOrDefault(p => p.IsGroup);
            if (groupParameter == null)
            {
                // Handle case where no group parameter is found
                return groupCount;
            }

            // Count the distinct group values
            groupCount = Math.Max(groupParameter.DIC_LablesIfNomainal.Keys.Distinct().Count() , groupParameter.ParameterValues.Distinct().Count());

            return groupCount;

            
        }

        public (int baseCount, int totalCount) CountPairwiseRows_NewComparativeGroups_Fn(ComparativeTable comparativeTable)
        {
            int numberofgroups = CountGroupValues_NewComparativeGroups_Fn(comparativeTable);
            int baseCount = numberofgroups > 2 ? numberofgroups - 2 : 0;

            int scaleParameterCount = comparativeTable.Parameters
                .Count(p => !p.IsGroup && p.NominalOrScale == "Scale");

            int totalCount = baseCount * scaleParameterCount;

            return (baseCount, totalCount);
        }



        public int CountCols_NewComparativeGroups_Fn(ComparativeTable comparativeTable, int numberofgroups, bool HasNominal, bool hastotal)
        {
            int totalcol = hastotal ? (HasNominal ? 2 : 1) : 0;

            int colCount = HasNominal
                ? 3 + (numberofgroups * 2) + totalcol
                : 3 + numberofgroups + totalcol;

            return colCount;
        }


        public void NewComparativeGroups_Fn()
        {
            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                if (ComparativeTables[tableindex].FormatType == "Default")
                {
                    bool HasNominal = DetermineHasNominal(ComparativeTables[tableindex]);
                    bool HasScale = DetermineHasScale(ComparativeTables[tableindex]);
                    Dictionary<string, bool> CheckedDataprimary = FormDataTransfer.Get<Dictionary<string, bool>>("nodeCheckedStatusPrimary");
                    Dictionary<string, bool> CheckedDataExtra = FormDataTransfer.Get<Dictionary<string, bool>>("nodeCheckedStatusExtra");
                    string LeftMarginValue = FormDataTransfer.Get<string>("LeftMarginValue");
                    string RightMarginValue = FormDataTransfer.Get<string>("RightMarginValue");
                    Total_Column_Comparative = CheckedDataExtra["TotalColumn"];

                    List<string> CheckedPrimaryNeeded = new List<string>();
                    string[] keysToCheck = { "NumberOfCasesFirst", "MinMaxFirst", "MeanSDFirst", "MedianIQRFirst", "MedianMinMaxSecond" };
                    foreach (string key in keysToCheck)
                    {
                        if (CheckedDataprimary.TryGetValue(key, out bool isChecked) && isChecked)
                        {
                            CheckedPrimaryNeeded.Add(key);
                        }
                    }

                    IWSection section = wordObj.CreatePortraitSection();

                    int numberofgroups = CountGroupValues_NewComparativeGroups_Fn(ComparativeTables[tableindex]);
                    wordObj.AddComparativeTitle(section, ComparativeTables[tableindex].TableName, numberofgroups);

                    int Variablerows = CountRows_NewComparativeGroups_Fn(ComparativeTables[tableindex] , CheckedPrimaryNeeded);
                    (int PairwiseCount,int TotalPairwiseCount)  = CountPairwiseRows_NewComparativeGroups_Fn(ComparativeTables[tableindex]);
                    int WordTableRows = 2 + Variablerows + TotalPairwiseCount;
                    int WordTableColumns = CountCols_NewComparativeGroups_Fn(ComparativeTables[tableindex] , numberofgroups , HasNominal , Total_Column_Comparative); 

                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);
                    wordObj.GeneralTableFormat(table);


                    //Merges
                    if(HasNominal)
                    {
                        wordObj.ApplyGeneralComparativeMerges_NewComparativeGroups_Fn(table, WordTableColumns, numberofgroups);
                    }
                    wordObj.ApplyGeneralComparativeBorders_NewComparativeGroups_Fn(table, WordTableRows, WordTableColumns, numberofgroups, HasNominal);


                    wordObj.SetComparativeWidths_NewComparativeGroups_Fn(table, WordTableRows, WordTableColumns, numberofgroups, ComparativeTables[tableindex] , HasScale , HasNominal , Total_Column_Comparative);

                    wordObj.Add_GeneralHeaders_Comparative_Center_NewComparativeGroups_Fn(table, WordTableRows, WordTableColumns, numberofgroups, HasNominal);

                    ComparativeParamaeterBorders_NewComparativeGroups_Fn(table, WordTableRows, WordTableColumns, ComparativeTables[tableindex], numberofgroups, HasScale, HasNominal  , CheckedPrimaryNeeded , PairwiseCount , CheckedDataExtra);
                    
                    (string testtype, string NominalOrScale, bool issame) = InsertSeperateTest_TestOfSig(ComparativeTables[tableindex], numberofgroups);

                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, testtype);



                    //InsertData_NewComparativeGroups_Fn(ComparativeTables[tableindex]);


                }


            }


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

                    

                    (string testtype, string NominalOrScale, bool issame) = InsertSeperateTest_TestOfSig(ComparativeTables[tableindex], numberofgroups);

                    wordObj.InsertHighlightTestName(table, 0, WordTableColumns - 2, testtype);

                    int startingRow;
                    int newRowCount = 3;
                    Parameter GroupPara = null;
                    foreach (Parameter parameter in ComparativeTables[tableindex].Parameters)
                    {
                        startingRow = newRowCount;
                        int column = 1;
                        int count = 0;

                        
                        if(parameter.IsGroup)
                        {
                            GroupPara = parameter;
                        }

                        if (parameter.IsGroup || parameter.FormattedValues.Count == 0)
                        {
                            continue;
                        }

                        

                        // Determine the count based on parameter type
                        if (parameter.NominalOrScale == "Nominal")
                        {
                            count = parameter.DIC_LablesIfNomainal.Keys.Count() + 1;
                            ComparativeTables[tableindex].TestsDone.Add("Chi");
                            pythonStat.InitPython();
                            string[] values = pythonStat.PerformChiSquareTest(parameter);
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
                                values[0] = "χ²="+ Convert.ToChar(11) + values[0];
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

                           


                            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();

                            var SortedGroupKeysAll = GroupPara.ParameterValues.Distinct().OrderBy(key => key).ToList();



                            foreach (int groupValue in SortedGroupKeysAll)
                            {
                                if(sortedKeys.Contains(groupValue))
                                {
                                    Dictionary<string, string> formattedValues = parameter.FormattedValues[groupValue];
                                    foreach (var distinctValue in parameter.DIC_LablesIfNomainal.Keys)
                                    {

                                        string frequency = formattedValues.ContainsKey($"Frequency_{distinctValue}") ? formattedValues[$"Frequency_{distinctValue}"] : "0";
                                        wordObj.Addpara_CenterNoBOLD(table, startingRow, column, frequency);
                                        column++;

                                        string percentage = formattedValues.ContainsKey($"Percentage_{distinctValue}") ? formattedValues[$"Percentage_{distinctValue}"] : "0.0";
                                        percentage = percentage.Replace("%", "");


                                        if (percentage.EndsWith("0"))
                                        {
                                            wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage);
                                        }

                                        else
                                        {
                                            wordObj.Addpara_CenterNoBOLD(table, startingRow, column, percentage.ToString());
                                        }
                                        startingRow++; // Move to the next row
                                        column--;
                                    }

                                    startingRow = newRowCount;
                                    column = column + 2;
                                }
                                else
                                {
                                    column = column + 2;
                                }

                            }
                            //var contingencyTable = new List<List<int>>();
                            //foreach (var tparameter in ComparativeTables[tableindex].Parameters)
                            //{
                            //    if(tparameter.IsGroup)
                            //    {
                            //        continue;
                            //    }

                            //    var allValues = parameter.ParameterValues.Distinct().ToList();


                            //    // Prepare contingency table
                                
                            //    var categoryCounts = new List<int>();
                            //    foreach (var value in allValues)
                            //    {   
                            //        int countctr = parameter.ParameterValues.Count(x => x == value);
                            //        categoryCounts.Add(countctr);
                            //    }
                            //    contingencyTable.Add(categoryCounts);

                            //    foreach (var countsd in contingencyTable)
                            //    {
                            //       MessageBox.Show(countsd.ToString());
                            //    }
                            //}
                            



                        }

                        else if (parameter.NominalOrScale == "Scale")
                        {


                            var sortedKeys = parameter.FormattedValues.Keys.OrderBy(key => key).ToList();
                            var SortedGroupKeysAll = GroupPara.ParameterValues.Distinct().OrderBy(key => key).ToList();



                            count = 4; // Count for scale parameters



                            foreach (int groupValue in SortedGroupKeysAll)
                            {
                                if (sortedKeys.Contains(groupValue))
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

                                else
                                {
                                    column = column + 2;
                                }



                            }
                            PerformTest(table, numberofgroups, parameter, startingRow, WordTableColumns - 2, WordTableColumns, issame, ComparativeTables[tableindex]);
                        }

                        // Increment newRowCount for the next parameter
                        newRowCount += count;
                    }

                    //Remove No and %
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
                    //InsertNeachGroup(table, eachgroupcount, ComparativeTables[tableindex]);

                    InsertNeachGroup_New(table, eachgroupcount, ComparativeTables[tableindex]);

                    InsertPairwiseF(table, ComparativeTables[tableindex], WordTableColumns);




                    SetMarginBasedOngroups(table, numberofgroups);


                    string groupText = numberofgroups.ToWords();

                    InsertYesonly(table, ComparativeTables[tableindex], WordTableColumns);


                    if (ComparativeTables[tableindex].HasTotalColumn)
                    {
                        AddTotalColumn(table, ComparativeTables[tableindex]);
                    }
                    //ChisquarePairwise(table, ComparativeTables[tableindex], numberofgroups , WordTableColumns);
                    wordObj.FormatTable(table, 12);

                    //ModifyGroupSpaces(table, ComparativeTables[tableindex], numberofgroups);


                    //if (!issame)
                    //{
                    //    ModifyTestSpaces(table, ComparativeTables[tableindex], WordTableColumns);
                    //}


                    InsertDefUnderTableNew(table, section, ComparativeTables[tableindex], groupText);


                    //Designed for maha till now
                    CustomMarginFormat(table);
                }

               



            }

        }

        public void ChisquarePairwise(IWTable table , ComparativeTable comparativeTable , int numberofgroups , int WordTableColumns)
        {
            int startrow = 2;
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

                else if (parameter.NominalOrScale == "Nominal")
                {
                    //count = parameter.ParameterValues.Distinct().Count() + 1;
                    if (parameter.NominalIsYes)
                    {
                        count = 1;
                    }
                    else if(parameter.ParameterNominalSig)
                    {
                        count = parameter.DIC_LablesIfNomainal.Keys.Count() + 2;
                    }
                    else
                    {
                        count = parameter.DIC_LablesIfNomainal.Keys.Count() + 1;
                    }

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


                if(numberofgroups >2)
                {
                    if(parameter.NominalOrScale == "Nominal")
                    {
                        if(parameter.ParameterNominalSig)
                        {
                            List<string> pairwise = new List<string>();
                            pairwise = pythonStat.ChisquarePairwise(parameter);

                           

                            WTableRow row;

                            int insert = 0;

                            insert = startrow + count -1;



                            row = table.AddRow();
                            table.Rows.Insert(insert, row);

                            int rowindex = table.Rows.IndexOf(row);

                            wordObj.AddParaCombined(table, rowindex, 0, "Sig. bet. grps.", true, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                            table.ApplyHorizontalMerge(rowindex, 1, WordTableColumns - 3);


                            for (int i = 0; i < WordTableColumns; i++)
                            {
                                if (parameter == lastParameter)
                                {

                                    table.Rows[rowindex].Cells[i].CellFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                                    table.Rows[rowindex].Cells[i].CellFormat.Borders.Top.LineWidth = 0.5f;
                                }
                                else if (!(parameter == lastParameter))
                                {
                                    //MessageBox.Show(rowindex.ToString());
                                    table.Rows[rowindex].Cells[i].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                                    table.Rows[rowindex].Cells[i].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                                }
                            }

                            WParagraph paragraph = (WParagraph)table[rowindex, 1].AddParagraph();
                            for (int i = 0; i < pairwise.Count; i++)
                            {
                                string groupnumber = "";

                                bool issig = false;
                                try
                                {
                                    double checkforP = double.Parse(pairwise[i]);
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


                                //wordObj.AddParaCombined(table, rowindex, 1, "p" , false , true , Syncfusion.Drawing.Color.Empty , Syncfusion.Drawing.Color.Black);
                                //wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, groupnumber, "Sub");

                                wordObj.InsertPairwiseF(table, rowindex, 1, "p", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                                wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, groupnumber, "Sub");
                                if (!pairwise[i].Contains("<"))
                                {
                                    wordObj.InsertPairwiseF(table, rowindex, 1, "=", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                                }

                                wordObj.InsertPairwiseF(table, rowindex, 1, pairwise[i], paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);

                                if (issig)
                                {
                                    wordObj.SubSuperScriptText(table, rowindex, 1, Syncfusion.Drawing.Color.Empty, "*", "Super");
                                }

                                if (i < pairwise.Count - 1)
                                {
                                    wordObj.InsertPairwiseF(table, rowindex, 1, ",", paragraph, false, true, Syncfusion.Drawing.Color.Empty, Syncfusion.Drawing.Color.Black);
                                }

                            }

                        }
                    }
                }







                startrow = startrow + count;




            }








        }

        public void CustomMarginFormat(IWTable table)
        {
            //Designed for maha till now
            //check_Maha.Checked
            if (Maha_Comparative) 
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
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);

            // Accessing the first worksheet in the Excel file
            worksheet = workbook.Worksheets[0];
            Sheet2 = workbook.Worksheets[1];


            AddHeadersToParameter();

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
                pic_ClearNominalList.Visible = true;
            }
            else if(cmb_ChooseTableFormat.Text == "Periods Groups")
            {
                

                list_Nominal.Visible = false;
                pic_RemoveNominalList.Visible = false;
                pic_AllParaToNominal.Visible=false;
                pic_ifyes.Visible=false;
                lbl_Nominal.Visible = false;
                pic_ClearNominalList.Visible = false;


            }
            else if(cmb_ChooseTableFormat.Text == "Groups 2 periods")
            {
                list_Nominal.Visible = false;
                pic_RemoveNominalList.Visible = false;
                pic_AllParaToNominal.Visible = false;
                pic_ifyes.Visible = false;
                lbl_Nominal.Visible = false;
                pic_ClearNominalList.Visible = false;
            }
            else if(cmb_ChooseTableFormat.Text == "Descriptive Periods No Test")
            {
                lbl_Nominal.Visible = true;
                list_Nominal.Visible = true;
                pic_RemoveNominalList.Visible = true;
                pic_AllParaToNominal.Visible = true;
                pic_ifyes.Visible = true;
                pic_ClearNominalList.Visible = true;

                lbl_PeriodsCount.Visible = true;
                txt_PeriodCount.Visible = true; 
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
                else if (cmb_ChooseTableFormat.Text == "Periods Groups")
                {
                    ChooseTable chooseTable = new ChooseTable();
                    chooseTable.Table_ChooseTable = Resource.Comparative_PeriodsUp_Groups;
                    chooseTable.ShowDialog();

                }
                else if (cmb_ChooseTableFormat.Text == "Groups 2 periods")
                {
                    ChooseTable chooseTable = new ChooseTable();
                    chooseTable.Table_ChooseTable = Resource.Groups_2Periods;
                    chooseTable.ShowDialog();

                }

                else if (cmb_ChooseTableFormat.Text == "Paper")
                {
                    ChooseTable chooseTable = new ChooseTable();
                    chooseTable.Table_ChooseTable = Resource.Paper_Comparative;
                    chooseTable.ShowDialog();

                }

                else if (cmb_ChooseTableFormat.Text == "Pathology Relation")
                {
                    ChooseTable chooseTable = new ChooseTable();
                    chooseTable.Table_ChooseTable = Resource.Pathology;
                    chooseTable.ShowDialog();

                }
                else if (cmb_ChooseTableFormat.Text == "Pathology Comparative")
                {
                    ChooseTable chooseTable = new ChooseTable();
                    chooseTable.Table_ChooseTable = Resource.Pathology_Comparative;
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
        private ParameterInfo parameterInfoForm; // Store the form instance

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
                    if (parameterInfoForm == null || parameterInfoForm.IsDisposed)
                    {
                        parameterInfoForm = new ParameterInfo();
                        parameterInfoForm.Show();
                    }

                    // Update the existing form with the new data
                    parameterInfoForm.UpdateParameter(paraName);
                }
            }
        }

        private void pic_AllLists_Click(object sender, EventArgs e)
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
        public void ClearAllLists()
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

            var selectedItemsGroups = new List<object>();
            foreach (var selectedItemGroup in list_Groups.Items)
            {
                selectedItemsGroups.Add(selectedItemGroup);
            }


            foreach (var selectedItemGroup in selectedItemsGroups)
            {
                list_Groups.Items.Remove(selectedItemGroup);
            }
        }
        private void pic_ClearNominalList_Click(object sender, EventArgs e)
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

            foreach (var item in selectedItemsNominal)
            {
                OrderedParameters.Remove(item.ToString());
            }
        }

        private void pic_clearNormalList_Click(object sender, EventArgs e)
        {
            var selectedItemsNormal = new List<object>();
            foreach (var selectedItemNormal in list_NormalScale.Items)
            {
                selectedItemsNormal.Add(selectedItemNormal);
            }


            foreach (var selectedItemNormal in selectedItemsNormal)
            {
                list_NormalScale.Items.Remove(selectedItemNormal);
            }

            foreach (var item in selectedItemsNormal)
            {
                OrderedParameters.Remove(item.ToString());
            }
        }

        private void pic_ClearAbnormalList_Click(object sender, EventArgs e)
        {
            var selectedItemsAbnormal = new List<object>();
            foreach (var selectedItemAbnormal in list_AbnormalScale.Items)
            {
                selectedItemsAbnormal.Add(selectedItemAbnormal);
            }


            foreach (var selectedItemAbnormal in selectedItemsAbnormal)
            {
                list_AbnormalScale.Items.Remove(selectedItemAbnormal);
            }

            foreach (var item in selectedItemsAbnormal)
            {
                OrderedParameters.Remove(item.ToString());
            }
        }

        private void txt_TableName_TextChanged(object sender, EventArgs e)
        {

        }

        private void pic_Minimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void pic_removeallGroupsList_Click(object sender, EventArgs e)
        {
            var selectedItemsGroups = new List<object>();
            foreach (var selectedItemGroup in list_Groups.Items)
            {
                selectedItemsGroups.Add(selectedItemGroup);
            }


            foreach (var selectedItemGroup in selectedItemsGroups)
            {
                list_Groups.Items.Remove(selectedItemGroup);
            }
        }

        private void btn_Options_Click(object sender, EventArgs e)
        {
            Options_Frm options_ = new Options_Frm();
            options_.Load += options_.Options_Frm_Load; // Ensure values are loaded
            options_.ShowDialog();
        }

        private void pic_DoneUpdatingTableType_Click(object sender, EventArgs e)
        {
            if(cmb_TableNames.SelectedIndex != -1)
            {
                string UpdatedTableType = cmb_UpdateTableType.Text;

                string selectedTableName = cmb_TableNames.SelectedItem.ToString();

                // Find the ComparativeTable instance corresponding to the selected table name
                var selectedTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                if (selectedTable.FormatType != UpdatedTableType)
                {
                    selectedTable.FormatType = UpdatedTableType;
                    MessageBox.Show($"Update table {selectedTableName} type to : {selectedTable.FormatType}", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Table {selectedTableName} already with same Format", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show($"Please Select a table", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            


            
        }

        private void pic_UpdateTableParameters_Click(object sender, EventArgs e)
        {
            if (cmb_TableNames.SelectedIndex != -1)
            {
                string selectedTableName = cmb_TableNames.SelectedItem.ToString();
                
                var comparativeTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                List<Parameter> oldParameters = comparativeTable.Parameters.ToList();

                bool HasGroup = oldParameters.Any(p => p.IsGroup);
                bool HasNominal = oldParameters.Any(p => p.NominalOrScale == "Nominal");
                bool HasNormalScale = oldParameters.Any(p => p.NormalOrAbnormal == "Normal");
                bool HasAbnormalScale= oldParameters.Any(p => p.NormalOrAbnormal == "Abnormal");

                int flag = 0;

                // Build the message string with conditional formatting for non-null lists
                StringBuilder messageBuilder = new StringBuilder("Are you sure you want to add the following Table?\n");

                if (list_Groups.Items.Count > 0)
                {
                    if (!HasGroup)
                    {
                        flag = 1;
                        //messageBuilder.AppendLine("Groups:\n" + String.Join("\n", list_Groups.Items.Cast<string>()));
                    }

                    foreach (var listItem in list_Groups.Items)
                    {
                        foreach (var ParameterItem in oldParameters)
                        {
                            if(ParameterItem.IsGroup)
                            {
                                if (!(listItem.ToString() == ParameterItem.Name))
                                {
                                    flag = 1;
                                    //messageBuilder.AppendLine("Groups:\n" + String.Join("\n", list_Groups.Items.Cast<string>()));
                                }
                            }
                        }
                    }

                    

                }

                if (list_Nominal.Items.Count > 0)
                {
                    if (!HasNominal)
                    {
                        flag = 1;
                        //messageBuilder.AppendLine("Nominal:\n" + String.Join("\n", list_Nominal.Items.Cast<string>()));
                    }

                    foreach (var listItem in list_Nominal.Items)
                    {
                        foreach (var ParameterItem in oldParameters)
                        {
                            if (ParameterItem.NominalOrScale == "Nominal")
                            {
                                if (!(listItem.ToString() == ParameterItem.Name))
                                {
                                    flag = 1;
                                    //messageBuilder.AppendLine("Nominal:\n" + String.Join("\n", list_Nominal.Items.Cast<string>()));
                                }
                            }
                                
                        }
                    }
                }

                if (list_NormalScale.Items.Count > 0)
                {
                    if (!HasNormalScale)
                    {
                        flag = 1;
                        //messageBuilder.AppendLine("Normal Scale:\n" + String.Join("\n", list_NormalScale.Items.Cast<string>()));
                    }

                    foreach (var listItem in list_NormalScale.Items)
                    {
                        foreach (var ParameterItem in oldParameters)
                        {
                            if (ParameterItem.NominalOrScale == "Scale")
                            {
                                if (ParameterItem.NormalOrAbnormal == "Normal")
                                {
                                    if (!(listItem.ToString() == ParameterItem.Name))
                                    {
                                        flag = 1;
                                        //messageBuilder.AppendLine("Normal Scale:\n" + String.Join("\n", list_NormalScale.Items.Cast<string>()));
                                    }
                                }
                            }
                                
                        }
                    }
                    
                }

                if (list_AbnormalScale.Items.Count > 0)
                {
                    if (!HasAbnormalScale)
                    {
                        flag = 1;
                        //messageBuilder.AppendLine("Abnormal Scale:\n" + String.Join("\n", list_AbnormalScale.Items.Cast<string>()));
                    }

                    foreach (var listItem in list_AbnormalScale.Items)
                    {
                        foreach (var ParameterItem in oldParameters)
                        {
                            if (ParameterItem.NominalOrScale == "Scale")
                            {
                                if (ParameterItem.NormalOrAbnormal == "Abnormal")
                                {
                                    if (!(listItem.ToString() == ParameterItem.Name))
                                    {
                                        flag = 1;
                                       //messageBuilder.AppendLine("Abnormal Scale:\n" + String.Join("\n", list_AbnormalScale.Items.Cast<string>()));
                                    }
                                }
                            }
                                
                        }
                    }
                    
                }
                messageBuilder.AppendLine("Groups:\n" + String.Join("\n", list_Groups.Items.Cast<string>()));
                messageBuilder.AppendLine("Nominal:\n" + String.Join("\n", list_Nominal.Items.Cast<string>()));
                messageBuilder.AppendLine("Normal Scale:\n" + String.Join("\n", list_NormalScale.Items.Cast<string>()));
                messageBuilder.AppendLine("Abnormal Scale:\n" + String.Join("\n", list_AbnormalScale.Items.Cast<string>()));

                string message = messageBuilder.ToString();

                if (flag == 1)
                {
                    // Show the MessageBox only if there are parameters to display
                    if (MessageBox.Show(message, "Adding Table", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {


                        RemoveTable();
                        oldParameters.Clear();



                        string TableName = AddTableUI();


                        if (cmb_ChooseTableFormat.Text == "Relation" || cmb_ChooseTableFormat.Text == "Relation Scale Pathology")
                        {
                            ComparativeBasic_Relation(TableName);
                        }
                        else
                        {
                            ComparativeBasic(TableName);
                        }



                        CheckFullEmptyParameters(TableName);
                        CheckForOthers_inNominal(TableName);




                    }
                }
                else
                {
                    MessageBox.Show("No Parameters selected to be added", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }






            }
            else
            {
                MessageBox.Show("No table selected to be Updated", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }

        private void pic_RefreshListsUpdate_Click(object sender, EventArgs e)
        {
            if (cmb_TableNames.SelectedIndex != -1)
            {
                ClearAllLists();
                



                string selectedTableName = cmb_TableNames.SelectedItem.ToString();

                var comparativeTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                foreach (var parameter in comparativeTable.Parameters)
                {
                    if(parameter.IsGroup)
                    {
                        list_Groups.Items.Add(parameter.Name);
                    }
                    else if (parameter.NominalOrScale == "Nominal")
                    {
                        list_Nominal.Items.Add(parameter.Name);
                    }
                    else if (parameter.NominalOrScale == "Scale")
                    {
                        if(parameter.NormalOrAbnormal == "Normal")
                        {
                            list_NormalScale.Items.Add(parameter.Name);
                        }
                        else if(parameter.NormalOrAbnormal == "Abnormal")
                        {
                            list_AbnormalScale.Items.Add(parameter.Name);
                        }
                    }
                }


                MessageBox.Show("Lists Updated", "Update Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void pic_SigAdj_Click(object sender, EventArgs e)
        {
            if(cmb_TableNames.SelectedIndex != -1)
            {
                string selectedTableName = cmb_TableNames.SelectedItem.ToString();
                var selectedTable = ComparativeTables.FirstOrDefault(table => table.TableName == selectedTableName);

                Significant_Adjust_Frm significant_Adjust_ = new Significant_Adjust_Frm();
                significant_Adjust_.comparativeTable_ToModify_SignifAdj = selectedTable;
                significant_Adjust_.ShowDialog();
            }



            
        }

        private void list_AbnormalScale_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pic_RemoveSubGroupsList_Click(object sender, EventArgs e)
        {
            if (list_SubGroups.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_SubGroups.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_SubGroups.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_ClearAllSubGroups_Click(object sender, EventArgs e)
        {
            var selectedItemsGroups = new List<object>();
            foreach (var selectedItemGroup in list_SubGroups.Items)
            {
                selectedItemsGroups.Add(selectedItemGroup);
            }


            foreach (var selectedItemGroup in selectedItemsGroups)
            {
                list_SubGroups.Items.Remove(selectedItemGroup);
            }
        }

        private void pic_AllParaToSubGroups_Click(object sender, EventArgs e)
        {
            for (int i = data_allPara.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = data_allPara.SelectedRows[i];
                var cellValue = row.Cells[0].Value;
                if (cellValue != null)
                {
                    string item = cellValue.ToString();
                    list_SubGroups.Items.Add(item);
                }

            }
        }

        private void cmb_ChooseTableFormat_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            FormManager.ShowStandaloneForm<MainFormsDesign>();
        }



        private void pic_AddNew_Click(object sender, EventArgs e)
        {
            // Restore saved design type (if needed)
            var designType = FormDataTransfer.Get<string>("TableDesignType");
            if (!string.IsNullOrEmpty(designType))
            {
                cmb_ChooseTableFormat.Text = designType;
            }

            string tableName = AddTableUINew();

            if (!string.IsNullOrEmpty(tableName))
            {
                StatBasic(tableName); // Only call if table creation succeeded
            }
        }


        private void label5_Click(object sender, EventArgs e)
        {

        }
    }
}
