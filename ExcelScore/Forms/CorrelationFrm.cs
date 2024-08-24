using Aspose.Cells;
using ExcelScore.Classes;
using MathNet.Numerics.Distributions;
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
using MathNet.Numerics.Statistics;
using Accord.Statistics.Distributions.Univariate;
using Syncfusion.DocIO.DLS;

namespace ExcelScore.Forms
{
    public partial class CorrelationFrm : Form
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
        public CorrelationFrm()
        {
            InitializeComponent();
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
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
        private void Correlation_Load(object sender, EventArgs e)
        {
            AddHeadersToParameter();
            lbl_CurrentCorr.Text = "Correlation : Matrix";
        }
        public void AddTableClass(string tableName)
        {
            var comparativeTable = new ComparativeTable
            {
                TableName = tableName,
                Parameters = new List<Parameter>()
            };

            // Add parameters from list_Groups
            

            comparativeTable.CorreType = cmb_CorreType.SelectedItem.ToString();





            // Add parameters from list_NormalScale
            foreach (var item in list_selectedParameters.Items)
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
        public string AddTableUI()
        {
            string TableName = null;
            if (!string.IsNullOrWhiteSpace(txt_TableName.Text) &&

                 list_selectedParameters.Items.Count > 0 && cmb_CorreType.SelectedIndex != -1)
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
                                        continue;
                                    }

                                    if (cellData != null && double.TryParse(cellData.ToString(), out cellValueDouble))
                                    {
                                        //MessageBox.Show(cellValueDouble.ToString());
                                        // Add the cell value to the parameter's values list
                                        parameter.ParameterValues.Add(cellValueDouble);
                                        //MessageBox.Show(cellValueDouble.ToString());
                                    }










                                }
                            }
                        }
                    }
                }
                // Iterate over each parameter in the table

            }
        }

        private void pic_addTable_Click(object sender, EventArgs e)
        {
            string TableName = AddTableUI();
            GetDataValues(TableName);
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
        pythonStat pythonStat = new pythonStat();
        WordClass wordObj = new WordClass();
        ManualTests manualTests = new ManualTests();    

        public void CalculateCorrelations(ComparativeTable table)
        {
            // Iterate over each pair of parameters
            
        }
        public void DrawMatrixCorr()
        {
            int lightenPercentage = 70;

            int baseRed = 226;
            int baseGreen = 107;
            int baseBlue = 10;

            int red = baseRed + (255 - baseRed) * lightenPercentage / 100;
            int green = baseGreen + (255 - baseGreen) * lightenPercentage / 100;
            int blue = baseBlue + (255 - baseBlue) * lightenPercentage / 100;

            int argbLighter = (255 << 24) | (red << 16) | (green << 8) | blue;

            Syncfusion.Drawing.Color lighterOrange = Syncfusion.Drawing.Color.FromArgb(argbLighter);
            Sheet2 = workbook.Worksheets[1];

            for (int tableindex = 0; tableindex < ComparativeTables.Count; tableindex++)
            {
                IWSection section = wordObj.CreatePortraitSection();
                wordObj.AddCorrTitle(section, ComparativeTables[tableindex].TableName, 1);

                int Variablerows = ComparativeTables[tableindex].Parameters.Count;

                int WordTableColumns = 2+ Variablerows;

                int WordTableRows = 1 + (2*Variablerows);

                IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                wordObj.GeneralTableFormat(table);

                wordObj.ApplyGeneralMatrixCorrBorders(table, WordTableRows, WordTableColumns);

                wordObj.SetCorrWidths(table, WordTableRows, WordTableColumns);


                wordObj.CorrCustom(table, WordTableRows, WordTableColumns, ComparativeTables[tableindex]);


                int row = 1;
                int emptyCol = 2;

                for (int i = 0; i < ComparativeTables[tableindex].Parameters.Count - 1; i++)
                {
                    
                    
                    int col = 3+i;
                    //int column = row + 2;

                    string[] CorrOutput = new string[2];
                    for (int j = i+1; j < ComparativeTables[tableindex].Parameters.Count; j++)
                    {
                        bool issig = false;
                        if ( i != j)
                        {
                            var param1 = ComparativeTables[tableindex].Parameters[i];
                            var param2 = ComparativeTables[tableindex].Parameters[j];

                            // Convert parameter values to arrays for MathNet Numerics
                            double[] x = param1.ParameterValues.ToArray();
                            double[] y = param2.ParameterValues.ToArray();

                            CorrOutput = manualTests.CalculateCorrelation(x, y, ComparativeTables[tableindex].CorreType);

                            //MessageBox.Show(CorrOutput[0].ToString());

                            if (CorrOutput[1] == "<0.001")
                            {
                                issig = true;
                            }
                            else
                            {
                                if (double.Parse(CorrOutput[1]) < 0.05)
                                {
                                    issig = true;
                                }
                                else if (double.Parse(CorrOutput[1]) >= 0.05)
                                {
                                    issig = false;
                                }
                                        
                            }

                            wordObj.Addpara_CenterNoBOLD(table, row, col, CorrOutput[0]);

                            wordObj.Addpara_CenterNoBOLD(table, row + 1, col, CorrOutput[1]);
                            

                            if (issig) 
                            {
                                wordObj.SubSuperScriptText(table, row, col, Syncfusion.Drawing.Color.Empty, "*", "Super");
                                wordObj.SubSuperScriptText(table, row + 1, col, Syncfusion.Drawing.Color.Empty, "*", "Super");
                                table.Rows[row].Cells[col].CellFormat.BackColor = lighterOrange;
                                table.Rows[row+1].Cells[col].CellFormat.BackColor = lighterOrange;
                            }
                            //wordObj.InsertTest_P_Corr(table, row, col, CorrOutput);
                            col++;
                        }

                        

                        

                    }

                    for(int columnstart = 0; columnstart < WordTableColumns; columnstart++)
                    {
                        table.Rows[row + 1].Cells[columnstart].CellFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                        table.Rows[row + 1].Cells[columnstart].CellFormat.Borders.Bottom.LineWidth = 0.5f;
                    }

                    for(int innerrow = row; innerrow < WordTableRows; innerrow++)
                    {

                        if(innerrow == row)
                        {
                            wordObj.Addpara_CenterNoBOLD(table, innerrow, emptyCol, "1.000");
                        }
                        else
                        {
                            table.Rows[innerrow].Cells[emptyCol].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;
                            //wordObj.Addpara_CenterNoBOLD(table, innerrow, emptyCol, "Test");
                        }
                        
                        
                    }
                    emptyCol++;

                    row += 2;


                }

                wordObj.Addpara_CenterNoBOLD(table, WordTableRows-2, WordTableColumns-1, "1.000");

                table.Rows[WordTableRows - 1].Cells[WordTableColumns - 1].CellFormat.BackColor = Syncfusion.Drawing.Color.LightGray;
                //wordObj.Addpara_CenterNoBOLD(table, WordTableRows - 1, WordTableColumns - 1, "Test");



                wordObj.FormatTable(table, 10);
            }

        }
        private void btn_Done_Click(object sender, EventArgs e)
        {
            //GetDataValues();
            //pythonStat.InitPython();

            wordObj.InitWord();

            DrawMatrixCorr();

            wordObj.SaveWord();
        }
        public bool ISMatrixCorr = true;
        private void btn_Swap_Click(object sender, EventArgs e)
        {
            if(ISMatrixCorr)
            {
                lbl_CurrentCorr.Text = "Correlation : Normal";
                ISMatrixCorr = false;
            }
            else if (!ISMatrixCorr)
            {
                lbl_CurrentCorr.Text = "Correlation : Matrix";
                ISMatrixCorr = true;
            }


            
        }

        private void pic_RemoveDependentAbnormal_Click(object sender, EventArgs e)
        {
            if (list_selectedDepenAbnormal.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_selectedDepenAbnormal.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_selectedDepenAbnormal.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }
    }
}
