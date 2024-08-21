using Aspose.Cells;
using ExcelScore.Classes;

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelScore.Forms
{
    public partial class NewAddDomain : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        //Used to switch between Adding using From and To Or Adding Single By listBox in UI
        public bool FromToAddUI = true;

        //DataGrid From Form Before
        public DataGridView Dgv { get; set; }

        //List To Store DataGrid Titles = Questions
        List<string> DataGridHeaders = new List<string>();


        public List<NewDomainClass> AllDomains = new List<NewDomainClass>();
        public NewAddDomain()
        {
            InitializeComponent();
        }
        NewDomainClass NewDomainClassObj = new NewDomainClass();
        public int UI_RangeFrom;
        public int UI_RangeTo;
        public bool UiRangeFrom;
        public bool UiRangeTo;
        private void btn_AddDomain_Click(object sender, EventArgs e)
        {
            //AddDomainDatarefac();
            AddDomainData();

        }

        public void AddDomainDatarefac() 
        {
            if (!ValidateInputs())
            {
                MessageBox.Show("Please ensure all fields are correctly filled out.");
                return; // Exit early if validation fails
            }

            // Initialize a new domain object from validated inputs
            NewDomainClassObj = new NewDomainClass
            {
                DomainName = txt_DomainName.Text,
                RangeFrom = int.Parse(txt_RangeFrom.Text),
                RangeTo = int.Parse(txt_RangeTo.Text),
                Level_Percent_Or_TotalScore = cmb_LevelPercentOrTotalChoose.SelectedItem.ToString()
            };

            ParseLevels(list_Level.Items);
            ParseQuestions(list_QuestionsBeforeReverse.Items, NewDomainClassObj.Questions_Not_Reversed);
            ParseQuestions(list_QuestionsReversed.Items, NewDomainClassObj.Questions_Reversed, isReversed: true);
            ParseSortedQuestions(listbox_Sort.Items);

            AllDomains.Add(NewDomainClassObj);
            MessageBox.Show("Domain Added");
        }
        private bool ValidateInputs()
        {
            // Add more detailed validation as needed
            return !string.IsNullOrEmpty(txt_DomainName.Text) &&
                   int.TryParse(txt_RangeFrom.Text, out _) &&
                   int.TryParse(txt_RangeTo.Text, out _) &&
                   cmb_LevelPercentOrTotalChoose.SelectedItem != null;
        }
        private void ParseLevels(IEnumerable items)
        {
            foreach (string item in items)
            {
                // Simplified for brevity; implement parsing logic here
                // For example:
                string[] parts = item.Split(' ');
                // Add logic to parse and add level information to NewDomainClassObj
            }
        }

        private void ParseQuestions(IEnumerable items, List<string> targetList, bool isReversed = false)
        {
            foreach (string item in items)
            {
                // Add directly or handle reversal as needed
                targetList.Add(item);
            }
        }
        private void ParseSortedQuestions(IEnumerable items)
        {
            foreach (string item in items)
            {
                bool isReversed = list_QuestionsReversed.Items.Contains(item);
                NewDomainClassObj.QuestionsSorted_ReverseValue.Add(item, isReversed ? 1 : 0);
            }
        }

        public void AddDomainData()
        {
            NewDomainClassObj = new NewDomainClass();

            /*if (!string.IsNullOrEmpty(txt_DomainName.Text)
                && !string.IsNullOrEmpty(txt_RangeFrom.Text) && int.TryParse(txt_RangeFrom.Text, out int rangeFrom)
                && !string.IsNullOrEmpty(txt_RangeTo.Text) && int.TryParse(txt_RangeTo.Text, out int rangeTo)
                && cmb_LevelPercentOrTotalChoose.SelectedItem != null
                && cmb_LevelComparison.SelectedItem != null
                && !string.IsNullOrEmpty(txt_LevelRangeFrom.Text)
                && !string.IsNullOrEmpty(txt_LevelOrRangeTo.Text)
                && !string.IsNullOrEmpty(txt_LevelValue.Text))
            {*/
            NewDomainClassObj.DomainName = txt_DomainName.Text;
            NewDomainClassObj.RangeFrom = int.Parse(txt_RangeFrom.Text);
            NewDomainClassObj.RangeTo = int.Parse(txt_RangeTo.Text);
            NewDomainClassObj.Level_Percent_Or_TotalScore = cmb_LevelPercentOrTotalChoose.SelectedItem.ToString();

            foreach (string item in list_Level.Items)
            {
                if (item.Contains("Greater Equal") || item.Contains("Smaller Equal"))
                {
                    string[] parts1 = item.Split(' ');
                    // Extract the relevant parts from the first input string
                    string comparison1 = parts1[0];  // "Greater"
                    string comparison2 = parts1[1];
                    int value1 = int.Parse(parts1[3].Split('-')[0]);  // 20
                    int subValue1 = int.Parse(parts1[3].Split('-')[1]);  // 1

                    NewDomainClassObj.Level_Comparison_Choice.Add(comparison1 + " " + comparison2);
                    NewDomainClassObj.Level_From_To.Add(value1);
                    NewDomainClassObj.Level_Value.Add(subValue1);


                }
                else if (item.Contains("Range"))
                {
                    string[] parts2 = item.Split(' ');
                    string range2 = parts2[0];  // "Range"
                    int minValue2 = int.Parse(parts2[1]);  // 20
                    int maxValue2 = int.Parse(parts2[2].Split('-')[0]);  // 40
                    int subValue2 = int.Parse(parts2[2].Split('-')[1]);  // 2
                    NewDomainClassObj.Level_Comparison_Choice.Add(range2);
                    NewDomainClassObj.Level_From_To.Add(minValue2);
                    NewDomainClassObj.Level_From_To.Add(maxValue2);
                    NewDomainClassObj.Level_Value.Add(subValue2);


                }
                else if (item.Contains("Greater") || item.Contains("Smaller"))
                {
                    string[] parts1 = item.Split(' ');
                    // Extract the relevant parts from the first input string
                    string comparison1 = parts1[0];  // "Greater"
                    int value1 = int.Parse(parts1[2].Split('-')[0]);  // 20
                    int subValue1 = int.Parse(parts1[2].Split('-')[1]);  // 1
                    NewDomainClassObj.Level_Comparison_Choice.Add(comparison1);
                    NewDomainClassObj.Level_From_To.Add(value1);
                    NewDomainClassObj.Level_Value.Add(subValue1);

                }
                else if (item.Contains("Equal To"))
                {
                    string[] parts1 = item.Split(' ');
                    // Extract the relevant parts from the first input string
                    string comparison1 = parts1[0];  // "Greater"
                    int value1 = int.Parse(parts1[2].Split('-')[0]);  // 20
                    int subValue1 = int.Parse(parts1[2].Split('-')[1]);  // 1
                    NewDomainClassObj.Level_Comparison_Choice.Add(comparison1);
                    NewDomainClassObj.Level_From_To.Add(value1);
                    NewDomainClassObj.Level_Value.Add(subValue1);

                }

            }

            //NewDomainClassObj.Level_Comparison_Choose = cmb_LevelComparison.SelectedItem.ToString();
            //NewDomainClassObj.Level_RangeFrom = txt_LevelRangeFrom.Text;
            //NewDomainClassObj.Level_Comparison_Or_RangeTo = txt_LevelOrRangeTo.Text;
            //NewDomainClassObj.Level_Comparison_Value = txt_LevelValue.Text;

            foreach (string s in list_QuestionsBeforeReverse.Items)
            {
                NewDomainClassObj.Questions_Not_Reversed.Add(s);
            }
            foreach (string s in list_QuestionsReversed.Items)
            {
                NewDomainClassObj.Questions_Reversed.Add(s);
            }

            try
            {
                foreach (string s in listbox_Sort.Items)
                {
                    string Question = s;
                    int Isreverse = 0;
                    if (list_QuestionsBeforeReverse.Items.Contains(Question))
                    {

                        Isreverse = 0;
                        NewDomainClassObj.QuestionsSorted_ReverseValue.Add(Question, Isreverse);


                    }
                    else if (list_QuestionsReversed.Items.Contains(Question))
                    {
                        Isreverse = 1;
                        NewDomainClassObj.QuestionsSorted_ReverseValue.Add(Question, Isreverse);
                    }

                    
                }

                AllDomains.Add(NewDomainClassObj);
                MessageBox.Show("Domain Added");
            }
            catch(Exception)
            {
                MessageBox.Show("Sort data please to proceed!");
            }

            
            //bool result = NewDomainClassObj
            
            
        }
        
        public void DrawLines()
        {
            //Drawing Lines For UI
            label_line1.Text = "";
            label_line1.AutoSize = false;
            label_line1.Width = 2;
            label_line1.Height = 600;
            label_line1.BorderStyle = BorderStyle.Fixed3D;


            lbl_line2.Text = "";
            lbl_line2.AutoSize = false;
            lbl_line2.Width = 271;
            lbl_line2.Height = 2;
            lbl_line2.BorderStyle = BorderStyle.Fixed3D;

            lbl_Line3.Text = "";
            lbl_Line3.AutoSize = false;
            lbl_Line3.Width = 271;
            lbl_Line3.Height = 2;
            lbl_Line3.BorderStyle = BorderStyle.Fixed3D;

            lbl_line4.Text = "";
            lbl_line4.AutoSize = false;
            lbl_line4.Width = 2;
            lbl_line4.Height = 600;
            lbl_line4.BorderStyle = BorderStyle.Fixed3D;

            lbl_line5.Text = "";
            lbl_line5.AutoSize = false;
            lbl_line5.Width = 500;
            lbl_line5.Height = 2;
            lbl_line5.BorderStyle = BorderStyle.Fixed3D;
        }

        //Filling ComboBox From And To and Filling List Single , also Filling List DataGridHeaders
        public void FillComboAndList()
        {
            for (int i = 0; i < Dgv.Columns.Count; i++)
            {
                string header = Dgv.Columns[i].HeaderText.Trim();
                    
                DataGridHeaders.Add(header);
                cmb_From.Items.Add(header);
                cmb_To.Items.Add(header);
                list_Single.Items.Add(header);

            }
        }
        private void NewAddDomain_Load(object sender, EventArgs e)
        {
            DrawLines();
            list_Single.Hide();
            txt_LevelRangeFrom.Hide();
            FillComboAndList();
            lbl_sort.Hide();
            listbox_Sort.Hide();
            pic_SortQuestionsUP.Hide();
            pic_SortQuestionsDown.Hide();


        }
        public void ClearonSwitch()
        {
            txt_DomainName.Text = string.Empty;
            txt_RangeFrom.Text = string.Empty;
            txt_RangeTo.Text = string.Empty;
            txt_LevelOrRangeTo.Text = string.Empty;
            txt_LevelRangeFrom.Text = string.Empty;
            txt_LevelValue.Text = string.Empty;
            cmb_LevelComparison.Text = string.Empty;
            list_Level.Items.Clear();
            list_QuestionsBeforeReverse.Items.Clear();
            list_QuestionsReversed.Items.Clear();
            cmb_LevelPercentOrTotalChoose.Text = string.Empty;
        }

        private void btn_switch_Click(object sender, EventArgs e)
        {
            if(FromToAddUI)
            {
                ClearonSwitch();
                lbl_From.Hide();
                lbl_To.Hide(); 
                cmb_From.Hide();    
                cmb_To.Hide();
                list_Single.Show();
                FromToAddUI = false;
            }
            else if(!FromToAddUI)
            {
                ClearonSwitch();
                lbl_From.Show();
                lbl_To.Show();
                cmb_From.Show();
                cmb_To.Show();
                list_Single.Hide();
                FromToAddUI = true;
            }    

        }

        private void cmb_LevelComparison_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Selected = cmb_LevelComparison.SelectedItem.ToString();
            if(Selected == "Range")
            {
                txt_LevelRangeFrom.Show();
            }
            else
            {
                txt_LevelRangeFrom.Hide();
            }
        }

        private void pic_AddLevel_Click(object sender, EventArgs e)
        {
            string Comparison = cmb_LevelComparison.SelectedItem.ToString();
            string PercentOrTotalScore = cmb_LevelPercentOrTotalChoose.Text;
            if (Comparison == "Range")
            {
                
                string RangeFrom = txt_LevelRangeFrom.Text;
                string RangeTo = txt_LevelOrRangeTo.Text;
                string LevelValue = txt_LevelValue.Text;
                list_Level.Items.Add(Comparison + " " + RangeFrom + " " + RangeTo + "-" + LevelValue);
            }
            else
            {
                string RangeTo = txt_LevelOrRangeTo.Text;
                string LevelValue = txt_LevelValue.Text;
                list_Level.Items.Add(Comparison + " " + RangeTo + "-" + LevelValue);
            }
            


            
        }

        private void pic_AddQuestions_Click(object sender, EventArgs e)
        {
            if (FromToAddUI)
            {
                list_QuestionsBeforeReverse.Items.Clear();
                list_QuestionsReversed.Items.Clear();
                string QuestionFrom = cmb_From.Text;
                string QuestionTo = cmb_To.Text;

               

                int indexFrom = DataGridHeaders.FindIndex(a => a.Contains(QuestionFrom));
                int indexTo = DataGridHeaders.FindIndex(a => a.Contains(QuestionTo));


                for(int i = indexFrom; i <= indexTo; i++)
                {
                    
                    list_QuestionsBeforeReverse.Items.Add(DataGridHeaders[i]);
                    listbox_Sort.Items.Add(DataGridHeaders[i]);
                }
            }
            else if (!FromToAddUI)
            {
                
                foreach (string s in list_Single.SelectedItems)
                {
                    list_QuestionsBeforeReverse.Items.Add(s);
                    listbox_Sort.Items.Add(s);
                }
               
            }
        }

        List<string> itemsToRemoveBeforeReverse = new List<string>();
        private void pic_AddReverseQuestions_Click(object sender, EventArgs e)
        {
            foreach (string s in list_QuestionsBeforeReverse.SelectedItems)
            {
                list_QuestionsReversed.Items.Add(s);
                itemsToRemoveBeforeReverse.Add(s);
            }
            foreach(string s in itemsToRemoveBeforeReverse)
            {
                list_QuestionsBeforeReverse.Items.Remove(s);
            }

        }

        private void pic_deleteSingleLevel_Click(object sender, EventArgs e)
        {
            if (list_Level.SelectedIndex != -1)
            {
                list_Level.Items.Remove(list_Level.SelectedItem);

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void pic_ClearLevelList_Click(object sender, EventArgs e)
        {
            list_Level.Items.Clear();
        }

        List<string> itemsToRemoveReverse = new List<string>();
        private void pic_RemoveReverseQuestions_Click(object sender, EventArgs e)
        {
            foreach (string s in list_QuestionsReversed.SelectedItems)
            {
                list_QuestionsBeforeReverse.Items.Add(s);
                itemsToRemoveReverse.Add(s);
            }
            foreach (string s in itemsToRemoveReverse)
            {
                list_QuestionsReversed.Items.Remove(s);
            }
        }


        List<string> deleteQuestionsBeforeReverse = new List<string>();
        private void pic_deleteQuestionsBeforeReverse_Click(object sender, EventArgs e)
        {
            foreach (string s in list_QuestionsBeforeReverse.SelectedItems)
            {

                deleteQuestionsBeforeReverse.Add(s);
            }
            foreach (string s in deleteQuestionsBeforeReverse)
            {
                list_QuestionsBeforeReverse.Items.Remove(s);
            }
        }

        private void pic_clearAllQuestionsBeforeReverse_Click(object sender, EventArgs e)
        {
            list_QuestionsBeforeReverse.Items.Clear();
        }
        private void MoveItemUp()
        {
            if (listbox_Sort.SelectedIndex > 0)
            {
                int selectedIndex = listbox_Sort.SelectedIndex;
                object selectedItem = listbox_Sort.SelectedItem;

                listbox_Sort.Items.RemoveAt(selectedIndex);
                listbox_Sort.Items.Insert(selectedIndex - 1, selectedItem);
                listbox_Sort.SelectedIndex = selectedIndex - 1;
            }
        }

        private void MoveItemDown()
        {
            if (listbox_Sort.SelectedIndex < listbox_Sort.Items.Count - 1)
            {
                int selectedIndex = listbox_Sort.SelectedIndex;
                object selectedItem = listbox_Sort.SelectedItem;

                listbox_Sort.Items.RemoveAt(selectedIndex);
                listbox_Sort.Items.Insert(selectedIndex + 1, selectedItem);
                listbox_Sort.SelectedIndex = selectedIndex + 1;
            }
        }

        private void btn_Sort_Click(object sender, EventArgs e)
        {
            lbl_NotReversed.Hide();
            lbl_Reversed.Hide();
            list_QuestionsBeforeReverse.Hide();
            list_QuestionsReversed.Hide();
            pic_AddReverseQuestions.Hide();
            pic_RemoveReverseQuestions.Hide();
            pic_clearAllQuestionsBeforeReverse.Hide();
            pic_deleteQuestionsBeforeReverse.Hide();


            lbl_sort.Show();
            listbox_Sort.Show();

        }
        public bool QuestionView = true;
        private void btn_QuestionsView_Click(object sender, EventArgs e)
        {
            if(QuestionView)
            {
                lbl_NotReversed.Hide();
                lbl_Reversed.Hide();
                list_QuestionsBeforeReverse.Hide();
                list_QuestionsReversed.Hide();
                pic_AddReverseQuestions.Hide();
                pic_RemoveReverseQuestions.Hide();
                pic_clearAllQuestionsBeforeReverse.Hide();
                pic_deleteQuestionsBeforeReverse.Hide();


                lbl_sort.Show();
                listbox_Sort.Show();
                pic_SortQuestionsUP.Show();
                pic_SortQuestionsDown.Show();
                QuestionView = false;
                btn_QuestionsView.Text = "Questions";
                listbox_Sort.Items.Clear();
                foreach(string s in list_QuestionsBeforeReverse.Items)
                {
                    listbox_Sort.Items.Add(s);
                }
                foreach(string s in list_QuestionsReversed.Items)
                {
                    listbox_Sort.Items.Add(s);
                }
            }

            else if(!QuestionView)
            {
                lbl_NotReversed.Show();
                lbl_Reversed.Show();
                list_QuestionsBeforeReverse.Show();
                list_QuestionsReversed.Show();
                pic_AddReverseQuestions.Show();
                pic_RemoveReverseQuestions.Show();
                pic_clearAllQuestionsBeforeReverse.Show();
                pic_deleteQuestionsBeforeReverse.Show();


                lbl_sort.Hide();
                listbox_Sort.Hide();
                pic_SortQuestionsUP.Hide();
                pic_SortQuestionsDown.Hide();
                QuestionView = true;
                btn_QuestionsView.Text = "Sort";
            }



            
        }

        private void pic_SortQuestionsUP_Click(object sender, EventArgs e)
        {
            MoveItemUp();
        }

        private void pic_SortQuestionsDown_Click(object sender, EventArgs e)
        {
            MoveItemDown();
        }

        private void cmb_LevelPercentOrTotalChoose_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txt_LevelOrRangeTo_TextChanged(object sender, EventArgs e)
        {

        }
        List<string> listTitles = new List<string>
        {
            "Total" , "Avg" , "Perc" ,"Level"
        };
        
        ColorClass colorClass = new ColorClass();
        //public int trackCtr = 0;
        int myalldomains;
        Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(ExcelFunctions.filepath);
        public void AddData()
        {
            int trackCtr = 0;
            int TableVlookctr = 0;
            
            
            workbook.Worksheets.RemoveAt("Sheet2");
            workbook.Worksheets.Insert(1, SheetType.Worksheet, "Sheet2");

            workbook.Worksheets.RemoveAt("Sheet3");
            workbook.Worksheets.Insert(2, SheetType.Worksheet, "Sheet3");

            Aspose.Cells.Worksheet worksheet = workbook.Worksheets[1];

            for (myalldomains = 0; myalldomains < AllDomains.Count; myalldomains++)
            {
                int TrackListFromToctr = 0;
                int TrackListValuectr = 0;
                //
                Color MyRandomColor = colorClass.GenerateRandomColor();
                int firstRow = 0;
                int firstColumn = trackCtr;
                //int TotalColumns = data.indexToObj - data.indexFromObj + 1;
                int TotalColumns = AllDomains[myalldomains].QuestionsSorted_ReverseValue.Count;
               // MessageBox.Show(TotalColumns.ToString());
                //MessageBox.Show(TotalColumns.ToString());

                int TotalRows = 1;

                //(Inserting Name)
                worksheet.Cells.Merge(firstRow, firstColumn, TotalRows, TotalColumns);

                Cell mergedCell = worksheet.Cells[firstRow, firstColumn];
                mergedCell.PutValue(AllDomains[myalldomains].DomainName);
                Aspose.Cells.Style style = mergedCell.GetStyle();

                // Set the text alignment to center
                style.HorizontalAlignment = TextAlignmentType.Center;
                style.VerticalAlignment = TextAlignmentType.Center;



                // Set the background color of the cell
                style.ForegroundColor = MyRandomColor;
                style.Pattern = BackgroundType.Solid;

                mergedCell.SetStyle(style);


                //(Inserting Number of Questions)
                int j = 1;
                for (int i = firstColumn; i <= firstColumn + TotalColumns; i++)

                {
                    
                    if (j <= TotalColumns)
                    {
                        
                        worksheet.Cells[1, i].Value = j;
                        j++;

                        Aspose.Cells.Style SecondRowStyle = worksheet.Cells[1, i].GetStyle();
                        SecondRowStyle.ForegroundColor = MyRandomColor;
                        SecondRowStyle.Pattern = BackgroundType.Solid;

                        worksheet.Cells[1, i].SetStyle(style);
                    }
                    
                }
                List<string> MyQuestions = new List<string>();
                List<object> columnData = new List<object>();
                //List<int> MyQuestionsIndex = new List<int>();
                //Inserting Data
                
                    
                    foreach(string MyQuestion in AllDomains[myalldomains].QuestionsSorted_ReverseValue.Keys)
                    {
                        foreach (DataGridViewColumn column in Dgv.Columns)
                        {
                            string header = column.HeaderText.Trim();
                            if (header == MyQuestion)
                            {

                                MyQuestions.Add(MyQuestion);

                            }
                        }
                        
                    }
                
                for (int i = 0; i < MyQuestions.Count; i++)
                {
                    string question = MyQuestions[i];
                    if (AllDomains[myalldomains].QuestionsSorted_ReverseValue.ContainsKey(question))
                    {
                        int value = AllDomains[myalldomains].QuestionsSorted_ReverseValue[question];
                        if (value == 1)
                        {
                            MyQuestions[i] += "R";
                            //MessageBox.Show(MyQuestions[i]);
                        }
                    }
                }

                int Questionctr = 0;
                
                int columnIndex = trackCtr;
                int rowCount = Dgv.RowCount;
                foreach (string Question in MyQuestions)
                {
                    //MessageBox.Show(Question);
                    if (!Question.EndsWith("R"))
                    {
                        foreach (DataGridViewRow row in Dgv.Rows)
                        {
                            int columnIndexctr = Dgv.Columns[Question].Index;
                            object cellValue = row.Cells[columnIndexctr].Value;

                            columnData.Add(cellValue);
                        }
                    }
                    else if (Question.EndsWith("R"))
                    {
                        string MyQuestionRemovedR = Question.Remove(Question.Length - 1, 1);
                        //int rangeSize = AllDomains[myalldomains].RangeFrom - AllDomains[myalldomains].RangeTo;
                        foreach (DataGridViewRow row in Dgv.Rows)
                        {
                            int columnIndexctr = Dgv.Columns[MyQuestionRemovedR].Index;
                            object cellValue = row.Cells[columnIndexctr].Value;
                            int reversedValue = AllDomains[myalldomains].RangeTo - (Convert.ToInt32(cellValue) - AllDomains[myalldomains].RangeFrom);
                            columnData.Add(reversedValue);
                            //MessageBox.Show(reversedValue.ToString());
                        }
                    }
                        if (Questionctr < MyQuestions.Count)
                        {
                            if (columnIndex <= trackCtr + TotalColumns)
                            {
                                int DataListCount = 0;


                                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                                {

                                    if (DataListCount < columnData.Count)
                                    {
                                        string stringValue = columnData[DataListCount].ToString();
                                        int intValue = Convert.ToInt32(stringValue);
                                        //MessageBox.Show(rowIndex.ToString());

                                        worksheet.Cells[rowIndex, columnIndex].PutValue(intValue);
                                        DataListCount++;
                                        Aspose.Cells.Style Datastyle = worksheet.Cells[rowIndex, columnIndex].GetStyle();
                                        Datastyle.ForegroundColor = MyRandomColor;
                                        Datastyle.Pattern = BackgroundType.Solid;
                                        Datastyle.HorizontalAlignment = TextAlignmentType.Center;
                                        Datastyle.VerticalAlignment = TextAlignmentType.Center;
                                        worksheet.Cells[rowIndex, columnIndex].SetStyle(Datastyle);
                                    }

                                }
                                columnIndex++;
                            }
                            Questionctr++;
                            columnData.Clear();
                        }
                    
                    
                    
                    
                }

                
                trackCtr = trackCtr + TotalColumns;
                //Inserting Titles
                int titlesctr = 0;
                
                for (int i = trackCtr; i < trackCtr + listTitles.Count; i++)
                {

                    if (titlesctr <= 3)
                    {
                        worksheet.Cells.Merge(firstRow, trackCtr, TotalRows, listTitles.Count);

                        Cell newmergedCell = worksheet.Cells[firstRow, trackCtr];
                        newmergedCell.PutValue(AllDomains[myalldomains].DomainName);
                        Aspose.Cells.Style style1 = newmergedCell.GetStyle();

                        // Set the text alignment to center
                        style1.HorizontalAlignment = TextAlignmentType.Center;
                        style1.VerticalAlignment = TextAlignmentType.Center;
                        



                        // Set the background color of the cell
                        style1.ForegroundColor = MyRandomColor;
                        style1.Pattern = BackgroundType.Solid;
                        style1.Font.IsBold = true;

                        newmergedCell.SetStyle(style);
                        
                        


                        if (titlesctr == 0)
                        {

                            int destinationColumn = trackCtr;
                            
                            int startColumn = trackCtr - AllDomains[myalldomains].QuestionsSorted_ReverseValue.Count;


                            for (int QuestionRow = 2; QuestionRow <= Dgv.Rows.Count; QuestionRow++)
                            {
                                string destinationCell = worksheet.Cells[QuestionRow, destinationColumn].Name;
                                string range = worksheet.Cells[QuestionRow, startColumn].Name + ":" + worksheet.Cells[QuestionRow, destinationColumn-1].Name;

                                string formula = "=SUM(" + range + ")";
                                worksheet.Cells[destinationCell].Formula = formula;
                                worksheet.Cells[destinationCell].SetStyle(style1);

                            }
                            
                            worksheet.Cells[1, destinationColumn].Value = listTitles[titlesctr];
                            worksheet.Cells[1, destinationColumn].SetStyle(style1);
                            titlesctr++;

                        }
                        
                        
                        if (titlesctr == 1)
                        {

                            int destinationColumn = trackCtr+1;

                            int startColumn = trackCtr - AllDomains[myalldomains].QuestionsSorted_ReverseValue.Count;


                            for (int QuestionRow = 2; QuestionRow <= Dgv.Rows.Count; QuestionRow++)
                            {
                                string destinationCell = worksheet.Cells[QuestionRow, destinationColumn].Name;
                                string range = worksheet.Cells[QuestionRow, startColumn].Name + ":" + worksheet.Cells[QuestionRow, destinationColumn - 2].Name;

                                string formula = "=AVERAGE(" + range + ")";
                                worksheet.Cells[destinationCell].Formula = formula;
                                worksheet.Cells[destinationCell].SetStyle(style1);

                            }
                           
                            worksheet.Cells[1, destinationColumn].Value = listTitles[titlesctr];
                            worksheet.Cells[1, destinationColumn].SetStyle(style1);
                            titlesctr++;
                        }
                        
                        if (titlesctr == 2)
                        {
                            int RangeFrom = AllDomains[myalldomains].RangeFrom;
                            int RangeTo = AllDomains[myalldomains].RangeTo;
                            int destinationColumn = trackCtr + 2;
                            for (int QuestionRow = 2; QuestionRow <= Dgv.Rows.Count; QuestionRow++)
                            {
                                string destinationCell = worksheet.Cells[QuestionRow, destinationColumn].Name;
                               
                                string AvgCell = worksheet.Cells[QuestionRow, destinationColumn-1].Name;
                                if(RangeFrom == 0)
                                {
                                    string formula = "=(" + AvgCell + ")/"+ RangeTo+ "*100";
                                    worksheet.Cells[destinationCell].Formula = formula;
                                    worksheet.Cells[destinationCell].SetStyle(style1);
                                }
                                else if(RangeFrom == 1)
                                {
                                    int NewRangeTo = RangeTo - 1;
                                    string formula = "=(" + AvgCell + "-1)" + "/"+ NewRangeTo + "*100";
                                    worksheet.Cells[destinationCell].Formula = formula;
                                    worksheet.Cells[destinationCell].SetStyle(style1);
                                }

                            }
                            //Cell PercentCell =  worksheet.Cells[1, destinationColumn];

                            //PercentCell.Value = "Percent";
                            //PercentCell.SetStyle(SecondRowStyle);
                            
                            worksheet.Cells[1, destinationColumn].Value = listTitles[titlesctr];
                            worksheet.Cells[1, destinationColumn].SetStyle(style1);
                            titlesctr++;
                        }
                        Aspose.Cells.Worksheet Tableworksheet = workbook.Worksheets[2];
                        if (titlesctr == 3)
                        {
                            int destinationColumn = trackCtr + 3;
                            // insert name in sheet 3
                            Tableworksheet.Cells.Merge(TableVlookctr, 0, 1, 3);
                            Cell TableMergedCell = Tableworksheet.Cells[TableVlookctr, 0];
                            TableMergedCell.PutValue(AllDomains[myalldomains].DomainName);
                            // Set the text alignment to center
                            style.HorizontalAlignment = TextAlignmentType.Center;
                            style.VerticalAlignment = TextAlignmentType.Center;
                            // Set the background color of the cell
                            style.ForegroundColor = MyRandomColor;
                            style.Pattern = BackgroundType.Solid;
                            TableMergedCell.SetStyle(style);
                            TableVlookctr++;

                            foreach (string item in AllDomains[myalldomains].Level_Comparison_Choice)
                            {
                                double greaterThan = 0.00;
                                if (item.Contains("Greater Equal") || item.Contains("Smaller Equal"))
                                {
                                    //string[] parts1 = item.Split(' ');
                                      // Extract the relevant parts from the first input string
                                    //string comparison1 = parts1[0];  // "Greater"
                                   
                                    //int value1 = int.Parse(parts1[3].Split('-')[0]);  // 20
                                    //int subValue1 = int.Parse(parts1[3].Split('-')[1]);  // 1

                                    if(item == "Greater Equal")
                                    {
                                        greaterThan = AllDomains[myalldomains].Level_From_To[TrackListFromToctr];

                                        Tableworksheet.Cells[TableVlookctr, 0].Value = "≥"+ AllDomains[myalldomains].Level_From_To[TrackListFromToctr];
                                        Tableworksheet.Cells[TableVlookctr, 0].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 1].Value = greaterThan;
                                        Tableworksheet.Cells[TableVlookctr, 1].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 2].Value = AllDomains[myalldomains].Level_Value[TrackListValuectr];
                                        Tableworksheet.Cells[TableVlookctr, 2].SetStyle(style);

                                        TrackListFromToctr++;
                                        TrackListValuectr++;
                                        TableVlookctr++;
                                    }
                                    else if(item == "Smaller Equal")
                                    {
                                        Tableworksheet.Cells[TableVlookctr, 0].Value = "≤" + AllDomains[myalldomains].Level_From_To[TrackListFromToctr];
                                        Tableworksheet.Cells[TableVlookctr, 0].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 1].Value = 0;
                                        Tableworksheet.Cells[TableVlookctr, 1].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 2].Value = AllDomains[myalldomains].Level_Value[TrackListValuectr];
                                        Tableworksheet.Cells[TableVlookctr, 2].SetStyle(style);

                                        TrackListFromToctr++;
                                        TrackListValuectr++;
                                        TableVlookctr++;
                                    }
                                    


                                    //NewDomainClassObj.Level_Comparison_Choose = comparison1 + "Equal";
                                    //NewDomainClassObj.Level_RangeFrom = txt_LevelRangeFrom.Text;
                                    //NewDomainClassObj.Level_Comparison_Or_RangeTo = txt_LevelOrRangeTo.Text;
                                    //NewDomainClassObj.Level_Comparison_Value = txt_LevelValue.Text;


                                }
                                else if(item.Contains("Range"))
                                {
                                    //string[] parts2 = item.Split(' ');
                                    ///string range2 = parts2[0];  // "Range"
                                    //int minValue2 = int.Parse(parts2[1]);  // 20
                                    //int maxValue2 = int.Parse(parts2[2].Split('-')[0]);  // 40
                                    //int subValue2 = int.Parse(parts2[2].Split('-')[1]);  // 2
                                    Tableworksheet.Cells[TableVlookctr, 0].Value = AllDomains[myalldomains].Level_From_To[TrackListFromToctr] + " - "+ AllDomains[myalldomains].Level_From_To[TrackListFromToctr+1];
                                    Tableworksheet.Cells[TableVlookctr, 0].SetStyle(style);

                                    Tableworksheet.Cells[TableVlookctr, 1].Value = AllDomains[myalldomains].Level_From_To[TrackListFromToctr];
                                    Tableworksheet.Cells[TableVlookctr, 1].SetStyle(style);

                                    Tableworksheet.Cells[TableVlookctr, 2].Value = AllDomains[myalldomains].Level_Value[TrackListValuectr];
                                    Tableworksheet.Cells[TableVlookctr, 2].SetStyle(style);

                                    TrackListFromToctr= TrackListFromToctr + 2;
                                    TrackListValuectr++;
                                    TableVlookctr++;

                                }
                                else if(item.Contains("Greater") || item.Contains("Smaller"))
                                {
                                    
                                    if (item == "Greater")
                                    {
                                        greaterThan = AllDomains[myalldomains].Level_From_To[TrackListFromToctr] + 0.0009;

                                        Tableworksheet.Cells[TableVlookctr, 0].Value = ">" + AllDomains[myalldomains].Level_From_To[TrackListFromToctr];
                                        Tableworksheet.Cells[TableVlookctr, 0].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 1].Value = greaterThan;
                                        Tableworksheet.Cells[TableVlookctr, 1].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 2].Value = AllDomains[myalldomains].Level_Value[TrackListValuectr];
                                        Tableworksheet.Cells[TableVlookctr, 2].SetStyle(style);

                                        TrackListFromToctr++;
                                        TrackListValuectr++;
                                        TableVlookctr++;
                                    }
                                    else if (item == "Smaller")
                                    {
                                        Tableworksheet.Cells[TableVlookctr, 0].Value = "<" + AllDomains[myalldomains].Level_From_To[TrackListFromToctr];
                                        Tableworksheet.Cells[TableVlookctr, 0].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 1].Value = 0;
                                        Tableworksheet.Cells[TableVlookctr, 1].SetStyle(style);

                                        Tableworksheet.Cells[TableVlookctr, 2].Value = AllDomains[myalldomains].Level_Value[TrackListValuectr];
                                        Tableworksheet.Cells[TableVlookctr, 2].SetStyle(style);

                                        TrackListFromToctr++;
                                        TrackListValuectr++;
                                        TableVlookctr++;
                                    }

                                }
                                else if (item.Contains("Equal To"))
                                {
                                        //string[] parts1 = item.Split(' ');
                                        // Extract the relevant parts from the first input string
                                        //string comparison1 = parts1[0];  // "Greater"
                                        //int value1 = int.Parse(parts1[2].Split('-')[0]);  // 20
                                        //int subValue1 = int.Parse(parts1[2].Split('-')[1]);  // 1
                    
                                }

                            }

                            worksheet.Cells[1, destinationColumn].Value = listTitles[titlesctr];
                            worksheet.Cells[1, destinationColumn].SetStyle(style1);

                            if (AllDomains[myalldomains].Level_Percent_Or_TotalScore == "Percent Score")
                            {
                                int LevelQuestionlength = AllDomains[myalldomains].Level_Value.Count;
                                int PercentScoreColumn =  trackCtr + 2;
                                string Datarange = worksheet.Cells[2, PercentScoreColumn].Name + ":" + worksheet.Cells[Dgv.RowCount, PercentScoreColumn].Name;

                                string rangeStart = Tableworksheet.Cells[TableVlookctr - LevelQuestionlength, 1].Name;
                                string rangeEnd = Tableworksheet.Cells[TableVlookctr - 1, 2].Name;

                                string rangeStartModified = "$" + rangeStart.Insert(1, "$");
                                string rangeEndModified = "$" + rangeEnd.Insert(1, "$");

                                string TableArrayRange = rangeStartModified + ":" + rangeEndModified;
                                //MessageBox.Show(rangeStartModified);
                                //Aspose.Cells.Range cellRange = Tableworksheet.Cells.CreateRange(TableArrayRange);

                                // Iterate over the cells and retrieve the data
                                //foreach (Cell cell in cellRange)
                                //{
                                //MessageBox.Show(cell.Name + " " + cell.Value);
                                //}
                                for (int QuestionRow = 2; QuestionRow <= Dgv.Rows.Count; QuestionRow++)
                                {
                                    string destinationCell = worksheet.Cells[QuestionRow, trackCtr+3].Name;

                                    
                                    string formula = "=VLOOKUP(" + Datarange + "," + "Sheet3!"+ TableArrayRange  + ",2"+")";

                                    //MessageBox.Show(destinationCell);
                                    
                                    worksheet.Cells[destinationCell].Formula = formula;
                                    worksheet.Cells[destinationCell].SetStyle(style1);



                                }
                            }
                            else if (AllDomains[myalldomains].Level_Percent_Or_TotalScore == "Total Score")
                            {
                                int LevelQuestionlength = AllDomains[myalldomains].Level_Value.Count;
                                int TotalScoreColumn = trackCtr;
                                string Datarange = worksheet.Cells[2, TotalScoreColumn].Name + ":" + worksheet.Cells[Dgv.RowCount, TotalScoreColumn].Name;

                                string rangeStart = Tableworksheet.Cells[TableVlookctr - LevelQuestionlength, 1].Name;
                                string rangeEnd = Tableworksheet.Cells[TableVlookctr - 1, 2].Name;

                                string rangeStartModified = "$" + rangeStart.Insert(1, "$");
                                string rangeEndModified = "$" + rangeEnd.Insert(1, "$");

                                string TableArrayRange = rangeStartModified + ":" + rangeEndModified;
                                //MessageBox.Show(rangeStartModified);
                                //Aspose.Cells.Range cellRange = Tableworksheet.Cells.CreateRange(TableArrayRange);

                                // Iterate over the cells and retrieve the data
                                //foreach (Cell cell in cellRange)
                                //{
                                //MessageBox.Show(cell.Name + " " + cell.Value);
                                //}
                                for (int QuestionRow = 2; QuestionRow <= Dgv.Rows.Count; QuestionRow++)
                                {
                                    string destinationCell = worksheet.Cells[QuestionRow, trackCtr + 3].Name;


                                    string formula = "=VLOOKUP(" + Datarange + "," + "Sheet3!" + TableArrayRange + ",2" + ")";

                                    //MessageBox.Show(destinationCell);

                                    worksheet.Cells[destinationCell].Formula = formula;
                                    worksheet.Cells[destinationCell].SetStyle(style1);

                                }
                            }

                            titlesctr++;
                        }



                    }


                }
                
                trackCtr = trackCtr + listTitles.Count+1;
                worksheet.Cells.InsertColumn(trackCtr, true);


                MyQuestions.Clear();
            }
            worksheet.AutoFitColumns();
            workbook.Save(ExcelFunctions.filepath);
            MessageBox.Show("Calculations Done!");
        }

        private void btn_Done_Click(object sender, EventArgs e)
        {
            try
            {
                AddData();
            }
            catch(Exception)
            {
                MessageBox.Show("Error!");
            }
            
            //TestClearExcel();
            
        }
        private void CombineDomains_ListModified(object sender, List<NewDomainClass> modifiedList)
        {
            // Update the list with the modified list
            AllDomains = modifiedList;

            // Do any other necessary operations with the modified list
            // ...
        }
        private void btn_View_Click(object sender, EventArgs e)
        {
            var childforms = Application.OpenForms.OfType<CombineDomains>().ToList();
            if (childforms.Count() == 1)
            {
                childforms.FirstOrDefault().Close();
            }
            else
            {
                CombineDomains combineDomains = new CombineDomains();
                combineDomains.AllDomainsCombined = AllDomains;
                combineDomains.Show();
            }


        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Word w = new Word();
            w.AllDomainsWord = AllDomains;
            w.ExcelWorksheet_ToWord = workbook;
            w.Show();
        }

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
            ChooseFrm chooseFrm = new ChooseFrm();
            chooseFrm.Dgv = Dgv;
            chooseFrm.Show();
        }

        private void panelmove_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
