using DocumentFormat.OpenXml.Wordprocessing;
using ExcelScore.Classes;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class NursingWordNew : Form
    {
        public NursingWordNew()
        {
            InitializeComponent();
        }
        WordDocument document;
        WordClass wordObj = new WordClass();

        public List<Tool> Alltools_Tables = new List<Tool>();

        private void NursingWordNew_Load(object sender, EventArgs e)
        {

        }

        public void Items(Tool CurrentTool , int ToolN)
        {
            int likertScaleCount = CurrentTool.LikertScale.Keys.Count; 

            if(CurrentTool.hasscales)
            {

                foreach (var CurrentScale in CurrentTool.Scales)
                {
                    IWSection section = wordObj.CreatePortraitSection();

                    int WordTableRows = 2 + CurrentScale.Items.Count;

                    int WordTableColumns = 2 + (CurrentTool.LikertScale.Keys.Count * 2);


                    IWTable table = wordObj.Createtable(section, WordTableRows, WordTableColumns);

                    wordObj.GeneralTableFormat(table);


                    wordObj.ApplyMerges(table, likertScaleCount);

                    wordObj.ApplyBorders(table, WordTableRows, WordTableColumns);

                    wordObj.SetWidths(table, WordTableRows, likertScaleCount);

                    wordObj.Add_Header_Items_Nusring(table, likertScaleCount, CurrentScale.Scale_Full_Name, CurrentTool);


                    wordObj.AddNo_perc_Center(table, WordTableColumns);

                    wordObj.AddQuestionNo(table, CurrentScale.Items.Count);

                    int Insertrow = 2;

                    foreach (var CurrentItem in CurrentScale.Items)
                    {
                        int InsertColumn = 2;

                        foreach (var CurrentLikertScore in CurrentTool.LikertScale.Keys)
                        {
                            int totalLikertFreq = 0;
                            foreach (var ParticpantResponse in CurrentItem.ParticipantResponses)
                            {
                                if(ParticpantResponse == CurrentLikertScore)
                                {
                                    totalLikertFreq++;
                                }
                            }
                            double totalLikertPerc = ((double)totalLikertFreq / ToolN)*100;


                            wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertFreq.ToString());
                            InsertColumn++;
                            wordObj.Addpara_CenterNoBOLD(table, Insertrow, InsertColumn, totalLikertPerc.ToString("0.0"));
                            InsertColumn++;
                        }

                        Insertrow++;
                    }




                    wordObj.FormatTableCustom(table, 12, 4, 2);


                }
            }
            else if(!(CurrentTool.hasscales))
            {

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            document = wordObj.InitWord();

            foreach (var CurrentTool in Alltools_Tables)
            {
                int toolN = CurrentTool.ToolItems[0].ParticipantResponses.Count;
                Items(CurrentTool , toolN);



            }

            string filepath = wordObj.SaveWord();

            wordObj.removeHeader(filepath);




        }
    }
}
