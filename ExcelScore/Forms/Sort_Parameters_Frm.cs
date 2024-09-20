using ExcelScore.Classes;
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
    public partial class Sort_Parameters_Frm : Form
    {
        public Sort_Parameters_Frm()
        {
            InitializeComponent();
        }


        public ComparativeTable SortcomparativeTable { get; set; } 


        private void list_AllParameters_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void MoveItemUp()
        {
            if (list_AllParameters.SelectedIndex > 0)
            {
                int selectedIndex = list_AllParameters.SelectedIndex;
                object selectedItem = list_AllParameters.SelectedItem;

                list_AllParameters.Items.RemoveAt(selectedIndex);
                list_AllParameters.Items.Insert(selectedIndex - 1, selectedItem);
                list_AllParameters.SelectedIndex = selectedIndex - 1;
            }
        }

        private void MoveItemDown()
        {
            if (list_AllParameters.SelectedIndex < list_AllParameters.Items.Count - 1)
            {
                int selectedIndex = list_AllParameters.SelectedIndex;
                object selectedItem = list_AllParameters.SelectedItem;

                list_AllParameters.Items.RemoveAt(selectedIndex);
                list_AllParameters.Items.Insert(selectedIndex + 1, selectedItem);
                list_AllParameters.SelectedIndex = selectedIndex + 1;
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

        private void Sort_Parameters_Frm_Load(object sender, EventArgs e)
        {
            foreach (var parameter in SortcomparativeTable.Parameters)
            {
                list_AllParameters.Items.Add(parameter.Name);
            }
        }

        public string TableType { get; set; }
        private void pic_DoneSorting_Click(object sender, EventArgs e)
        {
            List<string> sortedParameterNames = list_AllParameters.Items.Cast<string>().ToList();
            SortcomparativeTable.Parameters = SortcomparativeTable.Parameters.OrderBy(p => sortedParameterNames.IndexOf(p.Name)).ToList();

            if(TableType == "Comparative")
            {
                ComparativeGroups.DoneSortedComparative = SortcomparativeTable;
                this.Close();
            }
            else if (TableType == "Descriptive")
            {
                DescriptiveFrm.DoneSortedComparative = SortcomparativeTable;
                this.Close();


            }
            else if(TableType == "Regression")
            {
                Regression_Frm.DoneSortedComparative = SortcomparativeTable;
                this.Close();
            }

        }
    }
}
