using ExcelScore.App_UI;
using ExcelScore.Classes;
using ExcelScore.Forms;
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
    public partial class SortFrm_Child : Form
    {
        public SortFrm_Child()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }
        StatTable mySortedtable;
        private void SortFrm_Child_Load(object sender, EventArgs e)
        {
            mySortedtable =  FormDataTransfer.Get<StatTable>("SortStatTable");

            foreach (var parameter in mySortedtable.Parameters)
            {
                list_AllParameters.Items.Add(parameter.Name);
            }
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

        private void pic_DoneSorting_Click(object sender, EventArgs e)
        {
            List<string> sortedParameterNames = list_AllParameters.Items.Cast<string>().ToList();
            mySortedtable.Parameters = mySortedtable.Parameters.OrderBy(p => sortedParameterNames.IndexOf(p.Name)).ToList();

            //ComparativeGroups.DoneSortedComparative = SortcomparativeTable;
            this.Close();
            
        }
    }
}
