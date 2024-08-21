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
    public partial class Select_Groups : Form
    {
        public Select_Groups()
        {
            InitializeComponent();
        }


        public ComparativeTable SelectGroupcomparativeTable {  get; set; } 
        private void Select_Groups_Load(object sender, EventArgs e)
        {
            foreach (var parameter in SelectGroupcomparativeTable.Parameters)
            {
                if(parameter.IsGroup)
                {
                    lbl_parameterName.Text = parameter.Name;    
                    foreach (var value in parameter.ParameterValues.Distinct())
                    {
                        list_AllValues.Items.Add(value);
                    }
                }
            }
        }

        private void pic_AllNominalToSelected_Click(object sender, EventArgs e)
        {
            foreach (object selectedItem in list_AllValues.SelectedItems)
            {
                list_SelectedValues.Items.Add(selectedItem.ToString());
            }
        }

        private void pic_RemoveSelectedList_Click(object sender, EventArgs e)
        {
            if (list_SelectedValues.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_SelectedValues.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_SelectedValues.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void btn_DoneSelection_Click(object sender, EventArgs e)
        {
            SelectGroupcomparativeTable.HasSelect = true;
            List<int> selectedValues = new List<int>();
            foreach (var item in list_SelectedValues.Items)
            {
                if (int.TryParse(item.ToString(), out int intValue))
                {
                    selectedValues.Add(intValue);
                }
                   
                
                
            }
            SelectGroupcomparativeTable.SelectedValues = selectedValues;
            ComparativeGroups.DoneSelectionComparaTable = SelectGroupcomparativeTable;
            this.Close();




        }
    }
}
