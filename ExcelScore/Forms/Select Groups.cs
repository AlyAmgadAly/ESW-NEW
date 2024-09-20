using Accord.Math;
using ExcelScore.Classes;
using Microsoft.SolverFoundation.Services;
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
        private ComparativeGroups _comparativeGroups;
        private Letters _letters;
        private Regression_Frm _regression_Frm;

        public Select_Groups(ComparativeGroups comparativeGroups)
        {
            InitializeComponent();
            _comparativeGroups = comparativeGroups;
        }

        public Select_Groups(Letters letters)
        {
            InitializeComponent();
            _letters = letters;
        }

        public Select_Groups(Regression_Frm regression_Frm)
        {
            InitializeComponent();
            _regression_Frm = regression_Frm;
        }

        public string SelectParaName_SelectGrFrm { get; set; }

        public List<double> SelectParaValues_SelectGrFrm { get; set; }
        private void Select_Groups_Load(object sender, EventArgs e)
        {
            lbl_parameterName.Text = SelectParaName_SelectGrFrm;
            foreach (double item in SelectParaValues_SelectGrFrm.Distinct())
            {
                list_AllValues.Items.Add(item); 
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
            
            List<double> selectedValues = new List<double>();
            foreach (var item in list_SelectedValues.Items)
            {
                if (double.TryParse(item.ToString(), out double doublevalue))
                {
                    selectedValues.Add(doublevalue);
                }
                
            }
            if(_comparativeGroups != null)
            {
                _comparativeGroups.SelectedParameterValues_CompaFrm = selectedValues;
            }
            if(_letters != null)
            {
                _letters.SelectedParameterValues_LetterFrm = selectedValues;
            }
            if(_regression_Frm != null)
            {
                _regression_Frm.SelectedParameterValues_CompaFrm = selectedValues;
            }

            this.Close();




        }
    }
}
