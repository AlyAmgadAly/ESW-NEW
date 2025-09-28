using ExcelScore.App_UI;
using ExcelScore.Classes;
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
    public partial class EditCodesNominalFrm_Child : Form
    {
        public EditCodesNominalFrm_Child()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }
        StatParameter EditedCodeparameter;
        private void groupbx_ValueLabels_Enter(object sender, EventArgs e)
        {

        }
        public void LoadParameterCodes(StatParameter EditedParameter)
        {
            foreach (var ValueLabel in EditedParameter.ValueLabels)
            {
                string ValueLabelCombined = $"{ValueLabel.Key} = {ValueLabel.Value}";
                listbox_valueLabels.Items.Add(ValueLabelCombined);
            }
        }
        private void EditCodesNominalFrm_Child_Load(object sender, EventArgs e)
        {
            btn_Add.Enabled = false;
            btn_Change.Enabled = false;
            btn_Remove.Enabled = false;

            btn_Add.BackColor = Color.Gray;
            btn_Change.BackColor = Color.Gray;
            btn_Remove.BackColor = Color.Gray;

            EditedCodeparameter = FormDataTransfer.Get<StatParameter>("EditedCodeParameter");
            LoadParameterCodes(EditedCodeparameter);
        }

        public bool ValidateTextBoxes()
        {
            bool textboxesFilled = !string.IsNullOrWhiteSpace(txt_value.Text) && !string.IsNullOrWhiteSpace(txt_Label.Text);
            return textboxesFilled;
        }

        private void listbox_valueLabels_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_Label.Text = "";
            txt_value.Text = "";

            string CurrentItem = listbox_valueLabels.SelectedItem as string;
            string[] Kvp = CurrentItem.Split('=');


            string valuePart = Kvp[0].Trim(); // "1"
            string labelPart = Kvp[1].Trim(); // "GA"

            txt_value.Text = valuePart;
            txt_Label.Text = labelPart;


        }
        

        private void txt_value_TextChanged(object sender, EventArgs e)
        {
            bool textfilled = ValidateTextBoxes();
            if (EditedCodeparameter.ValueLabels.ContainsKey(int.Parse(txt_value.Text)) && EditedCodeparameter.ValueLabels.ContainsValue(txt_Label.Text))
            {
                btn_Remove.Enabled = true; 
            }
        }

        private void txt_Label_TextChanged(object sender, EventArgs e)
        {
            bool textfilled = ValidateTextBoxes();
        }
    }
}
