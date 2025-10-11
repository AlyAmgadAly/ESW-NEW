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
using ExcelScore.App_UI;

namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    public partial class SelectFrm_Child : Form
    {
        public SelectFrm_Child()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }
        public List<StatParameter> SpssParameters;
        
        private void SelectFrm_Child_Load(object sender, EventArgs e)
        {
            SpssParameters = FormDataTransfer.Get<List<StatParameter>>("SPSS_Parameters");
            PopDataGrid_SpssParams();
        }

        public void PopDataGrid_SpssParams()
        {
            Custom_UI_Functions.PopulateWithSpssParams(data_allPara, SpssParameters);
            Custom_UI_Functions.AttachFilter(txt_ParaName, data_allPara);
        }

        private void radiobtn_IF_CheckedChanged(object sender, EventArgs e)
        {
            if(radiobtn_IF.Checked)
            {
                radiobtn_Filtervar.Checked = false;
                pic_AllParaToIF.Enabled = true;
                pic_AllParatoFilter.Enabled = false;
            }
        }

        private void radiobtn_Filtervar_CheckedChanged(object sender, EventArgs e)
        {
            if(radiobtn_Filtervar.Checked)
            {
                radiobtn_IF.Checked = false;
                pic_AllParaToIF.Enabled = false;
                pic_AllParatoFilter.Enabled = true;
            }
        }

        private void pic_AllParaToIF_Click(object sender, EventArgs e)
        {
            if(radiobtn_IF.Checked)
            {
                App_UI.Custom_UI_Functions.AddTo_RichTextBox_From_DataGrid_Custom(txt_SelectStatement, data_allPara);
                Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
            }
          
        }

        private void btn_Equal_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " = ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_NotEqual_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " ~= ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_And_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " & ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_Or_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " | ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_lessthan_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " < ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_greaterthan_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " > ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_lessthanEqual_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " <= ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_greaterthanEqual_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " >= ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_plus_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text += " + ";
            Custom_UI_Functions.RichBoxHandleFocus(txt_SelectStatement);
        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            txt_SelectStatement.Text = "";
        }

        private void btn_Done_Click(object sender, EventArgs e)
        {
           
            if(radiobtn_IF.Checked)
            {
                string IfSelectCommandstr = txt_SelectStatement.Text;
                bool IfComm = true;
                FormDataTransfer.Set("IfSelectCommand", IfSelectCommandstr);
                FormDataTransfer.Set("SelectIfOrVar", IfComm);
                this.Hide();
            }
            else if (radiobtn_Filtervar.Checked)
            {
                string FilterVar = txt_filtervar.Text;
                bool IfComm = false;
                FormDataTransfer.Set("IfSelectCommand", FilterVar);
                FormDataTransfer.Set("SelectIfOrVar", IfComm);
                this.Hide();
            }
            else
            {
                this.Hide();
            }
            
        }

        private void pic_AllParatoFilter_Click(object sender, EventArgs e)
        {
            if (radiobtn_Filtervar.Checked)
            {
                txt_filtervar.Text = "";
                Custom_UI_Functions.AddTo_TextBox_From_DataGrid_Custom(txt_filtervar, data_allPara);
            }
        }
    }
}
