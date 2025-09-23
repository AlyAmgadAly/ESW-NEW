using ExcelScore.App_UI;
using ExcelScore.Classes;
using ExcelScore.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelScore.App_Forms.Primary_Forms.Child_Forms.Format_Forms;

namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    public partial class ChooseAnalyticFormat_Child : Form
    {
        public ChooseAnalyticFormat_Child()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }

        private void pnl_LoadFormatForms_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tree_ChooseFormatType_AfterSelect(object sender, TreeViewEventArgs e)
        {
            lbl_FormatSelected.Text = e.Node.Text;


            pnl_LoadFormatForms.Controls.Clear();

            if (e.Node.Text == "Default")
            {
                FormManager.ShowForm<Format_Forms.Format_Comparative.Default_Comparative_Format>(pnl_LoadFormatForms);
            }
            else
            {
                FormManager.ShowForm<testdesign>(pnl_LoadFormatForms);
            }
        }
    }
}
