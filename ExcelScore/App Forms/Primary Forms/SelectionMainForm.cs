using ExcelScore.App_Forms.Main_Menu_Forms;
using ExcelScore.App_Forms.Primary_Forms.Child_Forms;
using ExcelScore.App_UI;
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

namespace ExcelScore.App_Forms.Primary_Forms
{
    public partial class SelectionMainForm : Form
    {
        public SelectionMainForm()
        {
            InitializeComponent();
            Custom_UI_Functions.Make_Panel_Draggable(pnl_Top_bar, this);
        }

        private void btn_Comparative_Click(object sender, EventArgs e)
        {
            
        }

        private void SelectionMainForm_Load(object sender, EventArgs e)
        {

        }

        private void btn_Analytical_Click(object sender, EventArgs e)
        {
            Analytical_Child analytical_Child = new Analytical_Child();
            analytical_Child.Show();
            this.Hide();
        }
    }
}
