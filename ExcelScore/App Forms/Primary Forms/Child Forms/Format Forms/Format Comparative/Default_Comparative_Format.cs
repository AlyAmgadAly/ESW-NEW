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
using ExcelScore.Classes;

namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms.Format_Forms.Format_Comparative
{
    public partial class Default_Comparative_Format : Form
    {
        public Default_Comparative_Format()
        {
            InitializeComponent();
        }

        private void btn_Done_Click(object sender, EventArgs e)
        {
            Dictionary<string, bool> nodeCheckedStatusPrimary = Custom_UI_Functions.BuildNodeCheckedDictionary(Primary_TV);
            Dictionary<string, bool> nodeCheckedStatusExtra = Custom_UI_Functions.BuildNodeCheckedDictionary(Extra_TV);


            
            string TableDesignType = "Default";
            string TableType = "comparative";

            FormDataTransfer.Set("nodeCheckedStatusPrimary", nodeCheckedStatusPrimary);
            FormDataTransfer.Set("nodeCheckedStatusExtra", nodeCheckedStatusExtra);
            
            FormDataTransfer.Set("TableDesignType", TableDesignType);
            FormDataTransfer.Set("TableType", TableType);


            MessageBox.Show("Done");
        }
    }
}
