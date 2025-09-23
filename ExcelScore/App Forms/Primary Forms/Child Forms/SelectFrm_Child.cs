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
    }
}
