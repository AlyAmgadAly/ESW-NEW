using Accord.Statistics.Distributions.Univariate;
using Aspose.Cells;
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
    public partial class ParameterInfo : Form
    {
        public ParameterInfo()
        {
            InitializeComponent();
        }

        public Worksheet Sheet1 = ExcelFunctions.worksheet;

        public Worksheet Sheet2 = ExcelFunctions.Sheet2;

        public string ParameterName_ParaInfoFrm { get; set; }


        private void ParameterInfo_Load(object sender, EventArgs e)
        {
            
        }
    }
}
