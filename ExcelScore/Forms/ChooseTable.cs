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
    public partial class ChooseTable : Form
    {
        public ChooseTable()
        {
            InitializeComponent();
        }
        public  Image Table_ChooseTable { get; set; } 
        private void pic_CloseChooseTable_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ChooseTable_Load(object sender, EventArgs e)
        {
            if (Table_ChooseTable != null)
            {
                pictureBox1.Image = Table_ChooseTable;
            }
        }
    }
}
