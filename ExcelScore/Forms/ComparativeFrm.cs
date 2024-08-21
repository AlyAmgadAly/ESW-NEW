using ExcelScore.Classes;
using ExcelScore.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Forms
{
    public partial class ComparativeFrm : Form
    {

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public DataGridView Dgv { get; set; }
        public ComparativeFrm()
        {
            InitializeComponent();
        }

        
        private void ComparativeFrm_Load(object sender, EventArgs e)
        {

            



        }

  

        private void pic_back_Click(object sender, EventArgs e)
        {
            this.Close();
            ChooseFrm chooseFrm = new ChooseFrm();
            chooseFrm.Dgv = Dgv;
            chooseFrm.Show();
        }

        private void pic_preTwoGroups_Click(object sender, EventArgs e)
        {
            var childforms = Application.OpenForms.OfType<PreivewImage>().ToList();
            if (childforms.Count() == 1)
            {
                childforms.FirstOrDefault().Close();
            }
            else
            {
                Image image = Resources.Compartive_2_groups;
                PreivewImage preivewImageobj = new PreivewImage();
                preivewImageobj.SelectedImage = image;

                preivewImageobj.Show();
            }




            
        }

        private void btn_twoGroups_Click(object sender, EventArgs e)
        {
            ComparativeGroups comparativeGroupsobj = new ComparativeGroups();
            comparativeGroupsobj.Dgv = Dgv; 
            comparativeGroupsobj.Show();
            this.Close();
        }

        private void panelmove_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
    }
}
