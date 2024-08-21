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
    public partial class NominalYesOnly : Form
    {
        public NominalYesOnly()
        {
            InitializeComponent();
        }

        public List<string> AllNominal = new List<string>();


        public List<string> NominalYes = new List<string>();


        private void NominalYesOnly_Load(object sender, EventArgs e)
        {
            foreach (string nominal in AllNominal) 
            {
                list_AllNominal.Items.Add(nominal); 
            }
        }

        private void pic_AllNominalToNominalYesOnly_Click(object sender, EventArgs e)
        {
            foreach (object selectedItem in list_AllNominal.SelectedItems)
            {
                list_NominalYesOnly.Items.Add(selectedItem.ToString());
            }
        }

        private void pic_RemoveNominalList_Click(object sender, EventArgs e)
        {
            if (list_NominalYesOnly.SelectedIndex != -1)
            {

                var selectedItems = new List<object>();
                foreach (var selectedItem in list_NominalYesOnly.SelectedItems)
                {
                    selectedItems.Add(selectedItem);
                }


                foreach (var selectedItem in selectedItems)
                {
                    list_NominalYesOnly.Items.Remove(selectedItem);
                }

            }
            else
                MessageBox.Show("Please Select Item!");
        }

        private void btn_DoneNominalYes_Click(object sender, EventArgs e)
        {
            foreach (string item in list_NominalYesOnly.Items)
            {
                NominalYes.Add(item);
            }



            ComparativeGroups.NominalYes_Compara = NominalYes;
            this.Close();



        }
    }
}
