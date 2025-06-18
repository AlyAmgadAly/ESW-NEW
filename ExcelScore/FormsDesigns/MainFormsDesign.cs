using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;
using ExcelScore.Classes;
using ExcelScore.Forms;
using MathNet.Numerics.LinearAlgebra.Factorization;
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
using static ExcelScore.Classes.Tool;
using Font = System.Drawing.Font;

namespace ExcelScore.FormsDesigns
{
    public partial class MainFormsDesign : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        public MainFormsDesign()
        {
            InitializeComponent();
        }


        //Load the form and prevent reloading unneccessary
        private Form currentForm;

        private void ShowForm(Form form)
        {
            if (currentForm != null)
            {
                currentForm.Close();
            }

            panelContainer.Controls.Clear();
            currentForm = form;
            form.TopLevel = false;
            form.FormBorderStyle = FormBorderStyle.None;
            form.Dock = DockStyle.Fill;
            panelContainer.Controls.Add(form);
            form.Show();
        }

        private void MainFormsDesign_Load(object sender, EventArgs e)
        {

        }

      

        private void treeView1_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            
                Font nodeFont = e.Node.NodeFont ?? ((TreeView)sender).Font;
                System.Drawing.Color textColor = e.Node.ForeColor != System.Drawing.Color.Empty ? e.Node.ForeColor : TableTreeview.ForeColor;

                // Highlight selected node
                if ((e.State & TreeNodeStates.Selected) != 0)
                {
                    e.Graphics.FillRectangle(Brushes.LightSkyBlue, e.Bounds);
                    textColor = System.Drawing.Color.White;
                }
                else
                {
                    e.Graphics.FillRectangle(Brushes.White, e.Bounds);
                }

                TextRenderer.DrawText(e.Graphics, e.Node.Text, nodeFont, e.Bounds, textColor, TextFormatFlags.GlyphOverhangPadding);
            

        }

        private void TableTreeview_AfterSelect(object sender, TreeViewEventArgs e)
        {
            lbl_SelectedTable.Text = e.Node.Text;


            // Update label color
            lbl_SelectedTable.ForeColor = (lbl_SelectedTable.Text != "NA")
                ? System.Drawing.Color.Green
                : System.Drawing.Color.Black;


            panelContainer.Controls.Clear();

            if (e.Node.Text == "Default")
            {
                //ShowForm(new Comparative.Default());
                FormManager.ShowForm<Comparative.Default>(panelContainer);

            }
            else
            {
                FormManager.ShowForm<testdesign>(panelContainer);
            }



        }

        private void panelmove_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void panelmove_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void panelContainer_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pic_Close_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
