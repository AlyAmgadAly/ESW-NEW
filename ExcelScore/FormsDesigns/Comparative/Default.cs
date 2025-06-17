using ExcelScore.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelScore.FormsDesigns.Comparative
{
    public partial class Default : Form
    {
        public Default()
        {
            InitializeComponent();
        }

        
        private void Default_Load(object sender, EventArgs e)
        {
            
        }

        private void DefaultTreeview_AfterCheck(object sender, TreeViewEventArgs e)
        {
            // Prevent stack overflow due to recursive triggers
            Primary_TV.AfterCheck -= DefaultTreeview_AfterCheck;

            // Update all child nodes
            CheckAllChildNodes(e.Node, e.Node.Checked);

            Primary_TV.AfterCheck += DefaultTreeview_AfterCheck;
        }

        private void CheckAllChildNodes(TreeNode node, bool isChecked)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = isChecked;
                CheckAllChildNodes(child, isChecked); // recursively go deeper
            }
        }

        private void lbl_PrimarySettings_Click(object sender, EventArgs e)
        {

        }
    }
}
