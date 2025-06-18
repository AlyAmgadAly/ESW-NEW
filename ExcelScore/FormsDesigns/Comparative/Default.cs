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
        Dictionary<string, bool> BuildNodeCheckedDictionary(System.Windows.Forms.TreeView treeView)
        {
            Dictionary<string, bool> result = new Dictionary<string, bool>();
            AddNodesToDictionary(treeView.Nodes, result);
            return result;
        }


        public string GetNodeTextByName(System.Windows.Forms.TreeView treeView, string nodeName)
        {
            return GetNodeTextByNameRecursive(treeView.Nodes, nodeName);
        }

        private string GetNodeTextByNameRecursive(TreeNodeCollection nodes, string nodeName)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Name == nodeName)
                    return node.Text;

                if (node.Nodes.Count > 0)
                {
                    string found = GetNodeTextByNameRecursive(node.Nodes, nodeName);
                    if (found != null)
                        return found;
                }
            }

            return null; // Not found
        }

        void AddNodesToDictionary(TreeNodeCollection nodes, Dictionary<string, bool> dict)
        {
            foreach (TreeNode node in nodes)
            {
                if (!dict.ContainsKey(node.Name))
                {
                    dict[node.Name] = node.Checked;
                }
                else
                {
                    // If exists, update with the current checked value
                    dict[node.Name] = node.Checked;
                }

                if (node.Nodes.Count > 0)
                {
                    AddNodesToDictionary(node.Nodes, dict);
                }
            }
        }

        private void pic_DoneSettings_Click(object sender, EventArgs e)
        {
            Dictionary<string, bool> nodeCheckedStatusPrimary = BuildNodeCheckedDictionary(Primary_TV);
            Dictionary<string, bool> nodeCheckedStatusExtra = BuildNodeCheckedDictionary(Extra_TV);

            string LeftMarginValue = GetNodeTextByName(Extra_TV, "LeftCellMarginValue");
            string RightMarginValue = GetNodeTextByName(Extra_TV, "RightCellMarginValue");
            string Type = "Default";

            FormDataTransfer.Set("nodeCheckedStatusPrimary", nodeCheckedStatusPrimary);
            FormDataTransfer.Set("nodeCheckedStatusExtra", nodeCheckedStatusExtra);
            FormDataTransfer.Set("LeftMarginValue", LeftMarginValue);
            FormDataTransfer.Set("RightMarginValue", RightMarginValue);
            FormDataTransfer.Set("Type", Type);
            

            MessageBox.Show("Done");
            //MessageBox.Show(LeftMarginValue);
            //MessageBox.Show(RightMarginValue);


        }
    }
}
