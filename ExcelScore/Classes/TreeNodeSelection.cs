using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Classes
{
    public class TreeNodeSelection
    {
        public string ParentText { get; set; }
        public List<string> ChildTexts { get; set; } = null; // null or empty → check all



        public void ApplySelections(Dictionary<TreeView, TreeNodeSelection> selections)
        {
            foreach (var pair in selections)
            {
                TreeView treeView = pair.Key;
                TreeNodeSelection selection = pair.Value;

                TreeNode parentNode = FindNodeByText(treeView.Nodes, selection.ParentText);
                if (parentNode == null) continue;

                if (selection.ChildTexts == null || selection.ChildTexts.Count == 0)
                {
                    // Check parent only, your code will handle child checking
                    parentNode.Checked = true;
                }
                else
                {
                    // Check only specific children recursively
                    foreach (string childText in selection.ChildTexts)
                    {
                        TreeNode target = FindNodeByText(parentNode.Nodes, childText);
                        if (target != null)
                            target.Checked = true;
                    }
                }
            }
        }

        private TreeNode FindNodeByText(TreeNodeCollection nodes, string targetText)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Name == targetText)
                    return node;

                TreeNode found = FindNodeByText(node.Nodes, targetText);
                if (found != null)
                    return found;
            }

            return null;
        }


    }

}
