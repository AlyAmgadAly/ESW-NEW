using ExcelScore.StatClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.App_UI
{
    public class Custom_UI_Functions
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);
        public static void Make_Panel_Draggable(Control control, Form form)
        {
            control.MouseDown += (sender, e) =>
            {
                ReleaseCapture();
                SendMessage(form.Handle, 0x112, 0xf012, 0);
            };
        }


        /// <summary>
        /// Populates a DataGridView with SPSS parameters.
        /// </summary>
        public static void PopulateWithSpssParams(DataGridView dgv, List<StatParameter> spssParams)
        {
            if (dgv == null || spssParams == null) return;

            // Clear and load parameters
            dgv.Rows.Clear();
            foreach (var param in spssParams)
            {
                dgv.Rows.Add(param.Name);
            }

            // Disable sorting
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            // Enable mouse wheel scrolling and hide scrollbars
            dgv.MouseWheel += (s, e) => HandleMouseWheel(dgv, e);

            // Column resize rules
            dgv.AllowUserToResizeColumns = false;
            if (dgv.Columns.Count > 0)
                dgv.Columns[0].Resizable = DataGridViewTriState.True;

            // Insert parameter type images
            //InsertPictureTypes(dgv, spssParams);
        }

        /// <summary>
        /// Mouse wheel scrolling without showing scrollbars.
        /// </summary>
        private static void HandleMouseWheel(DataGridView dgv, MouseEventArgs e)
        {
            if (dgv.RowCount == 0) return;

            int currentIndex = dgv.FirstDisplayedScrollingRowIndex;
            int scrollLines = SystemInformation.MouseWheelScrollLines;

            if (e.Delta > 0) // Scroll up
            {
                int newIndex = Math.Max(0, currentIndex - scrollLines);
                dgv.FirstDisplayedScrollingRowIndex = newIndex;
            }
            else if (e.Delta < 0) // Scroll down
            {
                int newIndex = Math.Min(dgv.RowCount - 1, currentIndex + scrollLines);
                dgv.FirstDisplayedScrollingRowIndex = newIndex;
            }
        }

        /// <summary>
        /// Inserts images into the DataGridView based on parameter type.
        /// </summary>
        public static void InsertPictureTypes(DataGridView dgv, List<StatParameter> spssParams)
        {
            if (dgv == null || spssParams == null) return;

            // Map row names for fast lookup
            var paraNameToRowMap = new Dictionary<string, DataGridViewRow>();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    string paraName = row.Cells[0].Value.ToString();
                    if (!paraNameToRowMap.ContainsKey(paraName))
                    {
                        paraNameToRowMap[paraName] = row;
                    }
                }
            }

            // Apply type icons
            foreach (var spssparam in spssParams)
            {
                if (paraNameToRowMap.TryGetValue(spssparam.Name, out var targetRow))
                {
                    if (spssparam.Type == "Nominal")
                        targetRow.Cells["ColMeasure"].Value = Resource.Final_Nominal_Color;
                    else if (spssparam.Type == "Scale")
                        targetRow.Cells["ColMeasure"].Value = Resource.FinalScale_Color;
                }
            }
        }

        /// <summary>
        /// Adds filtering behavior to a TextBox for the given DataGridView.
        /// </summary>
        public static void AttachFilter(TextBox filterBox, DataGridView dgv)
        {
            if (filterBox == null || dgv == null) return;

            filterBox.TextChanged += (s, e) =>
            {
                string filterText = filterBox.Text.ToLower();

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;

                    string cellValue = row.Cells[0].Value?.ToString().ToLower() ?? "";
                    row.Visible = cellValue.Contains(filterText);
                }
            };
        }


        //Tree View

        public static void AddNodesToDictionary(TreeNodeCollection nodes, Dictionary<string, bool> dict)
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
        public static Dictionary<string, bool> BuildNodeCheckedDictionary(System.Windows.Forms.TreeView treeView)
        {
            Dictionary<string, bool> result = new Dictionary<string, bool>();
            AddNodesToDictionary(treeView.Nodes, result);
            return result;
        }

        public static void ClearListBox(ListBox listBox) 
        {
            var selectedItemsNominal = new List<object>();
            foreach (var selectedItemNominal in listBox.Items)
            {
                selectedItemsNominal.Add(selectedItemNominal);
            }


            foreach (var selectedItemNominal in selectedItemsNominal)
            {
                listBox.Items.Remove(selectedItemNominal);
            }


        }

        private void TextBox_Enter(object sender, EventArgs e)
        {
            if (sender is TextBox tb)
            {
                tb.SelectAll();
            }
        }
    }
}
