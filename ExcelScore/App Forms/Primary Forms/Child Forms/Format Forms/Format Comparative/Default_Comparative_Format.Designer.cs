namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms.Format_Forms.Format_Comparative
{
    partial class Default_Comparative_Format
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Total Column");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Row");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Column");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Total");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Percentages", new System.Windows.Forms.TreeNode[] {
            treeNode2,
            treeNode3,
            treeNode4});
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("0.19");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Left", new System.Windows.Forms.TreeNode[] {
            treeNode6});
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("0.19");
            System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("Right", new System.Windows.Forms.TreeNode[] {
            treeNode8});
            System.Windows.Forms.TreeNode treeNode10 = new System.Windows.Forms.TreeNode("Cell Margin", new System.Windows.Forms.TreeNode[] {
            treeNode7,
            treeNode9});
            System.Windows.Forms.TreeNode treeNode11 = new System.Windows.Forms.TreeNode("First Column");
            System.Windows.Forms.TreeNode treeNode12 = new System.Windows.Forms.TreeNode("Last Column");
            System.Windows.Forms.TreeNode treeNode13 = new System.Windows.Forms.TreeNode("Pcontrol", new System.Windows.Forms.TreeNode[] {
            treeNode11,
            treeNode12});
            System.Windows.Forms.TreeNode treeNode14 = new System.Windows.Forms.TreeNode("Paper");
            System.Windows.Forms.TreeNode treeNode15 = new System.Windows.Forms.TreeNode("Group Title");
            System.Windows.Forms.TreeNode treeNode16 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode17 = new System.Windows.Forms.TreeNode("Min. - Max.");
            System.Windows.Forms.TreeNode treeNode18 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode19 = new System.Windows.Forms.TreeNode("Median (IQR)");
            System.Windows.Forms.TreeNode treeNode20 = new System.Windows.Forms.TreeNode("1st Option", new System.Windows.Forms.TreeNode[] {
            treeNode16,
            treeNode17,
            treeNode18,
            treeNode19});
            System.Windows.Forms.TreeNode treeNode21 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode22 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode23 = new System.Windows.Forms.TreeNode("Median (Min. - Max.)");
            System.Windows.Forms.TreeNode treeNode24 = new System.Windows.Forms.TreeNode("2nd Option", new System.Windows.Forms.TreeNode[] {
            treeNode21,
            treeNode22,
            treeNode23});
            System.Windows.Forms.TreeNode treeNode25 = new System.Windows.Forms.TreeNode("Normality Based");
            System.Windows.Forms.TreeNode treeNode26 = new System.Windows.Forms.TreeNode("Empty");
            System.Windows.Forms.TreeNode treeNode27 = new System.Windows.Forms.TreeNode("Parameter Options", new System.Windows.Forms.TreeNode[] {
            treeNode20,
            treeNode24,
            treeNode25,
            treeNode26});
            this.Extra_TV = new System.Windows.Forms.TreeView();
            this.Primary_TV = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // Extra_TV
            // 
            this.Extra_TV.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.Extra_TV.CheckBoxes = true;
            this.Extra_TV.Dock = System.Windows.Forms.DockStyle.Right;
            this.Extra_TV.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Extra_TV.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.Extra_TV.FullRowSelect = true;
            this.Extra_TV.LabelEdit = true;
            this.Extra_TV.Location = new System.Drawing.Point(278, 0);
            this.Extra_TV.Name = "Extra_TV";
            treeNode1.Name = "TotalColumn";
            treeNode1.Text = "Total Column";
            treeNode2.Name = "RowPercentage";
            treeNode2.Text = "Row";
            treeNode3.Checked = true;
            treeNode3.Name = "ColumnPercentage";
            treeNode3.Text = "Column";
            treeNode4.Name = "TotalPercentage";
            treeNode4.Text = "Total";
            treeNode5.Name = "Percentages";
            treeNode5.Text = "Percentages";
            treeNode6.Name = "LeftCellMarginValue";
            treeNode6.Text = "0.19";
            treeNode7.Name = "Left";
            treeNode7.Text = "Left";
            treeNode8.Name = "RightCellMarginValue";
            treeNode8.Text = "0.19";
            treeNode9.Name = "Right";
            treeNode9.Text = "Right";
            treeNode10.Name = "CellMargin";
            treeNode10.Text = "Cell Margin";
            treeNode11.Name = "PcontrolFirst";
            treeNode11.Text = "First Column";
            treeNode12.Name = "PcontrolLast";
            treeNode12.Text = "Last Column";
            treeNode13.Name = "Pcontrol";
            treeNode13.Text = "Pcontrol";
            treeNode14.Name = "PaperFormat";
            treeNode14.Text = "Paper";
            treeNode15.Name = "GT";
            treeNode15.Text = "Group Title";
            this.Extra_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode5,
            treeNode10,
            treeNode13,
            treeNode14,
            treeNode15});
            this.Extra_TV.ShowLines = false;
            this.Extra_TV.Size = new System.Drawing.Size(283, 543);
            this.Extra_TV.TabIndex = 23;
            // 
            // Primary_TV
            // 
            this.Primary_TV.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.Primary_TV.CheckBoxes = true;
            this.Primary_TV.Dock = System.Windows.Forms.DockStyle.Left;
            this.Primary_TV.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Primary_TV.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.Primary_TV.FullRowSelect = true;
            this.Primary_TV.Location = new System.Drawing.Point(0, 0);
            this.Primary_TV.Name = "Primary_TV";
            treeNode16.Name = "NumberOfCasesFirst";
            treeNode16.Text = "N";
            treeNode17.Checked = true;
            treeNode17.Name = "MinMaxFirst";
            treeNode17.Text = "Min. - Max.";
            treeNode18.Checked = true;
            treeNode18.Name = "MeanSDFirst";
            treeNode18.Text = "Mean ± SD.";
            treeNode19.Checked = true;
            treeNode19.Name = "MedianIQRFirst";
            treeNode19.Text = "Median (IQR)";
            treeNode20.Name = "1stOption";
            treeNode20.Text = "1st Option";
            treeNode21.Name = "NumberOfCasesSecond";
            treeNode21.Text = "N";
            treeNode22.Name = "MeanSDSecond";
            treeNode22.Text = "Mean ± SD.";
            treeNode23.Name = "MedianMinMaxSecond";
            treeNode23.Text = "Median (Min. - Max.)";
            treeNode24.Name = "2ndOption";
            treeNode24.Text = "2nd Option";
            treeNode25.Name = "Normal_Abnormal";
            treeNode25.Text = "Normality Based";
            treeNode26.Name = "Empty";
            treeNode26.Text = "Empty";
            treeNode27.Name = "Parameter_Options";
            treeNode27.Text = "Parameter Options";
            this.Primary_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode27});
            this.Primary_TV.ShowLines = false;
            this.Primary_TV.Size = new System.Drawing.Size(283, 543);
            this.Primary_TV.TabIndex = 22;
            // 
            // Default_Comparative_Format
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.ClientSize = new System.Drawing.Size(561, 543);
            this.Controls.Add(this.Extra_TV);
            this.Controls.Add(this.Primary_TV);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Default_Comparative_Format";
            this.Text = "Default_Comparative_Format";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView Extra_TV;
        private System.Windows.Forms.TreeView Primary_TV;
    }
}