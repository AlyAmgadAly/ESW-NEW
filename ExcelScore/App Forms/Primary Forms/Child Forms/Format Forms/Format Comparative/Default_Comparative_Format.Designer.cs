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
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("First Column");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Last Column");
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("Pcontrol", new System.Windows.Forms.TreeNode[] {
            treeNode6,
            treeNode7});
            System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("Paper");
            System.Windows.Forms.TreeNode treeNode10 = new System.Windows.Forms.TreeNode("Group Title");
            System.Windows.Forms.TreeNode treeNode11 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode12 = new System.Windows.Forms.TreeNode("Min. - Max.");
            System.Windows.Forms.TreeNode treeNode13 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode14 = new System.Windows.Forms.TreeNode("Median (IQR)");
            System.Windows.Forms.TreeNode treeNode15 = new System.Windows.Forms.TreeNode("1st Option", new System.Windows.Forms.TreeNode[] {
            treeNode11,
            treeNode12,
            treeNode13,
            treeNode14});
            System.Windows.Forms.TreeNode treeNode16 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode17 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode18 = new System.Windows.Forms.TreeNode("Median (Min. - Max.)");
            System.Windows.Forms.TreeNode treeNode19 = new System.Windows.Forms.TreeNode("2nd Option", new System.Windows.Forms.TreeNode[] {
            treeNode16,
            treeNode17,
            treeNode18});
            System.Windows.Forms.TreeNode treeNode20 = new System.Windows.Forms.TreeNode("Normality Based");
            System.Windows.Forms.TreeNode treeNode21 = new System.Windows.Forms.TreeNode("Empty");
            System.Windows.Forms.TreeNode treeNode22 = new System.Windows.Forms.TreeNode("Parameter Options", new System.Windows.Forms.TreeNode[] {
            treeNode15,
            treeNode19,
            treeNode20,
            treeNode21});
            this.Extra_TV = new System.Windows.Forms.TreeView();
            this.Primary_TV = new System.Windows.Forms.TreeView();
            this.btn_Done = new System.Windows.Forms.Button();
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
            treeNode6.Name = "PcontrolFirst";
            treeNode6.Text = "First Column";
            treeNode7.Name = "PcontrolLast";
            treeNode7.Text = "Last Column";
            treeNode8.Name = "Pcontrol";
            treeNode8.Text = "Pcontrol";
            treeNode9.Name = "PaperFormat";
            treeNode9.Text = "Paper";
            treeNode10.Name = "GT";
            treeNode10.Text = "Group Title";
            this.Extra_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode5,
            treeNode8,
            treeNode9,
            treeNode10});
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
            treeNode11.Name = "NumberOfCasesFirst";
            treeNode11.Text = "N";
            treeNode12.Checked = true;
            treeNode12.Name = "MinMaxFirst";
            treeNode12.Text = "Min. - Max.";
            treeNode13.Checked = true;
            treeNode13.Name = "MeanSDFirst";
            treeNode13.Text = "Mean ± SD.";
            treeNode14.Checked = true;
            treeNode14.Name = "MedianIQRFirst";
            treeNode14.Text = "Median (IQR)";
            treeNode15.Name = "1stOption";
            treeNode15.Text = "1st Option";
            treeNode16.Name = "NumberOfCasesSecond";
            treeNode16.Text = "N";
            treeNode17.Name = "MeanSDSecond";
            treeNode17.Text = "Mean ± SD.";
            treeNode18.Name = "MedianMinMaxSecond";
            treeNode18.Text = "Median (Min. - Max.)";
            treeNode19.Name = "2ndOption";
            treeNode19.Text = "2nd Option";
            treeNode20.Name = "Normal_Abnormal";
            treeNode20.Text = "Normality Based";
            treeNode21.Name = "Empty";
            treeNode21.Text = "Empty";
            treeNode22.Name = "Parameter_Options";
            treeNode22.Text = "Parameter Options";
            this.Primary_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode22});
            this.Primary_TV.ShowLines = false;
            this.Primary_TV.Size = new System.Drawing.Size(283, 543);
            this.Primary_TV.TabIndex = 22;
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(173)))), ((int)(((byte)(181)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.Location = new System.Drawing.Point(435, 495);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(118, 40);
            this.btn_Done.TabIndex = 24;
            this.btn_Done.Text = "Done";
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // Default_Comparative_Format
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.ClientSize = new System.Drawing.Size(561, 543);
            this.Controls.Add(this.btn_Done);
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
        private System.Windows.Forms.Button btn_Done;
    }
}