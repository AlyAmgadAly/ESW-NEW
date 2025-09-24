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
            System.Windows.Forms.TreeNode treeNode55 = new System.Windows.Forms.TreeNode("Total Column");
            System.Windows.Forms.TreeNode treeNode56 = new System.Windows.Forms.TreeNode("Row");
            System.Windows.Forms.TreeNode treeNode57 = new System.Windows.Forms.TreeNode("Column");
            System.Windows.Forms.TreeNode treeNode58 = new System.Windows.Forms.TreeNode("Total");
            System.Windows.Forms.TreeNode treeNode59 = new System.Windows.Forms.TreeNode("Percentages", new System.Windows.Forms.TreeNode[] {
            treeNode56,
            treeNode57,
            treeNode58});
            System.Windows.Forms.TreeNode treeNode60 = new System.Windows.Forms.TreeNode("0.19");
            System.Windows.Forms.TreeNode treeNode61 = new System.Windows.Forms.TreeNode("Left", new System.Windows.Forms.TreeNode[] {
            treeNode60});
            System.Windows.Forms.TreeNode treeNode62 = new System.Windows.Forms.TreeNode("0.19");
            System.Windows.Forms.TreeNode treeNode63 = new System.Windows.Forms.TreeNode("Right", new System.Windows.Forms.TreeNode[] {
            treeNode62});
            System.Windows.Forms.TreeNode treeNode64 = new System.Windows.Forms.TreeNode("Cell Margin", new System.Windows.Forms.TreeNode[] {
            treeNode61,
            treeNode63});
            System.Windows.Forms.TreeNode treeNode65 = new System.Windows.Forms.TreeNode("First Column");
            System.Windows.Forms.TreeNode treeNode66 = new System.Windows.Forms.TreeNode("Last Column");
            System.Windows.Forms.TreeNode treeNode67 = new System.Windows.Forms.TreeNode("Pcontrol", new System.Windows.Forms.TreeNode[] {
            treeNode65,
            treeNode66});
            System.Windows.Forms.TreeNode treeNode68 = new System.Windows.Forms.TreeNode("Paper");
            System.Windows.Forms.TreeNode treeNode69 = new System.Windows.Forms.TreeNode("Group Title");
            System.Windows.Forms.TreeNode treeNode70 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode71 = new System.Windows.Forms.TreeNode("Min. - Max.");
            System.Windows.Forms.TreeNode treeNode72 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode73 = new System.Windows.Forms.TreeNode("Median (IQR)");
            System.Windows.Forms.TreeNode treeNode74 = new System.Windows.Forms.TreeNode("1st Option", new System.Windows.Forms.TreeNode[] {
            treeNode70,
            treeNode71,
            treeNode72,
            treeNode73});
            System.Windows.Forms.TreeNode treeNode75 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode76 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode77 = new System.Windows.Forms.TreeNode("Median (Min. - Max.)");
            System.Windows.Forms.TreeNode treeNode78 = new System.Windows.Forms.TreeNode("2nd Option", new System.Windows.Forms.TreeNode[] {
            treeNode75,
            treeNode76,
            treeNode77});
            System.Windows.Forms.TreeNode treeNode79 = new System.Windows.Forms.TreeNode("Normality Based");
            System.Windows.Forms.TreeNode treeNode80 = new System.Windows.Forms.TreeNode("Empty");
            System.Windows.Forms.TreeNode treeNode81 = new System.Windows.Forms.TreeNode("Parameter Options", new System.Windows.Forms.TreeNode[] {
            treeNode74,
            treeNode78,
            treeNode79,
            treeNode80});
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
            treeNode55.Name = "TotalColumn";
            treeNode55.Text = "Total Column";
            treeNode56.Name = "RowPercentage";
            treeNode56.Text = "Row";
            treeNode57.Checked = true;
            treeNode57.Name = "ColumnPercentage";
            treeNode57.Text = "Column";
            treeNode58.Name = "TotalPercentage";
            treeNode58.Text = "Total";
            treeNode59.Name = "Percentages";
            treeNode59.Text = "Percentages";
            treeNode60.Name = "LeftCellMarginValue";
            treeNode60.Text = "0.19";
            treeNode61.Name = "Left";
            treeNode61.Text = "Left";
            treeNode62.Name = "RightCellMarginValue";
            treeNode62.Text = "0.19";
            treeNode63.Name = "Right";
            treeNode63.Text = "Right";
            treeNode64.Name = "CellMargin";
            treeNode64.Text = "Cell Margin";
            treeNode65.Name = "PcontrolFirst";
            treeNode65.Text = "First Column";
            treeNode66.Name = "PcontrolLast";
            treeNode66.Text = "Last Column";
            treeNode67.Name = "Pcontrol";
            treeNode67.Text = "Pcontrol";
            treeNode68.Name = "PaperFormat";
            treeNode68.Text = "Paper";
            treeNode69.Name = "GT";
            treeNode69.Text = "Group Title";
            this.Extra_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode55,
            treeNode59,
            treeNode64,
            treeNode67,
            treeNode68,
            treeNode69});
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
            treeNode70.Name = "NumberOfCasesFirst";
            treeNode70.Text = "N";
            treeNode71.Checked = true;
            treeNode71.Name = "MinMaxFirst";
            treeNode71.Text = "Min. - Max.";
            treeNode72.Checked = true;
            treeNode72.Name = "MeanSDFirst";
            treeNode72.Text = "Mean ± SD.";
            treeNode73.Checked = true;
            treeNode73.Name = "MedianIQRFirst";
            treeNode73.Text = "Median (IQR)";
            treeNode74.Name = "1stOption";
            treeNode74.Text = "1st Option";
            treeNode75.Name = "NumberOfCasesSecond";
            treeNode75.Text = "N";
            treeNode76.Name = "MeanSDSecond";
            treeNode76.Text = "Mean ± SD.";
            treeNode77.Name = "MedianMinMaxSecond";
            treeNode77.Text = "Median (Min. - Max.)";
            treeNode78.Name = "2ndOption";
            treeNode78.Text = "2nd Option";
            treeNode79.Name = "Normal_Abnormal";
            treeNode79.Text = "Normality Based";
            treeNode80.Name = "Empty";
            treeNode80.Text = "Empty";
            treeNode81.Name = "Parameter_Options";
            treeNode81.Text = "Parameter Options";
            this.Primary_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode81});
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