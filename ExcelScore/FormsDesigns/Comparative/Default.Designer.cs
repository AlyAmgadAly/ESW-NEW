namespace ExcelScore.FormsDesigns.Comparative
{
    partial class Default
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
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Min. - Max.");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Median (IQR)");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("1st Option", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4});
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("N");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Mean ± SD.");
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("Median (Min. - Max.)");
            System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("2nd Option", new System.Windows.Forms.TreeNode[] {
            treeNode6,
            treeNode7,
            treeNode8});
            System.Windows.Forms.TreeNode treeNode10 = new System.Windows.Forms.TreeNode("Parameter Options", new System.Windows.Forms.TreeNode[] {
            treeNode5,
            treeNode9});
            System.Windows.Forms.TreeNode treeNode11 = new System.Windows.Forms.TreeNode("Total Column");
            System.Windows.Forms.TreeNode treeNode12 = new System.Windows.Forms.TreeNode("Row");
            System.Windows.Forms.TreeNode treeNode13 = new System.Windows.Forms.TreeNode("Column");
            System.Windows.Forms.TreeNode treeNode14 = new System.Windows.Forms.TreeNode("Percentages", new System.Windows.Forms.TreeNode[] {
            treeNode12,
            treeNode13});
            System.Windows.Forms.TreeNode treeNode15 = new System.Windows.Forms.TreeNode("0.19");
            System.Windows.Forms.TreeNode treeNode16 = new System.Windows.Forms.TreeNode("Left", new System.Windows.Forms.TreeNode[] {
            treeNode15});
            System.Windows.Forms.TreeNode treeNode17 = new System.Windows.Forms.TreeNode("0.19");
            System.Windows.Forms.TreeNode treeNode18 = new System.Windows.Forms.TreeNode("Right", new System.Windows.Forms.TreeNode[] {
            treeNode17});
            System.Windows.Forms.TreeNode treeNode19 = new System.Windows.Forms.TreeNode("Cell Margin", new System.Windows.Forms.TreeNode[] {
            treeNode16,
            treeNode18});
            System.Windows.Forms.TreeNode treeNode20 = new System.Windows.Forms.TreeNode("First Column");
            System.Windows.Forms.TreeNode treeNode21 = new System.Windows.Forms.TreeNode("Last Column");
            System.Windows.Forms.TreeNode treeNode22 = new System.Windows.Forms.TreeNode("Pcontrol", new System.Windows.Forms.TreeNode[] {
            treeNode20,
            treeNode21});
            System.Windows.Forms.TreeNode treeNode23 = new System.Windows.Forms.TreeNode("Paper");
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Default));
            this.Primary_TV = new System.Windows.Forms.TreeView();
            this.lbl_PrimarySettings = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Extra_TV = new System.Windows.Forms.TreeView();
            this.pic_DoneSettings = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneSettings)).BeginInit();
            this.SuspendLayout();
            // 
            // Primary_TV
            // 
            this.Primary_TV.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.Primary_TV.CheckBoxes = true;
            this.Primary_TV.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Primary_TV.ForeColor = System.Drawing.SystemColors.WindowText;
            this.Primary_TV.FullRowSelect = true;
            this.Primary_TV.Location = new System.Drawing.Point(12, 38);
            this.Primary_TV.Name = "Primary_TV";
            treeNode1.Name = "NumberOfCasesFirst";
            treeNode1.Text = "N";
            treeNode2.Checked = true;
            treeNode2.Name = "MinMaxFirst";
            treeNode2.Text = "Min. - Max.";
            treeNode3.Checked = true;
            treeNode3.Name = "MeanSDFirst";
            treeNode3.Text = "Mean ± SD.";
            treeNode4.Checked = true;
            treeNode4.Name = "MedianIQRFirst";
            treeNode4.Text = "Median (IQR)";
            treeNode5.Name = "1stOption";
            treeNode5.Text = "1st Option";
            treeNode6.Name = "NumberOfCasesSecond";
            treeNode6.Text = "N";
            treeNode7.Name = "MeanSDSecond";
            treeNode7.Text = "Mean ± SD.";
            treeNode8.Name = "MedianMinMaxSecond";
            treeNode8.Text = "Median (Min. - Max.)";
            treeNode9.Name = "2ndOption";
            treeNode9.Text = "2nd Option";
            treeNode10.Name = "Parameter_Options";
            treeNode10.Text = "Parameter Options";
            this.Primary_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode10});
            this.Primary_TV.ShowLines = false;
            this.Primary_TV.Size = new System.Drawing.Size(283, 452);
            this.Primary_TV.TabIndex = 6;
            this.Primary_TV.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.DefaultTreeview_AfterCheck);
            // 
            // lbl_PrimarySettings
            // 
            this.lbl_PrimarySettings.AutoSize = true;
            this.lbl_PrimarySettings.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_PrimarySettings.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_PrimarySettings.Location = new System.Drawing.Point(11, 5);
            this.lbl_PrimarySettings.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_PrimarySettings.Name = "lbl_PrimarySettings";
            this.lbl_PrimarySettings.Size = new System.Drawing.Size(199, 30);
            this.lbl_PrimarySettings.TabIndex = 19;
            this.lbl_PrimarySettings.Text = "Primary Settings :";
            this.lbl_PrimarySettings.Click += new System.EventHandler(this.lbl_PrimarySettings_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(317, 5);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(170, 30);
            this.label1.TabIndex = 20;
            this.label1.Text = "Extra Settings :";
            // 
            // Extra_TV
            // 
            this.Extra_TV.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.Extra_TV.CheckBoxes = true;
            this.Extra_TV.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Extra_TV.ForeColor = System.Drawing.SystemColors.WindowText;
            this.Extra_TV.FullRowSelect = true;
            this.Extra_TV.LabelEdit = true;
            this.Extra_TV.Location = new System.Drawing.Point(322, 38);
            this.Extra_TV.Name = "Extra_TV";
            treeNode11.Name = "TotalColumn";
            treeNode11.Text = "Total Column";
            treeNode12.Name = "Row";
            treeNode12.Text = "Row";
            treeNode13.Checked = true;
            treeNode13.Name = "Column";
            treeNode13.Text = "Column";
            treeNode14.Name = "Percentages";
            treeNode14.Text = "Percentages";
            treeNode15.Name = "LeftCellMarginValue";
            treeNode15.Text = "0.19";
            treeNode16.Name = "Left";
            treeNode16.Text = "Left";
            treeNode17.Name = "RightCellMarginValue";
            treeNode17.Text = "0.19";
            treeNode18.Name = "Right";
            treeNode18.Text = "Right";
            treeNode19.Name = "CellMargin";
            treeNode19.Text = "Cell Margin";
            treeNode20.Name = "FirstColumn";
            treeNode20.Text = "First Column";
            treeNode21.Name = "LastColumn";
            treeNode21.Text = "Last Column";
            treeNode22.Name = "Pcontrol";
            treeNode22.Text = "Pcontrol";
            treeNode23.Name = "PaperFormat";
            treeNode23.Text = "Paper";
            this.Extra_TV.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode11,
            treeNode14,
            treeNode19,
            treeNode22,
            treeNode23});
            this.Extra_TV.ShowLines = false;
            this.Extra_TV.Size = new System.Drawing.Size(283, 452);
            this.Extra_TV.TabIndex = 21;
            // 
            // pic_DoneSettings
            // 
            this.pic_DoneSettings.Image = ((System.Drawing.Image)(resources.GetObject("pic_DoneSettings.Image")));
            this.pic_DoneSettings.Location = new System.Drawing.Point(559, 5);
            this.pic_DoneSettings.Name = "pic_DoneSettings";
            this.pic_DoneSettings.Size = new System.Drawing.Size(46, 30);
            this.pic_DoneSettings.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_DoneSettings.TabIndex = 66;
            this.pic_DoneSettings.TabStop = false;
            this.pic_DoneSettings.Click += new System.EventHandler(this.pic_DoneSettings_Click);
            // 
            // Default
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(855, 502);
            this.Controls.Add(this.pic_DoneSettings);
            this.Controls.Add(this.Extra_TV);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbl_PrimarySettings);
            this.Controls.Add(this.Primary_TV);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Default";
            this.Text = "Default";
            this.Load += new System.EventHandler(this.Default_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneSettings)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView Primary_TV;
        private System.Windows.Forms.Label lbl_PrimarySettings;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TreeView Extra_TV;
        private System.Windows.Forms.PictureBox pic_DoneSettings;
    }
}