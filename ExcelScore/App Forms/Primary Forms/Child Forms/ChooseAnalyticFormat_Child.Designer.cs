namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    partial class ChooseAnalyticFormat_Child
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
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Default");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Comparative", new System.Windows.Forms.TreeNode[] {
            treeNode1});
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChooseAnalyticFormat_Child));
            this.pnl_Top_bar = new System.Windows.Forms.Panel();
            this.lbl_FormatSelected = new System.Windows.Forms.Label();
            this.lbl_ChooseFormat = new System.Windows.Forms.Label();
            this.tree_ChooseFormatType = new System.Windows.Forms.TreeView();
            this.pnl_LoadFormatForms = new System.Windows.Forms.Panel();
            this.pic_Exit = new System.Windows.Forms.PictureBox();
            this.pnl_Top_bar.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_Exit)).BeginInit();
            this.SuspendLayout();
            // 
            // pnl_Top_bar
            // 
            this.pnl_Top_bar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.pnl_Top_bar.Controls.Add(this.pic_Exit);
            this.pnl_Top_bar.Controls.Add(this.lbl_FormatSelected);
            this.pnl_Top_bar.Controls.Add(this.lbl_ChooseFormat);
            this.pnl_Top_bar.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_Top_bar.Location = new System.Drawing.Point(0, 0);
            this.pnl_Top_bar.Name = "pnl_Top_bar";
            this.pnl_Top_bar.Size = new System.Drawing.Size(800, 46);
            this.pnl_Top_bar.TabIndex = 9;
            // 
            // lbl_FormatSelected
            // 
            this.lbl_FormatSelected.AutoSize = true;
            this.lbl_FormatSelected.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_FormatSelected.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.lbl_FormatSelected.Location = new System.Drawing.Point(107, 7);
            this.lbl_FormatSelected.Name = "lbl_FormatSelected";
            this.lbl_FormatSelected.Size = new System.Drawing.Size(45, 30);
            this.lbl_FormatSelected.TabIndex = 6;
            this.lbl_FormatSelected.Text = "NA";
            // 
            // lbl_ChooseFormat
            // 
            this.lbl_ChooseFormat.AutoSize = true;
            this.lbl_ChooseFormat.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_ChooseFormat.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.lbl_ChooseFormat.Location = new System.Drawing.Point(14, 7);
            this.lbl_ChooseFormat.Name = "lbl_ChooseFormat";
            this.lbl_ChooseFormat.Size = new System.Drawing.Size(101, 30);
            this.lbl_ChooseFormat.TabIndex = 5;
            this.lbl_ChooseFormat.Text = "Format : ";
            // 
            // tree_ChooseFormatType
            // 
            this.tree_ChooseFormatType.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.tree_ChooseFormatType.Dock = System.Windows.Forms.DockStyle.Left;
            this.tree_ChooseFormatType.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tree_ChooseFormatType.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.tree_ChooseFormatType.HideSelection = false;
            this.tree_ChooseFormatType.Location = new System.Drawing.Point(0, 46);
            this.tree_ChooseFormatType.Name = "tree_ChooseFormatType";
            treeNode1.Name = "Default_Comparative";
            treeNode1.Text = "Default";
            treeNode2.Name = "Comparative";
            treeNode2.Text = "Comparative";
            this.tree_ChooseFormatType.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode2});
            this.tree_ChooseFormatType.Size = new System.Drawing.Size(239, 543);
            this.tree_ChooseFormatType.TabIndex = 10;
            this.tree_ChooseFormatType.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tree_ChooseFormatType_AfterSelect);
            // 
            // pnl_LoadFormatForms
            // 
            this.pnl_LoadFormatForms.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnl_LoadFormatForms.Location = new System.Drawing.Point(239, 46);
            this.pnl_LoadFormatForms.Name = "pnl_LoadFormatForms";
            this.pnl_LoadFormatForms.Size = new System.Drawing.Size(561, 543);
            this.pnl_LoadFormatForms.TabIndex = 11;
            this.pnl_LoadFormatForms.Paint += new System.Windows.Forms.PaintEventHandler(this.pnl_LoadFormatForms_Paint);
            // 
            // pic_Exit
            // 
            this.pic_Exit.Image = ((System.Drawing.Image)(resources.GetObject("pic_Exit.Image")));
            this.pic_Exit.Location = new System.Drawing.Point(743, 5);
            this.pic_Exit.Name = "pic_Exit";
            this.pic_Exit.Size = new System.Drawing.Size(45, 35);
            this.pic_Exit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_Exit.TabIndex = 75;
            this.pic_Exit.TabStop = false;
            this.pic_Exit.Click += new System.EventHandler(this.pic_Exit_Click);
            // 
            // ChooseAnalyticFormat_Child
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.ClientSize = new System.Drawing.Size(800, 589);
            this.Controls.Add(this.pnl_LoadFormatForms);
            this.Controls.Add(this.tree_ChooseFormatType);
            this.Controls.Add(this.pnl_Top_bar);
            this.ForeColor = System.Drawing.SystemColors.ControlLight;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ChooseAnalyticFormat_Child";
            this.Text = "ChooseAnalyticFormat_Child";
            this.pnl_Top_bar.ResumeLayout(false);
            this.pnl_Top_bar.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_Exit)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_Top_bar;
        private System.Windows.Forms.Label lbl_ChooseFormat;
        private System.Windows.Forms.TreeView tree_ChooseFormatType;
        private System.Windows.Forms.Panel pnl_LoadFormatForms;
        private System.Windows.Forms.Label lbl_FormatSelected;
        private System.Windows.Forms.PictureBox pic_Exit;
    }
}