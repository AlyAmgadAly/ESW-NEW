namespace ExcelScore.FormsDesigns
{
    partial class MainFormsDesign
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
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Node2");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Comparative", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2});
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Node4");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Descriptive", new System.Windows.Forms.TreeNode[] {
            treeNode4});
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("Node1");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Relation", new System.Windows.Forms.TreeNode[] {
            treeNode6});
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainFormsDesign));
            this.TableTreeview = new System.Windows.Forms.TreeView();
            this.lbl_SelectedTable = new System.Windows.Forms.Label();
            this.panelmove = new System.Windows.Forms.Panel();
            this.panelContainer = new System.Windows.Forms.Panel();
            this.pic_Close = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_Close)).BeginInit();
            this.SuspendLayout();
            // 
            // TableTreeview
            // 
            this.TableTreeview.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.TableTreeview.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TableTreeview.ForeColor = System.Drawing.SystemColors.WindowText;
            this.TableTreeview.FullRowSelect = true;
            this.TableTreeview.Location = new System.Drawing.Point(7, 63);
            this.TableTreeview.Name = "TableTreeview";
            treeNode1.Name = "Default_Comparative";
            treeNode1.Text = "Default";
            treeNode2.Name = "Node2";
            treeNode2.Text = "Node2";
            treeNode3.Name = "Comparative";
            treeNode3.Text = "Comparative";
            treeNode4.Name = "Node4";
            treeNode4.Text = "Node4";
            treeNode5.Name = "Descriptive";
            treeNode5.Text = "Descriptive";
            treeNode6.Name = "Node1";
            treeNode6.Text = "Node1";
            treeNode7.Name = "Relation";
            treeNode7.Text = "Relation";
            this.TableTreeview.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode3,
            treeNode5,
            treeNode7});
            this.TableTreeview.ShowLines = false;
            this.TableTreeview.Size = new System.Drawing.Size(234, 495);
            this.TableTreeview.TabIndex = 5;
            this.TableTreeview.DrawNode += new System.Windows.Forms.DrawTreeNodeEventHandler(this.treeView1_DrawNode);
            this.TableTreeview.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.TableTreeview_AfterSelect);
            // 
            // lbl_SelectedTable
            // 
            this.lbl_SelectedTable.AutoSize = true;
            this.lbl_SelectedTable.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_SelectedTable.ForeColor = System.Drawing.Color.Red;
            this.lbl_SelectedTable.Location = new System.Drawing.Point(39, 25);
            this.lbl_SelectedTable.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_SelectedTable.Name = "lbl_SelectedTable";
            this.lbl_SelectedTable.Size = new System.Drawing.Size(45, 30);
            this.lbl_SelectedTable.TabIndex = 19;
            this.lbl_SelectedTable.Text = "NA";
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(2, 1);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(898, 25);
            this.panelmove.TabIndex = 50;
            this.panelmove.Paint += new System.Windows.Forms.PaintEventHandler(this.panelmove_Paint);
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // panelContainer
            // 
            this.panelContainer.Location = new System.Drawing.Point(242, 63);
            this.panelContainer.Name = "panelContainer";
            this.panelContainer.Size = new System.Drawing.Size(618, 495);
            this.panelContainer.TabIndex = 51;
            this.panelContainer.Paint += new System.Windows.Forms.PaintEventHandler(this.panelContainer_Paint);
            // 
            // pic_Close
            // 
            this.pic_Close.Image = ((System.Drawing.Image)(resources.GetObject("pic_Close.Image")));
            this.pic_Close.Location = new System.Drawing.Point(825, 26);
            this.pic_Close.Name = "pic_Close";
            this.pic_Close.Size = new System.Drawing.Size(35, 30);
            this.pic_Close.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_Close.TabIndex = 52;
            this.pic_Close.TabStop = false;
            this.pic_Close.Click += new System.EventHandler(this.pic_Close_Click);
            // 
            // MainFormsDesign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(862, 560);
            this.Controls.Add(this.pic_Close);
            this.Controls.Add(this.panelContainer);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.lbl_SelectedTable);
            this.Controls.Add(this.TableTreeview);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainFormsDesign";
            this.Text = "MainFormsDesign";
            this.Load += new System.EventHandler(this.MainFormsDesign_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_Close)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView TableTreeview;
        private System.Windows.Forms.Label lbl_SelectedTable;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.Panel panelContainer;
        private System.Windows.Forms.PictureBox pic_Close;
    }
}