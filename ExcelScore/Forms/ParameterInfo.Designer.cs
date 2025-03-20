namespace ExcelScore.Forms
{
    partial class ParameterInfo
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ParameterInfo));
            this.lbl_ParameterName = new System.Windows.Forms.Label();
            this.panelmove = new System.Windows.Forms.Panel();
            this.pic_CloseChooseTable = new System.Windows.Forms.PictureBox();
            this.pic_Minimize = new System.Windows.Forms.PictureBox();
            this.lbl_Measure = new System.Windows.Forms.Label();
            this.list_Level = new System.Windows.Forms.ListBox();
            this.panelmove.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_CloseChooseTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_Minimize)).BeginInit();
            this.SuspendLayout();
            // 
            // lbl_ParameterName
            // 
            this.lbl_ParameterName.AutoSize = true;
            this.lbl_ParameterName.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_ParameterName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_ParameterName.Location = new System.Drawing.Point(20, 70);
            this.lbl_ParameterName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_ParameterName.Name = "lbl_ParameterName";
            this.lbl_ParameterName.Size = new System.Drawing.Size(181, 30);
            this.lbl_ParameterName.TabIndex = 19;
            this.lbl_ParameterName.Text = "Parameter : XXX";
            // 
            // panelmove
            // 
            this.panelmove.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(235)))), ((int)(((byte)(230)))), ((int)(((byte)(221)))));
            this.panelmove.Controls.Add(this.pic_CloseChooseTable);
            this.panelmove.Controls.Add(this.pic_Minimize);
            this.panelmove.Location = new System.Drawing.Point(0, -1);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(540, 42);
            this.panelmove.TabIndex = 50;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // pic_CloseChooseTable
            // 
            this.pic_CloseChooseTable.Image = ((System.Drawing.Image)(resources.GetObject("pic_CloseChooseTable.Image")));
            this.pic_CloseChooseTable.Location = new System.Drawing.Point(486, 4);
            this.pic_CloseChooseTable.Name = "pic_CloseChooseTable";
            this.pic_CloseChooseTable.Size = new System.Drawing.Size(42, 35);
            this.pic_CloseChooseTable.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_CloseChooseTable.TabIndex = 51;
            this.pic_CloseChooseTable.TabStop = false;
            this.pic_CloseChooseTable.Click += new System.EventHandler(this.pic_CloseChooseTable_Click);
            // 
            // pic_Minimize
            // 
            this.pic_Minimize.Image = ((System.Drawing.Image)(resources.GetObject("pic_Minimize.Image")));
            this.pic_Minimize.Location = new System.Drawing.Point(418, 4);
            this.pic_Minimize.Name = "pic_Minimize";
            this.pic_Minimize.Size = new System.Drawing.Size(62, 35);
            this.pic_Minimize.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_Minimize.TabIndex = 75;
            this.pic_Minimize.TabStop = false;
            this.pic_Minimize.Click += new System.EventHandler(this.pic_Minimize_Click);
            // 
            // lbl_Measure
            // 
            this.lbl_Measure.AutoSize = true;
            this.lbl_Measure.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Measure.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Measure.Location = new System.Drawing.Point(312, 70);
            this.lbl_Measure.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Measure.Name = "lbl_Measure";
            this.lbl_Measure.Size = new System.Drawing.Size(162, 30);
            this.lbl_Measure.TabIndex = 51;
            this.lbl_Measure.Text = "Measure : XXX";
            // 
            // list_Level
            // 
            this.list_Level.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Level.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Level.FormattingEnabled = true;
            this.list_Level.ItemHeight = 21;
            this.list_Level.Location = new System.Drawing.Point(275, 112);
            this.list_Level.Name = "list_Level";
            this.list_Level.Size = new System.Drawing.Size(253, 613);
            this.list_Level.TabIndex = 52;
            // 
            // ParameterInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(540, 761);
            this.Controls.Add(this.list_Level);
            this.Controls.Add(this.lbl_Measure);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.lbl_ParameterName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ParameterInfo";
            this.Text = "ParameterInfo";
            this.Load += new System.EventHandler(this.ParameterInfo_Load);
            this.panelmove.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pic_CloseChooseTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_Minimize)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_ParameterName;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.PictureBox pic_Minimize;
        private System.Windows.Forms.PictureBox pic_CloseChooseTable;
        private System.Windows.Forms.Label lbl_Measure;
        private System.Windows.Forms.ListBox list_Level;
    }
}