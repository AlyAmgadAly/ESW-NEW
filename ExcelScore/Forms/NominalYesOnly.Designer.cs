namespace ExcelScore.Forms
{
    partial class NominalYesOnly
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NominalYesOnly));
            this.list_AllNominal = new System.Windows.Forms.ListBox();
            this.list_NominalYesOnly = new System.Windows.Forms.ListBox();
            this.pic_AllNominalToNominalYesOnly = new System.Windows.Forms.PictureBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_DoneNominalYes = new System.Windows.Forms.Button();
            this.pic_RemoveNominalList = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllNominalToNominalYesOnly)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).BeginInit();
            this.SuspendLayout();
            // 
            // list_AllNominal
            // 
            this.list_AllNominal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AllNominal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllNominal.FormattingEnabled = true;
            this.list_AllNominal.ItemHeight = 17;
            this.list_AllNominal.Location = new System.Drawing.Point(12, 45);
            this.list_AllNominal.Name = "list_AllNominal";
            this.list_AllNominal.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllNominal.Size = new System.Drawing.Size(181, 344);
            this.list_AllNominal.TabIndex = 9;
            // 
            // list_NominalYesOnly
            // 
            this.list_NominalYesOnly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_NominalYesOnly.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_NominalYesOnly.FormattingEnabled = true;
            this.list_NominalYesOnly.ItemHeight = 17;
            this.list_NominalYesOnly.Location = new System.Drawing.Point(272, 45);
            this.list_NominalYesOnly.Name = "list_NominalYesOnly";
            this.list_NominalYesOnly.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_NominalYesOnly.Size = new System.Drawing.Size(181, 344);
            this.list_NominalYesOnly.TabIndex = 10;
            // 
            // pic_AllNominalToNominalYesOnly
            // 
            this.pic_AllNominalToNominalYesOnly.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllNominalToNominalYesOnly.Image")));
            this.pic_AllNominalToNominalYesOnly.Location = new System.Drawing.Point(202, 181);
            this.pic_AllNominalToNominalYesOnly.Name = "pic_AllNominalToNominalYesOnly";
            this.pic_AllNominalToNominalYesOnly.Size = new System.Drawing.Size(59, 38);
            this.pic_AllNominalToNominalYesOnly.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllNominalToNominalYesOnly.TabIndex = 15;
            this.pic_AllNominalToNominalYesOnly.TabStop = false;
            this.pic_AllNominalToNominalYesOnly.Click += new System.EventHandler(this.pic_AllNominalToNominalYesOnly_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(267, 11);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(112, 30);
            this.label5.TabIndex = 48;
            this.label5.Text = "Yes Only :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(11, 11);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 30);
            this.label1.TabIndex = 49;
            this.label1.Text = "Nominal :";
            // 
            // btn_DoneNominalYes
            // 
            this.btn_DoneNominalYes.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_DoneNominalYes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_DoneNominalYes.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_DoneNominalYes.ForeColor = System.Drawing.Color.Black;
            this.btn_DoneNominalYes.Image = ((System.Drawing.Image)(resources.GetObject("btn_DoneNominalYes.Image")));
            this.btn_DoneNominalYes.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_DoneNominalYes.Location = new System.Drawing.Point(354, 395);
            this.btn_DoneNominalYes.Name = "btn_DoneNominalYes";
            this.btn_DoneNominalYes.Size = new System.Drawing.Size(99, 43);
            this.btn_DoneNominalYes.TabIndex = 50;
            this.btn_DoneNominalYes.Text = "Done";
            this.btn_DoneNominalYes.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_DoneNominalYes.UseVisualStyleBackColor = false;
            this.btn_DoneNominalYes.Click += new System.EventHandler(this.btn_DoneNominalYes_Click);
            // 
            // pic_RemoveNominalList
            // 
            this.pic_RemoveNominalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveNominalList.Image")));
            this.pic_RemoveNominalList.Location = new System.Drawing.Point(381, 11);
            this.pic_RemoveNominalList.Name = "pic_RemoveNominalList";
            this.pic_RemoveNominalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveNominalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveNominalList.TabIndex = 51;
            this.pic_RemoveNominalList.TabStop = false;
            this.pic_RemoveNominalList.Click += new System.EventHandler(this.pic_RemoveNominalList_Click);
            // 
            // NominalYesOnly
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(465, 442);
            this.Controls.Add(this.pic_RemoveNominalList);
            this.Controls.Add(this.btn_DoneNominalYes);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.pic_AllNominalToNominalYesOnly);
            this.Controls.Add(this.list_NominalYesOnly);
            this.Controls.Add(this.list_AllNominal);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "NominalYesOnly";
            this.Text = "NominalYesOnly";
            this.Load += new System.EventHandler(this.NominalYesOnly_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllNominalToNominalYesOnly)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox list_AllNominal;
        private System.Windows.Forms.ListBox list_NominalYesOnly;
        private System.Windows.Forms.PictureBox pic_AllNominalToNominalYesOnly;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_DoneNominalYes;
        private System.Windows.Forms.PictureBox pic_RemoveNominalList;
    }
}