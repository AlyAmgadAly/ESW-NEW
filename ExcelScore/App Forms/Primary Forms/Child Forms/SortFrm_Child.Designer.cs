namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    partial class SortFrm_Child
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SortFrm_Child));
            this.pnl_Top_bar = new System.Windows.Forms.Panel();
            this.pic_DoneSorting = new System.Windows.Forms.PictureBox();
            this.pic_SortQuestionsDown = new System.Windows.Forms.PictureBox();
            this.lbl_Sort = new System.Windows.Forms.Label();
            this.pic_SortQuestionsUP = new System.Windows.Forms.PictureBox();
            this.list_AllParameters = new System.Windows.Forms.ListBox();
            this.pnl_Top_bar.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneSorting)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsUP)).BeginInit();
            this.SuspendLayout();
            // 
            // pnl_Top_bar
            // 
            this.pnl_Top_bar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.pnl_Top_bar.Controls.Add(this.pic_DoneSorting);
            this.pnl_Top_bar.Controls.Add(this.pic_SortQuestionsDown);
            this.pnl_Top_bar.Controls.Add(this.lbl_Sort);
            this.pnl_Top_bar.Controls.Add(this.pic_SortQuestionsUP);
            this.pnl_Top_bar.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_Top_bar.Location = new System.Drawing.Point(0, 0);
            this.pnl_Top_bar.Name = "pnl_Top_bar";
            this.pnl_Top_bar.Size = new System.Drawing.Size(272, 41);
            this.pnl_Top_bar.TabIndex = 17;
            // 
            // pic_DoneSorting
            // 
            this.pic_DoneSorting.Image = ((System.Drawing.Image)(resources.GetObject("pic_DoneSorting.Image")));
            this.pic_DoneSorting.Location = new System.Drawing.Point(230, 5);
            this.pic_DoneSorting.Name = "pic_DoneSorting";
            this.pic_DoneSorting.Size = new System.Drawing.Size(39, 33);
            this.pic_DoneSorting.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_DoneSorting.TabIndex = 29;
            this.pic_DoneSorting.TabStop = false;
            this.pic_DoneSorting.Click += new System.EventHandler(this.pic_DoneSorting_Click);
            // 
            // pic_SortQuestionsDown
            // 
            this.pic_SortQuestionsDown.Image = ((System.Drawing.Image)(resources.GetObject("pic_SortQuestionsDown.Image")));
            this.pic_SortQuestionsDown.Location = new System.Drawing.Point(154, 5);
            this.pic_SortQuestionsDown.Name = "pic_SortQuestionsDown";
            this.pic_SortQuestionsDown.Size = new System.Drawing.Size(39, 33);
            this.pic_SortQuestionsDown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_SortQuestionsDown.TabIndex = 28;
            this.pic_SortQuestionsDown.TabStop = false;
            this.pic_SortQuestionsDown.Click += new System.EventHandler(this.pic_SortQuestionsDown_Click);
            // 
            // lbl_Sort
            // 
            this.lbl_Sort.AutoSize = true;
            this.lbl_Sort.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Sort.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.lbl_Sort.Location = new System.Drawing.Point(12, 3);
            this.lbl_Sort.Name = "lbl_Sort";
            this.lbl_Sort.Size = new System.Drawing.Size(55, 30);
            this.lbl_Sort.TabIndex = 5;
            this.lbl_Sort.Text = "Sort";
            // 
            // pic_SortQuestionsUP
            // 
            this.pic_SortQuestionsUP.Image = ((System.Drawing.Image)(resources.GetObject("pic_SortQuestionsUP.Image")));
            this.pic_SortQuestionsUP.Location = new System.Drawing.Point(109, 5);
            this.pic_SortQuestionsUP.Name = "pic_SortQuestionsUP";
            this.pic_SortQuestionsUP.Size = new System.Drawing.Size(39, 33);
            this.pic_SortQuestionsUP.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_SortQuestionsUP.TabIndex = 27;
            this.pic_SortQuestionsUP.TabStop = false;
            this.pic_SortQuestionsUP.Click += new System.EventHandler(this.pic_SortQuestionsUP_Click);
            // 
            // list_AllParameters
            // 
            this.list_AllParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.list_AllParameters.Dock = System.Windows.Forms.DockStyle.Fill;
            this.list_AllParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllParameters.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.list_AllParameters.FormattingEnabled = true;
            this.list_AllParameters.ItemHeight = 17;
            this.list_AllParameters.Location = new System.Drawing.Point(0, 41);
            this.list_AllParameters.Name = "list_AllParameters";
            this.list_AllParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllParameters.Size = new System.Drawing.Size(272, 540);
            this.list_AllParameters.TabIndex = 18;
            // 
            // SortFrm_Child
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.ClientSize = new System.Drawing.Size(272, 581);
            this.Controls.Add(this.list_AllParameters);
            this.Controls.Add(this.pnl_Top_bar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "SortFrm_Child";
            this.Text = "SortFrm_Child";
            this.Load += new System.EventHandler(this.SortFrm_Child_Load);
            this.pnl_Top_bar.ResumeLayout(false);
            this.pnl_Top_bar.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneSorting)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsUP)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_Top_bar;
        private System.Windows.Forms.Label lbl_Sort;
        private System.Windows.Forms.ListBox list_AllParameters;
        private System.Windows.Forms.PictureBox pic_SortQuestionsDown;
        private System.Windows.Forms.PictureBox pic_SortQuestionsUP;
        private System.Windows.Forms.PictureBox pic_DoneSorting;
    }
}