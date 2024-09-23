namespace ExcelScore.Forms
{
    partial class Sort_Parameters_Frm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Sort_Parameters_Frm));
            this.list_AllParameters = new System.Windows.Forms.ListBox();
            this.pic_SortQuestionsUP = new System.Windows.Forms.PictureBox();
            this.pic_DoneSorting = new System.Windows.Forms.PictureBox();
            this.pic_SortQuestionsDown = new System.Windows.Forms.PictureBox();
            this.panelmove = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsUP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneSorting)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsDown)).BeginInit();
            this.SuspendLayout();
            // 
            // list_AllParameters
            // 
            this.list_AllParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AllParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllParameters.FormattingEnabled = true;
            this.list_AllParameters.ItemHeight = 17;
            this.list_AllParameters.Location = new System.Drawing.Point(12, 46);
            this.list_AllParameters.Name = "list_AllParameters";
            this.list_AllParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllParameters.Size = new System.Drawing.Size(263, 446);
            this.list_AllParameters.TabIndex = 8;
            this.list_AllParameters.SelectedIndexChanged += new System.EventHandler(this.list_AllParameters_SelectedIndexChanged);
            // 
            // pic_SortQuestionsUP
            // 
            this.pic_SortQuestionsUP.Image = ((System.Drawing.Image)(resources.GetObject("pic_SortQuestionsUP.Image")));
            this.pic_SortQuestionsUP.Location = new System.Drawing.Point(281, 156);
            this.pic_SortQuestionsUP.Name = "pic_SortQuestionsUP";
            this.pic_SortQuestionsUP.Size = new System.Drawing.Size(39, 50);
            this.pic_SortQuestionsUP.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_SortQuestionsUP.TabIndex = 25;
            this.pic_SortQuestionsUP.TabStop = false;
            this.pic_SortQuestionsUP.Click += new System.EventHandler(this.pic_SortQuestionsUP_Click);
            // 
            // pic_DoneSorting
            // 
            this.pic_DoneSorting.Image = ((System.Drawing.Image)(resources.GetObject("pic_DoneSorting.Image")));
            this.pic_DoneSorting.Location = new System.Drawing.Point(280, 445);
            this.pic_DoneSorting.Name = "pic_DoneSorting";
            this.pic_DoneSorting.Size = new System.Drawing.Size(46, 47);
            this.pic_DoneSorting.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_DoneSorting.TabIndex = 9;
            this.pic_DoneSorting.TabStop = false;
            this.pic_DoneSorting.Click += new System.EventHandler(this.pic_DoneSorting_Click);
            // 
            // pic_SortQuestionsDown
            // 
            this.pic_SortQuestionsDown.Image = ((System.Drawing.Image)(resources.GetObject("pic_SortQuestionsDown.Image")));
            this.pic_SortQuestionsDown.Location = new System.Drawing.Point(281, 240);
            this.pic_SortQuestionsDown.Name = "pic_SortQuestionsDown";
            this.pic_SortQuestionsDown.Size = new System.Drawing.Size(39, 50);
            this.pic_SortQuestionsDown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_SortQuestionsDown.TabIndex = 26;
            this.pic_SortQuestionsDown.TabStop = false;
            this.pic_SortQuestionsDown.Click += new System.EventHandler(this.pic_SortQuestionsDown_Click);
            // 
            // panelmove
            // 
            this.panelmove.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.panelmove.Location = new System.Drawing.Point(12, 2);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(314, 39);
            this.panelmove.TabIndex = 72;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // Sort_Parameters_Frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(331, 507);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.pic_SortQuestionsDown);
            this.Controls.Add(this.pic_SortQuestionsUP);
            this.Controls.Add(this.pic_DoneSorting);
            this.Controls.Add(this.list_AllParameters);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Sort_Parameters_Frm";
            this.Text = "Sort";
            this.Load += new System.EventHandler(this.Sort_Parameters_Frm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsUP)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneSorting)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_SortQuestionsDown)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox list_AllParameters;
        private System.Windows.Forms.PictureBox pic_DoneSorting;
        private System.Windows.Forms.PictureBox pic_SortQuestionsUP;
        private System.Windows.Forms.PictureBox pic_SortQuestionsDown;
        private System.Windows.Forms.Panel panelmove;
    }
}