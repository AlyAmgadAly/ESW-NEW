namespace ExcelScore.Forms
{
    partial class Options_Frm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Options_Frm));
            this.lbl_Add_Columns = new System.Windows.Forms.Label();
            this.lbl_Percent = new System.Windows.Forms.Label();
            this.lbl_Format = new System.Windows.Forms.Label();
            this.pic_DoneOptions = new System.Windows.Forms.PictureBox();
            this.panelmove = new System.Windows.Forms.Panel();
            this.check_TotalColumn_new = new ExcelScore.Custom_Controls.RJToggleButton();
            this.label1 = new System.Windows.Forms.Label();
            this.check_Perc_Row_new = new ExcelScore.Custom_Controls.RJToggleButton();
            this.check_Maha_new = new ExcelScore.Custom_Controls.RJToggleButton();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneOptions)).BeginInit();
            this.SuspendLayout();
            // 
            // lbl_Add_Columns
            // 
            this.lbl_Add_Columns.AutoSize = true;
            this.lbl_Add_Columns.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Add_Columns.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Add_Columns.Location = new System.Drawing.Point(22, 79);
            this.lbl_Add_Columns.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Add_Columns.Name = "lbl_Add_Columns";
            this.lbl_Add_Columns.Size = new System.Drawing.Size(168, 32);
            this.lbl_Add_Columns.TabIndex = 19;
            this.lbl_Add_Columns.Text = "Add Columns";
            // 
            // lbl_Percent
            // 
            this.lbl_Percent.AutoSize = true;
            this.lbl_Percent.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Percent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Percent.Location = new System.Drawing.Point(20, 188);
            this.lbl_Percent.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Percent.Name = "lbl_Percent";
            this.lbl_Percent.Size = new System.Drawing.Size(100, 32);
            this.lbl_Percent.TabIndex = 54;
            this.lbl_Percent.Text = "Percent";
            // 
            // lbl_Format
            // 
            this.lbl_Format.AutoSize = true;
            this.lbl_Format.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Format.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Format.Location = new System.Drawing.Point(23, 300);
            this.lbl_Format.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Format.Name = "lbl_Format";
            this.lbl_Format.Size = new System.Drawing.Size(95, 32);
            this.lbl_Format.TabIndex = 64;
            this.lbl_Format.Text = "Format";
            // 
            // pic_DoneOptions
            // 
            this.pic_DoneOptions.Image = ((System.Drawing.Image)(resources.GetObject("pic_DoneOptions.Image")));
            this.pic_DoneOptions.Location = new System.Drawing.Point(172, 377);
            this.pic_DoneOptions.Name = "pic_DoneOptions";
            this.pic_DoneOptions.Size = new System.Drawing.Size(46, 47);
            this.pic_DoneOptions.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_DoneOptions.TabIndex = 65;
            this.pic_DoneOptions.TabStop = false;
            this.pic_DoneOptions.Click += new System.EventHandler(this.pic_DoneOptions_Click);
            // 
            // panelmove
            // 
            this.panelmove.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.panelmove.Location = new System.Drawing.Point(5, 1);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(225, 28);
            this.panelmove.TabIndex = 73;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // check_TotalColumn_new
            // 
            this.check_TotalColumn_new.AutoSize = true;
            this.check_TotalColumn_new.Location = new System.Drawing.Point(24, 132);
            this.check_TotalColumn_new.MinimumSize = new System.Drawing.Size(45, 22);
            this.check_TotalColumn_new.Name = "check_TotalColumn_new";
            this.check_TotalColumn_new.OffBackColor = System.Drawing.Color.Gray;
            this.check_TotalColumn_new.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.check_TotalColumn_new.OnBackColor = System.Drawing.Color.Black;
            this.check_TotalColumn_new.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.check_TotalColumn_new.Size = new System.Drawing.Size(45, 22);
            this.check_TotalColumn_new.TabIndex = 74;
            this.check_TotalColumn_new.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(74, 127);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 30);
            this.label1.TabIndex = 75;
            this.label1.Text = "Total Column";
            // 
            // check_Perc_Row_new
            // 
            this.check_Perc_Row_new.AutoSize = true;
            this.check_Perc_Row_new.Location = new System.Drawing.Point(24, 244);
            this.check_Perc_Row_new.MinimumSize = new System.Drawing.Size(45, 22);
            this.check_Perc_Row_new.Name = "check_Perc_Row_new";
            this.check_Perc_Row_new.OffBackColor = System.Drawing.Color.Gray;
            this.check_Perc_Row_new.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.check_Perc_Row_new.OnBackColor = System.Drawing.Color.Black;
            this.check_Perc_Row_new.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.check_Perc_Row_new.Size = new System.Drawing.Size(45, 22);
            this.check_Perc_Row_new.TabIndex = 76;
            this.check_Perc_Row_new.UseVisualStyleBackColor = true;
            // 
            // check_Maha_new
            // 
            this.check_Maha_new.AutoSize = true;
            this.check_Maha_new.Location = new System.Drawing.Point(24, 354);
            this.check_Maha_new.MinimumSize = new System.Drawing.Size(45, 22);
            this.check_Maha_new.Name = "check_Maha_new";
            this.check_Maha_new.OffBackColor = System.Drawing.Color.Gray;
            this.check_Maha_new.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.check_Maha_new.OnBackColor = System.Drawing.Color.Black;
            this.check_Maha_new.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.check_Maha_new.Size = new System.Drawing.Size(45, 22);
            this.check_Maha_new.TabIndex = 77;
            this.check_Maha_new.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(73, 239);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 30);
            this.label2.TabIndex = 78;
            this.label2.Text = "% Row";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(74, 350);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 30);
            this.label3.TabIndex = 79;
            this.label3.Text = "Maha";
            // 
            // Options_Frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(236, 446);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.check_Maha_new);
            this.Controls.Add(this.check_Perc_Row_new);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.check_TotalColumn_new);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.pic_DoneOptions);
            this.Controls.Add(this.lbl_Format);
            this.Controls.Add(this.lbl_Percent);
            this.Controls.Add(this.lbl_Add_Columns);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Options_Frm";
            this.Text = "Options_Frm";
            this.Load += new System.EventHandler(this.Options_Frm_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.Options_Frm_Paint);
            ((System.ComponentModel.ISupportInitialize)(this.pic_DoneOptions)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_Add_Columns;
        private System.Windows.Forms.Label lbl_Percent;
        private System.Windows.Forms.Label lbl_Format;
        private System.Windows.Forms.PictureBox pic_DoneOptions;
        private System.Windows.Forms.Panel panelmove;
        private Custom_Controls.RJToggleButton check_TotalColumn_new;
        private System.Windows.Forms.Label label1;
        private Custom_Controls.RJToggleButton check_Perc_Row_new;
        private Custom_Controls.RJToggleButton check_Maha_new;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}