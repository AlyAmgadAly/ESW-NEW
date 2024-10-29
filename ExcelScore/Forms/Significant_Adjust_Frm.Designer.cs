namespace ExcelScore.Forms
{
    partial class Significant_Adjust_Frm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Significant_Adjust_Frm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.data_allPara = new System.Windows.Forms.DataGridView();
            this.btn_Update = new System.Windows.Forms.Button();
            this.data_Pvalues = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.cmb_ChooseTableFormat = new System.Windows.Forms.ComboBox();
            this.panelmove = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_TableName = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btn_Done = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.data_Pvalues)).BeginInit();
            this.SuspendLayout();
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(986, 9);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(61, 33);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 24;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // data_allPara
            // 
            this.data_allPara.AllowUserToAddRows = false;
            this.data_allPara.AllowUserToDeleteRows = false;
            this.data_allPara.AllowUserToResizeColumns = false;
            this.data_allPara.AllowUserToResizeRows = false;
            this.data_allPara.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.data_allPara.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.data_allPara.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.data_allPara.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.data_allPara.ColumnHeadersHeight = 54;
            this.data_allPara.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.data_allPara.DefaultCellStyle = dataGridViewCellStyle6;
            this.data_allPara.EnableHeadersVisualStyles = false;
            this.data_allPara.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.Location = new System.Drawing.Point(12, 31);
            this.data_allPara.Name = "data_allPara";
            this.data_allPara.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.data_allPara.RowHeadersVisible = false;
            this.data_allPara.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.data_allPara.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.data_allPara.Size = new System.Drawing.Size(500, 620);
            this.data_allPara.TabIndex = 63;
            this.data_allPara.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.data_allPara_CellContentClick);
            this.data_allPara.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.data_allPara_CellValueChanged);
            
            // 
            // btn_Update
            // 
            this.btn_Update.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Update.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Update.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Update.ForeColor = System.Drawing.Color.Black;
            this.btn_Update.Image = ((System.Drawing.Image)(resources.GetObject("btn_Update.Image")));
            this.btn_Update.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Update.Location = new System.Drawing.Point(930, 655);
            this.btn_Update.Name = "btn_Update";
            this.btn_Update.Size = new System.Drawing.Size(108, 42);
            this.btn_Update.TabIndex = 64;
            this.btn_Update.Text = "Update";
            this.btn_Update.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Update.UseVisualStyleBackColor = false;
            this.btn_Update.Click += new System.EventHandler(this.btn_Update_Click_1);
            // 
            // data_Pvalues
            // 
            this.data_Pvalues.AllowUserToAddRows = false;
            this.data_Pvalues.AllowUserToDeleteRows = false;
            this.data_Pvalues.AllowUserToResizeColumns = false;
            this.data_Pvalues.AllowUserToResizeRows = false;
            this.data_Pvalues.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.data_Pvalues.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_Pvalues.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.data_Pvalues.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.data_Pvalues.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.data_Pvalues.ColumnHeadersHeight = 54;
            this.data_Pvalues.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.data_Pvalues.DefaultCellStyle = dataGridViewCellStyle8;
            this.data_Pvalues.EnableHeadersVisualStyles = false;
            this.data_Pvalues.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_Pvalues.Location = new System.Drawing.Point(518, 437);
            this.data_Pvalues.Name = "data_Pvalues";
            this.data_Pvalues.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.data_Pvalues.RowHeadersVisible = false;
            this.data_Pvalues.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.data_Pvalues.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.data_Pvalues.Size = new System.Drawing.Size(520, 214);
            this.data_Pvalues.TabIndex = 65;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(520, 28);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 30);
            this.label1.TabIndex = 66;
            this.label1.Text = "Result for :";
            // 
            // cmb_ChooseTableFormat
            // 
            this.cmb_ChooseTableFormat.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_ChooseTableFormat.FormattingEnabled = true;
            this.cmb_ChooseTableFormat.Items.AddRange(new object[] {
            "Default",
            "Groups 2 periods",
            "Paper",
            "Pathology Relation",
            "Pathology Comparative",
            "Relation",
            "Relation Scale Pathology",
            "Descriptive"});
            this.cmb_ChooseTableFormat.Location = new System.Drawing.Point(668, 31);
            this.cmb_ChooseTableFormat.Name = "cmb_ChooseTableFormat";
            this.cmb_ChooseTableFormat.Size = new System.Drawing.Size(183, 30);
            this.cmb_ChooseTableFormat.TabIndex = 67;
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(12, 1);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(839, 25);
            this.panelmove.TabIndex = 68;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label2.Location = new System.Drawing.Point(520, 100);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(125, 30);
            this.label2.TabIndex = 69;
            this.label2.Text = "Min. - Max";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label3.Location = new System.Drawing.Point(520, 160);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(132, 30);
            this.label3.TabIndex = 70;
            this.label3.Text = "Mean ± SD.";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label4.Location = new System.Drawing.Point(520, 220);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(145, 30);
            this.label4.TabIndex = 71;
            this.label4.Text = "Median(IQR)";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(520, 280);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(124, 30);
            this.label5.TabIndex = 72;
            this.label5.Text = "Test Of Sig";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label6.Location = new System.Drawing.Point(520, 340);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(88, 30);
            this.label6.TabIndex = 73;
            this.label6.Text = "P value";
            // 
            // txt_TableName
            // 
            this.txt_TableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_TableName.Location = new System.Drawing.Point(668, 283);
            this.txt_TableName.Margin = new System.Windows.Forms.Padding(2);
            this.txt_TableName.Name = "txt_TableName";
            this.txt_TableName.Size = new System.Drawing.Size(183, 28);
            this.txt_TableName.TabIndex = 74;
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(668, 222);
            this.textBox1.Margin = new System.Windows.Forms.Padding(2);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(183, 28);
            this.textBox1.TabIndex = 75;
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(668, 162);
            this.textBox2.Margin = new System.Windows.Forms.Padding(2);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(183, 28);
            this.textBox2.TabIndex = 76;
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox3.Location = new System.Drawing.Point(668, 100);
            this.textBox3.Margin = new System.Windows.Forms.Padding(2);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(183, 28);
            this.textBox3.TabIndex = 77;
            // 
            // textBox4
            // 
            this.textBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox4.Location = new System.Drawing.Point(668, 340);
            this.textBox4.Margin = new System.Windows.Forms.Padding(2);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(183, 28);
            this.textBox4.TabIndex = 78;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label7.Location = new System.Drawing.Point(520, 404);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(112, 30);
            this.label7.TabIndex = 79;
            this.label7.Text = "Pairwise :";
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.Black;
            this.btn_Done.Image = ((System.Drawing.Image)(resources.GetObject("btn_Done.Image")));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Done.Location = new System.Drawing.Point(816, 655);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(108, 42);
            this.btn_Done.TabIndex = 80;
            this.btn_Done.Text = "Done";
            this.btn_Done.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // Significant_Adjust_Frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1050, 700);
            this.Controls.Add(this.btn_Done);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.txt_TableName);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.cmb_ChooseTableFormat);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.data_Pvalues);
            this.Controls.Add(this.btn_Update);
            this.Controls.Add(this.data_allPara);
            this.Controls.Add(this.pic_back);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Significant_Adjust_Frm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Significant_Adjust_Frm";
            this.Load += new System.EventHandler(this.Significant_Adjust_Frm_Load);
            
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.data_Pvalues)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.DataGridView data_allPara;
        private System.Windows.Forms.Button btn_Update;
        private System.Windows.Forms.DataGridView data_Pvalues;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmb_ChooseTableFormat;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_TableName;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btn_Done;
    }
}