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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.data_allPara = new System.Windows.Forms.DataGridView();
            this.btn_Update = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).BeginInit();
            this.SuspendLayout();
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(952, 12);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(86, 50);
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
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.data_allPara.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.data_allPara.ColumnHeadersHeight = 54;
            this.data_allPara.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.data_allPara.DefaultCellStyle = dataGridViewCellStyle2;
            this.data_allPara.EnableHeadersVisualStyles = false;
            this.data_allPara.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.Location = new System.Drawing.Point(12, 12);
            this.data_allPara.Name = "data_allPara";
            this.data_allPara.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.data_allPara.RowHeadersVisible = false;
            this.data_allPara.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.data_allPara.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.data_allPara.Size = new System.Drawing.Size(453, 676);
            this.data_allPara.TabIndex = 63;
            // 
            // btn_Update
            // 
            this.btn_Update.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Update.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Update.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Update.ForeColor = System.Drawing.Color.Black;
            this.btn_Update.Image = ((System.Drawing.Image)(resources.GetObject("btn_Update.Image")));
            this.btn_Update.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Update.Location = new System.Drawing.Point(930, 634);
            this.btn_Update.Name = "btn_Update";
            this.btn_Update.Size = new System.Drawing.Size(108, 54);
            this.btn_Update.TabIndex = 64;
            this.btn_Update.Text = "Update";
            this.btn_Update.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Update.UseVisualStyleBackColor = false;
            this.btn_Update.Click += new System.EventHandler(this.btn_Update_Click_1);
            // 
            // Significant_Adjust_Frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1050, 700);
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
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.DataGridView data_allPara;
        private System.Windows.Forms.Button btn_Update;
    }
}