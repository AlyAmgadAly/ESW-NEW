namespace ExcelScore
{
    partial class LoadExcel
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LoadExcel));
            this.datagrid_excelsheet = new System.Windows.Forms.DataGridView();
            this.pic_ImportExcel = new System.Windows.Forms.PictureBox();
            this.pic_AddDomain = new System.Windows.Forms.PictureBox();
            this.pic_appExit = new System.Windows.Forms.PictureBox();
            this.panelmove = new System.Windows.Forms.Panel();
            this.pic_exit = new System.Windows.Forms.PictureBox();
            this.pic_ConvertSpssToExcel = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.datagrid_excelsheet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ImportExcel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AddDomain)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_appExit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_exit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ConvertSpssToExcel)).BeginInit();
            this.SuspendLayout();
            // 
            // datagrid_excelsheet
            // 
            this.datagrid_excelsheet.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle13.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(102)))), ((int)(((byte)(63)))), ((int)(((byte)(5)))));
            dataGridViewCellStyle13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle13.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(178)))), ((int)(((byte)(102)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.datagrid_excelsheet.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle13;
            this.datagrid_excelsheet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            dataGridViewCellStyle14.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle14.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle14.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle14.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.datagrid_excelsheet.DefaultCellStyle = dataGridViewCellStyle14;
            this.datagrid_excelsheet.Location = new System.Drawing.Point(12, 95);
            this.datagrid_excelsheet.Name = "datagrid_excelsheet";
            dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle15.SelectionBackColor = System.Drawing.Color.White;
            dataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.datagrid_excelsheet.RowHeadersDefaultCellStyle = dataGridViewCellStyle15;
            this.datagrid_excelsheet.RowHeadersVisible = false;
            this.datagrid_excelsheet.RowHeadersWidth = 51;
            this.datagrid_excelsheet.RowTemplate.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(223)))), ((int)(((byte)(188)))), ((int)(((byte)(136)))));
            this.datagrid_excelsheet.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.datagrid_excelsheet.Size = new System.Drawing.Size(934, 601);
            this.datagrid_excelsheet.TabIndex = 0;
            this.datagrid_excelsheet.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.datagrid_excelsheet_CellContentClick);
            // 
            // pic_ImportExcel
            // 
            this.pic_ImportExcel.Image = ((System.Drawing.Image)(resources.GetObject("pic_ImportExcel.Image")));
            this.pic_ImportExcel.Location = new System.Drawing.Point(112, 14);
            this.pic_ImportExcel.Name = "pic_ImportExcel";
            this.pic_ImportExcel.Size = new System.Drawing.Size(80, 67);
            this.pic_ImportExcel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_ImportExcel.TabIndex = 10;
            this.pic_ImportExcel.TabStop = false;
            this.pic_ImportExcel.Click += new System.EventHandler(this.pic_ImportExcel_Click);
            // 
            // pic_AddDomain
            // 
            this.pic_AddDomain.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(242)))), ((int)(((byte)(203)))), ((int)(((byte)(182)))));
            this.pic_AddDomain.Image = ((System.Drawing.Image)(resources.GetObject("pic_AddDomain.Image")));
            this.pic_AddDomain.Location = new System.Drawing.Point(224, 14);
            this.pic_AddDomain.Name = "pic_AddDomain";
            this.pic_AddDomain.Size = new System.Drawing.Size(67, 65);
            this.pic_AddDomain.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pic_AddDomain.TabIndex = 11;
            this.pic_AddDomain.TabStop = false;
            this.pic_AddDomain.Click += new System.EventHandler(this.pic_AddDomain_Click);
            // 
            // pic_appExit
            // 
            this.pic_appExit.Image = ((System.Drawing.Image)(resources.GetObject("pic_appExit.Image")));
            this.pic_appExit.Location = new System.Drawing.Point(1312, 12);
            this.pic_appExit.Name = "pic_appExit";
            this.pic_appExit.Size = new System.Drawing.Size(71, 64);
            this.pic_appExit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_appExit.TabIndex = 15;
            this.pic_appExit.TabStop = false;
            this.pic_appExit.Click += new System.EventHandler(this.pic_appExit_Click);
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(296, 2);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(562, 88);
            this.panelmove.TabIndex = 16;
            this.panelmove.Paint += new System.Windows.Forms.PaintEventHandler(this.panelmove_Paint);
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // pic_exit
            // 
            this.pic_exit.Image = ((System.Drawing.Image)(resources.GetObject("pic_exit.Image")));
            this.pic_exit.Location = new System.Drawing.Point(863, 14);
            this.pic_exit.Name = "pic_exit";
            this.pic_exit.Size = new System.Drawing.Size(67, 65);
            this.pic_exit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_exit.TabIndex = 17;
            this.pic_exit.TabStop = false;
            this.pic_exit.Click += new System.EventHandler(this.pic_exit_Click);
            // 
            // pic_ConvertSpssToExcel
            // 
            this.pic_ConvertSpssToExcel.Image = ((System.Drawing.Image)(resources.GetObject("pic_ConvertSpssToExcel.Image")));
            this.pic_ConvertSpssToExcel.Location = new System.Drawing.Point(12, 14);
            this.pic_ConvertSpssToExcel.Name = "pic_ConvertSpssToExcel";
            this.pic_ConvertSpssToExcel.Size = new System.Drawing.Size(80, 67);
            this.pic_ConvertSpssToExcel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_ConvertSpssToExcel.TabIndex = 18;
            this.pic_ConvertSpssToExcel.TabStop = false;
            this.pic_ConvertSpssToExcel.Click += new System.EventHandler(this.pic_ConvertSpssToExcel_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(252, 86);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 19;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // LoadExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(242)))), ((int)(((byte)(203)))), ((int)(((byte)(182)))));
            this.ClientSize = new System.Drawing.Size(974, 708);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.pic_ConvertSpssToExcel);
            this.Controls.Add(this.pic_exit);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.pic_appExit);
            this.Controls.Add(this.pic_AddDomain);
            this.Controls.Add(this.pic_ImportExcel);
            this.Controls.Add(this.datagrid_excelsheet);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "LoadExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.datagrid_excelsheet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ImportExcel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AddDomain)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_appExit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_exit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ConvertSpssToExcel)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView datagrid_excelsheet;
        private System.Windows.Forms.PictureBox pic_ImportExcel;
        private System.Windows.Forms.PictureBox pic_AddDomain;
        private System.Windows.Forms.PictureBox pic_appExit;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.PictureBox pic_exit;
        private System.Windows.Forms.PictureBox pic_ConvertSpssToExcel;
        private System.Windows.Forms.Button button1;
    }
}

