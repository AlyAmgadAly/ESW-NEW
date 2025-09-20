namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    partial class Comparative_Child
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.data_allPara = new System.Windows.Forms.DataGridView();
            this.Col_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMeasure = new System.Windows.Forms.DataGridViewImageColumn();
            this.ColNormality = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ViewCode_Col = new System.Windows.Forms.DataGridViewButtonColumn();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).BeginInit();
            this.SuspendLayout();
            // 
            // data_allPara
            // 
            this.data_allPara.AllowUserToAddRows = false;
            this.data_allPara.AllowUserToDeleteRows = false;
            this.data_allPara.AllowUserToResizeColumns = false;
            this.data_allPara.AllowUserToResizeRows = false;
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
            this.data_allPara.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Col_Name,
            this.ColMeasure,
            this.ColNormality,
            this.ViewCode_Col});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.data_allPara.DefaultCellStyle = dataGridViewCellStyle2;
            this.data_allPara.Dock = System.Windows.Forms.DockStyle.Left;
            this.data_allPara.EnableHeadersVisualStyles = false;
            this.data_allPara.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.Location = new System.Drawing.Point(0, 0);
            this.data_allPara.Name = "data_allPara";
            this.data_allPara.ReadOnly = true;
            this.data_allPara.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.data_allPara.RowHeadersVisible = false;
            this.data_allPara.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.data_allPara.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.data_allPara.Size = new System.Drawing.Size(343, 654);
            this.data_allPara.TabIndex = 63;
            // 
            // Col_Name
            // 
            this.Col_Name.DataPropertyName = "ColName";
            this.Col_Name.HeaderText = "Name";
            this.Col_Name.Name = "Col_Name";
            this.Col_Name.ReadOnly = true;
            this.Col_Name.Width = 150;
            // 
            // ColMeasure
            // 
            this.ColMeasure.HeaderText = "Measure";
            this.ColMeasure.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom;
            this.ColMeasure.Name = "ColMeasure";
            this.ColMeasure.ReadOnly = true;
            this.ColMeasure.Width = 75;
            // 
            // ColNormality
            // 
            this.ColNormality.HeaderText = "Normality";
            this.ColNormality.Name = "ColNormality";
            this.ColNormality.ReadOnly = true;
            this.ColNormality.Width = 75;
            // 
            // ViewCode_Col
            // 
            this.ViewCode_Col.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.ViewCode_Col.FillWeight = 50F;
            this.ViewCode_Col.HeaderText = "View";
            this.ViewCode_Col.Name = "ViewCode_Col";
            this.ViewCode_Col.ReadOnly = true;
            this.ViewCode_Col.Width = 35;
            // 
            // Comparative_Child
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(820, 654);
            this.Controls.Add(this.data_allPara);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Comparative_Child";
            this.Text = "Comparative_Child";
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView data_allPara;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Name;
        private System.Windows.Forms.DataGridViewImageColumn ColMeasure;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColNormality;
        private System.Windows.Forms.DataGridViewButtonColumn ViewCode_Col;
    }
}