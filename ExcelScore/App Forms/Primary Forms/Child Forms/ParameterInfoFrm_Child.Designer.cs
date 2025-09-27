namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    partial class ParameterInfoFrm_Child
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pnl_Top_bar = new System.Windows.Forms.Panel();
            this.lbl_Measure = new System.Windows.Forms.Label();
            this.lbl_ParameterName = new System.Windows.Forms.Label();
            this.data_allPara = new System.Windows.Forms.DataGridView();
            this.list_Codes = new System.Windows.Forms.ListBox();
            this.Serial = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColValues = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnl_Top_bar.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).BeginInit();
            this.SuspendLayout();
            // 
            // pnl_Top_bar
            // 
            this.pnl_Top_bar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.pnl_Top_bar.Controls.Add(this.lbl_Measure);
            this.pnl_Top_bar.Controls.Add(this.lbl_ParameterName);
            this.pnl_Top_bar.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_Top_bar.Location = new System.Drawing.Point(0, 0);
            this.pnl_Top_bar.Name = "pnl_Top_bar";
            this.pnl_Top_bar.Size = new System.Drawing.Size(604, 39);
            this.pnl_Top_bar.TabIndex = 17;
            // 
            // lbl_Measure
            // 
            this.lbl_Measure.AutoSize = true;
            this.lbl_Measure.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Measure.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.lbl_Measure.Location = new System.Drawing.Point(361, 6);
            this.lbl_Measure.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Measure.Name = "lbl_Measure";
            this.lbl_Measure.Size = new System.Drawing.Size(162, 30);
            this.lbl_Measure.TabIndex = 21;
            this.lbl_Measure.Text = "Measure : XXX";
            // 
            // lbl_ParameterName
            // 
            this.lbl_ParameterName.AutoSize = true;
            this.lbl_ParameterName.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_ParameterName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.lbl_ParameterName.Location = new System.Drawing.Point(11, 6);
            this.lbl_ParameterName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_ParameterName.Name = "lbl_ParameterName";
            this.lbl_ParameterName.Size = new System.Drawing.Size(172, 30);
            this.lbl_ParameterName.TabIndex = 20;
            this.lbl_ParameterName.Text = "Parameter : XXX";
            // 
            // data_allPara
            // 
            this.data_allPara.AllowUserToAddRows = false;
            this.data_allPara.AllowUserToDeleteRows = false;
            this.data_allPara.AllowUserToResizeColumns = false;
            this.data_allPara.AllowUserToResizeRows = false;
            this.data_allPara.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.data_allPara.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.data_allPara.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.data_allPara.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.data_allPara.ColumnHeadersHeight = 54;
            this.data_allPara.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.data_allPara.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Serial,
            this.ColValues});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.data_allPara.DefaultCellStyle = dataGridViewCellStyle2;
            this.data_allPara.Dock = System.Windows.Forms.DockStyle.Left;
            this.data_allPara.EnableHeadersVisualStyles = false;
            this.data_allPara.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.Location = new System.Drawing.Point(0, 39);
            this.data_allPara.Name = "data_allPara";
            this.data_allPara.ReadOnly = true;
            this.data_allPara.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 8F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.data_allPara.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.data_allPara.RowHeadersVisible = false;
            this.data_allPara.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.data_allPara.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.data_allPara.Size = new System.Drawing.Size(305, 572);
            this.data_allPara.TabIndex = 64;
            // 
            // list_Codes
            // 
            this.list_Codes.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.list_Codes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.list_Codes.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Codes.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.list_Codes.FormattingEnabled = true;
            this.list_Codes.ItemHeight = 21;
            this.list_Codes.Location = new System.Drawing.Point(305, 39);
            this.list_Codes.Name = "list_Codes";
            this.list_Codes.Size = new System.Drawing.Size(299, 572);
            this.list_Codes.TabIndex = 65;
            // 
            // Serial
            // 
            this.Serial.HeaderText = "Serial";
            this.Serial.Name = "Serial";
            this.Serial.ReadOnly = true;
            // 
            // ColValues
            // 
            this.ColValues.DataPropertyName = "ColValues";
            this.ColValues.HeaderText = "Values";
            this.ColValues.Name = "ColValues";
            this.ColValues.ReadOnly = true;
            this.ColValues.Width = 150;
            // 
            // ParameterInfoFrm_Child
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.ClientSize = new System.Drawing.Size(604, 611);
            this.Controls.Add(this.list_Codes);
            this.Controls.Add(this.data_allPara);
            this.Controls.Add(this.pnl_Top_bar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ParameterInfoFrm_Child";
            this.Text = "ParameterInfoFrm_Child";
            this.Load += new System.EventHandler(this.ParameterInfoFrm_Child_Load);
            this.pnl_Top_bar.ResumeLayout(false);
            this.pnl_Top_bar.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_Top_bar;
        private System.Windows.Forms.Label lbl_Measure;
        private System.Windows.Forms.Label lbl_ParameterName;
        private System.Windows.Forms.DataGridView data_allPara;
        private System.Windows.Forms.ListBox list_Codes;
        private System.Windows.Forms.DataGridViewTextBoxColumn Serial;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColValues;
    }
}