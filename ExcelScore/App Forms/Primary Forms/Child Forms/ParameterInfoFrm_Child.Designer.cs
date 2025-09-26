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
            this.pnl_Top_bar = new System.Windows.Forms.Panel();
            this.lbl_ParameterName = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.pnl_Top_bar.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnl_Top_bar
            // 
            this.pnl_Top_bar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(57)))), ((int)(((byte)(62)))), ((int)(((byte)(70)))));
            this.pnl_Top_bar.Controls.Add(this.label2);
            this.pnl_Top_bar.Controls.Add(this.lbl_ParameterName);
            this.pnl_Top_bar.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_Top_bar.Location = new System.Drawing.Point(0, 0);
            this.pnl_Top_bar.Name = "pnl_Top_bar";
            this.pnl_Top_bar.Size = new System.Drawing.Size(491, 39);
            this.pnl_Top_bar.TabIndex = 17;
            // 
            // lbl_ParameterName
            // 
            this.lbl_ParameterName.AutoSize = true;
            this.lbl_ParameterName.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_ParameterName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.lbl_ParameterName.Location = new System.Drawing.Point(11, 5);
            this.lbl_ParameterName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_ParameterName.Name = "lbl_ParameterName";
            this.lbl_ParameterName.Size = new System.Drawing.Size(181, 30);
            this.lbl_ParameterName.TabIndex = 20;
            this.lbl_ParameterName.Text = "Parameter : XXX";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.label2.Location = new System.Drawing.Point(295, 5);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(181, 30);
            this.label2.TabIndex = 21;
            this.label2.Text = "Parameter : XXX";
            // 
            // ParameterInfoFrm_Child
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(34)))), ((int)(((byte)(40)))), ((int)(((byte)(49)))));
            this.ClientSize = new System.Drawing.Size(491, 596);
            this.Controls.Add(this.pnl_Top_bar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ParameterInfoFrm_Child";
            this.Text = "ParameterInfoFrm_Child";
            this.pnl_Top_bar.ResumeLayout(false);
            this.pnl_Top_bar.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_Top_bar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lbl_ParameterName;
    }
}