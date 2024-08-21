namespace ExcelScore.Forms
{
    partial class CombineDomains
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CombineDomains));
            this.btn_remove = new System.Windows.Forms.Button();
            this.listbox_Alldomains = new System.Windows.Forms.ListBox();
            this.panelmove = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // btn_remove
            // 
            this.btn_remove.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_remove.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_remove.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_remove.Image = ((System.Drawing.Image)(resources.GetObject("btn_remove.Image")));
            this.btn_remove.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_remove.Location = new System.Drawing.Point(328, 413);
            this.btn_remove.Name = "btn_remove";
            this.btn_remove.Size = new System.Drawing.Size(112, 40);
            this.btn_remove.TabIndex = 1;
            this.btn_remove.Text = "Delete";
            this.btn_remove.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_remove.UseVisualStyleBackColor = false;
            this.btn_remove.Click += new System.EventHandler(this.button1_Click);
            // 
            // listbox_Alldomains
            // 
            this.listbox_Alldomains.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listbox_Alldomains.FormattingEnabled = true;
            this.listbox_Alldomains.ItemHeight = 21;
            this.listbox_Alldomains.Location = new System.Drawing.Point(27, 46);
            this.listbox_Alldomains.Name = "listbox_Alldomains";
            this.listbox_Alldomains.Size = new System.Drawing.Size(413, 361);
            this.listbox_Alldomains.TabIndex = 2;
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(11, 2);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(429, 39);
            this.panelmove.TabIndex = 26;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // CombineDomains
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(454, 465);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.listbox_Alldomains);
            this.Controls.Add(this.btn_remove);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CombineDomains";
            this.Text = "CombineDomains";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CombineDomains_FormClosed);
            this.Load += new System.EventHandler(this.CombineDomains_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btn_remove;
        private System.Windows.Forms.ListBox listbox_Alldomains;
        private System.Windows.Forms.Panel panelmove;
    }
}