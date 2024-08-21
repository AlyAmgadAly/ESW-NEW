namespace ExcelScore.Forms
{
    partial class ComparativeFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ComparativeFrm));
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.btn_twoGroups = new System.Windows.Forms.Button();
            this.pic_preTwoGroups = new System.Windows.Forms.PictureBox();
            this.panelmove = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_preTwoGroups)).BeginInit();
            this.SuspendLayout();
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(704, 3);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(86, 52);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 0;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // btn_twoGroups
            // 
            this.btn_twoGroups.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_twoGroups.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_twoGroups.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_twoGroups.ForeColor = System.Drawing.Color.Black;
            this.btn_twoGroups.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_twoGroups.Location = new System.Drawing.Point(46, 36);
            this.btn_twoGroups.Name = "btn_twoGroups";
            this.btn_twoGroups.Size = new System.Drawing.Size(156, 77);
            this.btn_twoGroups.TabIndex = 16;
            this.btn_twoGroups.Text = "Groups";
            this.btn_twoGroups.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_twoGroups.UseVisualStyleBackColor = false;
            this.btn_twoGroups.Click += new System.EventHandler(this.btn_twoGroups_Click);
            // 
            // pic_preTwoGroups
            // 
            this.pic_preTwoGroups.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.pic_preTwoGroups.Image = ((System.Drawing.Image)(resources.GetObject("pic_preTwoGroups.Image")));
            this.pic_preTwoGroups.Location = new System.Drawing.Point(132, 53);
            this.pic_preTwoGroups.Name = "pic_preTwoGroups";
            this.pic_preTwoGroups.Size = new System.Drawing.Size(55, 50);
            this.pic_preTwoGroups.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_preTwoGroups.TabIndex = 17;
            this.pic_preTwoGroups.TabStop = false;
            this.pic_preTwoGroups.Click += new System.EventHandler(this.pic_preTwoGroups_Click);
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(11, 3);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(688, 39);
            this.panelmove.TabIndex = 26;
            this.panelmove.Paint += new System.Windows.Forms.PaintEventHandler(this.panelmove_Paint);
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // ComparativeFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(802, 522);
            this.Controls.Add(this.pic_preTwoGroups);
            this.Controls.Add(this.btn_twoGroups);
            this.Controls.Add(this.pic_back);
            this.Controls.Add(this.panelmove);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ComparativeFrm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ComparativeFrm";
            this.Load += new System.EventHandler(this.ComparativeFrm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_preTwoGroups)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.Button btn_twoGroups;
        private System.Windows.Forms.PictureBox pic_preTwoGroups;
        private System.Windows.Forms.Panel panelmove;
    }
}