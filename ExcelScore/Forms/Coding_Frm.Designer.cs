namespace ExcelScore.Forms
{
    partial class Coding_Frm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Coding_Frm));
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.btn_Update = new System.Windows.Forms.Button();
            this.btn_Done = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            this.SuspendLayout();
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(702, 12);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(86, 50);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 24;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // btn_Update
            // 
            this.btn_Update.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Update.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Update.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Update.ForeColor = System.Drawing.Color.Black;
            this.btn_Update.Image = ((System.Drawing.Image)(resources.GetObject("btn_Update.Image")));
            this.btn_Update.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Update.Location = new System.Drawing.Point(564, 396);
            this.btn_Update.Name = "btn_Update";
            this.btn_Update.Size = new System.Drawing.Size(108, 54);
            this.btn_Update.TabIndex = 61;
            this.btn_Update.Text = "Update";
            this.btn_Update.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Update.UseVisualStyleBackColor = false;
            this.btn_Update.Click += new System.EventHandler(this.btn_Update_Click);
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.Black;
            this.btn_Done.Image = ((System.Drawing.Image)(resources.GetObject("btn_Done.Image")));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Done.Location = new System.Drawing.Point(678, 396);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(108, 54);
            this.btn_Done.TabIndex = 62;
            this.btn_Done.Text = "Done";
            this.btn_Done.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // Coding_Frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(800, 472);
            this.Controls.Add(this.btn_Done);
            this.Controls.Add(this.btn_Update);
            this.Controls.Add(this.pic_back);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Coding_Frm";
            this.Text = "Coding_Frm";
            this.Load += new System.EventHandler(this.Coding_Frm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.Button btn_Update;
        private System.Windows.Forms.Button btn_Done;
    }
}