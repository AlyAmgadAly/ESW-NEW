namespace ExcelScore.Forms
{
    partial class PreivewImage
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PreivewImage));
            this.pic_ImageToPreview = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ImageToPreview)).BeginInit();
            this.SuspendLayout();
            // 
            // pic_ImageToPreview
            // 
            this.pic_ImageToPreview.Location = new System.Drawing.Point(12, 6);
            this.pic_ImageToPreview.Name = "pic_ImageToPreview";
            this.pic_ImageToPreview.Size = new System.Drawing.Size(417, 328);
            this.pic_ImageToPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_ImageToPreview.TabIndex = 0;
            this.pic_ImageToPreview.TabStop = false;
            this.pic_ImageToPreview.Click += new System.EventHandler(this.pic_ImageToPreview_Click);
            // 
            // PreivewImage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(441, 343);
            this.Controls.Add(this.pic_ImageToPreview);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "PreivewImage";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PreivewImage";
            this.Load += new System.EventHandler(this.PreivewImage_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_ImageToPreview)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pic_ImageToPreview;
    }
}