namespace ExcelScore.Forms
{
    partial class ChooseTable
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChooseTable));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pic_CloseChooseTable = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_CloseChooseTable)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(758, 506);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // pic_CloseChooseTable
            // 
            this.pic_CloseChooseTable.Image = ((System.Drawing.Image)(resources.GetObject("pic_CloseChooseTable.Image")));
            this.pic_CloseChooseTable.Location = new System.Drawing.Point(760, 3);
            this.pic_CloseChooseTable.Name = "pic_CloseChooseTable";
            this.pic_CloseChooseTable.Size = new System.Drawing.Size(37, 30);
            this.pic_CloseChooseTable.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_CloseChooseTable.TabIndex = 42;
            this.pic_CloseChooseTable.TabStop = false;
            this.pic_CloseChooseTable.Click += new System.EventHandler(this.pic_CloseChooseTable_Click);
            // 
            // ChooseTable
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(800, 530);
            this.Controls.Add(this.pic_CloseChooseTable);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ChooseTable";
            this.Text = "ChooseTable";
            this.Load += new System.EventHandler(this.ChooseTable_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_CloseChooseTable)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pic_CloseChooseTable;
    }
}