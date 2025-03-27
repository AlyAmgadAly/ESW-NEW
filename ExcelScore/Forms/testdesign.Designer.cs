namespace ExcelScore.Forms
{
    partial class testdesign
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
            this.rjToggleButton1 = new ExcelScore.Custom_Controls.RJToggleButton();
            this.SuspendLayout();
            // 
            // rjToggleButton1
            // 
            this.rjToggleButton1.AutoSize = true;
            this.rjToggleButton1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.rjToggleButton1.Location = new System.Drawing.Point(228, 208);
            this.rjToggleButton1.MinimumSize = new System.Drawing.Size(45, 22);
            this.rjToggleButton1.Name = "rjToggleButton1";
            this.rjToggleButton1.OffBackColor = System.Drawing.Color.Gray;
            this.rjToggleButton1.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.rjToggleButton1.OnBackColor = System.Drawing.Color.Black;
            this.rjToggleButton1.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.rjToggleButton1.Size = new System.Drawing.Size(45, 22);
            this.rjToggleButton1.TabIndex = 0;
            this.rjToggleButton1.UseVisualStyleBackColor = false;
            // 
            // testdesign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 482);
            this.Controls.Add(this.rjToggleButton1);
            this.Name = "testdesign";
            this.Text = "testdesign";
            this.Load += new System.EventHandler(this.testdesign_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Custom_Controls.RJToggleButton rjToggleButton1;
    }
}