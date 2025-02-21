namespace ExcelScore.Forms
{
    partial class NursingWordNew
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
            this.button1 = new System.Windows.Forms.Button();
            this.panelmove = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(819, 529);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 32);
            this.button1.TabIndex = 0;
            this.button1.Text = "Tables";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(5, 5);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(920, 25);
            this.panelmove.TabIndex = 50;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(819, 456);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 31);
            this.button2.TabIndex = 51;
            this.button2.Text = "SPSS Score";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // NursingWordNew
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(926, 592);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "NursingWordNew";
            this.Text = "NursingWordNew";
            this.Load += new System.EventHandler(this.NursingWordNew_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.Button button2;
    }
}