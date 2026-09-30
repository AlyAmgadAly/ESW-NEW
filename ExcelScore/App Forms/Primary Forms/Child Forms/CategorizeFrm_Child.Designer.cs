namespace ExcelScore.App_Forms.Primary_Forms.Child_Forms
{
    partial class CategorizeFrm_Child
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
            this.btnTestAI = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnTestAI
            // 
            this.btnTestAI.Location = new System.Drawing.Point(323, 124);
            this.btnTestAI.Name = "btnTestAI";
            this.btnTestAI.Size = new System.Drawing.Size(75, 23);
            this.btnTestAI.TabIndex = 0;
            this.btnTestAI.Text = "Test AI";
            this.btnTestAI.UseVisualStyleBackColor = true;
            this.btnTestAI.Click += new System.EventHandler(this.btnTestAI_Click);
            // 
            // CategorizeFrm_Child
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnTestAI);
            this.Name = "CategorizeFrm_Child";
            this.Text = "CategorizeParameters";
            this.Load += new System.EventHandler(this.CategorizeFrm_Child_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnTestAI;
    }
}