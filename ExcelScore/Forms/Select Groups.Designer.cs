namespace ExcelScore.Forms
{
    partial class Select_Groups
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Select_Groups));
            this.pic_RemoveSelectedList = new System.Windows.Forms.PictureBox();
            this.btn_DoneSelection = new System.Windows.Forms.Button();
            this.lbl_parameterName = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.pic_AllNominalToSelected = new System.Windows.Forms.PictureBox();
            this.list_SelectedValues = new System.Windows.Forms.ListBox();
            this.list_AllValues = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectedList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllNominalToSelected)).BeginInit();
            this.SuspendLayout();
            // 
            // pic_RemoveSelectedList
            // 
            this.pic_RemoveSelectedList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveSelectedList.Image")));
            this.pic_RemoveSelectedList.Location = new System.Drawing.Point(381, 8);
            this.pic_RemoveSelectedList.Name = "pic_RemoveSelectedList";
            this.pic_RemoveSelectedList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveSelectedList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveSelectedList.TabIndex = 58;
            this.pic_RemoveSelectedList.TabStop = false;
            this.pic_RemoveSelectedList.Click += new System.EventHandler(this.pic_RemoveSelectedList_Click);
            // 
            // btn_DoneSelection
            // 
            this.btn_DoneSelection.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_DoneSelection.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_DoneSelection.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_DoneSelection.ForeColor = System.Drawing.Color.Black;
            this.btn_DoneSelection.Image = ((System.Drawing.Image)(resources.GetObject("btn_DoneSelection.Image")));
            this.btn_DoneSelection.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_DoneSelection.Location = new System.Drawing.Point(354, 392);
            this.btn_DoneSelection.Name = "btn_DoneSelection";
            this.btn_DoneSelection.Size = new System.Drawing.Size(99, 43);
            this.btn_DoneSelection.TabIndex = 57;
            this.btn_DoneSelection.Text = "Done";
            this.btn_DoneSelection.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_DoneSelection.UseVisualStyleBackColor = false;
            this.btn_DoneSelection.Click += new System.EventHandler(this.btn_DoneSelection_Click);
            // 
            // lbl_parameterName
            // 
            this.lbl_parameterName.AutoSize = true;
            this.lbl_parameterName.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_parameterName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_parameterName.Location = new System.Drawing.Point(11, 8);
            this.lbl_parameterName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_parameterName.Name = "lbl_parameterName";
            this.lbl_parameterName.Size = new System.Drawing.Size(80, 30);
            this.lbl_parameterName.TabIndex = 56;
            this.lbl_parameterName.Text = "Empty";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(267, 8);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(82, 30);
            this.label5.TabIndex = 55;
            this.label5.Text = "Equal :";
            // 
            // pic_AllNominalToSelected
            // 
            this.pic_AllNominalToSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllNominalToSelected.Image")));
            this.pic_AllNominalToSelected.Location = new System.Drawing.Point(202, 178);
            this.pic_AllNominalToSelected.Name = "pic_AllNominalToSelected";
            this.pic_AllNominalToSelected.Size = new System.Drawing.Size(59, 38);
            this.pic_AllNominalToSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllNominalToSelected.TabIndex = 54;
            this.pic_AllNominalToSelected.TabStop = false;
            this.pic_AllNominalToSelected.Click += new System.EventHandler(this.pic_AllNominalToSelected_Click);
            // 
            // list_SelectedValues
            // 
            this.list_SelectedValues.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_SelectedValues.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_SelectedValues.FormattingEnabled = true;
            this.list_SelectedValues.ItemHeight = 17;
            this.list_SelectedValues.Location = new System.Drawing.Point(272, 42);
            this.list_SelectedValues.Name = "list_SelectedValues";
            this.list_SelectedValues.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_SelectedValues.Size = new System.Drawing.Size(181, 344);
            this.list_SelectedValues.TabIndex = 53;
            // 
            // list_AllValues
            // 
            this.list_AllValues.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AllValues.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllValues.FormattingEnabled = true;
            this.list_AllValues.ItemHeight = 17;
            this.list_AllValues.Location = new System.Drawing.Point(12, 42);
            this.list_AllValues.Name = "list_AllValues";
            this.list_AllValues.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllValues.Size = new System.Drawing.Size(181, 344);
            this.list_AllValues.TabIndex = 52;
            // 
            // Select_Groups
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(465, 442);
            this.Controls.Add(this.pic_RemoveSelectedList);
            this.Controls.Add(this.btn_DoneSelection);
            this.Controls.Add(this.lbl_parameterName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.pic_AllNominalToSelected);
            this.Controls.Add(this.list_SelectedValues);
            this.Controls.Add(this.list_AllValues);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Select_Groups";
            this.Text = "Select_Groups";
            this.Load += new System.EventHandler(this.Select_Groups_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectedList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllNominalToSelected)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pic_RemoveSelectedList;
        private System.Windows.Forms.Button btn_DoneSelection;
        private System.Windows.Forms.Label lbl_parameterName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.PictureBox pic_AllNominalToSelected;
        private System.Windows.Forms.ListBox list_SelectedValues;
        private System.Windows.Forms.ListBox list_AllValues;
    }
}