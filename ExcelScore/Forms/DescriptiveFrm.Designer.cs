namespace ExcelScore.Forms
{
    partial class DescriptiveFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DescriptiveFrm));
            this.panelmove = new System.Windows.Forms.Panel();
            this.list_AllParameters = new System.Windows.Forms.ListBox();
            this.list_Nominal = new System.Windows.Forms.ListBox();
            this.list_Scale = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_TableName = new System.Windows.Forms.TextBox();
            this.cmb_TableNames = new System.Windows.Forms.ComboBox();
            this.list_ViewTableParameters = new System.Windows.Forms.ListBox();
            this.lbl_Scale = new System.Windows.Forms.Label();
            this.lbl_Nominal = new System.Windows.Forms.Label();
            this.pic_removeScale = new System.Windows.Forms.PictureBox();
            this.pic_RemoveNominalList = new System.Windows.Forms.PictureBox();
            this.btn_Done = new System.Windows.Forms.Button();
            this.pic_removeTableSelected = new System.Windows.Forms.PictureBox();
            this.pic_addTable = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToScale = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToNominal = new System.Windows.Forms.PictureBox();
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.btn_SortTable = new System.Windows.Forms.Button();
            this.btn_SwicthDescrToComp = new System.Windows.Forms.Button();
            this.lbl_CurrentAction = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeScale)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToScale)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNominal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            this.SuspendLayout();
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(11, 4);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(899, 18);
            this.panelmove.TabIndex = 50;
            this.panelmove.Paint += new System.Windows.Forms.PaintEventHandler(this.panelmove_Paint);
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // list_AllParameters
            // 
            this.list_AllParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AllParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllParameters.FormattingEnabled = true;
            this.list_AllParameters.ItemHeight = 17;
            this.list_AllParameters.Location = new System.Drawing.Point(12, 54);
            this.list_AllParameters.Name = "list_AllParameters";
            this.list_AllParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllParameters.Size = new System.Drawing.Size(263, 582);
            this.list_AllParameters.TabIndex = 52;
            // 
            // list_Nominal
            // 
            this.list_Nominal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Nominal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Nominal.FormattingEnabled = true;
            this.list_Nominal.ItemHeight = 17;
            this.list_Nominal.Location = new System.Drawing.Point(349, 84);
            this.list_Nominal.Name = "list_Nominal";
            this.list_Nominal.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Nominal.Size = new System.Drawing.Size(263, 225);
            this.list_Nominal.TabIndex = 53;
            this.list_Nominal.SelectedIndexChanged += new System.EventHandler(this.list_Nominal_SelectedIndexChanged);
            // 
            // list_Scale
            // 
            this.list_Scale.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Scale.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Scale.FormattingEnabled = true;
            this.list_Scale.ItemHeight = 17;
            this.list_Scale.Location = new System.Drawing.Point(349, 354);
            this.list_Scale.Name = "list_Scale";
            this.list_Scale.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Scale.Size = new System.Drawing.Size(263, 276);
            this.list_Scale.TabIndex = 54;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(732, 71);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(147, 30);
            this.label5.TabIndex = 62;
            this.label5.Text = "Table Name :";
            // 
            // txt_TableName
            // 
            this.txt_TableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_TableName.Location = new System.Drawing.Point(737, 103);
            this.txt_TableName.Margin = new System.Windows.Forms.Padding(2);
            this.txt_TableName.Name = "txt_TableName";
            this.txt_TableName.Size = new System.Drawing.Size(183, 28);
            this.txt_TableName.TabIndex = 61;
            // 
            // cmb_TableNames
            // 
            this.cmb_TableNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_TableNames.FormattingEnabled = true;
            this.cmb_TableNames.Location = new System.Drawing.Point(737, 400);
            this.cmb_TableNames.Name = "cmb_TableNames";
            this.cmb_TableNames.Size = new System.Drawing.Size(183, 30);
            this.cmb_TableNames.TabIndex = 58;
            this.cmb_TableNames.SelectedIndexChanged += new System.EventHandler(this.cmb_TableNames_SelectedIndexChanged);
            // 
            // list_ViewTableParameters
            // 
            this.list_ViewTableParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_ViewTableParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_ViewTableParameters.FormattingEnabled = true;
            this.list_ViewTableParameters.ItemHeight = 17;
            this.list_ViewTableParameters.Location = new System.Drawing.Point(737, 169);
            this.list_ViewTableParameters.Name = "list_ViewTableParameters";
            this.list_ViewTableParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_ViewTableParameters.Size = new System.Drawing.Size(183, 208);
            this.list_ViewTableParameters.TabIndex = 57;
            // 
            // lbl_Scale
            // 
            this.lbl_Scale.AutoSize = true;
            this.lbl_Scale.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Scale.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Scale.Location = new System.Drawing.Point(344, 321);
            this.lbl_Scale.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Scale.Name = "lbl_Scale";
            this.lbl_Scale.Size = new System.Drawing.Size(78, 30);
            this.lbl_Scale.TabIndex = 65;
            this.lbl_Scale.Text = "Scale :";
            // 
            // lbl_Nominal
            // 
            this.lbl_Nominal.AutoSize = true;
            this.lbl_Nominal.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Nominal.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Nominal.Location = new System.Drawing.Point(344, 48);
            this.lbl_Nominal.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Nominal.Name = "lbl_Nominal";
            this.lbl_Nominal.Size = new System.Drawing.Size(112, 30);
            this.lbl_Nominal.TabIndex = 67;
            this.lbl_Nominal.Text = "Nominal :";
            // 
            // pic_removeScale
            // 
            this.pic_removeScale.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeScale.Image")));
            this.pic_removeScale.Location = new System.Drawing.Point(461, 321);
            this.pic_removeScale.Name = "pic_removeScale";
            this.pic_removeScale.Size = new System.Drawing.Size(37, 30);
            this.pic_removeScale.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeScale.TabIndex = 66;
            this.pic_removeScale.TabStop = false;
            this.pic_removeScale.Click += new System.EventHandler(this.pic_removeScale_Click);
            // 
            // pic_RemoveNominalList
            // 
            this.pic_RemoveNominalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveNominalList.Image")));
            this.pic_RemoveNominalList.Location = new System.Drawing.Point(461, 48);
            this.pic_RemoveNominalList.Name = "pic_RemoveNominalList";
            this.pic_RemoveNominalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveNominalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveNominalList.TabIndex = 64;
            this.pic_RemoveNominalList.TabStop = false;
            this.pic_RemoveNominalList.Click += new System.EventHandler(this.pic_RemoveNominalList_Click);
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.Black;
            this.btn_Done.Image = ((System.Drawing.Image)(resources.GetObject("btn_Done.Image")));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Done.Location = new System.Drawing.Point(869, 575);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(104, 55);
            this.btn_Done.TabIndex = 63;
            this.btn_Done.Text = "Done";
            this.btn_Done.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // pic_removeTableSelected
            // 
            this.pic_removeTableSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeTableSelected.Image")));
            this.pic_removeTableSelected.Location = new System.Drawing.Point(936, 400);
            this.pic_removeTableSelected.Name = "pic_removeTableSelected";
            this.pic_removeTableSelected.Size = new System.Drawing.Size(37, 30);
            this.pic_removeTableSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeTableSelected.TabIndex = 60;
            this.pic_removeTableSelected.TabStop = false;
            this.pic_removeTableSelected.Click += new System.EventHandler(this.pic_removeTableSelected_Click);
            // 
            // pic_addTable
            // 
            this.pic_addTable.Image = ((System.Drawing.Image)(resources.GetObject("pic_addTable.Image")));
            this.pic_addTable.Location = new System.Drawing.Point(936, 103);
            this.pic_addTable.Name = "pic_addTable";
            this.pic_addTable.Size = new System.Drawing.Size(37, 28);
            this.pic_addTable.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_addTable.TabIndex = 59;
            this.pic_addTable.TabStop = false;
            this.pic_addTable.Click += new System.EventHandler(this.pic_addTable_Click);
            // 
            // pic_AllParaToScale
            // 
            this.pic_AllParaToScale.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToScale.Image")));
            this.pic_AllParaToScale.Location = new System.Drawing.Point(281, 459);
            this.pic_AllParaToScale.Name = "pic_AllParaToScale";
            this.pic_AllParaToScale.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToScale.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToScale.TabIndex = 56;
            this.pic_AllParaToScale.TabStop = false;
            this.pic_AllParaToScale.Click += new System.EventHandler(this.pic_AllParaToScale_Click);
            // 
            // pic_AllParaToNominal
            // 
            this.pic_AllParaToNominal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToNominal.Image")));
            this.pic_AllParaToNominal.Location = new System.Drawing.Point(284, 179);
            this.pic_AllParaToNominal.Name = "pic_AllParaToNominal";
            this.pic_AllParaToNominal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToNominal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToNominal.TabIndex = 55;
            this.pic_AllParaToNominal.TabStop = false;
            this.pic_AllParaToNominal.Click += new System.EventHandler(this.pic_AllParaToNominal_Click);
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(915, 4);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(84, 50);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 51;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // btn_SortTable
            // 
            this.btn_SortTable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_SortTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_SortTable.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_SortTable.ForeColor = System.Drawing.Color.Black;
            this.btn_SortTable.Image = ((System.Drawing.Image)(resources.GetObject("btn_SortTable.Image")));
            this.btn_SortTable.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_SortTable.Location = new System.Drawing.Point(737, 575);
            this.btn_SortTable.Name = "btn_SortTable";
            this.btn_SortTable.Size = new System.Drawing.Size(114, 55);
            this.btn_SortTable.TabIndex = 68;
            this.btn_SortTable.Text = "Sort";
            this.btn_SortTable.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_SortTable.UseVisualStyleBackColor = false;
            this.btn_SortTable.Click += new System.EventHandler(this.btn_SortTable_Click);
            // 
            // btn_SwicthDescrToComp
            // 
            this.btn_SwicthDescrToComp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_SwicthDescrToComp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_SwicthDescrToComp.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_SwicthDescrToComp.ForeColor = System.Drawing.Color.Black;
            this.btn_SwicthDescrToComp.Image = ((System.Drawing.Image)(resources.GetObject("btn_SwicthDescrToComp.Image")));
            this.btn_SwicthDescrToComp.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_SwicthDescrToComp.Location = new System.Drawing.Point(737, 499);
            this.btn_SwicthDescrToComp.Name = "btn_SwicthDescrToComp";
            this.btn_SwicthDescrToComp.Size = new System.Drawing.Size(114, 55);
            this.btn_SwicthDescrToComp.TabIndex = 69;
            this.btn_SwicthDescrToComp.Text = "Compare";
            this.btn_SwicthDescrToComp.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_SwicthDescrToComp.UseVisualStyleBackColor = false;
            this.btn_SwicthDescrToComp.Click += new System.EventHandler(this.btn_SwicthDescrToComp_Click);
            // 
            // lbl_CurrentAction
            // 
            this.lbl_CurrentAction.AutoSize = true;
            this.lbl_CurrentAction.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_CurrentAction.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_CurrentAction.Location = new System.Drawing.Point(13, 24);
            this.lbl_CurrentAction.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_CurrentAction.Name = "lbl_CurrentAction";
            this.lbl_CurrentAction.Size = new System.Drawing.Size(177, 30);
            this.lbl_CurrentAction.TabIndex = 70;
            this.lbl_CurrentAction.Text = "Current Action :";
            this.lbl_CurrentAction.Click += new System.EventHandler(this.lbl_CurrentAction_Click);
            // 
            // DescriptiveFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1016, 648);
            this.Controls.Add(this.lbl_CurrentAction);
            this.Controls.Add(this.btn_SwicthDescrToComp);
            this.Controls.Add(this.btn_SortTable);
            this.Controls.Add(this.lbl_Nominal);
            this.Controls.Add(this.pic_removeScale);
            this.Controls.Add(this.lbl_Scale);
            this.Controls.Add(this.pic_RemoveNominalList);
            this.Controls.Add(this.btn_Done);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txt_TableName);
            this.Controls.Add(this.pic_removeTableSelected);
            this.Controls.Add(this.pic_addTable);
            this.Controls.Add(this.cmb_TableNames);
            this.Controls.Add(this.list_ViewTableParameters);
            this.Controls.Add(this.pic_AllParaToScale);
            this.Controls.Add(this.pic_AllParaToNominal);
            this.Controls.Add(this.list_Scale);
            this.Controls.Add(this.list_Nominal);
            this.Controls.Add(this.list_AllParameters);
            this.Controls.Add(this.pic_back);
            this.Controls.Add(this.panelmove);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "DescriptiveFrm";
            this.Text = "DescriptiveFrm";
            this.Load += new System.EventHandler(this.DescriptiveFrm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeScale)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToScale)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNominal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.ListBox list_AllParameters;
        private System.Windows.Forms.ListBox list_Nominal;
        private System.Windows.Forms.ListBox list_Scale;
        private System.Windows.Forms.PictureBox pic_AllParaToNominal;
        private System.Windows.Forms.PictureBox pic_AllParaToScale;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_TableName;
        private System.Windows.Forms.PictureBox pic_removeTableSelected;
        private System.Windows.Forms.PictureBox pic_addTable;
        private System.Windows.Forms.ComboBox cmb_TableNames;
        private System.Windows.Forms.ListBox list_ViewTableParameters;
        private System.Windows.Forms.Button btn_Done;
        private System.Windows.Forms.PictureBox pic_RemoveNominalList;
        private System.Windows.Forms.PictureBox pic_removeScale;
        private System.Windows.Forms.Label lbl_Scale;
        private System.Windows.Forms.Label lbl_Nominal;
        private System.Windows.Forms.Button btn_SortTable;
        private System.Windows.Forms.Button btn_SwicthDescrToComp;
        private System.Windows.Forms.Label lbl_CurrentAction;
    }
}