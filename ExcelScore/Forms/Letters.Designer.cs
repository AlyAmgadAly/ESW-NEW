namespace ExcelScore.Forms
{
    partial class Letters
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Letters));
            this.panelmove = new System.Windows.Forms.Panel();
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.list_AllParameters = new System.Windows.Forms.ListBox();
            this.pic_RemoveSelectedPara = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.pic_AllParaToSelected = new System.Windows.Forms.PictureBox();
            this.list_selectedParameters = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_TableName = new System.Windows.Forms.TextBox();
            this.pic_removeTableSelected = new System.Windows.Forms.PictureBox();
            this.pic_addTable = new System.Windows.Forms.PictureBox();
            this.cmb_TableNames = new System.Windows.Forms.ComboBox();
            this.list_ViewTableParameters = new System.Windows.Forms.ListBox();
            this.btn_Done = new System.Windows.Forms.Button();
            this.pic_removeGroup = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.list_Groups = new System.Windows.Forms.ListBox();
            this.pic_AllParaToGroups = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmb_letterType = new System.Windows.Forms.ComboBox();
            this.btn_ChooseTable = new System.Windows.Forms.Button();
            this.pic_RemoveSelectPara = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToSelect = new System.Windows.Forms.PictureBox();
            this.list_Select = new System.Windows.Forms.ListBox();
            this.lbl_Select = new System.Windows.Forms.Label();
            this.pic_groups_select = new System.Windows.Forms.PictureBox();
            this.cmb_NormalORAb = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectedPara)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeGroup)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToGroups)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectPara)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelect)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_groups_select)).BeginInit();
            this.SuspendLayout();
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(-1, 0);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(947, 36);
            this.panelmove.TabIndex = 50;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(943, 6);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(84, 50);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 52;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // list_AllParameters
            // 
            this.list_AllParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AllParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllParameters.FormattingEnabled = true;
            this.list_AllParameters.ItemHeight = 17;
            this.list_AllParameters.Location = new System.Drawing.Point(12, 41);
            this.list_AllParameters.Name = "list_AllParameters";
            this.list_AllParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllParameters.Size = new System.Drawing.Size(263, 599);
            this.list_AllParameters.TabIndex = 53;
            // 
            // pic_RemoveSelectedPara
            // 
            this.pic_RemoveSelectedPara.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveSelectedPara.Image")));
            this.pic_RemoveSelectedPara.Location = new System.Drawing.Point(552, 110);
            this.pic_RemoveSelectedPara.Name = "pic_RemoveSelectedPara";
            this.pic_RemoveSelectedPara.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveSelectedPara.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveSelectedPara.TabIndex = 57;
            this.pic_RemoveSelectedPara.TabStop = false;
            this.pic_RemoveSelectedPara.Click += new System.EventHandler(this.pic_RemoveSelectedPara_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label4.Location = new System.Drawing.Point(413, 110);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(143, 30);
            this.label4.TabIndex = 56;
            this.label4.Text = "Parameters :";
            // 
            // pic_AllParaToSelected
            // 
            this.pic_AllParaToSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToSelected.Image")));
            this.pic_AllParaToSelected.Location = new System.Drawing.Point(319, 227);
            this.pic_AllParaToSelected.Name = "pic_AllParaToSelected";
            this.pic_AllParaToSelected.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToSelected.TabIndex = 55;
            this.pic_AllParaToSelected.TabStop = false;
            this.pic_AllParaToSelected.Click += new System.EventHandler(this.pic_AllParaToSelected_Click);
            // 
            // list_selectedParameters
            // 
            this.list_selectedParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_selectedParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_selectedParameters.FormattingEnabled = true;
            this.list_selectedParameters.ItemHeight = 17;
            this.list_selectedParameters.Location = new System.Drawing.Point(418, 143);
            this.list_selectedParameters.Name = "list_selectedParameters";
            this.list_selectedParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_selectedParameters.Size = new System.Drawing.Size(274, 293);
            this.list_selectedParameters.TabIndex = 54;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(758, 83);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(147, 30);
            this.label5.TabIndex = 63;
            this.label5.Text = "Table Name :";
            // 
            // txt_TableName
            // 
            this.txt_TableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_TableName.Location = new System.Drawing.Point(763, 115);
            this.txt_TableName.Margin = new System.Windows.Forms.Padding(2);
            this.txt_TableName.Name = "txt_TableName";
            this.txt_TableName.Size = new System.Drawing.Size(183, 28);
            this.txt_TableName.TabIndex = 62;
            // 
            // pic_removeTableSelected
            // 
            this.pic_removeTableSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeTableSelected.Image")));
            this.pic_removeTableSelected.Location = new System.Drawing.Point(962, 412);
            this.pic_removeTableSelected.Name = "pic_removeTableSelected";
            this.pic_removeTableSelected.Size = new System.Drawing.Size(37, 30);
            this.pic_removeTableSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeTableSelected.TabIndex = 61;
            this.pic_removeTableSelected.TabStop = false;
            this.pic_removeTableSelected.Click += new System.EventHandler(this.pic_removeTableSelected_Click);
            // 
            // pic_addTable
            // 
            this.pic_addTable.Image = ((System.Drawing.Image)(resources.GetObject("pic_addTable.Image")));
            this.pic_addTable.Location = new System.Drawing.Point(962, 115);
            this.pic_addTable.Name = "pic_addTable";
            this.pic_addTable.Size = new System.Drawing.Size(37, 28);
            this.pic_addTable.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_addTable.TabIndex = 60;
            this.pic_addTable.TabStop = false;
            this.pic_addTable.Click += new System.EventHandler(this.pic_addTable_Click);
            // 
            // cmb_TableNames
            // 
            this.cmb_TableNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_TableNames.FormattingEnabled = true;
            this.cmb_TableNames.Location = new System.Drawing.Point(763, 412);
            this.cmb_TableNames.Name = "cmb_TableNames";
            this.cmb_TableNames.Size = new System.Drawing.Size(183, 30);
            this.cmb_TableNames.TabIndex = 59;
            this.cmb_TableNames.SelectedIndexChanged += new System.EventHandler(this.cmb_TableNames_SelectedIndexChanged);
            // 
            // list_ViewTableParameters
            // 
            this.list_ViewTableParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_ViewTableParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_ViewTableParameters.FormattingEnabled = true;
            this.list_ViewTableParameters.ItemHeight = 17;
            this.list_ViewTableParameters.Location = new System.Drawing.Point(763, 181);
            this.list_ViewTableParameters.Name = "list_ViewTableParameters";
            this.list_ViewTableParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_ViewTableParameters.Size = new System.Drawing.Size(183, 208);
            this.list_ViewTableParameters.TabIndex = 58;
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.Black;
            this.btn_Done.Image = ((System.Drawing.Image)(resources.GetObject("btn_Done.Image")));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Done.Location = new System.Drawing.Point(888, 581);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(111, 55);
            this.btn_Done.TabIndex = 64;
            this.btn_Done.Text = "Done";
            this.btn_Done.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // pic_removeGroup
            // 
            this.pic_removeGroup.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeGroup.Image")));
            this.pic_removeGroup.Location = new System.Drawing.Point(518, 462);
            this.pic_removeGroup.Name = "pic_removeGroup";
            this.pic_removeGroup.Size = new System.Drawing.Size(37, 30);
            this.pic_removeGroup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeGroup.TabIndex = 67;
            this.pic_removeGroup.TabStop = false;
            this.pic_removeGroup.Click += new System.EventHandler(this.pic_removeGroup_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(413, 462);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 30);
            this.label1.TabIndex = 66;
            this.label1.Text = "Groups :";
            // 
            // list_Groups
            // 
            this.list_Groups.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Groups.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Groups.FormattingEnabled = true;
            this.list_Groups.ItemHeight = 17;
            this.list_Groups.Location = new System.Drawing.Point(418, 495);
            this.list_Groups.Name = "list_Groups";
            this.list_Groups.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Groups.Size = new System.Drawing.Size(274, 38);
            this.list_Groups.TabIndex = 65;
            // 
            // pic_AllParaToGroups
            // 
            this.pic_AllParaToGroups.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToGroups.Image")));
            this.pic_AllParaToGroups.Location = new System.Drawing.Point(319, 495);
            this.pic_AllParaToGroups.Name = "pic_AllParaToGroups";
            this.pic_AllParaToGroups.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToGroups.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToGroups.TabIndex = 68;
            this.pic_AllParaToGroups.TabStop = false;
            this.pic_AllParaToGroups.Click += new System.EventHandler(this.pic_AllParaToGroups_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label2.Location = new System.Drawing.Point(413, 38);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 30);
            this.label2.TabIndex = 70;
            this.label2.Text = "Type :";
            // 
            // cmb_letterType
            // 
            this.cmb_letterType.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_letterType.FormattingEnabled = true;
            this.cmb_letterType.Items.AddRange(new object[] {
            "Easy",
            "Hard"});
            this.cmb_letterType.Location = new System.Drawing.Point(418, 71);
            this.cmb_letterType.Name = "cmb_letterType";
            this.cmb_letterType.Size = new System.Drawing.Size(183, 30);
            this.cmb_letterType.TabIndex = 71;
            // 
            // btn_ChooseTable
            // 
            this.btn_ChooseTable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_ChooseTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_ChooseTable.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_ChooseTable.ForeColor = System.Drawing.Color.Black;
            this.btn_ChooseTable.Image = ((System.Drawing.Image)(resources.GetObject("btn_ChooseTable.Image")));
            this.btn_ChooseTable.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_ChooseTable.Location = new System.Drawing.Point(888, 504);
            this.btn_ChooseTable.Name = "btn_ChooseTable";
            this.btn_ChooseTable.Size = new System.Drawing.Size(111, 55);
            this.btn_ChooseTable.TabIndex = 72;
            this.btn_ChooseTable.Text = "Table";
            this.btn_ChooseTable.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_ChooseTable.UseVisualStyleBackColor = false;
            this.btn_ChooseTable.Click += new System.EventHandler(this.btn_ChooseTable_Click);
            // 
            // pic_RemoveSelectPara
            // 
            this.pic_RemoveSelectPara.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveSelectPara.Image")));
            this.pic_RemoveSelectPara.Location = new System.Drawing.Point(562, 548);
            this.pic_RemoveSelectPara.Name = "pic_RemoveSelectPara";
            this.pic_RemoveSelectPara.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveSelectPara.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveSelectPara.TabIndex = 77;
            this.pic_RemoveSelectPara.TabStop = false;
            this.pic_RemoveSelectPara.Click += new System.EventHandler(this.pic_RemoveSelectPara_Click);
            // 
            // pic_AllParaToSelect
            // 
            this.pic_AllParaToSelect.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToSelect.Image")));
            this.pic_AllParaToSelect.Location = new System.Drawing.Point(319, 581);
            this.pic_AllParaToSelect.Name = "pic_AllParaToSelect";
            this.pic_AllParaToSelect.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToSelect.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToSelect.TabIndex = 76;
            this.pic_AllParaToSelect.TabStop = false;
            this.pic_AllParaToSelect.Click += new System.EventHandler(this.pic_AllParaToSelect_Click);
            // 
            // list_Select
            // 
            this.list_Select.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Select.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Select.FormattingEnabled = true;
            this.list_Select.ItemHeight = 17;
            this.list_Select.Location = new System.Drawing.Point(418, 581);
            this.list_Select.Name = "list_Select";
            this.list_Select.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.list_Select.Size = new System.Drawing.Size(274, 38);
            this.list_Select.TabIndex = 75;
            // 
            // lbl_Select
            // 
            this.lbl_Select.AutoSize = true;
            this.lbl_Select.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Select.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Select.Location = new System.Drawing.Point(413, 548);
            this.lbl_Select.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Select.Name = "lbl_Select";
            this.lbl_Select.Size = new System.Drawing.Size(87, 30);
            this.lbl_Select.TabIndex = 74;
            this.lbl_Select.Text = "Select :";
            // 
            // pic_groups_select
            // 
            this.pic_groups_select.Image = ((System.Drawing.Image)(resources.GetObject("pic_groups_select.Image")));
            this.pic_groups_select.Location = new System.Drawing.Point(518, 548);
            this.pic_groups_select.Name = "pic_groups_select";
            this.pic_groups_select.Size = new System.Drawing.Size(37, 30);
            this.pic_groups_select.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_groups_select.TabIndex = 73;
            this.pic_groups_select.TabStop = false;
            this.pic_groups_select.Click += new System.EventHandler(this.pic_groups_select_Click);
            // 
            // cmb_NormalORAb
            // 
            this.cmb_NormalORAb.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_NormalORAb.FormattingEnabled = true;
            this.cmb_NormalORAb.Items.AddRange(new object[] {
            "Normal",
            "Not Normal"});
            this.cmb_NormalORAb.Location = new System.Drawing.Point(607, 71);
            this.cmb_NormalORAb.Name = "cmb_NormalORAb";
            this.cmb_NormalORAb.Size = new System.Drawing.Size(146, 30);
            this.cmb_NormalORAb.TabIndex = 78;
            // 
            // Letters
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1031, 648);
            this.Controls.Add(this.cmb_NormalORAb);
            this.Controls.Add(this.pic_RemoveSelectPara);
            this.Controls.Add(this.pic_AllParaToSelect);
            this.Controls.Add(this.list_Select);
            this.Controls.Add(this.lbl_Select);
            this.Controls.Add(this.pic_groups_select);
            this.Controls.Add(this.btn_ChooseTable);
            this.Controls.Add(this.cmb_letterType);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pic_AllParaToGroups);
            this.Controls.Add(this.pic_removeGroup);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.list_Groups);
            this.Controls.Add(this.btn_Done);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txt_TableName);
            this.Controls.Add(this.pic_removeTableSelected);
            this.Controls.Add(this.pic_addTable);
            this.Controls.Add(this.cmb_TableNames);
            this.Controls.Add(this.list_ViewTableParameters);
            this.Controls.Add(this.pic_RemoveSelectedPara);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pic_AllParaToSelected);
            this.Controls.Add(this.list_selectedParameters);
            this.Controls.Add(this.list_AllParameters);
            this.Controls.Add(this.pic_back);
            this.Controls.Add(this.panelmove);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Letters";
            this.Text = "Letters";
            this.Load += new System.EventHandler(this.Letters_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectedPara)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeGroup)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToGroups)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectPara)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelect)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_groups_select)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.ListBox list_AllParameters;
        private System.Windows.Forms.PictureBox pic_RemoveSelectedPara;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.PictureBox pic_AllParaToSelected;
        private System.Windows.Forms.ListBox list_selectedParameters;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_TableName;
        private System.Windows.Forms.PictureBox pic_removeTableSelected;
        private System.Windows.Forms.PictureBox pic_addTable;
        private System.Windows.Forms.ComboBox cmb_TableNames;
        private System.Windows.Forms.ListBox list_ViewTableParameters;
        private System.Windows.Forms.Button btn_Done;
        private System.Windows.Forms.PictureBox pic_removeGroup;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox list_Groups;
        private System.Windows.Forms.PictureBox pic_AllParaToGroups;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmb_letterType;
        private System.Windows.Forms.Button btn_ChooseTable;
        private System.Windows.Forms.PictureBox pic_RemoveSelectPara;
        private System.Windows.Forms.PictureBox pic_AllParaToSelect;
        private System.Windows.Forms.ListBox list_Select;
        private System.Windows.Forms.Label lbl_Select;
        private System.Windows.Forms.PictureBox pic_groups_select;
        private System.Windows.Forms.ComboBox cmb_NormalORAb;
    }
}