namespace ExcelScore.Forms
{
    partial class CorrelationFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CorrelationFrm));
            this.list_AllParameters = new System.Windows.Forms.ListBox();
            this.panelmove = new System.Windows.Forms.Panel();
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.cmb_CorreType = new System.Windows.Forms.ComboBox();
            this.lbl_corrType = new System.Windows.Forms.Label();
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
            this.lbl_CurrentCorr = new System.Windows.Forms.Label();
            this.btn_Swap = new System.Windows.Forms.Button();
            this.pic_RemoveDependentNormal = new System.Windows.Forms.PictureBox();
            this.lbl_DepNormal = new System.Windows.Forms.Label();
            this.pic_AllParaToDependNormal = new System.Windows.Forms.PictureBox();
            this.list_selectedDepenNormal = new System.Windows.Forms.ListBox();
            this.pic_RemoveDependentAbnormal = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pic_AllParaToDependAbnormal = new System.Windows.Forms.PictureBox();
            this.list_selectedDepenAbnormal = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectedPara)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveDependentNormal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToDependNormal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveDependentAbnormal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToDependAbnormal)).BeginInit();
            this.SuspendLayout();
            // 
            // list_AllParameters
            // 
            this.list_AllParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AllParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllParameters.FormattingEnabled = true;
            this.list_AllParameters.ItemHeight = 17;
            this.list_AllParameters.Location = new System.Drawing.Point(12, 74);
            this.list_AllParameters.Name = "list_AllParameters";
            this.list_AllParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllParameters.Size = new System.Drawing.Size(263, 565);
            this.list_AllParameters.TabIndex = 54;
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(0, -1);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(947, 36);
            this.panelmove.TabIndex = 55;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(945, 8);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(84, 50);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 56;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // cmb_CorreType
            // 
            this.cmb_CorreType.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_CorreType.FormattingEnabled = true;
            this.cmb_CorreType.Items.AddRange(new object[] {
            "Pearson",
            "Spearman"});
            this.cmb_CorreType.Location = new System.Drawing.Point(410, 70);
            this.cmb_CorreType.Name = "cmb_CorreType";
            this.cmb_CorreType.Size = new System.Drawing.Size(183, 30);
            this.cmb_CorreType.TabIndex = 81;
            this.cmb_CorreType.SelectedIndexChanged += new System.EventHandler(this.cmb_CorreType_SelectedIndexChanged);
            // 
            // lbl_corrType
            // 
            this.lbl_corrType.AutoSize = true;
            this.lbl_corrType.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_corrType.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_corrType.Location = new System.Drawing.Point(405, 37);
            this.lbl_corrType.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_corrType.Name = "lbl_corrType";
            this.lbl_corrType.Size = new System.Drawing.Size(75, 30);
            this.lbl_corrType.TabIndex = 80;
            this.lbl_corrType.Text = "Type :";
            // 
            // pic_RemoveSelectedPara
            // 
            this.pic_RemoveSelectedPara.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveSelectedPara.Image")));
            this.pic_RemoveSelectedPara.Location = new System.Drawing.Point(544, 395);
            this.pic_RemoveSelectedPara.Name = "pic_RemoveSelectedPara";
            this.pic_RemoveSelectedPara.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveSelectedPara.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveSelectedPara.TabIndex = 75;
            this.pic_RemoveSelectedPara.TabStop = false;
            this.pic_RemoveSelectedPara.Click += new System.EventHandler(this.pic_RemoveSelectedPara_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label4.Location = new System.Drawing.Point(405, 395);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(143, 30);
            this.label4.TabIndex = 74;
            this.label4.Text = "Parameters :";
            // 
            // pic_AllParaToSelected
            // 
            this.pic_AllParaToSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToSelected.Image")));
            this.pic_AllParaToSelected.Location = new System.Drawing.Point(322, 511);
            this.pic_AllParaToSelected.Name = "pic_AllParaToSelected";
            this.pic_AllParaToSelected.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToSelected.TabIndex = 73;
            this.pic_AllParaToSelected.TabStop = false;
            this.pic_AllParaToSelected.Click += new System.EventHandler(this.pic_AllParaToSelected_Click);
            // 
            // list_selectedParameters
            // 
            this.list_selectedParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_selectedParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_selectedParameters.FormattingEnabled = true;
            this.list_selectedParameters.ItemHeight = 17;
            this.list_selectedParameters.Location = new System.Drawing.Point(410, 428);
            this.list_selectedParameters.Name = "list_selectedParameters";
            this.list_selectedParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_selectedParameters.Size = new System.Drawing.Size(274, 208);
            this.list_selectedParameters.TabIndex = 72;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(773, 99);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(147, 30);
            this.label5.TabIndex = 87;
            this.label5.Text = "Table Name :";
            // 
            // txt_TableName
            // 
            this.txt_TableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_TableName.Location = new System.Drawing.Point(778, 131);
            this.txt_TableName.Margin = new System.Windows.Forms.Padding(2);
            this.txt_TableName.Name = "txt_TableName";
            this.txt_TableName.Size = new System.Drawing.Size(183, 28);
            this.txt_TableName.TabIndex = 86;
            // 
            // pic_removeTableSelected
            // 
            this.pic_removeTableSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeTableSelected.Image")));
            this.pic_removeTableSelected.Location = new System.Drawing.Point(977, 428);
            this.pic_removeTableSelected.Name = "pic_removeTableSelected";
            this.pic_removeTableSelected.Size = new System.Drawing.Size(37, 30);
            this.pic_removeTableSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeTableSelected.TabIndex = 85;
            this.pic_removeTableSelected.TabStop = false;
            this.pic_removeTableSelected.Click += new System.EventHandler(this.pic_removeTableSelected_Click);
            // 
            // pic_addTable
            // 
            this.pic_addTable.Image = ((System.Drawing.Image)(resources.GetObject("pic_addTable.Image")));
            this.pic_addTable.Location = new System.Drawing.Point(977, 131);
            this.pic_addTable.Name = "pic_addTable";
            this.pic_addTable.Size = new System.Drawing.Size(37, 28);
            this.pic_addTable.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_addTable.TabIndex = 84;
            this.pic_addTable.TabStop = false;
            this.pic_addTable.Click += new System.EventHandler(this.pic_addTable_Click);
            // 
            // cmb_TableNames
            // 
            this.cmb_TableNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_TableNames.FormattingEnabled = true;
            this.cmb_TableNames.Location = new System.Drawing.Point(778, 428);
            this.cmb_TableNames.Name = "cmb_TableNames";
            this.cmb_TableNames.Size = new System.Drawing.Size(183, 30);
            this.cmb_TableNames.TabIndex = 83;
            this.cmb_TableNames.SelectedIndexChanged += new System.EventHandler(this.cmb_TableNames_SelectedIndexChanged);
            // 
            // list_ViewTableParameters
            // 
            this.list_ViewTableParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_ViewTableParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_ViewTableParameters.FormattingEnabled = true;
            this.list_ViewTableParameters.ItemHeight = 17;
            this.list_ViewTableParameters.Location = new System.Drawing.Point(778, 197);
            this.list_ViewTableParameters.Name = "list_ViewTableParameters";
            this.list_ViewTableParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_ViewTableParameters.Size = new System.Drawing.Size(183, 208);
            this.list_ViewTableParameters.TabIndex = 82;
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.Black;
            this.btn_Done.Image = ((System.Drawing.Image)(resources.GetObject("btn_Done.Image")));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Done.Location = new System.Drawing.Point(903, 581);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(111, 55);
            this.btn_Done.TabIndex = 88;
            this.btn_Done.Text = "Done";
            this.btn_Done.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // lbl_CurrentCorr
            // 
            this.lbl_CurrentCorr.AutoSize = true;
            this.lbl_CurrentCorr.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_CurrentCorr.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_CurrentCorr.Location = new System.Drawing.Point(11, 41);
            this.lbl_CurrentCorr.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_CurrentCorr.Name = "lbl_CurrentCorr";
            this.lbl_CurrentCorr.Size = new System.Drawing.Size(141, 30);
            this.lbl_CurrentCorr.TabIndex = 89;
            this.lbl_CurrentCorr.Text = "Correlation :";
            // 
            // btn_Swap
            // 
            this.btn_Swap.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Swap.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Swap.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Swap.ForeColor = System.Drawing.Color.Black;
            this.btn_Swap.Image = ((System.Drawing.Image)(resources.GetObject("btn_Swap.Image")));
            this.btn_Swap.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Swap.Location = new System.Drawing.Point(778, 581);
            this.btn_Swap.Name = "btn_Swap";
            this.btn_Swap.Size = new System.Drawing.Size(111, 55);
            this.btn_Swap.TabIndex = 90;
            this.btn_Swap.Text = "Swap";
            this.btn_Swap.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Swap.UseVisualStyleBackColor = false;
            this.btn_Swap.Click += new System.EventHandler(this.btn_Swap_Click);
            // 
            // pic_RemoveDependentNormal
            // 
            this.pic_RemoveDependentNormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveDependentNormal.Image")));
            this.pic_RemoveDependentNormal.Location = new System.Drawing.Point(624, 117);
            this.pic_RemoveDependentNormal.Name = "pic_RemoveDependentNormal";
            this.pic_RemoveDependentNormal.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveDependentNormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveDependentNormal.TabIndex = 94;
            this.pic_RemoveDependentNormal.TabStop = false;
            this.pic_RemoveDependentNormal.Click += new System.EventHandler(this.pic_RemoveDependentNormal_Click);
            // 
            // lbl_DepNormal
            // 
            this.lbl_DepNormal.AutoSize = true;
            this.lbl_DepNormal.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_DepNormal.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_DepNormal.Location = new System.Drawing.Point(405, 117);
            this.lbl_DepNormal.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_DepNormal.Name = "lbl_DepNormal";
            this.lbl_DepNormal.Size = new System.Drawing.Size(223, 30);
            this.lbl_DepNormal.TabIndex = 93;
            this.lbl_DepNormal.Text = "Dependent Normal :";
            // 
            // pic_AllParaToDependNormal
            // 
            this.pic_AllParaToDependNormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToDependNormal.Image")));
            this.pic_AllParaToDependNormal.Location = new System.Drawing.Point(322, 174);
            this.pic_AllParaToDependNormal.Name = "pic_AllParaToDependNormal";
            this.pic_AllParaToDependNormal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToDependNormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToDependNormal.TabIndex = 92;
            this.pic_AllParaToDependNormal.TabStop = false;
            this.pic_AllParaToDependNormal.Click += new System.EventHandler(this.pic_AllParaToDependNormal_Click);
            // 
            // list_selectedDepenNormal
            // 
            this.list_selectedDepenNormal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_selectedDepenNormal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_selectedDepenNormal.FormattingEnabled = true;
            this.list_selectedDepenNormal.ItemHeight = 17;
            this.list_selectedDepenNormal.Location = new System.Drawing.Point(410, 150);
            this.list_selectedDepenNormal.Name = "list_selectedDepenNormal";
            this.list_selectedDepenNormal.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_selectedDepenNormal.Size = new System.Drawing.Size(274, 89);
            this.list_selectedDepenNormal.TabIndex = 91;
            // 
            // pic_RemoveDependentAbnormal
            // 
            this.pic_RemoveDependentAbnormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveDependentAbnormal.Image")));
            this.pic_RemoveDependentAbnormal.Location = new System.Drawing.Point(647, 263);
            this.pic_RemoveDependentAbnormal.Name = "pic_RemoveDependentAbnormal";
            this.pic_RemoveDependentAbnormal.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveDependentAbnormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveDependentAbnormal.TabIndex = 98;
            this.pic_RemoveDependentAbnormal.TabStop = false;
            this.pic_RemoveDependentAbnormal.Click += new System.EventHandler(this.pic_RemoveDependentAbnormal_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(405, 263);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(248, 30);
            this.label1.TabIndex = 97;
            this.label1.Text = "Dependent Abnormal :";
            // 
            // pic_AllParaToDependAbnormal
            // 
            this.pic_AllParaToDependAbnormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToDependAbnormal.Image")));
            this.pic_AllParaToDependAbnormal.Location = new System.Drawing.Point(322, 320);
            this.pic_AllParaToDependAbnormal.Name = "pic_AllParaToDependAbnormal";
            this.pic_AllParaToDependAbnormal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToDependAbnormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToDependAbnormal.TabIndex = 96;
            this.pic_AllParaToDependAbnormal.TabStop = false;
            this.pic_AllParaToDependAbnormal.Click += new System.EventHandler(this.pic_AllParaToDependAbnormal_Click);
            // 
            // list_selectedDepenAbnormal
            // 
            this.list_selectedDepenAbnormal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_selectedDepenAbnormal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_selectedDepenAbnormal.FormattingEnabled = true;
            this.list_selectedDepenAbnormal.ItemHeight = 17;
            this.list_selectedDepenAbnormal.Location = new System.Drawing.Point(410, 296);
            this.list_selectedDepenAbnormal.Name = "list_selectedDepenAbnormal";
            this.list_selectedDepenAbnormal.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_selectedDepenAbnormal.Size = new System.Drawing.Size(274, 89);
            this.list_selectedDepenAbnormal.TabIndex = 95;
            // 
            // CorrelationFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1031, 648);
            this.Controls.Add(this.pic_RemoveDependentAbnormal);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pic_AllParaToDependAbnormal);
            this.Controls.Add(this.list_selectedDepenAbnormal);
            this.Controls.Add(this.pic_RemoveDependentNormal);
            this.Controls.Add(this.lbl_DepNormal);
            this.Controls.Add(this.pic_AllParaToDependNormal);
            this.Controls.Add(this.list_selectedDepenNormal);
            this.Controls.Add(this.btn_Swap);
            this.Controls.Add(this.lbl_CurrentCorr);
            this.Controls.Add(this.btn_Done);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txt_TableName);
            this.Controls.Add(this.pic_removeTableSelected);
            this.Controls.Add(this.pic_addTable);
            this.Controls.Add(this.cmb_TableNames);
            this.Controls.Add(this.list_ViewTableParameters);
            this.Controls.Add(this.cmb_CorreType);
            this.Controls.Add(this.lbl_corrType);
            this.Controls.Add(this.pic_RemoveSelectedPara);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pic_AllParaToSelected);
            this.Controls.Add(this.list_selectedParameters);
            this.Controls.Add(this.pic_back);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.list_AllParameters);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "CorrelationFrm";
            this.Text = "Correlation";
            this.Load += new System.EventHandler(this.Correlation_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectedPara)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveDependentNormal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToDependNormal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveDependentAbnormal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToDependAbnormal)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox list_AllParameters;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.ComboBox cmb_CorreType;
        private System.Windows.Forms.Label lbl_corrType;
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
        private System.Windows.Forms.Label lbl_CurrentCorr;
        private System.Windows.Forms.Button btn_Swap;
        private System.Windows.Forms.PictureBox pic_RemoveDependentNormal;
        private System.Windows.Forms.Label lbl_DepNormal;
        private System.Windows.Forms.PictureBox pic_AllParaToDependNormal;
        private System.Windows.Forms.ListBox list_selectedDepenNormal;
        private System.Windows.Forms.PictureBox pic_RemoveDependentAbnormal;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pic_AllParaToDependAbnormal;
        private System.Windows.Forms.ListBox list_selectedDepenAbnormal;
    }
}