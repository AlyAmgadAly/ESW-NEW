namespace ExcelScore.Forms
{
    partial class RelationsFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RelationsFrm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panelmove = new System.Windows.Forms.Panel();
            this.list_AllParameters = new System.Windows.Forms.ListBox();
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.lbl_DependentType = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_TableName = new System.Windows.Forms.TextBox();
            this.pic_removeTableSelected = new System.Windows.Forms.PictureBox();
            this.pic_addTable = new System.Windows.Forms.PictureBox();
            this.cmb_TableNames = new System.Windows.Forms.ComboBox();
            this.list_ViewTableParameters = new System.Windows.Forms.ListBox();
            this.btn_SortTable = new System.Windows.Forms.Button();
            this.btn_Done = new System.Windows.Forms.Button();
            this.pic_RemoveAbNormalList = new System.Windows.Forms.PictureBox();
            this.pic_RemoveNormalList = new System.Windows.Forms.PictureBox();
            this.pic_RemoveNominalList = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToAbnormal = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToNormal = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.pic_AllParaToNominal = new System.Windows.Forms.PictureBox();
            this.list_AbnormalScale = new System.Windows.Forms.ListBox();
            this.list_NormalScale = new System.Windows.Forms.ListBox();
            this.list_Nominal = new System.Windows.Forms.ListBox();
            this.pic_ifyes = new System.Windows.Forms.PictureBox();
            this.lbl_DependentT = new System.Windows.Forms.Label();
            this.data_allPara = new System.Windows.Forms.DataGridView();
            this.Col_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMeasure = new System.Windows.Forms.DataGridViewImageColumn();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveAbNormalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNormalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToAbnormal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNormal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNominal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ifyes)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).BeginInit();
            this.SuspendLayout();
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(9, 2);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(930, 36);
            this.panelmove.TabIndex = 50;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // list_AllParameters
            // 
            this.list_AllParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AllParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AllParameters.FormattingEnabled = true;
            this.list_AllParameters.ItemHeight = 17;
            this.list_AllParameters.Location = new System.Drawing.Point(9, 77);
            this.list_AllParameters.Name = "list_AllParameters";
            this.list_AllParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AllParameters.Size = new System.Drawing.Size(263, 565);
            this.list_AllParameters.TabIndex = 55;
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(944, 5);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(84, 50);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 57;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // lbl_DependentType
            // 
            this.lbl_DependentType.AutoSize = true;
            this.lbl_DependentType.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_DependentType.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_DependentType.Location = new System.Drawing.Point(146, 46);
            this.lbl_DependentType.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_DependentType.Name = "lbl_DependentType";
            this.lbl_DependentType.Size = new System.Drawing.Size(45, 30);
            this.lbl_DependentType.TabIndex = 61;
            this.lbl_DependentType.Text = "NA";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(751, 63);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(147, 30);
            this.label5.TabIndex = 71;
            this.label5.Text = "Table Name :";
            // 
            // txt_TableName
            // 
            this.txt_TableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_TableName.Location = new System.Drawing.Point(756, 95);
            this.txt_TableName.Margin = new System.Windows.Forms.Padding(2);
            this.txt_TableName.Name = "txt_TableName";
            this.txt_TableName.Size = new System.Drawing.Size(183, 28);
            this.txt_TableName.TabIndex = 70;
            // 
            // pic_removeTableSelected
            // 
            this.pic_removeTableSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeTableSelected.Image")));
            this.pic_removeTableSelected.Location = new System.Drawing.Point(955, 392);
            this.pic_removeTableSelected.Name = "pic_removeTableSelected";
            this.pic_removeTableSelected.Size = new System.Drawing.Size(37, 30);
            this.pic_removeTableSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeTableSelected.TabIndex = 69;
            this.pic_removeTableSelected.TabStop = false;
            this.pic_removeTableSelected.Click += new System.EventHandler(this.pic_removeTableSelected_Click);
            // 
            // pic_addTable
            // 
            this.pic_addTable.Image = ((System.Drawing.Image)(resources.GetObject("pic_addTable.Image")));
            this.pic_addTable.Location = new System.Drawing.Point(955, 95);
            this.pic_addTable.Name = "pic_addTable";
            this.pic_addTable.Size = new System.Drawing.Size(37, 28);
            this.pic_addTable.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_addTable.TabIndex = 68;
            this.pic_addTable.TabStop = false;
            this.pic_addTable.Click += new System.EventHandler(this.pic_addTable_Click);
            // 
            // cmb_TableNames
            // 
            this.cmb_TableNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_TableNames.FormattingEnabled = true;
            this.cmb_TableNames.Location = new System.Drawing.Point(756, 392);
            this.cmb_TableNames.Name = "cmb_TableNames";
            this.cmb_TableNames.Size = new System.Drawing.Size(183, 30);
            this.cmb_TableNames.TabIndex = 67;
            // 
            // list_ViewTableParameters
            // 
            this.list_ViewTableParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_ViewTableParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_ViewTableParameters.FormattingEnabled = true;
            this.list_ViewTableParameters.ItemHeight = 17;
            this.list_ViewTableParameters.Location = new System.Drawing.Point(756, 161);
            this.list_ViewTableParameters.Name = "list_ViewTableParameters";
            this.list_ViewTableParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_ViewTableParameters.Size = new System.Drawing.Size(183, 208);
            this.list_ViewTableParameters.TabIndex = 66;
            // 
            // btn_SortTable
            // 
            this.btn_SortTable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_SortTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_SortTable.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_SortTable.ForeColor = System.Drawing.Color.Black;
            this.btn_SortTable.Image = ((System.Drawing.Image)(resources.GetObject("btn_SortTable.Image")));
            this.btn_SortTable.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_SortTable.Location = new System.Drawing.Point(756, 560);
            this.btn_SortTable.Name = "btn_SortTable";
            this.btn_SortTable.Size = new System.Drawing.Size(102, 55);
            this.btn_SortTable.TabIndex = 73;
            this.btn_SortTable.Text = "Sort";
            this.btn_SortTable.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_SortTable.UseVisualStyleBackColor = false;
            this.btn_SortTable.Click += new System.EventHandler(this.btn_SortTable_Click);
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.Black;
            this.btn_Done.Image = ((System.Drawing.Image)(resources.GetObject("btn_Done.Image")));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Done.Location = new System.Drawing.Point(891, 560);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(101, 55);
            this.btn_Done.TabIndex = 72;
            this.btn_Done.Text = "Done";
            this.btn_Done.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // pic_RemoveAbNormalList
            // 
            this.pic_RemoveAbNormalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveAbNormalList.Image")));
            this.pic_RemoveAbNormalList.Location = new System.Drawing.Point(587, 395);
            this.pic_RemoveAbNormalList.Name = "pic_RemoveAbNormalList";
            this.pic_RemoveAbNormalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveAbNormalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveAbNormalList.TabIndex = 85;
            this.pic_RemoveAbNormalList.TabStop = false;
            this.pic_RemoveAbNormalList.Click += new System.EventHandler(this.pic_RemoveAbNormalList_Click_1);
            // 
            // pic_RemoveNormalList
            // 
            this.pic_RemoveNormalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveNormalList.Image")));
            this.pic_RemoveNormalList.Location = new System.Drawing.Point(562, 215);
            this.pic_RemoveNormalList.Name = "pic_RemoveNormalList";
            this.pic_RemoveNormalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveNormalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveNormalList.TabIndex = 84;
            this.pic_RemoveNormalList.TabStop = false;
            this.pic_RemoveNormalList.Click += new System.EventHandler(this.pic_RemoveNormalList_Click_1);
            // 
            // pic_RemoveNominalList
            // 
            this.pic_RemoveNominalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveNominalList.Image")));
            this.pic_RemoveNominalList.Location = new System.Drawing.Point(513, 46);
            this.pic_RemoveNominalList.Name = "pic_RemoveNominalList";
            this.pic_RemoveNominalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveNominalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveNominalList.TabIndex = 83;
            this.pic_RemoveNominalList.TabStop = false;
            this.pic_RemoveNominalList.Click += new System.EventHandler(this.pic_RemoveNominalList_Click_1);
            // 
            // pic_AllParaToAbnormal
            // 
            this.pic_AllParaToAbnormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToAbnormal.Image")));
            this.pic_AllParaToAbnormal.Location = new System.Drawing.Point(306, 457);
            this.pic_AllParaToAbnormal.Name = "pic_AllParaToAbnormal";
            this.pic_AllParaToAbnormal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToAbnormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToAbnormal.TabIndex = 82;
            this.pic_AllParaToAbnormal.TabStop = false;
            this.pic_AllParaToAbnormal.Click += new System.EventHandler(this.pic_AllParaToAbnormal_Click_1);
            // 
            // pic_AllParaToNormal
            // 
            this.pic_AllParaToNormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToNormal.Image")));
            this.pic_AllParaToNormal.Location = new System.Drawing.Point(306, 280);
            this.pic_AllParaToNormal.Name = "pic_AllParaToNormal";
            this.pic_AllParaToNormal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToNormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToNormal.TabIndex = 81;
            this.pic_AllParaToNormal.TabStop = false;
            this.pic_AllParaToNormal.Click += new System.EventHandler(this.pic_AllParaToNormal_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label2.Location = new System.Drawing.Point(396, 395);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(186, 30);
            this.label2.TabIndex = 80;
            this.label2.Text = "Abnormal Scale :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(396, 215);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(161, 30);
            this.label1.TabIndex = 79;
            this.label1.Text = "Normal Scale :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label4.Location = new System.Drawing.Point(396, 46);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(112, 30);
            this.label4.TabIndex = 78;
            this.label4.Text = "Nominal :";
            // 
            // pic_AllParaToNominal
            // 
            this.pic_AllParaToNominal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToNominal.Image")));
            this.pic_AllParaToNominal.Location = new System.Drawing.Point(306, 110);
            this.pic_AllParaToNominal.Name = "pic_AllParaToNominal";
            this.pic_AllParaToNominal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToNominal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToNominal.TabIndex = 77;
            this.pic_AllParaToNominal.TabStop = false;
            this.pic_AllParaToNominal.Click += new System.EventHandler(this.pic_AllParaToNominal_Click);
            // 
            // list_AbnormalScale
            // 
            this.list_AbnormalScale.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_AbnormalScale.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_AbnormalScale.FormattingEnabled = true;
            this.list_AbnormalScale.ItemHeight = 17;
            this.list_AbnormalScale.Location = new System.Drawing.Point(401, 428);
            this.list_AbnormalScale.Name = "list_AbnormalScale";
            this.list_AbnormalScale.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_AbnormalScale.Size = new System.Drawing.Size(263, 106);
            this.list_AbnormalScale.TabIndex = 76;
            // 
            // list_NormalScale
            // 
            this.list_NormalScale.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_NormalScale.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_NormalScale.FormattingEnabled = true;
            this.list_NormalScale.ItemHeight = 17;
            this.list_NormalScale.Location = new System.Drawing.Point(401, 248);
            this.list_NormalScale.Name = "list_NormalScale";
            this.list_NormalScale.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_NormalScale.Size = new System.Drawing.Size(263, 106);
            this.list_NormalScale.TabIndex = 75;
            // 
            // list_Nominal
            // 
            this.list_Nominal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Nominal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Nominal.FormattingEnabled = true;
            this.list_Nominal.ItemHeight = 17;
            this.list_Nominal.Location = new System.Drawing.Point(401, 79);
            this.list_Nominal.Name = "list_Nominal";
            this.list_Nominal.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Nominal.Size = new System.Drawing.Size(263, 106);
            this.list_Nominal.TabIndex = 74;
            // 
            // pic_ifyes
            // 
            this.pic_ifyes.Image = ((System.Drawing.Image)(resources.GetObject("pic_ifyes.Image")));
            this.pic_ifyes.Location = new System.Drawing.Point(670, 79);
            this.pic_ifyes.Name = "pic_ifyes";
            this.pic_ifyes.Size = new System.Drawing.Size(37, 30);
            this.pic_ifyes.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_ifyes.TabIndex = 86;
            this.pic_ifyes.TabStop = false;
            this.pic_ifyes.Click += new System.EventHandler(this.pic_ifyes_Click_1);
            // 
            // lbl_DependentT
            // 
            this.lbl_DependentT.AutoSize = true;
            this.lbl_DependentT.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_DependentT.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_DependentT.Location = new System.Drawing.Point(11, 44);
            this.lbl_DependentT.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_DependentT.Name = "lbl_DependentT";
            this.lbl_DependentT.Size = new System.Drawing.Size(140, 30);
            this.lbl_DependentT.TabIndex = 87;
            this.lbl_DependentT.Text = "Dependent :";
            // 
            // data_allPara
            // 
            this.data_allPara.AllowUserToAddRows = false;
            this.data_allPara.AllowUserToDeleteRows = false;
            this.data_allPara.AllowUserToResizeColumns = false;
            this.data_allPara.AllowUserToResizeRows = false;
            this.data_allPara.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.data_allPara.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.data_allPara.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.data_allPara.ColumnHeadersHeight = 54;
            this.data_allPara.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.data_allPara.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Col_Name,
            this.ColMeasure});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.data_allPara.DefaultCellStyle = dataGridViewCellStyle2;
            this.data_allPara.EnableHeadersVisualStyles = false;
            this.data_allPara.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.Location = new System.Drawing.Point(9, 77);
            this.data_allPara.Name = "data_allPara";
            this.data_allPara.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.data_allPara.RowHeadersVisible = false;
            this.data_allPara.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.data_allPara.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.data_allPara.Size = new System.Drawing.Size(289, 565);
            this.data_allPara.TabIndex = 88;
            // 
            // Col_Name
            // 
            this.Col_Name.HeaderText = "Name";
            this.Col_Name.Name = "Col_Name";
            this.Col_Name.Width = 175;
            // 
            // ColMeasure
            // 
            this.ColMeasure.HeaderText = "Measure";
            this.ColMeasure.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom;
            this.ColMeasure.Name = "ColMeasure";
            // 
            // RelationsFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1031, 648);
            this.Controls.Add(this.data_allPara);
            this.Controls.Add(this.lbl_DependentT);
            this.Controls.Add(this.pic_ifyes);
            this.Controls.Add(this.pic_RemoveAbNormalList);
            this.Controls.Add(this.pic_RemoveNormalList);
            this.Controls.Add(this.pic_RemoveNominalList);
            this.Controls.Add(this.pic_AllParaToAbnormal);
            this.Controls.Add(this.pic_AllParaToNormal);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pic_AllParaToNominal);
            this.Controls.Add(this.list_AbnormalScale);
            this.Controls.Add(this.list_NormalScale);
            this.Controls.Add(this.list_Nominal);
            this.Controls.Add(this.btn_SortTable);
            this.Controls.Add(this.btn_Done);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txt_TableName);
            this.Controls.Add(this.pic_removeTableSelected);
            this.Controls.Add(this.pic_addTable);
            this.Controls.Add(this.cmb_TableNames);
            this.Controls.Add(this.list_ViewTableParameters);
            this.Controls.Add(this.lbl_DependentType);
            this.Controls.Add(this.pic_back);
            this.Controls.Add(this.list_AllParameters);
            this.Controls.Add(this.panelmove);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "RelationsFrm";
            this.Text = "RelationsFrm";
            this.Load += new System.EventHandler(this.RelationsFrm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveAbNormalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNormalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToAbnormal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNormal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNominal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ifyes)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.ListBox list_AllParameters;
        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.Label lbl_DependentType;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_TableName;
        private System.Windows.Forms.PictureBox pic_removeTableSelected;
        private System.Windows.Forms.PictureBox pic_addTable;
        private System.Windows.Forms.ComboBox cmb_TableNames;
        private System.Windows.Forms.ListBox list_ViewTableParameters;
        private System.Windows.Forms.Button btn_SortTable;
        private System.Windows.Forms.Button btn_Done;
        private System.Windows.Forms.PictureBox pic_RemoveAbNormalList;
        private System.Windows.Forms.PictureBox pic_RemoveNormalList;
        private System.Windows.Forms.PictureBox pic_RemoveNominalList;
        private System.Windows.Forms.PictureBox pic_AllParaToAbnormal;
        private System.Windows.Forms.PictureBox pic_AllParaToNormal;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.PictureBox pic_AllParaToNominal;
        private System.Windows.Forms.ListBox list_AbnormalScale;
        private System.Windows.Forms.ListBox list_NormalScale;
        private System.Windows.Forms.ListBox list_Nominal;
        private System.Windows.Forms.PictureBox pic_ifyes;
        private System.Windows.Forms.Label lbl_DependentT;
        private System.Windows.Forms.DataGridView data_allPara;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Name;
        private System.Windows.Forms.DataGridViewImageColumn ColMeasure;
    }
}