namespace ExcelScore.Forms
{
    partial class Regression_Frm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Regression_Frm));
            this.txt_ParaName = new System.Windows.Forms.TextBox();
            this.data_allPara = new System.Windows.Forms.DataGridView();
            this.Col_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMeasure = new System.Windows.Forms.DataGridViewImageColumn();
            this.ColNormality = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pic_back = new System.Windows.Forms.PictureBox();
            this.panelmove = new System.Windows.Forms.Panel();
            this.pic_ClearAbnormalList = new System.Windows.Forms.PictureBox();
            this.pic_clearNormalList = new System.Windows.Forms.PictureBox();
            this.pic_ClearNominalList = new System.Windows.Forms.PictureBox();
            this.pic_AllLists = new System.Windows.Forms.PictureBox();
            this.lbl_Periods = new System.Windows.Forms.Label();
            this.pic_RemoveAbNormalList = new System.Windows.Forms.PictureBox();
            this.pic_RemoveNormalList = new System.Windows.Forms.PictureBox();
            this.pic_RemoveNominalList = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToAbnormal = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToNormal = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lbl_Nominal = new System.Windows.Forms.Label();
            this.pic_AllParaToNominal = new System.Windows.Forms.PictureBox();
            this.list_Seperated = new System.Windows.Forms.ListBox();
            this.list_NotSeperatedScale = new System.Windows.Forms.ListBox();
            this.list_Dependent = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_TableName = new System.Windows.Forms.TextBox();
            this.pic_removeTableSelected = new System.Windows.Forms.PictureBox();
            this.pic_addTable = new System.Windows.Forms.PictureBox();
            this.cmb_TableNames = new System.Windows.Forms.ComboBox();
            this.list_ViewTableParameters = new System.Windows.Forms.ListBox();
            this.btn_Update = new System.Windows.Forms.Button();
            this.btn_SortTable = new System.Windows.Forms.Button();
            this.btn_Done = new System.Windows.Forms.Button();
            this.cmb_ChooseTableFormat = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.pic_RemoveSelectPara = new System.Windows.Forms.PictureBox();
            this.pic_AllParaToSelect = new System.Windows.Forms.PictureBox();
            this.list_Select = new System.Windows.Forms.ListBox();
            this.lbl_Select = new System.Windows.Forms.Label();
            this.pic_groups_select = new System.Windows.Forms.PictureBox();
            this.pic_clearNotSeperatedNominal = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.list_NotSeperatedNominal = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ClearAbnormalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_clearNormalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ClearNominalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllLists)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveAbNormalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNormalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToAbnormal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNormal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNominal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectPara)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelect)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_groups_select)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_clearNotSeperatedNominal)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.SuspendLayout();
            // 
            // txt_ParaName
            // 
            this.txt_ParaName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_ParaName.Location = new System.Drawing.Point(54, 82);
            this.txt_ParaName.Margin = new System.Windows.Forms.Padding(2);
            this.txt_ParaName.Name = "txt_ParaName";
            this.txt_ParaName.Size = new System.Drawing.Size(96, 28);
            this.txt_ParaName.TabIndex = 69;
            this.txt_ParaName.TextChanged += new System.EventHandler(this.txt_ParaName_TextChanged);
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
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.data_allPara.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.data_allPara.ColumnHeadersHeight = 54;
            this.data_allPara.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.data_allPara.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Col_Name,
            this.ColMeasure,
            this.ColNormality});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.data_allPara.DefaultCellStyle = dataGridViewCellStyle4;
            this.data_allPara.EnableHeadersVisualStyles = false;
            this.data_allPara.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.data_allPara.Location = new System.Drawing.Point(10, 69);
            this.data_allPara.Name = "data_allPara";
            this.data_allPara.ReadOnly = true;
            this.data_allPara.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.data_allPara.RowHeadersVisible = false;
            this.data_allPara.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.data_allPara.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.data_allPara.Size = new System.Drawing.Size(334, 556);
            this.data_allPara.TabIndex = 68;
            // 
            // Col_Name
            // 
            this.Col_Name.DataPropertyName = "ColName";
            this.Col_Name.HeaderText = "Name";
            this.Col_Name.Name = "Col_Name";
            this.Col_Name.ReadOnly = true;
            this.Col_Name.Width = 150;
            // 
            // ColMeasure
            // 
            this.ColMeasure.HeaderText = "Measure";
            this.ColMeasure.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom;
            this.ColMeasure.Name = "ColMeasure";
            this.ColMeasure.ReadOnly = true;
            this.ColMeasure.Width = 75;
            // 
            // ColNormality
            // 
            this.ColNormality.HeaderText = "Normality";
            this.ColNormality.Name = "ColNormality";
            this.ColNormality.ReadOnly = true;
            this.ColNormality.Width = 75;
            // 
            // pic_back
            // 
            this.pic_back.Image = ((System.Drawing.Image)(resources.GetObject("pic_back.Image")));
            this.pic_back.Location = new System.Drawing.Point(925, 5);
            this.pic_back.Name = "pic_back";
            this.pic_back.Size = new System.Drawing.Size(86, 50);
            this.pic_back.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_back.TabIndex = 70;
            this.pic_back.TabStop = false;
            this.pic_back.Click += new System.EventHandler(this.pic_back_Click);
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(13, 5);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(907, 27);
            this.panelmove.TabIndex = 71;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // pic_ClearAbnormalList
            // 
            this.pic_ClearAbnormalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_ClearAbnormalList.Image")));
            this.pic_ClearAbnormalList.Location = new System.Drawing.Point(614, 410);
            this.pic_ClearAbnormalList.Name = "pic_ClearAbnormalList";
            this.pic_ClearAbnormalList.Size = new System.Drawing.Size(45, 31);
            this.pic_ClearAbnormalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_ClearAbnormalList.TabIndex = 99;
            this.pic_ClearAbnormalList.TabStop = false;
            this.pic_ClearAbnormalList.Click += new System.EventHandler(this.pic_ClearAbnormalList_Click);
            // 
            // pic_clearNormalList
            // 
            this.pic_clearNormalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_clearNormalList.Image")));
            this.pic_clearNormalList.Location = new System.Drawing.Point(563, 135);
            this.pic_clearNormalList.Name = "pic_clearNormalList";
            this.pic_clearNormalList.Size = new System.Drawing.Size(45, 31);
            this.pic_clearNormalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_clearNormalList.TabIndex = 98;
            this.pic_clearNormalList.TabStop = false;
            this.pic_clearNormalList.Click += new System.EventHandler(this.pic_clearNormalList_Click);
            // 
            // pic_ClearNominalList
            // 
            this.pic_ClearNominalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_ClearNominalList.Image")));
            this.pic_ClearNominalList.Location = new System.Drawing.Point(617, 46);
            this.pic_ClearNominalList.Name = "pic_ClearNominalList";
            this.pic_ClearNominalList.Size = new System.Drawing.Size(45, 31);
            this.pic_ClearNominalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_ClearNominalList.TabIndex = 97;
            this.pic_ClearNominalList.TabStop = false;
            this.pic_ClearNominalList.Click += new System.EventHandler(this.pic_ClearNominalList_Click);
            // 
            // pic_AllLists
            // 
            this.pic_AllLists.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllLists.Image")));
            this.pic_AllLists.Location = new System.Drawing.Point(374, 43);
            this.pic_AllLists.Name = "pic_AllLists";
            this.pic_AllLists.Size = new System.Drawing.Size(45, 31);
            this.pic_AllLists.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllLists.TabIndex = 96;
            this.pic_AllLists.TabStop = false;
            this.pic_AllLists.Click += new System.EventHandler(this.pic_AllLists_Click);
            // 
            // lbl_Periods
            // 
            this.lbl_Periods.AutoSize = true;
            this.lbl_Periods.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Periods.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Periods.Location = new System.Drawing.Point(405, 69);
            this.lbl_Periods.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Periods.Name = "lbl_Periods";
            this.lbl_Periods.Size = new System.Drawing.Size(0, 30);
            this.lbl_Periods.TabIndex = 95;
            // 
            // pic_RemoveAbNormalList
            // 
            this.pic_RemoveAbNormalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveAbNormalList.Image")));
            this.pic_RemoveAbNormalList.Location = new System.Drawing.Point(573, 410);
            this.pic_RemoveAbNormalList.Name = "pic_RemoveAbNormalList";
            this.pic_RemoveAbNormalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveAbNormalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveAbNormalList.TabIndex = 88;
            this.pic_RemoveAbNormalList.TabStop = false;
            this.pic_RemoveAbNormalList.Click += new System.EventHandler(this.pic_RemoveAbNormalList_Click);
            // 
            // pic_RemoveNormalList
            // 
            this.pic_RemoveNormalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveNormalList.Image")));
            this.pic_RemoveNormalList.Location = new System.Drawing.Point(520, 135);
            this.pic_RemoveNormalList.Name = "pic_RemoveNormalList";
            this.pic_RemoveNormalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveNormalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveNormalList.TabIndex = 87;
            this.pic_RemoveNormalList.TabStop = false;
            this.pic_RemoveNormalList.Click += new System.EventHandler(this.pic_RemoveNormalList_Click);
            // 
            // pic_RemoveNominalList
            // 
            this.pic_RemoveNominalList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveNominalList.Image")));
            this.pic_RemoveNominalList.Location = new System.Drawing.Point(574, 46);
            this.pic_RemoveNominalList.Name = "pic_RemoveNominalList";
            this.pic_RemoveNominalList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveNominalList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveNominalList.TabIndex = 86;
            this.pic_RemoveNominalList.TabStop = false;
            this.pic_RemoveNominalList.Click += new System.EventHandler(this.pic_RemoveNominalList_Click);
            // 
            // pic_AllParaToAbnormal
            // 
            this.pic_AllParaToAbnormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToAbnormal.Image")));
            this.pic_AllParaToAbnormal.Location = new System.Drawing.Point(360, 467);
            this.pic_AllParaToAbnormal.Name = "pic_AllParaToAbnormal";
            this.pic_AllParaToAbnormal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToAbnormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToAbnormal.TabIndex = 84;
            this.pic_AllParaToAbnormal.TabStop = false;
            this.pic_AllParaToAbnormal.Click += new System.EventHandler(this.pic_AllParaToAbnormal_Click);
            // 
            // pic_AllParaToNormal
            // 
            this.pic_AllParaToNormal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToNormal.Image")));
            this.pic_AllParaToNormal.Location = new System.Drawing.Point(360, 188);
            this.pic_AllParaToNormal.Name = "pic_AllParaToNormal";
            this.pic_AllParaToNormal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToNormal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToNormal.TabIndex = 83;
            this.pic_AllParaToNormal.TabStop = false;
            this.pic_AllParaToNormal.Click += new System.EventHandler(this.pic_AllParaToNormal_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label2.Location = new System.Drawing.Point(440, 408);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 30);
            this.label2.TabIndex = 81;
            this.label2.Text = "Seperated :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(443, 135);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 30);
            this.label1.TabIndex = 80;
            this.label1.Text = "Scale:";
            // 
            // lbl_Nominal
            // 
            this.lbl_Nominal.AutoSize = true;
            this.lbl_Nominal.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Nominal.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Nominal.Location = new System.Drawing.Point(440, 46);
            this.lbl_Nominal.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Nominal.Name = "lbl_Nominal";
            this.lbl_Nominal.Size = new System.Drawing.Size(140, 30);
            this.lbl_Nominal.TabIndex = 79;
            this.lbl_Nominal.Text = "Dependent :";
            // 
            // pic_AllParaToNominal
            // 
            this.pic_AllParaToNominal.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToNominal.Image")));
            this.pic_AllParaToNominal.Location = new System.Drawing.Point(360, 89);
            this.pic_AllParaToNominal.Name = "pic_AllParaToNominal";
            this.pic_AllParaToNominal.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToNominal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToNominal.TabIndex = 78;
            this.pic_AllParaToNominal.TabStop = false;
            this.pic_AllParaToNominal.Click += new System.EventHandler(this.pic_AllParaToNominal_Click);
            // 
            // list_Seperated
            // 
            this.list_Seperated.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Seperated.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Seperated.FormattingEnabled = true;
            this.list_Seperated.ItemHeight = 17;
            this.list_Seperated.Location = new System.Drawing.Point(445, 446);
            this.list_Seperated.Name = "list_Seperated";
            this.list_Seperated.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Seperated.Size = new System.Drawing.Size(263, 89);
            this.list_Seperated.TabIndex = 76;
            // 
            // list_NotSeperatedScale
            // 
            this.list_NotSeperatedScale.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_NotSeperatedScale.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_NotSeperatedScale.FormattingEnabled = true;
            this.list_NotSeperatedScale.ItemHeight = 17;
            this.list_NotSeperatedScale.Location = new System.Drawing.Point(445, 167);
            this.list_NotSeperatedScale.Name = "list_NotSeperatedScale";
            this.list_NotSeperatedScale.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_NotSeperatedScale.Size = new System.Drawing.Size(263, 89);
            this.list_NotSeperatedScale.TabIndex = 75;
            // 
            // list_Dependent
            // 
            this.list_Dependent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Dependent.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Dependent.FormattingEnabled = true;
            this.list_Dependent.ItemHeight = 17;
            this.list_Dependent.Location = new System.Drawing.Point(445, 85);
            this.list_Dependent.Name = "list_Dependent";
            this.list_Dependent.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Dependent.Size = new System.Drawing.Size(263, 38);
            this.list_Dependent.TabIndex = 74;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(754, 67);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(147, 30);
            this.label5.TabIndex = 105;
            this.label5.Text = "Table Name :";
            // 
            // txt_TableName
            // 
            this.txt_TableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_TableName.Location = new System.Drawing.Point(759, 99);
            this.txt_TableName.Margin = new System.Windows.Forms.Padding(2);
            this.txt_TableName.Name = "txt_TableName";
            this.txt_TableName.Size = new System.Drawing.Size(183, 28);
            this.txt_TableName.TabIndex = 104;
            // 
            // pic_removeTableSelected
            // 
            this.pic_removeTableSelected.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeTableSelected.Image")));
            this.pic_removeTableSelected.Location = new System.Drawing.Point(958, 396);
            this.pic_removeTableSelected.Name = "pic_removeTableSelected";
            this.pic_removeTableSelected.Size = new System.Drawing.Size(37, 30);
            this.pic_removeTableSelected.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeTableSelected.TabIndex = 103;
            this.pic_removeTableSelected.TabStop = false;
            this.pic_removeTableSelected.Click += new System.EventHandler(this.pic_removeTableSelected_Click);
            // 
            // pic_addTable
            // 
            this.pic_addTable.Image = ((System.Drawing.Image)(resources.GetObject("pic_addTable.Image")));
            this.pic_addTable.Location = new System.Drawing.Point(958, 99);
            this.pic_addTable.Name = "pic_addTable";
            this.pic_addTable.Size = new System.Drawing.Size(37, 28);
            this.pic_addTable.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_addTable.TabIndex = 102;
            this.pic_addTable.TabStop = false;
            this.pic_addTable.Click += new System.EventHandler(this.pic_addTable_Click);
            // 
            // cmb_TableNames
            // 
            this.cmb_TableNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_TableNames.FormattingEnabled = true;
            this.cmb_TableNames.Location = new System.Drawing.Point(759, 396);
            this.cmb_TableNames.Name = "cmb_TableNames";
            this.cmb_TableNames.Size = new System.Drawing.Size(183, 30);
            this.cmb_TableNames.TabIndex = 101;
            this.cmb_TableNames.SelectedIndexChanged += new System.EventHandler(this.cmb_TableNames_SelectedIndexChanged);
            // 
            // list_ViewTableParameters
            // 
            this.list_ViewTableParameters.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_ViewTableParameters.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_ViewTableParameters.FormattingEnabled = true;
            this.list_ViewTableParameters.ItemHeight = 17;
            this.list_ViewTableParameters.Location = new System.Drawing.Point(759, 165);
            this.list_ViewTableParameters.Name = "list_ViewTableParameters";
            this.list_ViewTableParameters.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_ViewTableParameters.Size = new System.Drawing.Size(183, 208);
            this.list_ViewTableParameters.TabIndex = 100;
            // 
            // btn_Update
            // 
            this.btn_Update.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Update.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Update.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Update.ForeColor = System.Drawing.Color.Black;
            this.btn_Update.Image = ((System.Drawing.Image)(resources.GetObject("btn_Update.Image")));
            this.btn_Update.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Update.Location = new System.Drawing.Point(768, 579);
            this.btn_Update.Name = "btn_Update";
            this.btn_Update.Size = new System.Drawing.Size(108, 54);
            this.btn_Update.TabIndex = 108;
            this.btn_Update.Text = "Update";
            this.btn_Update.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Update.UseVisualStyleBackColor = false;
            this.btn_Update.Click += new System.EventHandler(this.btn_Update_Click);
            // 
            // btn_SortTable
            // 
            this.btn_SortTable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_SortTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_SortTable.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_SortTable.ForeColor = System.Drawing.Color.Black;
            this.btn_SortTable.Image = ((System.Drawing.Image)(resources.GetObject("btn_SortTable.Image")));
            this.btn_SortTable.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_SortTable.Location = new System.Drawing.Point(768, 519);
            this.btn_SortTable.Name = "btn_SortTable";
            this.btn_SortTable.Size = new System.Drawing.Size(108, 54);
            this.btn_SortTable.TabIndex = 107;
            this.btn_SortTable.Text = "Sort";
            this.btn_SortTable.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_SortTable.UseVisualStyleBackColor = false;
            this.btn_SortTable.Click += new System.EventHandler(this.btn_SortTable_Click);
            // 
            // btn_Done
            // 
            this.btn_Done.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.btn_Done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Done.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Done.ForeColor = System.Drawing.Color.Black;
            this.btn_Done.Image = ((System.Drawing.Image)(resources.GetObject("btn_Done.Image")));
            this.btn_Done.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btn_Done.Location = new System.Drawing.Point(882, 579);
            this.btn_Done.Name = "btn_Done";
            this.btn_Done.Size = new System.Drawing.Size(108, 54);
            this.btn_Done.TabIndex = 106;
            this.btn_Done.Text = "Done";
            this.btn_Done.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.btn_Done.UseVisualStyleBackColor = false;
            this.btn_Done.Click += new System.EventHandler(this.btn_Done_Click);
            // 
            // cmb_ChooseTableFormat
            // 
            this.cmb_ChooseTableFormat.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F);
            this.cmb_ChooseTableFormat.FormattingEnabled = true;
            this.cmb_ChooseTableFormat.Items.AddRange(new object[] {
            "Linear",
            "Logistic"});
            this.cmb_ChooseTableFormat.Location = new System.Drawing.Point(102, 37);
            this.cmb_ChooseTableFormat.Name = "cmb_ChooseTableFormat";
            this.cmb_ChooseTableFormat.Size = new System.Drawing.Size(202, 30);
            this.cmb_ChooseTableFormat.TabIndex = 110;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label6.Location = new System.Drawing.Point(8, 35);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(99, 30);
            this.label6.TabIndex = 109;
            this.label6.Text = "Format :";
            // 
            // pic_RemoveSelectPara
            // 
            this.pic_RemoveSelectPara.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemoveSelectPara.Image")));
            this.pic_RemoveSelectPara.Location = new System.Drawing.Point(575, 551);
            this.pic_RemoveSelectPara.Name = "pic_RemoveSelectPara";
            this.pic_RemoveSelectPara.Size = new System.Drawing.Size(37, 30);
            this.pic_RemoveSelectPara.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemoveSelectPara.TabIndex = 115;
            this.pic_RemoveSelectPara.TabStop = false;
            this.pic_RemoveSelectPara.Click += new System.EventHandler(this.pic_RemoveSelectPara_Click);
            // 
            // pic_AllParaToSelect
            // 
            this.pic_AllParaToSelect.Image = ((System.Drawing.Image)(resources.GetObject("pic_AllParaToSelect.Image")));
            this.pic_AllParaToSelect.Location = new System.Drawing.Point(363, 587);
            this.pic_AllParaToSelect.Name = "pic_AllParaToSelect";
            this.pic_AllParaToSelect.Size = new System.Drawing.Size(59, 38);
            this.pic_AllParaToSelect.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AllParaToSelect.TabIndex = 114;
            this.pic_AllParaToSelect.TabStop = false;
            this.pic_AllParaToSelect.Click += new System.EventHandler(this.pic_AllParaToSelect_Click);
            // 
            // list_Select
            // 
            this.list_Select.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Select.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Select.FormattingEnabled = true;
            this.list_Select.ItemHeight = 17;
            this.list_Select.Location = new System.Drawing.Point(448, 587);
            this.list_Select.Name = "list_Select";
            this.list_Select.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.list_Select.Size = new System.Drawing.Size(263, 38);
            this.list_Select.TabIndex = 113;
            // 
            // lbl_Select
            // 
            this.lbl_Select.AutoSize = true;
            this.lbl_Select.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Select.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.lbl_Select.Location = new System.Drawing.Point(446, 551);
            this.lbl_Select.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbl_Select.Name = "lbl_Select";
            this.lbl_Select.Size = new System.Drawing.Size(87, 30);
            this.lbl_Select.TabIndex = 112;
            this.lbl_Select.Text = "Select :";
            // 
            // pic_groups_select
            // 
            this.pic_groups_select.Image = ((System.Drawing.Image)(resources.GetObject("pic_groups_select.Image")));
            this.pic_groups_select.Location = new System.Drawing.Point(532, 551);
            this.pic_groups_select.Name = "pic_groups_select";
            this.pic_groups_select.Size = new System.Drawing.Size(37, 30);
            this.pic_groups_select.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_groups_select.TabIndex = 111;
            this.pic_groups_select.TabStop = false;
            this.pic_groups_select.Click += new System.EventHandler(this.pic_groups_select_Click);
            // 
            // pic_clearNotSeperatedNominal
            // 
            this.pic_clearNotSeperatedNominal.Image = ((System.Drawing.Image)(resources.GetObject("pic_clearNotSeperatedNominal.Image")));
            this.pic_clearNotSeperatedNominal.Location = new System.Drawing.Point(593, 273);
            this.pic_clearNotSeperatedNominal.Name = "pic_clearNotSeperatedNominal";
            this.pic_clearNotSeperatedNominal.Size = new System.Drawing.Size(45, 31);
            this.pic_clearNotSeperatedNominal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_clearNotSeperatedNominal.TabIndex = 120;
            this.pic_clearNotSeperatedNominal.TabStop = false;
            this.pic_clearNotSeperatedNominal.Click += new System.EventHandler(this.pic_clearNotSeperatedNominal_Click);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(550, 273);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(37, 30);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox2.TabIndex = 119;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
            this.pictureBox3.Location = new System.Drawing.Point(360, 335);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(59, 38);
            this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox3.TabIndex = 118;
            this.pictureBox3.TabStop = false;
            this.pictureBox3.Click += new System.EventHandler(this.pictureBox3_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label3.Location = new System.Drawing.Point(443, 273);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 30);
            this.label3.TabIndex = 117;
            this.label3.Text = "Nominal :";
            // 
            // list_NotSeperatedNominal
            // 
            this.list_NotSeperatedNominal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_NotSeperatedNominal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_NotSeperatedNominal.FormattingEnabled = true;
            this.list_NotSeperatedNominal.ItemHeight = 17;
            this.list_NotSeperatedNominal.Location = new System.Drawing.Point(445, 305);
            this.list_NotSeperatedNominal.Name = "list_NotSeperatedNominal";
            this.list_NotSeperatedNominal.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_NotSeperatedNominal.Size = new System.Drawing.Size(263, 89);
            this.list_NotSeperatedNominal.TabIndex = 116;
            // 
            // Regression_Frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1015, 650);
            this.Controls.Add(this.pic_clearNotSeperatedNominal);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.list_NotSeperatedNominal);
            this.Controls.Add(this.pic_RemoveSelectPara);
            this.Controls.Add(this.pic_AllParaToSelect);
            this.Controls.Add(this.list_Select);
            this.Controls.Add(this.lbl_Select);
            this.Controls.Add(this.pic_groups_select);
            this.Controls.Add(this.cmb_ChooseTableFormat);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btn_Update);
            this.Controls.Add(this.btn_SortTable);
            this.Controls.Add(this.btn_Done);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txt_TableName);
            this.Controls.Add(this.pic_removeTableSelected);
            this.Controls.Add(this.pic_addTable);
            this.Controls.Add(this.cmb_TableNames);
            this.Controls.Add(this.list_ViewTableParameters);
            this.Controls.Add(this.pic_ClearAbnormalList);
            this.Controls.Add(this.pic_clearNormalList);
            this.Controls.Add(this.pic_ClearNominalList);
            this.Controls.Add(this.pic_AllLists);
            this.Controls.Add(this.lbl_Periods);
            this.Controls.Add(this.pic_RemoveAbNormalList);
            this.Controls.Add(this.pic_RemoveNormalList);
            this.Controls.Add(this.pic_RemoveNominalList);
            this.Controls.Add(this.pic_AllParaToAbnormal);
            this.Controls.Add(this.pic_AllParaToNormal);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbl_Nominal);
            this.Controls.Add(this.pic_AllParaToNominal);
            this.Controls.Add(this.list_Seperated);
            this.Controls.Add(this.list_NotSeperatedScale);
            this.Controls.Add(this.list_Dependent);
            this.Controls.Add(this.pic_back);
            this.Controls.Add(this.panelmove);
            this.Controls.Add(this.txt_ParaName);
            this.Controls.Add(this.data_allPara);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Regression_Frm";
            this.Text = "Regression_Frm";
            this.Load += new System.EventHandler(this.Regression_Frm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.data_allPara)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_back)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ClearAbnormalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_clearNormalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_ClearNominalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllLists)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveAbNormalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNormalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveNominalList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToAbnormal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNormal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToNominal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTableSelected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemoveSelectPara)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AllParaToSelect)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_groups_select)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_clearNotSeperatedNominal)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txt_ParaName;
        private System.Windows.Forms.DataGridView data_allPara;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Name;
        private System.Windows.Forms.DataGridViewImageColumn ColMeasure;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColNormality;
        private System.Windows.Forms.PictureBox pic_back;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.PictureBox pic_ClearAbnormalList;
        private System.Windows.Forms.PictureBox pic_clearNormalList;
        private System.Windows.Forms.PictureBox pic_ClearNominalList;
        private System.Windows.Forms.PictureBox pic_AllLists;
        private System.Windows.Forms.Label lbl_Periods;
        private System.Windows.Forms.PictureBox pic_RemoveAbNormalList;
        private System.Windows.Forms.PictureBox pic_RemoveNormalList;
        private System.Windows.Forms.PictureBox pic_RemoveNominalList;
        private System.Windows.Forms.PictureBox pic_AllParaToAbnormal;
        private System.Windows.Forms.PictureBox pic_AllParaToNormal;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbl_Nominal;
        private System.Windows.Forms.PictureBox pic_AllParaToNominal;
        private System.Windows.Forms.ListBox list_Seperated;
        private System.Windows.Forms.ListBox list_NotSeperatedScale;
        private System.Windows.Forms.ListBox list_Dependent;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_TableName;
        private System.Windows.Forms.PictureBox pic_removeTableSelected;
        private System.Windows.Forms.PictureBox pic_addTable;
        private System.Windows.Forms.ComboBox cmb_TableNames;
        private System.Windows.Forms.ListBox list_ViewTableParameters;
        private System.Windows.Forms.Button btn_Update;
        private System.Windows.Forms.Button btn_SortTable;
        private System.Windows.Forms.Button btn_Done;
        private System.Windows.Forms.ComboBox cmb_ChooseTableFormat;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.PictureBox pic_RemoveSelectPara;
        private System.Windows.Forms.PictureBox pic_AllParaToSelect;
        private System.Windows.Forms.ListBox list_Select;
        private System.Windows.Forms.Label lbl_Select;
        private System.Windows.Forms.PictureBox pic_groups_select;
        private System.Windows.Forms.PictureBox pic_clearNotSeperatedNominal;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox list_NotSeperatedNominal;
    }
}