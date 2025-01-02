namespace ExcelScore.Forms
{
    partial class Word
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Word));
            this.button1 = new System.Windows.Forms.Button();
            this.pic_exit = new System.Windows.Forms.PictureBox();
            this.list_allDomains = new System.Windows.Forms.ListBox();
            this.list_Periods = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.pic_AddPeriodsList = new System.Windows.Forms.PictureBox();
            this.pic_RemovePeriodsList = new System.Windows.Forms.PictureBox();
            this.panelmove = new System.Windows.Forms.Panel();
            this.list_descriptive = new System.Windows.Forms.ListBox();
            this.Descriptive = new System.Windows.Forms.Label();
            this.pic_AddDescriptive = new System.Windows.Forms.PictureBox();
            this.pic_removeDescriptive = new System.Windows.Forms.PictureBox();
            this.list_TotalScorePeriods = new System.Windows.Forms.ListBox();
            this.pic_removeTotalScoreP = new System.Windows.Forms.PictureBox();
            this.pic_addTotalScoreP = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.list_Overall = new System.Windows.Forms.ListBox();
            this.pic_removeOverall = new System.Windows.Forms.PictureBox();
            this.pic_addOverall = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.periods_check = new System.Windows.Forms.CheckBox();
            this.pic_clearAllLists = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pic_exit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AddPeriodsList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemovePeriodsList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AddDescriptive)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeDescriptive)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTotalScoreP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTotalScoreP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeOverall)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addOverall)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_clearAllLists)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.Black;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.button1.Location = new System.Drawing.Point(1100, 552);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 44);
            this.button1.TabIndex = 13;
            this.button1.Text = "Done";
            this.button1.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.btnDone_Click);
            // 
            // pic_exit
            // 
            this.pic_exit.Image = ((System.Drawing.Image)(resources.GetObject("pic_exit.Image")));
            this.pic_exit.Location = new System.Drawing.Point(1149, 1);
            this.pic_exit.Name = "pic_exit";
            this.pic_exit.Size = new System.Drawing.Size(60, 47);
            this.pic_exit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_exit.TabIndex = 18;
            this.pic_exit.TabStop = false;
            this.pic_exit.Click += new System.EventHandler(this.pic_exit_Click);
            // 
            // list_allDomains
            // 
            this.list_allDomains.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_allDomains.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_allDomains.FormattingEnabled = true;
            this.list_allDomains.ItemHeight = 19;
            this.list_allDomains.Location = new System.Drawing.Point(31, 45);
            this.list_allDomains.Name = "list_allDomains";
            this.list_allDomains.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_allDomains.Size = new System.Drawing.Size(263, 498);
            this.list_allDomains.TabIndex = 19;
            this.list_allDomains.SelectedIndexChanged += new System.EventHandler(this.list_allDomains_SelectedIndexChanged);
            // 
            // list_Periods
            // 
            this.list_Periods.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Periods.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Periods.FormattingEnabled = true;
            this.list_Periods.ItemHeight = 17;
            this.list_Periods.Location = new System.Drawing.Point(539, 100);
            this.list_Periods.Name = "list_Periods";
            this.list_Periods.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Periods.Size = new System.Drawing.Size(173, 446);
            this.list_Periods.TabIndex = 37;
            this.list_Periods.SelectedIndexChanged += new System.EventHandler(this.list_Periods_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label5.Location = new System.Drawing.Point(534, 67);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(102, 30);
            this.label5.TabIndex = 38;
            this.label5.Text = "Periods :";
            // 
            // pic_AddPeriodsList
            // 
            this.pic_AddPeriodsList.Image = ((System.Drawing.Image)(resources.GetObject("pic_AddPeriodsList.Image")));
            this.pic_AddPeriodsList.Location = new System.Drawing.Point(629, 68);
            this.pic_AddPeriodsList.Name = "pic_AddPeriodsList";
            this.pic_AddPeriodsList.Size = new System.Drawing.Size(35, 30);
            this.pic_AddPeriodsList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AddPeriodsList.TabIndex = 39;
            this.pic_AddPeriodsList.TabStop = false;
            this.pic_AddPeriodsList.Click += new System.EventHandler(this.pic_AddPeriodsList_Click);
            // 
            // pic_RemovePeriodsList
            // 
            this.pic_RemovePeriodsList.Image = ((System.Drawing.Image)(resources.GetObject("pic_RemovePeriodsList.Image")));
            this.pic_RemovePeriodsList.Location = new System.Drawing.Point(670, 68);
            this.pic_RemovePeriodsList.Name = "pic_RemovePeriodsList";
            this.pic_RemovePeriodsList.Size = new System.Drawing.Size(37, 30);
            this.pic_RemovePeriodsList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_RemovePeriodsList.TabIndex = 40;
            this.pic_RemovePeriodsList.TabStop = false;
            this.pic_RemovePeriodsList.Click += new System.EventHandler(this.pic_RemovePeriodsList_Click);
            // 
            // panelmove
            // 
            this.panelmove.Location = new System.Drawing.Point(11, 1);
            this.panelmove.Margin = new System.Windows.Forms.Padding(2);
            this.panelmove.Name = "panelmove";
            this.panelmove.Size = new System.Drawing.Size(1110, 23);
            this.panelmove.TabIndex = 44;
            this.panelmove.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelmove_MouseDown);
            // 
            // list_descriptive
            // 
            this.list_descriptive.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_descriptive.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_descriptive.FormattingEnabled = true;
            this.list_descriptive.ItemHeight = 19;
            this.list_descriptive.Location = new System.Drawing.Point(321, 100);
            this.list_descriptive.Name = "list_descriptive";
            this.list_descriptive.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_descriptive.Size = new System.Drawing.Size(173, 441);
            this.list_descriptive.TabIndex = 46;
            // 
            // Descriptive
            // 
            this.Descriptive.AutoSize = true;
            this.Descriptive.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Descriptive.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.Descriptive.Location = new System.Drawing.Point(316, 67);
            this.Descriptive.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Descriptive.Name = "Descriptive";
            this.Descriptive.Size = new System.Drawing.Size(142, 30);
            this.Descriptive.TabIndex = 48;
            this.Descriptive.Text = "Descriptive :";
            this.Descriptive.Click += new System.EventHandler(this.Descriptive_Click);
            // 
            // pic_AddDescriptive
            // 
            this.pic_AddDescriptive.Image = ((System.Drawing.Image)(resources.GetObject("pic_AddDescriptive.Image")));
            this.pic_AddDescriptive.Location = new System.Drawing.Point(454, 67);
            this.pic_AddDescriptive.Name = "pic_AddDescriptive";
            this.pic_AddDescriptive.Size = new System.Drawing.Size(34, 30);
            this.pic_AddDescriptive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_AddDescriptive.TabIndex = 48;
            this.pic_AddDescriptive.TabStop = false;
            this.pic_AddDescriptive.Click += new System.EventHandler(this.pic_AddDescriptive_Click);
            // 
            // pic_removeDescriptive
            // 
            this.pic_removeDescriptive.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeDescriptive.Image")));
            this.pic_removeDescriptive.Location = new System.Drawing.Point(494, 68);
            this.pic_removeDescriptive.Name = "pic_removeDescriptive";
            this.pic_removeDescriptive.Size = new System.Drawing.Size(35, 30);
            this.pic_removeDescriptive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeDescriptive.TabIndex = 49;
            this.pic_removeDescriptive.TabStop = false;
            this.pic_removeDescriptive.Click += new System.EventHandler(this.pic_removeDescriptive_Click);
            // 
            // list_TotalScorePeriods
            // 
            this.list_TotalScorePeriods.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_TotalScorePeriods.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_TotalScorePeriods.FormattingEnabled = true;
            this.list_TotalScorePeriods.ItemHeight = 17;
            this.list_TotalScorePeriods.Location = new System.Drawing.Point(747, 100);
            this.list_TotalScorePeriods.Name = "list_TotalScorePeriods";
            this.list_TotalScorePeriods.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_TotalScorePeriods.Size = new System.Drawing.Size(173, 446);
            this.list_TotalScorePeriods.TabIndex = 50;
            // 
            // pic_removeTotalScoreP
            // 
            this.pic_removeTotalScoreP.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeTotalScoreP.Image")));
            this.pic_removeTotalScoreP.Location = new System.Drawing.Point(907, 68);
            this.pic_removeTotalScoreP.Name = "pic_removeTotalScoreP";
            this.pic_removeTotalScoreP.Size = new System.Drawing.Size(30, 30);
            this.pic_removeTotalScoreP.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeTotalScoreP.TabIndex = 53;
            this.pic_removeTotalScoreP.TabStop = false;
            this.pic_removeTotalScoreP.Click += new System.EventHandler(this.pic_removeTotalScoreP_Click);
            // 
            // pic_addTotalScoreP
            // 
            this.pic_addTotalScoreP.Image = ((System.Drawing.Image)(resources.GetObject("pic_addTotalScoreP.Image")));
            this.pic_addTotalScoreP.Location = new System.Drawing.Point(871, 67);
            this.pic_addTotalScoreP.Name = "pic_addTotalScoreP";
            this.pic_addTotalScoreP.Size = new System.Drawing.Size(30, 30);
            this.pic_addTotalScoreP.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_addTotalScoreP.TabIndex = 52;
            this.pic_addTotalScoreP.TabStop = false;
            this.pic_addTotalScoreP.Click += new System.EventHandler(this.pic_addTotalScoreP_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label1.Location = new System.Drawing.Point(727, 67);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 30);
            this.label1.TabIndex = 51;
            this.label1.Text = "Total Score :";
            // 
            // list_Overall
            // 
            this.list_Overall.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.list_Overall.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.list_Overall.FormattingEnabled = true;
            this.list_Overall.ItemHeight = 17;
            this.list_Overall.Location = new System.Drawing.Point(984, 100);
            this.list_Overall.Name = "list_Overall";
            this.list_Overall.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Overall.Size = new System.Drawing.Size(173, 446);
            this.list_Overall.TabIndex = 54;
            // 
            // pic_removeOverall
            // 
            this.pic_removeOverall.Image = ((System.Drawing.Image)(resources.GetObject("pic_removeOverall.Image")));
            this.pic_removeOverall.Location = new System.Drawing.Point(1119, 67);
            this.pic_removeOverall.Name = "pic_removeOverall";
            this.pic_removeOverall.Size = new System.Drawing.Size(30, 30);
            this.pic_removeOverall.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_removeOverall.TabIndex = 57;
            this.pic_removeOverall.TabStop = false;
            this.pic_removeOverall.Click += new System.EventHandler(this.pic_removeOverall_Click);
            // 
            // pic_addOverall
            // 
            this.pic_addOverall.Image = ((System.Drawing.Image)(resources.GetObject("pic_addOverall.Image")));
            this.pic_addOverall.Location = new System.Drawing.Point(1083, 67);
            this.pic_addOverall.Name = "pic_addOverall";
            this.pic_addOverall.Size = new System.Drawing.Size(30, 30);
            this.pic_addOverall.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_addOverall.TabIndex = 56;
            this.pic_addOverall.TabStop = false;
            this.pic_addOverall.Click += new System.EventHandler(this.pic_addOverall_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.label2.Location = new System.Drawing.Point(979, 67);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 30);
            this.label2.TabIndex = 55;
            this.label2.Text = "Overall :";
            // 
            // periods_check
            // 
            this.periods_check.AutoSize = true;
            this.periods_check.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.periods_check.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(205)))), ((int)(((byte)(137)))), ((int)(((byte)(36)))));
            this.periods_check.Location = new System.Drawing.Point(984, 560);
            this.periods_check.Name = "periods_check";
            this.periods_check.Size = new System.Drawing.Size(97, 29);
            this.periods_check.TabIndex = 58;
            this.periods_check.Text = "Periods";
            this.periods_check.UseVisualStyleBackColor = true;
            // 
            // pic_clearAllLists
            // 
            this.pic_clearAllLists.Image = ((System.Drawing.Image)(resources.GetObject("pic_clearAllLists.Image")));
            this.pic_clearAllLists.Location = new System.Drawing.Point(300, 29);
            this.pic_clearAllLists.Name = "pic_clearAllLists";
            this.pic_clearAllLists.Size = new System.Drawing.Size(51, 35);
            this.pic_clearAllLists.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pic_clearAllLists.TabIndex = 59;
            this.pic_clearAllLists.TabStop = false;
            this.pic_clearAllLists.Click += new System.EventHandler(this.pic_clearAllLists_Click);
            // 
            // Word
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(245)))), ((int)(((byte)(237)))));
            this.ClientSize = new System.Drawing.Size(1212, 608);
            this.Controls.Add(this.pic_clearAllLists);
            this.Controls.Add(this.periods_check);
            this.Controls.Add(this.pic_removeOverall);
            this.Controls.Add(this.pic_addOverall);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.list_Overall);
            this.Controls.Add(this.pic_removeTotalScoreP);
            this.Controls.Add(this.pic_addTotalScoreP);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.list_TotalScorePeriods);
            this.Controls.Add(this.pic_removeDescriptive);
            this.Controls.Add(this.pic_AddDescriptive);
            this.Controls.Add(this.Descriptive);
            this.Controls.Add(this.list_descriptive);
            this.Controls.Add(this.pic_RemovePeriodsList);
            this.Controls.Add(this.pic_AddPeriodsList);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.list_Periods);
            this.Controls.Add(this.list_allDomains);
            this.Controls.Add(this.pic_exit);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panelmove);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Word";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Word";
            this.Load += new System.EventHandler(this.Word_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pic_exit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AddPeriodsList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_RemovePeriodsList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_AddDescriptive)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeDescriptive)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeTotalScoreP)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addTotalScoreP)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_removeOverall)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_addOverall)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic_clearAllLists)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.PictureBox pic_exit;
        private System.Windows.Forms.ListBox list_allDomains;
        private System.Windows.Forms.ListBox list_Periods;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.PictureBox pic_AddPeriodsList;
        private System.Windows.Forms.PictureBox pic_RemovePeriodsList;
        private System.Windows.Forms.Panel panelmove;
        private System.Windows.Forms.ListBox list_descriptive;
        private System.Windows.Forms.Label Descriptive;
        private System.Windows.Forms.PictureBox pic_AddDescriptive;
        private System.Windows.Forms.PictureBox pic_removeDescriptive;
        private System.Windows.Forms.ListBox list_TotalScorePeriods;
        private System.Windows.Forms.PictureBox pic_removeTotalScoreP;
        private System.Windows.Forms.PictureBox pic_addTotalScoreP;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox list_Overall;
        private System.Windows.Forms.PictureBox pic_removeOverall;
        private System.Windows.Forms.PictureBox pic_addOverall;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox periods_check;
        private System.Windows.Forms.PictureBox pic_clearAllLists;
    }
}