namespace Alta3_PPA
{
    partial class SlideMetadata
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.CBTitle = new System.Windows.Forms.ComboBox();
            this.CBChapSub = new System.Windows.Forms.ComboBox();
            this.CBType = new System.Windows.Forms.ComboBox();
            this.TBGuid = new System.Windows.Forms.TextBox();
            this.BtnSave = new System.Windows.Forms.Button();
            this.TBTitle = new System.Windows.Forms.TextBox();
            this.TBChapter = new System.Windows.Forms.TextBox();
            this.GrpGUID = new System.Windows.Forms.GroupBox();
            this.CBHistoricGuid = new System.Windows.Forms.ComboBox();
            this.BtnCopyGuid = new System.Windows.Forms.Button();
            this.BtnShowGuids = new System.Windows.Forms.Button();
            this.BtnHistoricToActive = new System.Windows.Forms.Button();
            this.BtnNewGuid = new System.Windows.Forms.Button();
            this.BtnPreviousSlide = new System.Windows.Forms.Button();
            this.BtnNextSlide = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.BtnSaveAndProceed = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.GrpGUID.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "TYPE";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "GUID";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "HISTORIC GUID";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 46);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(37, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "TITLE";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 71);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "CHAPTER";
            // 
            // CBTitle
            // 
            this.CBTitle.FormattingEnabled = true;
            this.CBTitle.Location = new System.Drawing.Point(500, 43);
            this.CBTitle.Name = "CBTitle";
            this.CBTitle.Size = new System.Drawing.Size(160, 21);
            this.CBTitle.TabIndex = 10;
            this.CBTitle.SelectedIndexChanged += new System.EventHandler(this.CBScrubberKey_SelectedIndexChanged);
            // 
            // CBChapSub
            // 
            this.CBChapSub.FormattingEnabled = true;
            this.CBChapSub.Location = new System.Drawing.Point(500, 68);
            this.CBChapSub.Name = "CBChapSub";
            this.CBChapSub.Size = new System.Drawing.Size(160, 21);
            this.CBChapSub.TabIndex = 11;
            this.CBChapSub.SelectedIndexChanged += new System.EventHandler(this.CBTitleKey_SelectedIndexChanged);
            // 
            // CBType
            // 
            this.CBType.FormattingEnabled = true;
            this.CBType.Location = new System.Drawing.Point(102, 18);
            this.CBType.Name = "CBType";
            this.CBType.Size = new System.Drawing.Size(558, 21);
            this.CBType.TabIndex = 13;
            this.CBType.SelectedIndexChanged += new System.EventHandler(this.CBType_SelectedIndexChanged);
            // 
            // TBGuid
            // 
            this.TBGuid.Location = new System.Drawing.Point(102, 21);
            this.TBGuid.Name = "TBGuid";
            this.TBGuid.Size = new System.Drawing.Size(557, 20);
            this.TBGuid.TabIndex = 16;
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(221, 265);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(238, 25);
            this.BtnSave.TabIndex = 22;
            this.BtnSave.Text = "Save";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // TBTitle
            // 
            this.TBTitle.Location = new System.Drawing.Point(102, 43);
            this.TBTitle.Name = "TBTitle";
            this.TBTitle.Size = new System.Drawing.Size(392, 20);
            this.TBTitle.TabIndex = 24;
            // 
            // TBChapter
            // 
            this.TBChapter.Location = new System.Drawing.Point(102, 68);
            this.TBChapter.Name = "TBChapter";
            this.TBChapter.Size = new System.Drawing.Size(392, 20);
            this.TBChapter.TabIndex = 25;
            // 
            // GrpGUID
            // 
            this.GrpGUID.Controls.Add(this.CBHistoricGuid);
            this.GrpGUID.Controls.Add(this.BtnCopyGuid);
            this.GrpGUID.Controls.Add(this.BtnShowGuids);
            this.GrpGUID.Controls.Add(this.BtnHistoricToActive);
            this.GrpGUID.Controls.Add(this.BtnNewGuid);
            this.GrpGUID.Controls.Add(this.label3);
            this.GrpGUID.Controls.Add(this.label2);
            this.GrpGUID.Controls.Add(this.TBGuid);
            this.GrpGUID.Location = new System.Drawing.Point(7, 12);
            this.GrpGUID.Name = "GrpGUID";
            this.GrpGUID.Size = new System.Drawing.Size(666, 104);
            this.GrpGUID.TabIndex = 29;
            this.GrpGUID.TabStop = false;
            this.GrpGUID.Text = "GUIDs";
            // 
            // CBHistoricGuid
            // 
            this.CBHistoricGuid.FormattingEnabled = true;
            this.CBHistoricGuid.Location = new System.Drawing.Point(102, 47);
            this.CBHistoricGuid.Name = "CBHistoricGuid";
            this.CBHistoricGuid.Size = new System.Drawing.Size(557, 21);
            this.CBHistoricGuid.TabIndex = 33;
            // 
            // BtnCopyGuid
            // 
            this.BtnCopyGuid.Location = new System.Drawing.Point(500, 74);
            this.BtnCopyGuid.Name = "BtnCopyGuid";
            this.BtnCopyGuid.Size = new System.Drawing.Size(160, 23);
            this.BtnCopyGuid.TabIndex = 23;
            this.BtnCopyGuid.Text = "Copy Active GUID";
            this.BtnCopyGuid.UseVisualStyleBackColor = true;
            this.BtnCopyGuid.Click += new System.EventHandler(this.BtnCopyGuid_Click);
            // 
            // BtnShowGuids
            // 
            this.BtnShowGuids.Location = new System.Drawing.Point(6, 74);
            this.BtnShowGuids.Name = "BtnShowGuids";
            this.BtnShowGuids.Size = new System.Drawing.Size(160, 23);
            this.BtnShowGuids.TabIndex = 22;
            this.BtnShowGuids.Text = "Show/Hide GUIDs";
            this.BtnShowGuids.UseVisualStyleBackColor = true;
            this.BtnShowGuids.Click += new System.EventHandler(this.BtnShowGuids_Click);
            // 
            // BtnHistoricToActive
            // 
            this.BtnHistoricToActive.Location = new System.Drawing.Point(172, 74);
            this.BtnHistoricToActive.Name = "BtnHistoricToActive";
            this.BtnHistoricToActive.Size = new System.Drawing.Size(156, 23);
            this.BtnHistoricToActive.TabIndex = 19;
            this.BtnHistoricToActive.Text = "Historic To Active";
            this.BtnHistoricToActive.UseVisualStyleBackColor = true;
            this.BtnHistoricToActive.Click += new System.EventHandler(this.BtnHistoricToWorking_Click);
            // 
            // BtnNewGuid
            // 
            this.BtnNewGuid.Location = new System.Drawing.Point(334, 74);
            this.BtnNewGuid.Name = "BtnNewGuid";
            this.BtnNewGuid.Size = new System.Drawing.Size(160, 23);
            this.BtnNewGuid.TabIndex = 17;
            this.BtnNewGuid.Text = "Create New GUID";
            this.BtnNewGuid.UseVisualStyleBackColor = true;
            this.BtnNewGuid.Click += new System.EventHandler(this.BtnNewGuid_Click);
            // 
            // BtnPreviousSlide
            // 
            this.BtnPreviousSlide.Location = new System.Drawing.Point(7, 265);
            this.BtnPreviousSlide.Name = "BtnPreviousSlide";
            this.BtnPreviousSlide.Size = new System.Drawing.Size(208, 25);
            this.BtnPreviousSlide.TabIndex = 31;
            this.BtnPreviousSlide.Text = "Previous Slide";
            this.BtnPreviousSlide.UseVisualStyleBackColor = true;
            this.BtnPreviousSlide.Click += new System.EventHandler(this.BtnPreviousSlide_Click);
            // 
            // BtnNextSlide
            // 
            this.BtnNextSlide.Location = new System.Drawing.Point(464, 265);
            this.BtnNextSlide.Name = "BtnNextSlide";
            this.BtnNextSlide.Size = new System.Drawing.Size(208, 25);
            this.BtnNextSlide.TabIndex = 32;
            this.BtnNextSlide.Text = "Next Slide";
            this.BtnNextSlide.UseVisualStyleBackColor = true;
            this.BtnNextSlide.Click += new System.EventHandler(this.BtnNextSlide_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.textBox1);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.CBType);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.CBTitle);
            this.groupBox2.Controls.Add(this.TBTitle);
            this.groupBox2.Controls.Add(this.TBChapter);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.CBChapSub);
            this.groupBox2.Controls.Add(this.menuStrip1);
            this.groupBox2.Location = new System.Drawing.Point(7, 122);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(666, 137);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Metadata";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Location = new System.Drawing.Point(3, 16);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(660, 24);
            this.menuStrip1.TabIndex = 39;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // BtnSaveAndProceed
            // 
            this.BtnSaveAndProceed.Location = new System.Drawing.Point(7, 294);
            this.BtnSaveAndProceed.Name = "BtnSaveAndProceed";
            this.BtnSaveAndProceed.Size = new System.Drawing.Size(666, 25);
            this.BtnSaveAndProceed.TabIndex = 35;
            this.BtnSaveAndProceed.Text = "Save and Proceed";
            this.BtnSaveAndProceed.UseVisualStyleBackColor = true;
            this.BtnSaveAndProceed.Click += new System.EventHandler(this.BtnSaveAndProceed_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(102, 94);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(392, 20);
            this.textBox1.TabIndex = 40;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 97);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(80, 13);
            this.label4.TabIndex = 41;
            this.label4.Text = "SUBCHAPTER";
            // 
            // SlideMetadata
            // 
            this.AcceptButton = this.BtnSaveAndProceed;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(681, 335);
            this.Controls.Add(this.BtnSaveAndProceed);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.BtnNextSlide);
            this.Controls.Add(this.BtnPreviousSlide);
            this.Controls.Add(this.GrpGUID);
            this.Controls.Add(this.BtnSave);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "SlideMetadata";
            this.Text = "Slide Metadata";
            this.GrpGUID.ResumeLayout(false);
            this.GrpGUID.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox CBTitle;
        private System.Windows.Forms.ComboBox CBChapSub;
        private System.Windows.Forms.ComboBox CBType;
        private System.Windows.Forms.TextBox TBGuid;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.TextBox TBTitle;
        private System.Windows.Forms.TextBox TBChapter;
        private System.Windows.Forms.GroupBox GrpGUID;
        private System.Windows.Forms.Button BtnHistoricToActive;
        private System.Windows.Forms.Button BtnNewGuid;
        private System.Windows.Forms.Button BtnShowGuids;
        private System.Windows.Forms.Button BtnCopyGuid;
        private System.Windows.Forms.Button BtnPreviousSlide;
        private System.Windows.Forms.Button BtnNextSlide;
        private System.Windows.Forms.ComboBox CBHistoricGuid;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button BtnSaveAndProceed;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label4;
    }
}