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
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.CBScrubberKey = new System.Windows.Forms.ComboBox();
            this.CBTitleKey = new System.Windows.Forms.ComboBox();
            this.CBType = new System.Windows.Forms.ComboBox();
            this.TBActiveGuid = new System.Windows.Forms.TextBox();
            this.BtnSave = new System.Windows.Forms.Button();
            this.TBScrubberValue = new System.Windows.Forms.TextBox();
            this.TBTitleValue = new System.Windows.Forms.TextBox();
            this.TBDay = new System.Windows.Forms.TextBox();
            this.TBCriteria = new System.Windows.Forms.TextBox();
            this.GrpGUID = new System.Windows.Forms.GroupBox();
            this.CBHistoricGuid = new System.Windows.Forms.ComboBox();
            this.BtnCopyActiveGuid = new System.Windows.Forms.Button();
            this.BtnShowActiveGuids = new System.Windows.Forms.Button();
            this.TBWorkingGuid = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.BtnHistoricToWorking = new System.Windows.Forms.Button();
            this.BtnCommitGuid = new System.Windows.Forms.Button();
            this.BtnNewGuid = new System.Windows.Forms.Button();
            this.BtnPreviousSlide = new System.Windows.Forms.Button();
            this.BtnNextSlide = new System.Windows.Forms.Button();
            this.GrpSearch = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.CBFilter = new System.Windows.Forms.ComboBox();
            this.BtnSearch = new System.Windows.Forms.Button();
            this.BtnNextResult = new System.Windows.Forms.Button();
            this.BtnPreviousResult = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.BtnNewTitle = new System.Windows.Forms.Button();
            this.BtnSwapTitle = new System.Windows.Forms.Button();
            this.BtnNewScrubber = new System.Windows.Forms.Button();
            this.BtnSwapScrubber = new System.Windows.Forms.Button();
            this.BtnSaveAndProceed = new System.Windows.Forms.Button();
            this.GrpGUID.SuspendLayout();
            this.GrpSearch.SuspendLayout();
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
            this.label2.Size = new System.Drawing.Size(75, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "ACTIVE GUID";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "HISTORIC GUID";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "DAY";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 46);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "CHAP: SUB";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 71);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(71, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "SLIDE TITLE";
            // 
            // CBScrubberKey
            // 
            this.CBScrubberKey.FormattingEnabled = true;
            this.CBScrubberKey.Location = new System.Drawing.Point(364, 42);
            this.CBScrubberKey.Name = "CBScrubberKey";
            this.CBScrubberKey.Size = new System.Drawing.Size(160, 21);
            this.CBScrubberKey.TabIndex = 10;
            this.CBScrubberKey.SelectedIndexChanged += new System.EventHandler(this.CBScrubberKey_SelectedIndexChanged);
            // 
            // CBTitleKey
            // 
            this.CBTitleKey.FormattingEnabled = true;
            this.CBTitleKey.Location = new System.Drawing.Point(364, 68);
            this.CBTitleKey.Name = "CBTitleKey";
            this.CBTitleKey.Size = new System.Drawing.Size(160, 21);
            this.CBTitleKey.TabIndex = 11;
            this.CBTitleKey.SelectedIndexChanged += new System.EventHandler(this.CBTitleKey_SelectedIndexChanged);
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
            // TBActiveGuid
            // 
            this.TBActiveGuid.Location = new System.Drawing.Point(102, 21);
            this.TBActiveGuid.Name = "TBActiveGuid";
            this.TBActiveGuid.Size = new System.Drawing.Size(557, 20);
            this.TBActiveGuid.TabIndex = 16;
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(222, 375);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(238, 23);
            this.BtnSave.TabIndex = 22;
            this.BtnSave.Text = "Save";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // TBScrubberValue
            // 
            this.TBScrubberValue.Location = new System.Drawing.Point(102, 43);
            this.TBScrubberValue.Name = "TBScrubberValue";
            this.TBScrubberValue.Size = new System.Drawing.Size(256, 20);
            this.TBScrubberValue.TabIndex = 24;
            // 
            // TBTitleValue
            // 
            this.TBTitleValue.Location = new System.Drawing.Point(102, 68);
            this.TBTitleValue.Name = "TBTitleValue";
            this.TBTitleValue.Size = new System.Drawing.Size(256, 20);
            this.TBTitleValue.TabIndex = 25;
            // 
            // TBDay
            // 
            this.TBDay.Location = new System.Drawing.Point(102, 92);
            this.TBDay.Name = "TBDay";
            this.TBDay.Size = new System.Drawing.Size(557, 20);
            this.TBDay.TabIndex = 26;
            // 
            // TBCriteria
            // 
            this.TBCriteria.Location = new System.Drawing.Point(69, 19);
            this.TBCriteria.Name = "TBCriteria";
            this.TBCriteria.Size = new System.Drawing.Size(591, 20);
            this.TBCriteria.TabIndex = 28;
            // 
            // GrpGUID
            // 
            this.GrpGUID.Controls.Add(this.CBHistoricGuid);
            this.GrpGUID.Controls.Add(this.BtnCopyActiveGuid);
            this.GrpGUID.Controls.Add(this.BtnShowActiveGuids);
            this.GrpGUID.Controls.Add(this.TBWorkingGuid);
            this.GrpGUID.Controls.Add(this.label7);
            this.GrpGUID.Controls.Add(this.BtnHistoricToWorking);
            this.GrpGUID.Controls.Add(this.BtnCommitGuid);
            this.GrpGUID.Controls.Add(this.BtnNewGuid);
            this.GrpGUID.Controls.Add(this.label3);
            this.GrpGUID.Controls.Add(this.label2);
            this.GrpGUID.Controls.Add(this.TBActiveGuid);
            this.GrpGUID.Location = new System.Drawing.Point(8, 122);
            this.GrpGUID.Name = "GrpGUID";
            this.GrpGUID.Size = new System.Drawing.Size(666, 120);
            this.GrpGUID.TabIndex = 29;
            this.GrpGUID.TabStop = false;
            this.GrpGUID.Text = "GUIDs";
            // 
            // CBHistoricGuid
            // 
            this.CBHistoricGuid.FormattingEnabled = true;
            this.CBHistoricGuid.Location = new System.Drawing.Point(102, 43);
            this.CBHistoricGuid.Name = "CBHistoricGuid";
            this.CBHistoricGuid.Size = new System.Drawing.Size(557, 21);
            this.CBHistoricGuid.TabIndex = 33;
            // 
            // BtnCopyActiveGuid
            // 
            this.BtnCopyActiveGuid.Location = new System.Drawing.Point(530, 91);
            this.BtnCopyActiveGuid.Name = "BtnCopyActiveGuid";
            this.BtnCopyActiveGuid.Size = new System.Drawing.Size(125, 23);
            this.BtnCopyActiveGuid.TabIndex = 23;
            this.BtnCopyActiveGuid.Text = "Copy Active GUID";
            this.BtnCopyActiveGuid.UseVisualStyleBackColor = true;
            this.BtnCopyActiveGuid.Click += new System.EventHandler(this.BtnCopyActiveGuid_Click);
            // 
            // BtnShowActiveGuids
            // 
            this.BtnShowActiveGuids.Location = new System.Drawing.Point(399, 91);
            this.BtnShowActiveGuids.Name = "BtnShowActiveGuids";
            this.BtnShowActiveGuids.Size = new System.Drawing.Size(125, 23);
            this.BtnShowActiveGuids.TabIndex = 22;
            this.BtnShowActiveGuids.Text = "Show/Hide GUIDs";
            this.BtnShowActiveGuids.UseVisualStyleBackColor = true;
            this.BtnShowActiveGuids.Click += new System.EventHandler(this.BtnShowActiveGuids_Click);
            // 
            // TBWorkingGuid
            // 
            this.TBWorkingGuid.Location = new System.Drawing.Point(102, 67);
            this.TBWorkingGuid.Name = "TBWorkingGuid";
            this.TBWorkingGuid.Size = new System.Drawing.Size(557, 20);
            this.TBWorkingGuid.TabIndex = 21;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 70);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(90, 13);
            this.label7.TabIndex = 20;
            this.label7.Text = "WORKING GUID";
            // 
            // BtnHistoricToWorking
            // 
            this.BtnHistoricToWorking.Location = new System.Drawing.Point(268, 91);
            this.BtnHistoricToWorking.Name = "BtnHistoricToWorking";
            this.BtnHistoricToWorking.Size = new System.Drawing.Size(125, 23);
            this.BtnHistoricToWorking.TabIndex = 19;
            this.BtnHistoricToWorking.Text = "Historic To Working";
            this.BtnHistoricToWorking.UseVisualStyleBackColor = true;
            this.BtnHistoricToWorking.Click += new System.EventHandler(this.BtnHistoricToWorking_Click);
            // 
            // BtnCommitGuid
            // 
            this.BtnCommitGuid.Location = new System.Drawing.Point(137, 91);
            this.BtnCommitGuid.Name = "BtnCommitGuid";
            this.BtnCommitGuid.Size = new System.Drawing.Size(125, 23);
            this.BtnCommitGuid.TabIndex = 18;
            this.BtnCommitGuid.Text = "Commit Working GUID";
            this.BtnCommitGuid.UseVisualStyleBackColor = true;
            this.BtnCommitGuid.Click += new System.EventHandler(this.BtnCommitGuid_Click);
            // 
            // BtnNewGuid
            // 
            this.BtnNewGuid.Location = new System.Drawing.Point(6, 91);
            this.BtnNewGuid.Name = "BtnNewGuid";
            this.BtnNewGuid.Size = new System.Drawing.Size(125, 23);
            this.BtnNewGuid.TabIndex = 17;
            this.BtnNewGuid.Text = "New Working GUID";
            this.BtnNewGuid.UseVisualStyleBackColor = true;
            this.BtnNewGuid.Click += new System.EventHandler(this.BtnNewGuid_Click);
            // 
            // BtnPreviousSlide
            // 
            this.BtnPreviousSlide.Location = new System.Drawing.Point(8, 375);
            this.BtnPreviousSlide.Name = "BtnPreviousSlide";
            this.BtnPreviousSlide.Size = new System.Drawing.Size(208, 23);
            this.BtnPreviousSlide.TabIndex = 31;
            this.BtnPreviousSlide.Text = "Previous Slide";
            this.BtnPreviousSlide.UseVisualStyleBackColor = true;
            this.BtnPreviousSlide.Click += new System.EventHandler(this.BtnPreviousSlide_Click);
            // 
            // BtnNextSlide
            // 
            this.BtnNextSlide.Location = new System.Drawing.Point(465, 375);
            this.BtnNextSlide.Name = "BtnNextSlide";
            this.BtnNextSlide.Size = new System.Drawing.Size(208, 23);
            this.BtnNextSlide.TabIndex = 32;
            this.BtnNextSlide.Text = "Next Slide";
            this.BtnNextSlide.UseVisualStyleBackColor = true;
            this.BtnNextSlide.Click += new System.EventHandler(this.BtnNextSlide_Click);
            // 
            // GrpSearch
            // 
            this.GrpSearch.Controls.Add(this.label9);
            this.GrpSearch.Controls.Add(this.CBFilter);
            this.GrpSearch.Controls.Add(this.BtnSearch);
            this.GrpSearch.Controls.Add(this.BtnNextResult);
            this.GrpSearch.Controls.Add(this.BtnPreviousResult);
            this.GrpSearch.Controls.Add(this.label8);
            this.GrpSearch.Controls.Add(this.TBCriteria);
            this.GrpSearch.Location = new System.Drawing.Point(8, 12);
            this.GrpSearch.Name = "GrpSearch";
            this.GrpSearch.Size = new System.Drawing.Size(666, 104);
            this.GrpSearch.TabIndex = 33;
            this.GrpSearch.TabStop = false;
            this.GrpSearch.Text = "Search and Go";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 48);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(44, 13);
            this.label9.TabIndex = 37;
            this.label9.Text = "FILTER";
            // 
            // CBFilter
            // 
            this.CBFilter.FormattingEnabled = true;
            this.CBFilter.Location = new System.Drawing.Point(69, 45);
            this.CBFilter.Name = "CBFilter";
            this.CBFilter.Size = new System.Drawing.Size(591, 21);
            this.CBFilter.TabIndex = 34;
            // 
            // BtnSearch
            // 
            this.BtnSearch.Location = new System.Drawing.Point(214, 72);
            this.BtnSearch.Name = "BtnSearch";
            this.BtnSearch.Size = new System.Drawing.Size(238, 23);
            this.BtnSearch.TabIndex = 36;
            this.BtnSearch.Text = "Search";
            this.BtnSearch.UseVisualStyleBackColor = true;
            this.BtnSearch.Click += new System.EventHandler(this.BtnSearch_Click);
            // 
            // BtnNextResult
            // 
            this.BtnNextResult.Location = new System.Drawing.Point(457, 72);
            this.BtnNextResult.Name = "BtnNextResult";
            this.BtnNextResult.Size = new System.Drawing.Size(202, 23);
            this.BtnNextResult.TabIndex = 34;
            this.BtnNextResult.Text = "Next Result";
            this.BtnNextResult.UseVisualStyleBackColor = true;
            this.BtnNextResult.Click += new System.EventHandler(this.BtnNextResult_Click);
            // 
            // BtnPreviousResult
            // 
            this.BtnPreviousResult.Location = new System.Drawing.Point(9, 72);
            this.BtnPreviousResult.Name = "BtnPreviousResult";
            this.BtnPreviousResult.Size = new System.Drawing.Size(202, 23);
            this.BtnPreviousResult.TabIndex = 35;
            this.BtnPreviousResult.Text = "Previous Result";
            this.BtnPreviousResult.UseVisualStyleBackColor = true;
            this.BtnPreviousResult.Click += new System.EventHandler(this.BtnPreviousResult_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 22);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(57, 13);
            this.label8.TabIndex = 34;
            this.label8.Text = "CRITERIA";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.BtnNewTitle);
            this.groupBox2.Controls.Add(this.BtnSwapTitle);
            this.groupBox2.Controls.Add(this.BtnNewScrubber);
            this.groupBox2.Controls.Add(this.BtnSwapScrubber);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.CBType);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.CBScrubberKey);
            this.groupBox2.Controls.Add(this.TBScrubberValue);
            this.groupBox2.Controls.Add(this.TBDay);
            this.groupBox2.Controls.Add(this.TBTitleValue);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.CBTitleKey);
            this.groupBox2.Location = new System.Drawing.Point(8, 248);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(666, 121);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Metadata";
            // 
            // BtnNewTitle
            // 
            this.BtnNewTitle.Location = new System.Drawing.Point(530, 66);
            this.BtnNewTitle.Name = "BtnNewTitle";
            this.BtnNewTitle.Size = new System.Drawing.Size(63, 23);
            this.BtnNewTitle.TabIndex = 38;
            this.BtnNewTitle.Text = "New";
            this.BtnNewTitle.UseVisualStyleBackColor = true;
            this.BtnNewTitle.Click += new System.EventHandler(this.BtnNewTitle_Click);
            // 
            // BtnSwapTitle
            // 
            this.BtnSwapTitle.Location = new System.Drawing.Point(596, 66);
            this.BtnSwapTitle.Name = "BtnSwapTitle";
            this.BtnSwapTitle.Size = new System.Drawing.Size(63, 23);
            this.BtnSwapTitle.TabIndex = 37;
            this.BtnSwapTitle.Text = "Swap";
            this.BtnSwapTitle.UseVisualStyleBackColor = true;
            this.BtnSwapTitle.Click += new System.EventHandler(this.BtnSwapTitle_Click);
            // 
            // BtnNewScrubber
            // 
            this.BtnNewScrubber.Location = new System.Drawing.Point(530, 41);
            this.BtnNewScrubber.Name = "BtnNewScrubber";
            this.BtnNewScrubber.Size = new System.Drawing.Size(63, 23);
            this.BtnNewScrubber.TabIndex = 36;
            this.BtnNewScrubber.Text = "New";
            this.BtnNewScrubber.UseVisualStyleBackColor = true;
            this.BtnNewScrubber.Click += new System.EventHandler(this.BtnNewScrubber_Click);
            // 
            // BtnSwapScrubber
            // 
            this.BtnSwapScrubber.Location = new System.Drawing.Point(596, 41);
            this.BtnSwapScrubber.Name = "BtnSwapScrubber";
            this.BtnSwapScrubber.Size = new System.Drawing.Size(63, 23);
            this.BtnSwapScrubber.TabIndex = 34;
            this.BtnSwapScrubber.Text = "Swap";
            this.BtnSwapScrubber.UseVisualStyleBackColor = true;
            this.BtnSwapScrubber.Click += new System.EventHandler(this.BtnSwapScrubber_Click);
            // 
            // BtnSaveAndProceed
            // 
            this.BtnSaveAndProceed.Location = new System.Drawing.Point(8, 404);
            this.BtnSaveAndProceed.Name = "BtnSaveAndProceed";
            this.BtnSaveAndProceed.Size = new System.Drawing.Size(666, 23);
            this.BtnSaveAndProceed.TabIndex = 35;
            this.BtnSaveAndProceed.Text = "Save and Proceed";
            this.BtnSaveAndProceed.UseVisualStyleBackColor = true;
            this.BtnSaveAndProceed.Click += new System.EventHandler(this.BtnSaveAndProceed_Click);
            // 
            // SlideMetadata
            // 
            this.AcceptButton = this.BtnSaveAndProceed;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(681, 434);
            this.Controls.Add(this.BtnSaveAndProceed);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.GrpSearch);
            this.Controls.Add(this.BtnNextSlide);
            this.Controls.Add(this.BtnPreviousSlide);
            this.Controls.Add(this.GrpGUID);
            this.Controls.Add(this.BtnSave);
            this.Name = "SlideMetadata";
            this.Text = "Slide Metadata";
            this.GrpGUID.ResumeLayout(false);
            this.GrpGUID.PerformLayout();
            this.GrpSearch.ResumeLayout(false);
            this.GrpSearch.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox CBScrubberKey;
        private System.Windows.Forms.ComboBox CBTitleKey;
        private System.Windows.Forms.ComboBox CBType;
        private System.Windows.Forms.TextBox TBActiveGuid;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.TextBox TBScrubberValue;
        private System.Windows.Forms.TextBox TBTitleValue;
        private System.Windows.Forms.TextBox TBDay;
        private System.Windows.Forms.TextBox TBCriteria;
        private System.Windows.Forms.GroupBox GrpGUID;
        private System.Windows.Forms.Button BtnHistoricToWorking;
        private System.Windows.Forms.Button BtnCommitGuid;
        private System.Windows.Forms.Button BtnNewGuid;
        private System.Windows.Forms.TextBox TBWorkingGuid;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button BtnShowActiveGuids;
        private System.Windows.Forms.Button BtnCopyActiveGuid;
        private System.Windows.Forms.Button BtnPreviousSlide;
        private System.Windows.Forms.Button BtnNextSlide;
        private System.Windows.Forms.ComboBox CBHistoricGuid;
        private System.Windows.Forms.GroupBox GrpSearch;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox CBFilter;
        private System.Windows.Forms.Button BtnSearch;
        private System.Windows.Forms.Button BtnNextResult;
        private System.Windows.Forms.Button BtnPreviousResult;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button BtnSwapScrubber;
        private System.Windows.Forms.Button BtnSaveAndProceed;
        private System.Windows.Forms.Button BtnNewTitle;
        private System.Windows.Forms.Button BtnSwapTitle;
        private System.Windows.Forms.Button BtnNewScrubber;
    }
}