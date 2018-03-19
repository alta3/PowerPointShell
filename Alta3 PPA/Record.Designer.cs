namespace Alta3_PPA
{
    partial class Record
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
            this.components = new System.ComponentModel.Container();
            this.recordTime = new System.Windows.Forms.Timer(this.components);
            this.TBFqdn = new System.Windows.Forms.TextBox();
            this.TBPort = new System.Windows.Forms.TextBox();
            this.TBTime = new System.Windows.Forms.TextBox();
            this.TBSubchapter = new System.Windows.Forms.TextBox();
            this.TBLocation = new System.Windows.Forms.TextBox();
            this.TBIndex = new System.Windows.Forms.TextBox();
            this.TBGuid = new System.Windows.Forms.TextBox();
            this.TBChapter = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lbTitle = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.btnPrevious = new System.Windows.Forms.Button();
            this.btnStartStop = new System.Windows.Forms.Button();
            this.btnNextSlide = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.btnBrowser = new System.Windows.Forms.Button();
            this.TBTitle = new System.Windows.Forms.TextBox();
            this.TBType = new System.Windows.Forms.TextBox();
            this.fldBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // recordTime
            // 
            this.recordTime.Interval = 1000;
            this.recordTime.Tick += new System.EventHandler(this.recordTime_Tick);
            // 
            // TBFqdn
            // 
            this.TBFqdn.Location = new System.Drawing.Point(77, 184);
            this.TBFqdn.Name = "TBFqdn";
            this.TBFqdn.Size = new System.Drawing.Size(244, 20);
            this.TBFqdn.TabIndex = 0;
            // 
            // TBPort
            // 
            this.TBPort.Location = new System.Drawing.Point(362, 184);
            this.TBPort.Name = "TBPort";
            this.TBPort.Size = new System.Drawing.Size(66, 20);
            this.TBPort.TabIndex = 11;
            // 
            // TBTime
            // 
            this.TBTime.Location = new System.Drawing.Point(511, 184);
            this.TBTime.Name = "TBTime";
            this.TBTime.Size = new System.Drawing.Size(200, 20);
            this.TBTime.TabIndex = 12;
            // 
            // TBSubchapter
            // 
            this.TBSubchapter.Enabled = false;
            this.TBSubchapter.Location = new System.Drawing.Point(77, 86);
            this.TBSubchapter.Name = "TBSubchapter";
            this.TBSubchapter.Size = new System.Drawing.Size(634, 20);
            this.TBSubchapter.TabIndex = 13;
            // 
            // TBLocation
            // 
            this.TBLocation.Location = new System.Drawing.Point(77, 160);
            this.TBLocation.Name = "TBLocation";
            this.TBLocation.Size = new System.Drawing.Size(482, 20);
            this.TBLocation.TabIndex = 14;
            // 
            // TBIndex
            // 
            this.TBIndex.Enabled = false;
            this.TBIndex.Location = new System.Drawing.Point(77, 111);
            this.TBIndex.Name = "TBIndex";
            this.TBIndex.Size = new System.Drawing.Size(634, 20);
            this.TBIndex.TabIndex = 15;
            // 
            // TBGuid
            // 
            this.TBGuid.Enabled = false;
            this.TBGuid.Location = new System.Drawing.Point(77, 137);
            this.TBGuid.Name = "TBGuid";
            this.TBGuid.Size = new System.Drawing.Size(634, 20);
            this.TBGuid.TabIndex = 16;
            // 
            // TBChapter
            // 
            this.TBChapter.Enabled = false;
            this.TBChapter.Location = new System.Drawing.Point(77, 60);
            this.TBChapter.Name = "TBChapter";
            this.TBChapter.Size = new System.Drawing.Size(634, 20);
            this.TBChapter.TabIndex = 17;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 187);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "IP/FQDN:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(327, 187);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "Port:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(434, 187);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 13);
            this.label3.TabIndex = 20;
            this.label3.Text = "Record Time:";
            // 
            // lbTitle
            // 
            this.lbTitle.AutoSize = true;
            this.lbTitle.Location = new System.Drawing.Point(9, 11);
            this.lbTitle.Name = "lbTitle";
            this.lbTitle.Size = new System.Drawing.Size(56, 13);
            this.lbTitle.TabIndex = 21;
            this.lbTitle.Text = "Slide Title:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 163);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 13);
            this.label5.TabIndex = 22;
            this.label5.Text = "Location:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 114);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(62, 13);
            this.label6.TabIndex = 23;
            this.label6.Text = "Slide Index:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 37);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(60, 13);
            this.label7.TabIndex = 24;
            this.label7.Text = "Slide Type:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(9, 140);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "Slide GUID:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(9, 63);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(44, 13);
            this.label9.TabIndex = 26;
            this.label9.Text = "Chapter";
            // 
            // btnPrevious
            // 
            this.btnPrevious.Location = new System.Drawing.Point(12, 210);
            this.btnPrevious.Name = "btnPrevious";
            this.btnPrevious.Size = new System.Drawing.Size(200, 23);
            this.btnPrevious.TabIndex = 27;
            this.btnPrevious.Text = "Previous Slide";
            this.btnPrevious.UseVisualStyleBackColor = true;
            this.btnPrevious.Click += new System.EventHandler(this.btnPrevious_Click);
            // 
            // btnStartStop
            // 
            this.btnStartStop.Location = new System.Drawing.Point(218, 210);
            this.btnStartStop.Name = "btnStartStop";
            this.btnStartStop.Size = new System.Drawing.Size(287, 23);
            this.btnStartStop.TabIndex = 29;
            this.btnStartStop.Text = "Start/Stop Recording";
            this.btnStartStop.UseVisualStyleBackColor = true;
            this.btnStartStop.Click += new System.EventHandler(this.btnStartStop_Click);
            // 
            // btnNextSlide
            // 
            this.btnNextSlide.Location = new System.Drawing.Point(511, 210);
            this.btnNextSlide.Name = "btnNextSlide";
            this.btnNextSlide.Size = new System.Drawing.Size(200, 23);
            this.btnNextSlide.TabIndex = 30;
            this.btnNextSlide.Text = "Next Slide";
            this.btnNextSlide.UseVisualStyleBackColor = true;
            this.btnNextSlide.Click += new System.EventHandler(this.btnNextSlide_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(9, 89);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(62, 13);
            this.label11.TabIndex = 32;
            this.label11.Text = "Subchapter";
            // 
            // btnBrowser
            // 
            this.btnBrowser.Location = new System.Drawing.Point(568, 158);
            this.btnBrowser.Name = "btnBrowser";
            this.btnBrowser.Size = new System.Drawing.Size(143, 23);
            this.btnBrowser.TabIndex = 33;
            this.btnBrowser.Text = "Browse";
            this.btnBrowser.UseVisualStyleBackColor = true;
            this.btnBrowser.Click += new System.EventHandler(this.btnBrowser_Click);
            // 
            // TBTitle
            // 
            this.TBTitle.Enabled = false;
            this.TBTitle.Location = new System.Drawing.Point(77, 8);
            this.TBTitle.Name = "TBTitle";
            this.TBTitle.Size = new System.Drawing.Size(634, 20);
            this.TBTitle.TabIndex = 34;
            // 
            // TBType
            // 
            this.TBType.Enabled = false;
            this.TBType.Location = new System.Drawing.Point(77, 34);
            this.TBType.Name = "TBType";
            this.TBType.Size = new System.Drawing.Size(634, 20);
            this.TBType.TabIndex = 35;
            // 
            // Record
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(723, 241);
            this.Controls.Add(this.TBType);
            this.Controls.Add(this.TBTitle);
            this.Controls.Add(this.btnBrowser);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.btnNextSlide);
            this.Controls.Add(this.btnStartStop);
            this.Controls.Add(this.btnPrevious);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lbTitle);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TBChapter);
            this.Controls.Add(this.TBGuid);
            this.Controls.Add(this.TBIndex);
            this.Controls.Add(this.TBLocation);
            this.Controls.Add(this.TBSubchapter);
            this.Controls.Add(this.TBTime);
            this.Controls.Add(this.TBPort);
            this.Controls.Add(this.TBFqdn);
            this.Name = "Record";
            this.Text = "Record";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer recordTime;
        private System.Windows.Forms.TextBox TBFqdn;
        private System.Windows.Forms.TextBox TBPort;
        private System.Windows.Forms.TextBox TBTime;
        private System.Windows.Forms.TextBox TBSubchapter;
        private System.Windows.Forms.TextBox TBLocation;
        private System.Windows.Forms.TextBox TBIndex;
        private System.Windows.Forms.TextBox TBGuid;
        private System.Windows.Forms.TextBox TBChapter;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lbTitle;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnPrevious;
        private System.Windows.Forms.Button btnStartStop;
        private System.Windows.Forms.Button btnNextSlide;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btnBrowser;
        private System.Windows.Forms.TextBox TBTitle;
        private System.Windows.Forms.TextBox TBType;
        private System.Windows.Forms.FolderBrowserDialog fldBrowser;
    }
}