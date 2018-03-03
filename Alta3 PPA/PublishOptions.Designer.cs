namespace Alta3_PPA
{
    partial class PublishOptions
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
            this.chkPDF = new System.Windows.Forms.CheckBox();
            this.chkLatex = new System.Windows.Forms.CheckBox();
            this.chkMarkdown = new System.Windows.Forms.CheckBox();
            this.txtPubDir = new System.Windows.Forms.TextBox();
            this.btnFldBrowser = new System.Windows.Forms.Button();
            this.chkPNG = new System.Windows.Forms.CheckBox();
            this.chkPowerPoint = new System.Windows.Forms.CheckBox();
            this.chkVocab = new System.Windows.Forms.CheckBox();
            this.chkQuestion = new System.Windows.Forms.CheckBox();
            this.btnPublish = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.chkYAML = new System.Windows.Forms.CheckBox();
            this.fldBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // chkPDF
            // 
            this.chkPDF.AutoSize = true;
            this.chkPDF.Checked = true;
            this.chkPDF.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPDF.Location = new System.Drawing.Point(123, 12);
            this.chkPDF.Name = "chkPDF";
            this.chkPDF.Size = new System.Drawing.Size(47, 17);
            this.chkPDF.TabIndex = 0;
            this.chkPDF.Text = "PDF";
            this.chkPDF.UseVisualStyleBackColor = true;
            // 
            // chkLatex
            // 
            this.chkLatex.AutoSize = true;
            this.chkLatex.Checked = true;
            this.chkLatex.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLatex.Location = new System.Drawing.Point(176, 13);
            this.chkLatex.Name = "chkLatex";
            this.chkLatex.Size = new System.Drawing.Size(56, 17);
            this.chkLatex.TabIndex = 1;
            this.chkLatex.Text = "LaTex";
            this.chkLatex.UseVisualStyleBackColor = true;
            // 
            // chkMarkdown
            // 
            this.chkMarkdown.AutoSize = true;
            this.chkMarkdown.Checked = true;
            this.chkMarkdown.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMarkdown.Location = new System.Drawing.Point(238, 12);
            this.chkMarkdown.Name = "chkMarkdown";
            this.chkMarkdown.Size = new System.Drawing.Size(76, 17);
            this.chkMarkdown.TabIndex = 2;
            this.chkMarkdown.Text = "Markdown";
            this.chkMarkdown.UseVisualStyleBackColor = true;
            // 
            // txtPubDir
            // 
            this.txtPubDir.Location = new System.Drawing.Point(86, 35);
            this.txtPubDir.Name = "txtPubDir";
            this.txtPubDir.Size = new System.Drawing.Size(428, 20);
            this.txtPubDir.TabIndex = 3;
            // 
            // btnFldBrowser
            // 
            this.btnFldBrowser.Location = new System.Drawing.Point(520, 32);
            this.btnFldBrowser.Name = "btnFldBrowser";
            this.btnFldBrowser.Size = new System.Drawing.Size(75, 23);
            this.btnFldBrowser.TabIndex = 4;
            this.btnFldBrowser.Text = "Browse";
            this.btnFldBrowser.UseVisualStyleBackColor = true;
            this.btnFldBrowser.Click += new System.EventHandler(this.btnFldBrowser_Click);
            // 
            // chkPNG
            // 
            this.chkPNG.AutoSize = true;
            this.chkPNG.Checked = true;
            this.chkPNG.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPNG.Location = new System.Drawing.Point(320, 13);
            this.chkPNG.Name = "chkPNG";
            this.chkPNG.Size = new System.Drawing.Size(54, 17);
            this.chkPNG.TabIndex = 5;
            this.chkPNG.Text = "PNGs";
            this.chkPNG.UseVisualStyleBackColor = true;
            // 
            // chkPowerPoint
            // 
            this.chkPowerPoint.AutoSize = true;
            this.chkPowerPoint.Checked = true;
            this.chkPowerPoint.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPowerPoint.Location = new System.Drawing.Point(12, 12);
            this.chkPowerPoint.Name = "chkPowerPoint";
            this.chkPowerPoint.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.chkPowerPoint.Size = new System.Drawing.Size(105, 17);
            this.chkPowerPoint.TabIndex = 0;
            this.chkPowerPoint.Text = "Final PowerPoint";
            this.chkPowerPoint.UseVisualStyleBackColor = true;
            // 
            // chkVocab
            // 
            this.chkVocab.AutoSize = true;
            this.chkVocab.Checked = true;
            this.chkVocab.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkVocab.Location = new System.Drawing.Point(520, 13);
            this.chkVocab.Name = "chkVocab";
            this.chkVocab.Size = new System.Drawing.Size(79, 17);
            this.chkVocab.TabIndex = 7;
            this.chkVocab.Text = "Vocabulary";
            this.chkVocab.UseVisualStyleBackColor = true;
            // 
            // chkQuestion
            // 
            this.chkQuestion.AutoSize = true;
            this.chkQuestion.Checked = true;
            this.chkQuestion.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkQuestion.Location = new System.Drawing.Point(441, 13);
            this.chkQuestion.Name = "chkQuestion";
            this.chkQuestion.Size = new System.Drawing.Size(73, 17);
            this.chkQuestion.TabIndex = 8;
            this.chkQuestion.Text = "Questions";
            this.chkQuestion.UseVisualStyleBackColor = true;
            // 
            // btnPublish
            // 
            this.btnPublish.Location = new System.Drawing.Point(12, 61);
            this.btnPublish.Name = "btnPublish";
            this.btnPublish.Size = new System.Drawing.Size(583, 23);
            this.btnPublish.TabIndex = 9;
            this.btnPublish.Text = "Publish Products";
            this.btnPublish.UseVisualStyleBackColor = true;
            this.btnPublish.Click += new System.EventHandler(this.btnPublish_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Output Folder";
            // 
            // chkYAML
            // 
            this.chkYAML.AutoSize = true;
            this.chkYAML.Checked = true;
            this.chkYAML.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkYAML.Location = new System.Drawing.Point(380, 12);
            this.chkYAML.Name = "chkYAML";
            this.chkYAML.Size = new System.Drawing.Size(55, 17);
            this.chkYAML.TabIndex = 11;
            this.chkYAML.Text = "YAML";
            this.chkYAML.UseVisualStyleBackColor = true;
            // 
            // PublishOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 90);
            this.Controls.Add(this.chkYAML);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnPublish);
            this.Controls.Add(this.chkQuestion);
            this.Controls.Add(this.chkVocab);
            this.Controls.Add(this.chkPowerPoint);
            this.Controls.Add(this.chkPNG);
            this.Controls.Add(this.btnFldBrowser);
            this.Controls.Add(this.txtPubDir);
            this.Controls.Add(this.chkMarkdown);
            this.Controls.Add(this.chkLatex);
            this.Controls.Add(this.chkPDF);
            this.Name = "PublishOptions";
            this.Text = "PublishOptions";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkPDF;
        private System.Windows.Forms.CheckBox chkLatex;
        private System.Windows.Forms.CheckBox chkMarkdown;
        private System.Windows.Forms.TextBox txtPubDir;
        private System.Windows.Forms.Button btnFldBrowser;
        private System.Windows.Forms.CheckBox chkPNG;
        private System.Windows.Forms.CheckBox chkPowerPoint;
        private System.Windows.Forms.CheckBox chkVocab;
        private System.Windows.Forms.CheckBox chkQuestion;
        private System.Windows.Forms.Button btnPublish;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkYAML;
        private System.Windows.Forms.FolderBrowserDialog fldBrowserDialog;
    }
}