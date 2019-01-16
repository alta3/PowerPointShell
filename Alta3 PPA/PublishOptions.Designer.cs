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
            this.chkMarkdown = new System.Windows.Forms.CheckBox();
            this.txtPubDir = new System.Windows.Forms.TextBox();
            this.btnFldBrowser = new System.Windows.Forms.Button();
            this.chkPNG = new System.Windows.Forms.CheckBox();
            this.btnPublish = new System.Windows.Forms.Button();
            this.chkYAML = new System.Windows.Forms.CheckBox();
            this.fldBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // chkPDF
            // 
            this.chkPDF.AutoSize = true;
            this.chkPDF.Checked = true;
            this.chkPDF.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPDF.Location = new System.Drawing.Point(12, 9);
            this.chkPDF.Name = "chkPDF";
            this.chkPDF.Size = new System.Drawing.Size(47, 17);
            this.chkPDF.TabIndex = 0;
            this.chkPDF.Text = "PDF";
            this.chkPDF.UseVisualStyleBackColor = true;
            // 
            // chkMarkdown
            // 
            this.chkMarkdown.AutoSize = true;
            this.chkMarkdown.Checked = true;
            this.chkMarkdown.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMarkdown.Location = new System.Drawing.Point(215, 12);
            this.chkMarkdown.Name = "chkMarkdown";
            this.chkMarkdown.Size = new System.Drawing.Size(76, 17);
            this.chkMarkdown.TabIndex = 2;
            this.chkMarkdown.Text = "Markdown";
            this.chkMarkdown.UseVisualStyleBackColor = true;
            // 
            // txtPubDir
            // 
            this.txtPubDir.Location = new System.Drawing.Point(12, 35);
            this.txtPubDir.Name = "txtPubDir";
            this.txtPubDir.Size = new System.Drawing.Size(319, 20);
            this.txtPubDir.TabIndex = 3;
            // 
            // btnFldBrowser
            // 
            this.btnFldBrowser.Location = new System.Drawing.Point(337, 32);
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
            this.chkPNG.Location = new System.Drawing.Point(107, 12);
            this.chkPNG.Name = "chkPNG";
            this.chkPNG.Size = new System.Drawing.Size(54, 17);
            this.chkPNG.TabIndex = 5;
            this.chkPNG.Text = "PNGs";
            this.chkPNG.UseVisualStyleBackColor = true;
            // 
            // btnPublish
            // 
            this.btnPublish.Location = new System.Drawing.Point(12, 61);
            this.btnPublish.Name = "btnPublish";
            this.btnPublish.Size = new System.Drawing.Size(400, 23);
            this.btnPublish.TabIndex = 9;
            this.btnPublish.Text = "Publish Products";
            this.btnPublish.UseVisualStyleBackColor = true;
            this.btnPublish.Click += new System.EventHandler(this.btnPublish_Click);
            // 
            // chkYAML
            // 
            this.chkYAML.AutoSize = true;
            this.chkYAML.Checked = true;
            this.chkYAML.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkYAML.Location = new System.Drawing.Point(356, 9);
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
            this.ClientSize = new System.Drawing.Size(423, 90);
            this.Controls.Add(this.chkYAML);
            this.Controls.Add(this.btnPublish);
            this.Controls.Add(this.chkPNG);
            this.Controls.Add(this.btnFldBrowser);
            this.Controls.Add(this.txtPubDir);
            this.Controls.Add(this.chkMarkdown);
            this.Controls.Add(this.chkPDF);
            this.Name = "PublishOptions";
            this.Text = "PublishOptions";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkPDF;
        private System.Windows.Forms.CheckBox chkMarkdown;
        private System.Windows.Forms.TextBox txtPubDir;
        private System.Windows.Forms.Button btnFldBrowser;
        private System.Windows.Forms.CheckBox chkPNG;
        private System.Windows.Forms.Button btnPublish;
        private System.Windows.Forms.CheckBox chkYAML;
        private System.Windows.Forms.FolderBrowserDialog fldBrowserDialog;
    }
}