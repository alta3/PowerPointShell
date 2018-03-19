using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public partial class PublishOptions : Form
    {
        public PublishOptions()
        {
            InitializeComponent();
            txtPubDir.Text = A3Globals.A3_PUBLISH;
        }

        private void btnPublish_Click(object sender, EventArgs e)
        {
            A3Environment.StartUp();

            A3LogFile logFile = new A3LogFile();

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            A3Presentation a3Presentation = new A3Presentation(presentation);
            A3Outline outline = a3Presentation.ToOutline(logFile);

            if (chkPowerPoint.Checked) { A3Publish.PublishPowerPoint(); }
            if (chkPNG.Checked) { A3Publish.PublishPresentationPNGs(presentation); A3Publish.PublishBookPNGs(presentation); }
            if (chkYAML.Checked) { A3Publish.PublishYAML(logFile, outline); }
            if (chkMarkdown.Checked) { A3Publish.PublishMarkdown(outline); }
            if (chkLatex.Checked) { A3Publish.PublishLaTex(presentation, outline); }
            if (chkPDF.Checked) { A3Publish.PublishPDF(presentation, outline); }           
            if (chkQuestion.Checked) { A3Publish.PublishQuestions(); }
            if (chkVocab.Checked) { A3Publish.PublishVocabulary(); }

            this.Dispose();
        }

        private void btnFldBrowser_Click(object sender, EventArgs e)
        {
            if (fldBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                A3Globals.A3_PUBLISH = fldBrowserDialog.SelectedPath;
                txtPubDir.Text = A3Globals.A3_PUBLISH;
            }
        }
    }
}
