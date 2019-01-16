using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public partial class PublishOptions : Form
    {
        public PublishOptions()
        {
            InitializeComponent();
            txtPubDir.Text = A3Environment.A3_PUBLISH;
        }

        private void btnPublish_Click(object sender, EventArgs e)
        {
            A3Environment.StartUp();

            A3Log log = new A3Log(A3Log.Operations.Publish);

            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            A3Outline outline = presentation.GenerateOutline(log);

            if (chkPNG.Checked) presentation.PublishPNGs();
            if (chkMarkdown.Checked) presentation.PublishMarkdown();
            if (chkPDF.Checked) presentation.PublishPDF();
            if (chkYAML.Checked) presentation.PublishYaml();

            Dispose();
        }

        private void btnFldBrowser_Click(object sender, EventArgs e)
        {
            if (fldBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                A3Environment.A3_PUBLISH = fldBrowserDialog.SelectedPath;
                txtPubDir.Text = A3Environment.A3_PUBLISH;
            }
        }
    }
}
