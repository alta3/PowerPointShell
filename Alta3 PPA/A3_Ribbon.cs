using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public partial class A3_Ribbon
    {
        private void A3_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        #region GenerateFromYaml
        private void BtnGenFromYaml_Click(object sender, RibbonControlEventArgs e)
        {
            OpenYamlForGen.ShowDialog();
        }
        private void OpenYamlForGen_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string yamlPath = OpenYamlForGen.FileName;
            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            presentation.GenerateFromYaml(yamlPath);
        }
        #endregion

        private void BtnFixAllMetadata_Click(object sender, RibbonControlEventArgs e)
        {
            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            presentation.FixMetadata(true, true);
        }
        private void BtnShowSlideMetadata_Click(object sender, RibbonControlEventArgs e)
        {
            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            A3Environment.A3SLIDE = new A3Slide(slide);
            A3Environment.A3SLIDE.ShowMetadataForm();
        }

        private void BtnPublish_Click(object sender, RibbonControlEventArgs e)
        {
            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            presentation.FixMetadata(false, false);

            PublishOptions publish = new PublishOptions();
            publish.ShowDialog();
        }

        private void BtnNewBaseline_Click(object sender, RibbonControlEventArgs e)
        {
            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            presentation.NewBaseLine();
        }

        private void BtnFillSubChaps_Click(object sender, RibbonControlEventArgs e)
        {
            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            presentation.FillSubChapters();
        }
    }
}