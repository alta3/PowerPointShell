using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
            A3Presentation a3Presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            a3Presentation.GenerateFromYaml(yamlPath);
        }
        #endregion

        private void BtnFixAllMetadata_Click(object sender, RibbonControlEventArgs e)
        {
            A3Log log = new A3Log(A3Log.Operations.FixMetadata);
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            A3Presentation.FixMetadata(presentation, log);
        }
        private void BtnShowSlideMetadata_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            A3Slide.SetA3SlideFromPPTSlide(slide);
            A3Slide.ShowMetadataForm();
        }

        private void BtnPublish_Click(object sender, RibbonControlEventArgs e)
        {
            A3Log log = new A3Log(A3Log.Operations.PrePublish);
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            A3Presentation.FixMetadata(presentation, log);

            PublishOptions publish = new PublishOptions();
            publish.ShowDialog();
        }

        private void BtnNewBaseline_Click(object sender, RibbonControlEventArgs e)
        {
            A3Log log = new A3Log(A3Log.Operations.NewBaseline);
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            A3Presentation.NewBaseline(presentation, log);
        }

        private void BtnFillSubChaps_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            A3Presentation.FillSubChapters(presentation);
        }
    }
}