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
            A3Environment.Clean();
            // Grab the current global infer states and then switch them to true while generating the powerpoint from yaml
            bool inferState = A3Environment.ALLOW_INFER_FROM_SLIDE;
            bool defaultInferState = A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE;
            A3Environment.ALLOW_INFER_FROM_SLIDE = true;
            A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE = true;
            A3Log log = new A3Log(A3Log.Operations.GenerateFromYaml);
            string yamlPath = OpenYamlForGen.FileName;
            A3Yaml.GenerateFromYaml(log, yamlPath);
            // Return to the original state of the infer global states
            A3Environment.QUIT_FROM_CURRENT_LOOP = false;
            A3Environment.ALLOW_INFER_FROM_SLIDE = inferState;
            A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE = defaultInferState;
            A3Environment.Clean();
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