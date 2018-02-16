using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using YamlDotNet.Serialization;
using System.Windows.Forms;
using YamlDotNet.Serialization.NamingConventions;

namespace Alta3_PPA
{
    public partial class A3_Ribbon
    {
        private void A3_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnFixAllMetadata_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            A3Presentation.FixMetadata(presentation);
        }
        private void BtnShowSlideMetadata_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            A3Slide.ShowMetadataForm();
        }

        private void BtnPublish_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            A3Presentation.FixMetadata(presentation);

            System.IO.DirectoryInfo di = new DirectoryInfo(A3Globals.A3_PUBLISH);

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                dir.Delete(true);
            }

            A3Environment.StartUp();

            FirstChapter first = new FirstChapter();
            first.Show();
        }

        #region GenerateFromYaml
        private void BtnGenFromYaml_Click(object sender, RibbonControlEventArgs e)
        {
            // Open a file loader dialog
            this.OpenYamlForGen.ShowDialog();
        }
        private void OpenYamlForGen_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Initialize the environment
            A3Environment.Clean();

            // Create a logFile if errors occur store them at the indicated path.
            A3LogFile logFile = new A3LogFile();

            // Get the yaml path from the dialoge box
            string yamlPath = this.OpenYamlForGen.FileName;

            // Read the file into a string for processing
            string text = File.ReadAllText(yamlPath);

            // Lint the YAML file before attempting to deserialize the outline
            A3Yaml.Lint(logFile, text);

            // Log that we are about to try and desearilize this will help to see if our linting is effective or not
            logFile.WriteInfo("YAML lint complete. About to desearilize outline.");

            // Create the outline from the YAML file
            Deserializer deserializer = new DeserializerBuilder().WithNamingConvention(new CamelCaseNamingConvention()).Build();
            A3Outline outline = new A3Outline();
            try { outline = deserializer.Deserialize<A3Outline>(text); }
            catch (Exception ex) { logFile.WriteError(ex.Message); }

            // outline.Validate(logFile, "GenFromYaml");

            if (logFile.HasError())
            {
                string errorMsg = String.Concat("There were errors during the validation process.\r\n",
                    "The first error in the log is: ", logFile.Entries[0].Message,
                    "Please check the error file located at: ", logFile.Path, " for more information.\r\n",
                    "In order to successfully run the operation you must fix these errors.");
                MessageBox.Show(errorMsg, "Errors During Build", MessageBoxButtons.OK);
                this.OpenYamlForGen.Dispose();
                return;
            }

            // Open a copy of the blank PowerPoint in the current PowerPoint context
            PowerPoint.Presentation ppt = Globals.ThisAddIn.Application.Presentations.Open(A3Globals.BLANK_POWERPOINT, 0, 0, Microsoft.Office.Core.MsoTriState.msoTrue);
           
            // Save the powerpoint presentation to the working directory so that changes do not affect the model presentation
            string saveDir = String.Concat(A3Globals.A3_WORKING, "\\", outline.Course);
            try { Directory.CreateDirectory(saveDir); } catch { }
            string savePath = String.Concat(saveDir, "\\", outline.Course);
            int version = 0;
            while (File.Exists(String.Concat(savePath, ".pptm")))
            {
                version += 1;
                savePath = string.Concat(saveDir, "\\", outline.Course, version.ToString());
            }
            ppt.SaveAs(String.Concat(savePath, ".pptm"));

            // Generate the Presentation
            outline.GeneratePresentation(ppt);

            // Save the newly generated Presentation
            ppt.Save();

            // Alert the user the operation has concluded
            string message = String.Concat("The PowerPoint has been successfully built and saved to the following location:\r\n", savePath);
            MessageBox.Show(message, "Build Success", MessageBoxButtons.OK);
        }
        #endregion
    }
}