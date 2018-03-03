﻿using System;
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
            A3Slide.SetActiveSlide(slide);
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
            first.ShowDialog();
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

            // Generate from YAML file
            A3Yaml.GenerateFromYaml(logFile, yamlPath);
        }
        #endregion
    }
}