using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.RepresentationModel;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;

namespace PowerPointShell
{
    [Cmdlet("A3", "Generate")]
    public class PPSGenerateCmdlet : PSCmdlet
    {
        // The parameters for A3-Generate Cmdlet
        [Parameter(Position = 0)]
        public string YamlFile { get; set; }

        protected override void ProcessRecord()
        {
            // Get the YAML path
            string yamlPath = null;
            string currentPath = this.SessionState.Path.CurrentLocation.Path;

            if (YamlFile.StartsWith(".//") || YamlFile.StartsWith(".\\") || !YamlFile.StartsWith("C:"))
            {
                char[] remove = { '.', '/', '\\' };
                yamlPath = string.Concat(currentPath, ".//", YamlFile.TrimStart(remove));
            }
            else
            {
                yamlPath = YamlFile;
            }

            string text = File.ReadAllText(yamlPath);

            // Create an outline object from the YAML file by utilizing yamldotnet library
            ProgressRecord pr = new ProgressRecord(0, "Generating PowerPoint Presentation From YAML Outline", "Initializing Outline Object From YAML File");
            WriteProgress(pr);
            var input = new StringReader(text);
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(new CamelCaseNamingConvention())
                .Build();
            PowerPointShell.Outline outline = deserializer.Deserialize<Outline>(input);

            // Create a new instance of the powerpoint application
            pr.StatusDescription = "Opening PowerPoint";
            WriteProgress(pr);
            var powerpoint = new PowerPoint.Application();

            // Open a blank presentation that is not visable
            PowerPoint.Presentation ppt = powerpoint.Presentations.Open(GlobalVars.BLANK_POWERPOINT,0, 0, 0);

            // Save the powerpoint presentation to the working directory so that changes do not affect the model presentation
            string wrkPath = GlobalVars.WORKING_PATH;
            string pptName = outline.Course;
            string savePath = string.Concat(wrkPath, "\\", pptName, ".pptm");
            int version = 0;
            while (File.Exists(savePath))
            {
                version += 1;
                var ver = version.ToString();
                savePath = string.Concat(wrkPath, "\\", pptName, ver, ".pptm");
            }
            ppt.SaveAs(savePath);
            pr.StatusDescription = "PowerPoint Presentation Opened & Saved As: " + savePath;
            WriteProgress(pr);

            // Generate the presentation from the outline given
            pr.StatusDescription = "Creating Slides From Outline -- This Will Take A Few Seconds";
            WriteProgress(pr);
            outline.GeneratePresentation(ppt);

            // Save the presentation close it then reopen as visible for editing
            ppt.Save();
            ppt.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ppt);

            pr.RecordType = ProgressRecordType.Completed;
            WriteProgress(pr);
            ppt = powerpoint.Presentations.Open(savePath, 0, 0, Microsoft.Office.Core.MsoTriState.msoTrue);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ppt);
        }
    }
}
