using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    class A3Publish
    {
        public static void PublishPNGs(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                string guid = slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text;
                string path = String.Concat(A3Globals.A3_PNGS, "\\", guid, ".png");
                slide.Export(path, "png", 1920, 1080);
            }
        }
        public static void PublishMarkdown(A3Outline outline)
        {
            foreach (A3Chapter chapter in outline.Chapters)
            {
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    foreach (A3Content content in subchapter.Slides)
                    {
                        A3Notes.ToMarkdown(content.Notes, content.ActiveGuid);
                    }
                }
            }
        }
        public static void PublishLaTex(PowerPoint.Presentation presentation, A3Outline outline)
        {
            if (!Directory.EnumerateFiles(A3Globals.A3_MARKDOWN).Any())
            {
                A3Publish.PublishMarkdown(outline);
            }
            outline.GenerateLaTex(presentation, outline);
        }
        public static void PublishPDF(PowerPoint.Presentation presentation, A3Outline outline)
        {
            if (!Directory.EnumerateFiles(A3Globals.A3_LATEX).Any())
            {
                A3Publish.PublishLaTex(presentation, outline);
            }

            const int ERROR_CANCELLED = 1223;

            ProcessStartInfo build = new ProcessStartInfo(String.Concat(@"powershell.exe -file ", A3Globals.A3_RESOURCE, @"\latex_builder.ps1 -root ", A3Globals.A3_LATEX))
            {
                UseShellExecute = true,
                Verb = "runas"
            };
            try
            {
                Process.Start(build);
            }
            catch (Win32Exception ex)
            {
                if (ex.NativeErrorCode == ERROR_CANCELLED)
                    MessageBox.Show("You must select 'Yes' to the UAC prompt to continue");
                else
                    throw;
            }
        }
        public static void PublishQuestions()
        { }
        public static void PublishVocabulary()
        { }
        public static void PublishYAML(A3LogFile logFile, A3Outline outline)
        {
            A3Yaml.ProduceYaml(logFile, outline);
        }
    }
}
