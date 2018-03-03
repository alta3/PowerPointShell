using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    class A3Presentation
    {
        public List<A3Slide> Slides { get; set; }

        public A3Presentation(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                this.Slides.Add(new A3Slide(slide));
            }
        }

        public static void FixMetadata(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (!A3Globals.QUIT_FROM_CURRENT_LOOP)
                {
                    A3Slide.SetActiveSlide(slide);
                    A3Slide.FixNullMetadata(true);
                }
            }
            A3Globals.QUIT_FROM_CURRENT_LOOP = false;
        }
        public A3Outline ToOutline(A3LogFile logFile)
        {
            A3Outline outline = new A3Outline();

            // Based on
            this.ValidateCourseCount(logFile);

            return outline;
        }

        private void ValidateCourseCount(A3LogFile logFile)
        {
            List<string> guids = new List<string>();
            foreach (A3Slide slide in this.Slides)
            {
                if (slide.Type.ToLower() == "course")
                {
                    guids.Add(slide.ActiveGuid);
                }
            }
            if (guids.Count > 1)
            {
                string message = "More than one course slide found. The following slides active guid reports it is currently a course slide:\r\n";
                foreach (string guid in guids)
                {
                    logFile.WriteError(String.Concat(message, "ACTIVE GUID: ", guid, "\r\n"));
                }
            }
            else if (guids.Count < 1)
            {
                logFile.WriteError("No course slide found in the metadata fields\r\n");
            }
        }
    }
}
