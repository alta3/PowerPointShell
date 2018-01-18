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

        #region TODO: IMPLEMENT
        // Validate Presentation 
        public static void Validate(A3LogFile logFile, PowerPoint.Presentation presentation)
        {

        }
        public static void ValidateMetadata(A3LogFile logFile, PowerPoint.Presentation presentation)
        {

        }
        public static void ValidateStructure(A3LogFile logFile, PowerPoint.Presentation presentation)
        {
            // Construct the current presentation and then validate
            A3Presentation a3Presentation = new A3Presentation(presentation);
        }
        private static void ValidateCourseSlideNumber(A3LogFile logFile, A3Presentation a3Presentation)
        {
            int courseCount = 0;
            List<string> guids = new List<string>();
            foreach (A3Slide slide in a3Presentation.Slides)
            {
                if (slide.Type.ToLower() == "course")
                {
                    guids.Add(slide.ActiveGuid);
                    courseCount++;
                }
            }
            if (courseCount > 1)
            {
                string message = "More than one course slide found. The following slides active guid reports it is currently a course slide:\r\n";
                foreach (string guid in guids)
                {
                    logFile.WriteError(String.Concat(message, "ACTIVE GUID: ", guid, "\r\n"));
                }
            }
            else if (courseCount < 1)
            {
                logFile.WriteError("No course slide found in the metadata fields\r\n");
            }
        }
        #endregion
    }
}
