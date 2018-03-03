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

            List<A3Slide> a3SlidesCourse = this.GetCourse(logFile);
            try { outline.Course = a3SlidesCourse[0].Title; }
            catch { outline.Course = "!!!ERROR!!! -- SEE THE LOGS"; }

            outline.Chapters = new List<A3Chapter>();
            this.GetChapters(outline, logFile);

            foreach (A3Chapter a3Chapter in outline.Chapters)
            {
                a3Chapter.Subchapters = new List<A3Subchapter>();
                this.GetSubChapters(a3Chapter, logFile);
            }

            return outline;
        }

        private List<A3Slide> GetCourse(A3LogFile logFile)
        {
            List<A3Slide> a3SlidesCourse = new List<A3Slide>();
            a3SlidesCourse = this.Slides.FindAll(a3Slide => a3Slide.Type.ToLower() == "course");
            if (a3SlidesCourse.Count > 1)
            {
                string message = "More than one course slide found. The following slides active guid reports it is currently a course slide:\r\n";
                foreach (A3Slide a3Slide in a3SlidesCourse)
                {
                    logFile.WriteError(String.Concat(message, "ACTIVE GUID: ", a3Slide.ActiveGuid, "\r\n"));
                }
            }
            else if (a3SlidesCourse.Count < 1)
            {
                logFile.WriteError("No course slide found in the metadata fields\r\n");
            }
            return a3SlidesCourse;
        }
        private void GetChapters(A3Outline outline, A3LogFile logFile)
        {
            List<A3Slide> a3SlidesChapters = new List<A3Slide>();
            a3SlidesChapters = this.Slides.FindAll(a3Slide => a3Slide.Type.ToLower() == "chapter");
            foreach (A3Slide a3Chapter in a3SlidesChapters)
            {
                outline.Chapters.Add((A3Chapter)a3Chapter.TypeConversion());
            }
            if (outline.Chapters.Count < 1)
            {
                logFile.WriteWarn("NO CHAPTERS FOUND!!!");
            }
        }
        private void GetSubChapters(A3Chapter a3Chapter, A3LogFile logFile)
        {
            List<A3Slide> a3Slides = new List<A3Slide>();
            a3Slides = this.Slides.FindAll(a3Slide => (a3Slide.Type.ToLower() == "content" || a3Slide.Type.ToLower() == "no-pub") && a3Slide.ChapSub.Split(':')[0].Trim().ToLower() == a3Chapter.Title.Trim().ToLower());

            List<string> subTitles = new List<string>();
            foreach (A3Slide a3Slide in a3Slides)
            {
                subTitles.Add(a3Slide.ChapSub.Split(':')[1].Trim());
            }

            foreach (string subTitle in subTitles.Distinct().ToList())
            {
                List<A3Slide> a3SubSlides = new List<A3Slide>();
                a3SubSlides = a3Slides.FindAll(slide => slide.ChapSub.Split(':')[1].Trim() == subTitle);
                List<A3Content> a3SubContentSlides = new List<A3Content>();
                foreach (A3Slide a3SubSlide in a3SubSlides)
                {
                    a3SubContentSlides.Add((A3Content)a3SubSlide.TypeConversion());
                }
                A3Subchapter a3Subchapter = new A3Subchapter
                {
                    Title = subTitle,
                    Slides = a3SubContentSlides,
                    Questions = new List<A3Question>()
                };
                a3Chapter.Subchapters.Add(a3Subchapter);
            }
            if (a3Chapter.Subchapters.Count < 1)
            {
                logFile.WriteWarn("NO SUBCHAPTERS FOUND");
            }
        }
    }
}
