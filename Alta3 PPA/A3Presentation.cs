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
            this.Slides = new List<A3Slide>();
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                this.Slides.Add(new A3Slide(slide));
            }
        }

        public static void FillSubChapters(PowerPoint.Presentation presentation)
        {
            string subChapName = "Contents";
            bool after_chap = false;
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                A3Slide.SetA3SlideFromPPTSlide(slide);
                if (A3Globals.A3SLIDE.Type.ToLower() == "chapter")
                {
                    subChapName = "Contents";
                    after_chap = true;
                    continue;
                }
                if (A3Globals.A3SLIDE.Type.ToLower() == "content" && after_chap)
                {
                    if (A3Globals.A3SLIDE.Subchapter != "Contents" && A3Globals.A3SLIDE.Subchapter != subChapName)
                    {
                        subChapName = A3Globals.A3SLIDE.Subchapter;
                    }
                    else if (A3Globals.A3SLIDE.Subchapter == "Contents" && A3Globals.A3SLIDE.Subchapter != subChapName)
                    {
                        A3Globals.A3SLIDE.ChapSub = String.Concat(A3Globals.A3SLIDE.Chapter, @": ", subChapName);
                        A3Globals.A3SLIDE.WriteChapSub();
                    }
                }
                if (A3Globals.A3SLIDE.Type.ToLower() == "question")
                {
                    return;
                }
            }
        }
        public static void NewBaseline(PowerPoint.Presentation presentation, A3LogFile logFile)
        {
            string timestamp = DateTimeOffset.Now.DateTime.ToString().Replace('/', '.').Replace(':', '.').Replace(' ', '-');
            string saveDir = String.Concat(A3Globals.A3_WORKING, @"\new-baseline");
            try { Directory.CreateDirectory(saveDir); } catch { }
            string savePath = String.Concat(saveDir, @"\new-baseline", timestamp);
            int version = 0;
            while (File.Exists(String.Concat(savePath, ".pptm")))
            {
                version += 1;
                savePath = string.Concat(savePath, version.ToString());
            }
            presentation.SaveAs(String.Concat(savePath, ".pptm"));

            A3Globals.ALLOW_INFER_FROM_SLIDE = true;
            A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE = true;
            A3Globals.ENFORCE_CHAP_SUB_SPLITTING = false;
            string chapterName = null;
            bool before_chap = true;
            bool after_question = false;
            DialogResult dialogResult = MessageBox.Show("About to create a new baseline presentation. This will be saved to the working directory. A message box will pop up at the end of the process, please do not attempt to interact with PowerPoint while this operation completes.", "About to run!", MessageBoxButtons.OK);
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (!A3Globals.QUIT_FROM_CURRENT_LOOP)
                {
                    A3Slide.NewBaseline(slide, chapterName, before_chap, after_question,logFile);
                }
                if (A3Globals.A3SLIDE.Type.ToLower() == "chapter")
                {
                    chapterName = A3Globals.A3SLIDE.Chapter;
                    before_chap = false;
                }
                else if (A3Globals.A3SLIDE.Type.ToLower() == "question")
                {
                    after_question = true;
                }

            }
            A3Globals.QUIT_FROM_CURRENT_LOOP = false;
            A3Globals.ALLOW_INFER_FROM_SLIDE = false;
            A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE = false;
            A3Globals.ENFORCE_CHAP_SUB_SPLITTING = true;
            DialogResult dialog = MessageBox.Show("Finished running new baseline.", "Completed!", MessageBoxButtons.OK);
        }
        public static void FixMetadata(PowerPoint.Presentation presentation, A3LogFile logFile)
        {
            DialogResult dialogResult = MessageBox.Show("Allow the program to Infer Metadata information from the slide rather than forcing user to ensure the data?", "Allow Infer?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                A3Globals.ALLOW_INFER_FROM_SLIDE = true;
                DialogResult dialogResult2 = MessageBox.Show("If the program cannot Infer the metadata from the slide would you like to allow the program to default to types of metadata? -- Default: Content Slide; will make the textboxes for other information if not already present.", "Allow Default Infer?", MessageBoxButtons.YesNo);
                if (dialogResult2 == DialogResult.Yes)
                {
                    A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE = true;
                }
            }
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (!A3Globals.QUIT_FROM_CURRENT_LOOP)
                {
                    A3Slide.SetA3SlideFromPPTSlide(slide);
                    A3Slide.FixNullMetadata(true, logFile);
                }
            }
            A3Globals.QUIT_FROM_CURRENT_LOOP = false;
            A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE = false;
            A3Globals.ALLOW_INFER_FROM_SLIDE = false;
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
            a3SlidesCourse = this.Slides.FindAll(a3Slide => a3Slide.Type.ToUpper() == "COURSE");
            if (a3SlidesCourse == null)
            {
                logFile.WriteError("No course slide found in the metadata fields\r\n");
                return a3SlidesCourse;
            }
            if (a3SlidesCourse.Count > 1)
            {
                string message = "More than one course slide found. The following slides active guid reports it is currently a course slide:\r\n";
                foreach (A3Slide a3Slide in a3SlidesCourse)
                {
                    logFile.WriteError(String.Concat(message, "ACTIVE GUID: ", a3Slide.ActiveGuid, "\r\n"));
                }
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
            a3Slides = this.Slides.FindAll(a3Slide => (a3Slide.Type.ToLower() == "content" || a3Slide.Type.ToLower() == "no-pub") && a3Slide.Chapter == a3Chapter.Title);

            List<string> subTitles = new List<string>();
            foreach (A3Slide a3Slide in a3Slides)
            {
                try { subTitles.Add(a3Slide.Subchapter); }
                catch { logFile.WriteError(String.Concat("FAILED TO WRITE SLIDE INDEX: ", a3Slide.Slide.SlideIndex.ToString(), " TO DECK. CHECK THE METADATA!")); }
            }

            foreach (string subTitle in subTitles.Distinct().ToList())
            {
                List<A3Slide> a3SubSlides = new List<A3Slide>();
                a3SubSlides = a3Slides.FindAll(slide => slide.Subchapter == subTitle);
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
                logFile.WriteWarn(String.Concat(a3Chapter.Title, " NO SUBCHAPTERS FOUND"));
            }
        }
    }
}
