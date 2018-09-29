using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA {
    class A3Presentation {
        public List<A3Slide> Slides { get; set; }

        public A3Presentation(PowerPoint.Presentation presentation) {
            this.Slides = new List<A3Slide>();
            foreach (PowerPoint.Slide slide in presentation.Slides) {
                this.Slides.Add(new A3Slide(slide));
            }
        }

        public static void SavePresentation(PowerPoint.Presentation presentation, string directory, string fileName) {
            try { Directory.CreateDirectory(directory); } catch { }
            string savePath = String.Concat(directory, "\\", fileName);
            int version = 0;
            while (File.Exists(String.Concat(savePath, ".pptm"))) {
                version += 1;
                savePath = string.Concat(directory, "\\", fileName, version.ToString());
            }
            presentation.SaveAs(String.Concat(savePath, ".pptm"));
        }

        public static void FillSubChapters(PowerPoint.Presentation presentation) {
            string chapName = "";
            string subChapName = "Contents";
            foreach (PowerPoint.Slide slide in presentation.Slides) {
                A3Slide.SetA3SlideFromPPTSlide(slide);
                if (A3Globals.A3SLIDE.Type.ToLower() == "chapter") {
                    chapName = A3Globals.A3SLIDE.Title;
                    subChapName = "Contents";
                    A3Globals.SLIDE_ITTERATION_AFTER_CHAPTER = true;
                }
                if (A3Globals.A3SLIDE.Type.ToLower() == "content" && A3Globals.SLIDE_ITTERATION_AFTER_CHAPTER) {
                    if (A3Globals.A3SLIDE.Subchapter != subChapName && A3Globals.A3SLIDE.Subchapter != null)
                    {
                        subChapName = A3Globals.A3SLIDE.Subchapter;
                    }

                    A3Globals.A3SLIDE.ChapSub = String.Concat(chapName, @": ", subChapName);
                    A3Globals.A3SLIDE.WriteChapSub();
                }
                if (A3Globals.A3SLIDE.Type.ToLower() == "question") {
                    return;
                }
            }
        }
        public static void NewBaseline(PowerPoint.Presentation presentation, A3LogFile logFile) {
            string fileName = String.Concat("new-baseline-", DateTimeOffset.Now.DateTime.ToString().Replace('/', '.').Replace(':', '.').Replace(' ', '-'));
            string saveDir = String.Concat(A3Globals.A3_WORKING, @"\new-baseline");
            SavePresentation(presentation, saveDir, fileName);

            A3Environment.DefaultInfer();
            foreach (PowerPoint.Slide slide in presentation.Slides) {
                A3Slide.NewBaseline(slide, logFile);
                if (A3Globals.A3SLIDE.Type == A3Slide.TypeStrings[(int)A3Slide.SlideType.CHAPTER]) {
                    try { A3Globals.SLIDE_ITTERATION_CURRENT_CHAPTER = A3Globals.A3SLIDE.Title; }
                    catch { A3Globals.SLIDE_ITTERATION_CURRENT_CHAPTER = "Chapter"; }
                    A3Globals.SLIDE_ITTERATION_CURRENT_SUBCHAPTER = "Contents";
                    A3Globals.SLIDE_ITTERATION_AFTER_CHAPTER = true;
                }
                else if (A3Globals.A3SLIDE.Type == A3Slide.TypeStrings[(int)A3Slide.SlideType.QUESTION]) {
                    A3Globals.SLIDE_ITTERATION_AFTER_QUESTION = true;
                }

            }
            A3Environment.Clean();
            DialogResult dialog = MessageBox.Show("Finished running new baseline.", "Completed!", MessageBoxButtons.OK);
        }
        public static void FixMetadata(PowerPoint.Presentation presentation, A3LogFile logFile) {
            A3Environment.Clean();
            foreach (PowerPoint.Slide slide in presentation.Slides) {
                if (A3Globals.QUIT_FROM_CURRENT_LOOP) {
                    break;
                }
                A3Slide.SetA3SlideFromPPTSlide(slide);
                A3Slide.FixMetadataErrors(false, logFile);
            }
            A3Environment.Clean();
        }

        #region Outline
        public A3Outline ToOutline(A3LogFile logFile) {
            A3Outline outline = new A3Outline();

            List<A3Slide> a3SlidesCourse = this.GetCourse(logFile);
            try { outline.Name = a3SlidesCourse[0].Title; }
            catch { outline.Name = "!!!ERROR!!! -- SEE THE LOGS"; }

            outline.Chapters = new List<A3Chapter>();
            this.GetChapters(outline, logFile);

            foreach (A3Chapter a3Chapter in outline.Chapters) {
                a3Chapter.Subchapters = new List<A3Subchapter>();
                this.GetSubChapters(a3Chapter, logFile);
            }

            return outline;
        }

        private List<A3Slide> GetCourse(A3LogFile logFile) {
            List<A3Slide> a3SlidesCourse = new List<A3Slide>();
            a3SlidesCourse = this.Slides.FindAll(a3Slide => a3Slide.Type.ToUpper() == "COURSE");
            if (a3SlidesCourse == null) {
                logFile.WriteError("No course slide found in the metadata fields\r\n");
                return a3SlidesCourse;
            }
            if (a3SlidesCourse.Count > 1) {
                string message = "More than one course slide found. The following slides active guid reports it is currently a course slide:\r\n";
                foreach (A3Slide a3Slide in a3SlidesCourse) {
                    logFile.WriteError(String.Concat(message, "ACTIVE GUID: ", a3Slide.Guid, "\r\n"));
                }
            }
            return a3SlidesCourse;
        }
        private void GetChapters(A3Outline outline, A3LogFile logFile) {
            List<A3Slide> a3SlidesChapters = new List<A3Slide>();
            a3SlidesChapters = this.Slides.FindAll(a3Slide => a3Slide.Type.ToLower() == "chapter");
            foreach (A3Slide a3Chapter in a3SlidesChapters) {
                outline.Chapters.Add((A3Chapter)a3Chapter.ObjectConversion());
            }
            if (outline.Chapters.Count < 1) {
                logFile.WriteWarn("NO CHAPTERS FOUND!!!");
            }
        }
        private void GetSubChapters(A3Chapter a3Chapter, A3LogFile logFile) {
            List<A3Slide> a3Slides = new List<A3Slide>();
            a3Slides = this.Slides.FindAll(a3Slide => (a3Slide.Type.ToLower() == "content" || a3Slide.Type.ToLower() == "nopub") && a3Slide.Chapter == a3Chapter.Title);

            List<string> subTitles = new List<string>();
            foreach (A3Slide a3Slide in a3Slides) {
                try {
                    subTitles.Add(a3Slide.Subchapter);
                }
                catch {
                    logFile.WriteError(String.Concat("FAILED TO WRITE SLIDE INDEX: ", a3Slide.Slide.SlideIndex.ToString(), " TO DECK. CHECK THE METADATA!"));
                }
            }

            foreach (string subTitle in subTitles.Distinct().ToList()) {
                List<A3Slide> a3SubSlides = new List<A3Slide>();
                a3SubSlides = a3Slides.FindAll(slide => slide.Subchapter == subTitle);
                List<A3Content> a3SubContentSlides = new List<A3Content>();
                foreach (A3Slide a3SubSlide in a3SubSlides) {
                    a3SubContentSlides.Add((A3Content)a3SubSlide.ObjectConversion());
                }
                A3Subchapter a3Subchapter = new A3Subchapter {
                    Title = subTitle,
                    Slides = a3SubContentSlides,
                    Questions = new List<A3Question>()
                };
                a3Chapter.Subchapters.Add(a3Subchapter);
            }
            if (a3Chapter.Subchapters.Count < 1) {
                logFile.WriteWarn(String.Concat(a3Chapter.Title, " NO SUBCHAPTERS FOUND"));
            }
        }
        #endregion
    }
}
