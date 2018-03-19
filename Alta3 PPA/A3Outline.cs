using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using YamlDotNet.Core;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NodeDeserializers;
using Markdig;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Outline
    {
        #region Outline Properites
        public string Id { get; set; }
        public string Course { get; set; }
        public List<A3Content> TOC { get; set; }
        public List<A3Chapter> Chapters { get; set; }
        public List<A3Lab> Labs { get; set; }
        #endregion

        #region Generate Presentation
        public void GeneratePresentation(PowerPoint.Presentation presentation)
        {
            this.GenerateCourseSlide(presentation);
            this.GenerateTOCSlide(presentation);
            presentation.SectionProperties.AddBeforeSlide(1, this.Course);
            this.GenerateChapters(presentation);
            this.GenerateEndOfDeckSlide(presentation);
            this.GenerateQuizSlide(presentation);
        }

        private void GenerateChapters(PowerPoint.Presentation presentation)
        {
            int chapterCount = 1;
            foreach (A3Chapter chapter in this.Chapters)
            {
                chapter.Generate(presentation, chapterCount);
                chapterCount += 1;
            }
        }
        private void GenerateCourseSlide(PowerPoint.Presentation presentation)
        {
            // Insert the course slide from the model PowerPoint
            presentation.Slides.AddSlide(presentation.Slides.Count + 1, presentation.SlideMaster.CustomLayouts[1]);

            // Change the title to the course title given in the yaml file
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = this.Course,
                Type = "COURSE",
                ActiveGuid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateEndOfDeckSlide(PowerPoint.Presentation presentation)
        {
            // Insert a title slide from the model PowerPoint
            presentation.Slides.AddSlide(presentation.Slides.Count + 1, presentation.SlideMaster.CustomLayouts[4]);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);

            // Change the title, chapsub, type, and active guid to accurately reflect what is happening
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "End of Deck",
                ChapSub = String.Concat(this.Course, ": End Of Deck"),
                ActiveGuid = Guid.NewGuid().ToString(),
                Type = "CONTENT"
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateTOCSlide(PowerPoint.Presentation presentation)
        {
            // Insert a split slide from the model PowerPoint
            presentation.Slides.AddSlide(presentation.Slides.Count + 1, presentation.SlideMaster.CustomLayouts[4]);

            // Populate the appropriate values of the slide deck here
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Table of Contents",
                ChapSub = String.Concat(this.Course, ": TOC"),
                Type = "TOC",
                ActiveGuid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();

            // TODO: Create a linked list to the first chapter of each day and colorize the results
        }
        private void GenerateQuizSlide(PowerPoint.Presentation presentation)
        {
            // Insert a question slide from the model PowerPoint
            presentation.Slides.InsertFromFile(A3Globals.MODEL_POWERPOINT, presentation.Slides.Count + 1, 6);       

            // Ensure the title is Knowledge Check and move on 
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Knowledge Check",
                Type = "QUESTION",
                ActiveGuid = Guid.NewGuid().ToString(),
            };
            a3ActiveSlide.WriteFromMemory();
            presentation.SectionProperties.AddBeforeSlide(presentation.Slides.Count, "Knowledge Check");

        }
        #endregion

        #region Generate LaTex
        public void GenerateLaTex(PowerPoint.Presentation presentation, A3Outline outline)
        {
            this.GenerateLaTexMain();
            this.GenerateLaTexChapters();
            this.GenerateLaTexSubchapters(outline);
        }

        private void GenerateLaTexMain()
        {
            List<string> main = new List<string>
            {
                @"\documentclass[openany]{book}",
                @"",
                @"\usepackage{float}",
                @"\usepackage{graphicx}",
                @"\usepackage{fancyhdr}",
                @"\usepackage{hyperref}",
                @"\usepackage[utf8]{inputenc}",
                @"\usepackage[section] {placeins}",
                @"\usepackage[top = 0.5in, bottom = 0.5in, bmargin = 0.5in, left = 0.6in, right = 0.6in, headsep = 3mm ]{geometry}",
                @"",
                @"\providecommand{\tightlist}{\setlength{\itemsep}{0pt}\setlength{\parskip}{0pt}}",
                @"",
                @"\pagestyle{fancy}",
                @"\fancyfoot{}",
                @"\fancyfoot[C]{\thepage}",
                @"\fancyfoot[LR]{\copyright \ Stuart Feeser}",

                "",
                @"\begin{document}",
                @"",
                @"\begin{titlepage}",
                @"\vspace*{55mm}",
                @"\centering",
                String.Concat(@"\includegraphics[width=.5\textwidth]{", "\"", A3Globals.A3_RESOURCE.Replace('\\','/'), @"/a3logo", "\"}"),
                @"\linebreak",
                @"\linebreak",
                String.Concat(@"{\Huge\textbf{", this.Course, @"}}"),
                @"\linebreak",
                @"\linebreak",
                @"{\Large Alta3 Research, Inc.}",
                @"\\",
                @"{\today}",
                @"\vfill",
                @"\begin{flushright}",
                @"Alta3 Research, Inc. \\",
                @"sfeeser@alta3.com \\",
                @"https://alta3.com",
                @"\end{flushright}",
                @"\end{titlepage}",

                "",
                @"\frontmatter",
                @"\tableofcontents",

                "",
                @"\mainmatter"
            };
            foreach (A3Chapter chapter in this.Chapters)
            {
                try { Directory.CreateDirectory(String.Concat(A3Globals.A3_LATEX, @"\chapters\", chapter.Title)); } catch { }
                main.Add(String.Concat(@"\input{", "\"", A3Globals.A3_LATEX.Replace('\\','/'), @"/chapters/", chapter.Title, ".tex\"}"));
            }

            main.Add("");
            main.Add(@"\backmatter");

            main.Add("");
            main.Add(@"\end{document}");

            File.WriteAllLines(String.Concat(A3Globals.A3_LATEX, @"\", "main.tex"), main);
        }
        private void GenerateLaTexChapters()
        {
            foreach (A3Chapter chapter in this.Chapters)
            {
                List<string> chap = new List<string>
                {
                    String.Concat(@"\chapter{", chapter.Title, @"}"),
                    @"\newpage",
                    ""
                };
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    Directory.CreateDirectory(String.Concat(A3Globals.A3_LATEX, @"\chapters\", chapter.Title));
                    chap.Add(String.Concat(@"\input{", "\"", A3Globals.A3_LATEX.Replace('\\', '/'), @"/chapters/", chapter.Title, @"/", subchapter.Title, ".tex\"}"));
                }
                File.WriteAllLines(String.Concat(A3Globals.A3_LATEX, @"\chapters\", chapter.Title, @".tex"), chap);
            }
        }
        private void GenerateLaTexSubchapters(A3Outline outline)
        {
            string[] mdFiles = Directory.GetFiles(String.Concat(A3Globals.A3_MARKDOWN));
            List<string> htmlNotes = new List<string>();
            foreach (string filePath in mdFiles)
            {
                string aguid = filePath.Split('.')[0];
                string note = File.ReadAllText(filePath);
                if (note != null)
                {
                    if (note != "")
                    {
                        htmlNotes.Add(aguid);
                        htmlNotes.Add(Markdown.ToHtml(note));
                        htmlNotes.Add(aguid);
                    }
                }
            }
            File.WriteAllLines(String.Concat(A3Globals.A3_LATEX, @"notes.html"), htmlNotes);
            List<string> notes = A3Notes.ToLatex(outline, String.Concat(A3Globals.A3_LATEX, @"notes.html"));

            foreach (A3Chapter chapter in this.Chapters)
            {
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    List<string> sub = new List<string>
                    {
                        String.Concat(@"\section{", subchapter.Title, @"}")
                    };
                    foreach (A3Content a3Content in subchapter.Slides)
                    {
                        sub.Add(@"\begin{figure}[H]");
                        sub.Add(String.Concat(@"\includegraphics*[width=1\linewidth, height=.425\textheight, trim= 0 0 0 0, clip]{", "\"", A3Globals.A3_BOOK_PNGS.Replace('\\','/'),a3Content.ActiveGuid, "\"}"));
                        sub.Add(@"\end{figure}");
                        if (a3Content.Notes != null)
                        {
                            if (a3Content.Notes != "")
                            {
                                int startIndex = notes.FindIndex(s => s.Contains(a3Content.ActiveGuid));
                                int endIndex = notes.FindLastIndex(s => s.Contains(a3Content.ActiveGuid));
                                startIndex++;
                                endIndex--;

                                sub.Add(String.Concat(@"%SLIDE_INDEX_OF_ABOVE_FIGURE: ", a3Content.Index));
                                sub.Add(@"\begin{flushleft}");
                                for (int i = startIndex; i < endIndex; i++)
                                {
                                    sub.Add(notes[i]);
                                }
                                sub.Add(@"\end{flushleft}");
                                sub.Add(String.Concat(@"%SLIDE_INDEX_OF_ABOVE_TEXT: ", a3Content.Index));
                            }
                        }

                    }
                    sub.Add(@"\clearpage");
                    File.WriteAllLines(String.Concat(A3Globals.A3_LATEX, @"\chapters\", chapter.Title, @"\", subchapter.Title, @".tex"), sub);
                }
            }
        }
        #endregion

        #region Validation Methods
        public void Validate(A3LogFile logFile, int processingLevel)
        {
            this.ValidateTitle(logFile);
            // If process dicatates different checks make a switch statement here locally but pass the process varaible to the chapter and labs so that it can independently handle those
            this.ValidateChapters(logFile, processingLevel);
            this.ValidateLabs(logFile, processingLevel);
        }

        private void ValidateTitle(A3LogFile logFile)
        {
            if (this.Course == null || this.Course.Count(c => !Char.IsWhiteSpace(c)) == 0)
            {
                logFile.WriteEntry(new A3LogEntry("ERROR", "No Course Title Found"));
            }
        }
        private void ValidateChapters(A3LogFile logFile, int processingLevel)
        {
            int chapterCount = 1;
            foreach (A3Chapter chapter in this.Chapters)
            {
                chapter.Validate(chapterCount, logFile, processingLevel);
                chapterCount++;
            }
        }
        private void ValidateLabs(A3LogFile logFile, int processingLevel)
        {
            int labCount = 1;
            foreach (A3Lab lab in this.Labs)
            {
                // lab.Validate(logFile, process);
                labCount++;
            }
        }
        #endregion
    }
}
