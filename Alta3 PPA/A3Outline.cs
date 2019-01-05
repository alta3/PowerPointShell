﻿using System;
using System.Collections.Generic;
using System.IO;
using Markdig;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Outline
    {
        #region Outline Properites
        public enum Metadata
        {
            NAME,
            FILENAME,
            HASLABS,
            HASSLIDES,
            HASVIDEOS,
            WEBURL
        }

        public string Course { get; set; }
        public string Filename { get; set; }
        public bool HasLabs { get; set;}
        public bool HasSlides { get; set; }
        public bool HasVideos { get; set; }
        public string Weburl { get; set; }
        public List<A3Chapter> Chapters { get; set; }
        #endregion

        // move to A3Presentation. 
        #region Generate Presentation
        public void GeneratePresentation(PowerPoint.Presentation presentation)
        {
            GenerateCourseSlide(presentation);
            GenerateTOCSlide(presentation);
            presentation.SectionProperties.AddBeforeSlide(1, Course);
            GenerateChapters(presentation);
            GenerateEndOfDeckSlide(presentation);
            GenerateQuizSlide(presentation);
        }

        private void GenerateChapters(PowerPoint.Presentation presentation)
        {
            int chapterCount = 1;
            foreach (A3Chapter chapter in Chapters)
            {
                chapter.Generate(presentation, chapterCount);
                chapterCount += 1;
            }
        }
        private void GenerateCourseSlide(PowerPoint.Presentation presentation)
        {
            // Insert the course slide from the model PowerPoint
            presentation.Slides[1].Duplicate().MoveTo(presentation.Slides.Count);

            // Change the title to the course title given in the yaml file
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = Course,
                Type = A3Slide.Types.COURSE,
                Guid = Guid.NewGuid().ToString(),
                Notes = String.Concat("name: ",           Course, 
                                      "\r\nfilename: ",   Filename, 
                                      "\r\nhas-labs: ",   HasLabs, 
                                      "\r\nhas-slides: ", HasSlides, 
                                      "\r\nhas-videos: ", HasVideos, 
                                      "\r\nweburl: ",     Weburl)
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateEndOfDeckSlide(PowerPoint.Presentation presentation)
        {
            // Insert a title slide from the model PowerPoint
            presentation.Slides[3].Duplicate().MoveTo(presentation.Slides.Count);
            
            // Change the title, chapsub, type, and active guid to accurately reflect what is happening
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "End of Deck",
                ChapSub = string.Concat(Course, ": End Of Deck"),
                Guid = Guid.NewGuid().ToString(),
                Type = "CONTENT"
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateTOCSlide(PowerPoint.Presentation presentation)
        {
            // Insert a split slide from the model PowerPoint
            presentation.Slides[4].Duplicate().MoveTo(presentation.Slides.Count);

            // Populate the appropriate values of the slide deck here
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Table of Contents",
                ChapSub = String.Concat(Course, ": TOC"),
                Type = "TOC",
                Guid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();

            // TODO: Create a linked list to the first chapter of each day and colorize the results
        }
        private void GenerateQuizSlide(PowerPoint.Presentation presentation)
        {
            // Insert a question slide from the model PowerPoint
            presentation.Slides[6].Duplicate().MoveTo(presentation.Slides.Count);

            // Ensure the title is Knowledge Check and move on 
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Knowledge Check",
                Type = "QUESTION",
                Guid = Guid.NewGuid().ToString(),
            };
            a3ActiveSlide.WriteFromMemory();
            presentation.SectionProperties.AddBeforeSlide(presentation.Slides.Count, "Knowledge Check");

        }
        #endregion

        #region Generate LaTex
        public void GenerateLaTex()
        {
            GenerateLaTexMain();
            GenerateLaTexChapters();
            GenerateLaTexSubchapters(this);
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
                String.Concat(@"\includegraphics[width=.5\textwidth]{", "\"", A3Environment.A3_RESOURCE.Replace('\\','/'), @"/a3logo", "\"}"),
                @"\linebreak",
                @"\linebreak",
                String.Concat(@"{\Huge\textbf{", Course, @"}}"),
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
            foreach (A3Chapter chapter in Chapters)
            {
                try { Directory.CreateDirectory(String.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter.Title)); } catch { }
                main.Add(String.Concat(@"\input{", "\"", A3Environment.A3_LATEX.Replace('\\','/'), @"/chapters/", chapter.Title, ".tex\"}"));
            }

            main.Add("");
            main.Add(@"\backmatter");

            main.Add("");
            main.Add(@"\end{document}");

            File.WriteAllLines(String.Concat(A3Environment.A3_LATEX, @"\", "main.tex"), main);
        }
        private void GenerateLaTexChapters()
        {
            foreach (A3Chapter chapter in Chapters)
            {
                List<string> chap = new List<string>
                {
                    String.Concat(@"\chapter{", chapter.Title, @"}"),
                    @"\newpage",
                    ""
                };
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    Directory.CreateDirectory(String.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter.Title));
                    chap.Add(String.Concat(@"\input{", "\"", A3Environment.A3_LATEX.Replace('\\', '/'), @"/chapters/", chapter.Title, @"/", subchapter.Title, ".tex\"}"));
                }
                File.WriteAllLines(String.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter.Title, @".tex"), chap);
            }
        }
        private void GenerateLaTexSubchapters(A3Outline outline)
        {
            string[] mdFiles = Directory.GetFiles(String.Concat(A3Environment.A3_MARKDOWN));
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
            File.WriteAllLines(String.Concat(A3Environment.A3_LATEX, @"notes.html"), htmlNotes);
            List<string> notes = A3Notes.ToLatex(outline, String.Concat(A3Environment.A3_LATEX, @"notes.html"));

            foreach (A3Chapter chapter in Chapters)
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
                        sub.Add(String.Concat(@"\includegraphics*[width=1\linewidth, height=.425\textheight, trim= 0 0 0 0, clip]{", "\"", A3Environment.A3_BOOK_PNGS.Replace('\\','/'),a3Content.Guid, "\"}"));
                        sub.Add(@"\end{figure}");
                        if (a3Content.Notes != null)
                        {
                            if (a3Content.Notes != "")
                            {
                                int startIndex = notes.FindIndex(s => s.Contains(a3Content.Guid));
                                int endIndex = notes.FindLastIndex(s => s.Contains(a3Content.Guid));
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
                    File.WriteAllLines(String.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter.Title, @"\", subchapter.Title, @".tex"), sub);
                }
            }
        }
        #endregion

        #region Generate YAML
        public void GenerateYAML()
        {

        }
        #endregion
    }
}
