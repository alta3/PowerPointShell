﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Chapter
    {
        public string Title { get; set; }
        public List<A3Subchapter> Subchapters { get; set; }
        public string Guid { get; set; }
        public List<string> HGuids { get; set; }

        public A3Chapter(A3Slide slide) 
        {
            Guid    = slide.Guid;
            HGuids  = slide.HGuids;
            Title   = slide.Title;
            Subchapters = new List<A3Subchapter>();
        }

        public void WriteToPresentation(Presentation presentation, int chapter)
        {
            WriteChapterSlide(presentation);
            int index = presentation.Slides.Count;
            presentation.SectionProperties.AddBeforeSlide(index, Title);
            WriteSubChapterSlides(presentation);
        }

        private void WriteChapterSlide(Presentation presentation)
        {
            // Open the appropriate slide and set it to the active slide in the presentation
            presentation.Slides[2].Duplicate().MoveTo(presentation.Slides.Count);

            // Change the title of the slide and the scrubber to accurately reflect the outline
            A3Slide slide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Type = A3Slide.Types.CHAPTER,
                Chapter = "Vocabluary",
                Title = Title,
                Guid = System.Guid.NewGuid().ToString()
            };
            slide.WriteFromMemory();

            Shape toc = slide.GetShapeByTag(A3Slide.Tags.TOC);
            toc.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
            toc.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = null;
            toc.TextFrame.TextRange.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
            toc.TextFrame.TextRange.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = presentation.Slides[2].SlideID + "," + presentation.Slides[2].SlideIndex + "," + presentation.Slides[2].Name;

            try
            {
                foreach (Microsoft.Vbe.Interop.VBComponent component in presentation.VBProject.VBComponents)
                {
                    if (component.Name.ToLower().StartsWith("slide")) component.CodeModule.AddFromString(A3Environment.CHAPTER_VBA);
                }
            }
            catch
            {
                if (A3Environment.QUIT_FROM_CURRENT_LOOP is false) MessageBox.Show("You must give access to the VBA Object Model for this plugin to work: \r\n File -> Options -> Trust Center -> Trust Center Setttings -> Macro Settings -> Trust Access to the VBA Project object model. This build will fail.", "Security Setting Problem", MessageBoxButtons.OK);
            }
        }
        private void WriteSubChapterSlides(Presentation presentation)
        {
            Subchapters?.ForEach(sub => sub.WriteToPresentation(presentation, Title));
        }

        public void PublishLatex()
        {
            List<string> chap = new List<string>();
            chap.AddRange(A3Presentation.LatexMap[A3Presentation.LatexLines.CHAPTER]
                                        .Replace("_CHAPTER_TITLE_", Title)
                                        .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                                        .ToList());
            Subchapters.ForEach(sub =>
            {
                Directory.CreateDirectory(string.Concat(A3Environment.A3_LATEX, @"\chapters\", Title));
                chap.Add(A3Presentation.LatexMap[A3Presentation.LatexLines.CHAPTERSUBCHAPTER]
                                       .Replace("_LATEX_PATH_", A3Environment.A3_LATEX)
                                       .Replace("_CHAPTER_TITLE_", Title)
                                       .Replace("_SUBCHAPTER_TITLE_", sub.Title));
                sub.PublishLatex(Title);
            });
            File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"\chapters\", Title, @".tex"), chap);
        }
    }
}
