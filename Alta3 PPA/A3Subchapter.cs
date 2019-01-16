using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Subchapter
    {
        public string Title { get; set; }
        public List<A3Content> Slides { get; set; }

        public A3Subchapter(A3Slide slide)
        {
            Title       = slide.Subchapter;
            Slides      = new List<A3Content>();
        }

        public void WriteToPresentation(Presentation presentation, string chapter)
        {
            Slides?.ForEach(s => s.WriteToPresentation(presentation));
        }

        public void PublishLatex(string chapter)
        {
            List<string> latex = new List<string>(File.ReadAllLines(string.Concat(A3Environment.A3_LATEX, @"\raw.tex")));

            List<string> subchapter = new List<string>() { A3Presentation.LatexMap[A3Presentation.LatexLines.SECTION].Replace("_SUBCHAPTER_TITLE", Title) };

            string bookPNGs = A3Environment.A3_BOOK_PNGS.Replace('\\', '/');
            Slides.ForEach(s => {
                subchapter.AddRange(A3Presentation.LatexMap[A3Presentation.LatexLines.FIGURE].Replace("_BOOK_PNGS_", bookPNGs)
                                                                                             .Replace("_GUID_", s.Guid)
                                                                                             .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                                                                                             .ToList());
                int startIndex = latex.FindIndex(l => l.Contains(s.Guid)) + 1;
                int endIndex = latex.FindLastIndex(l => l.Contains(s.Guid)) - 1;
                subchapter.AddRange(A3Presentation.LatexMap[A3Presentation.LatexLines.NOTESTART].Replace("_SLIDE_INDEX_", s.Index.ToString())
                                                                                                .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                                                                                                .ToList());
                for (int i = startIndex; i < endIndex; i++) subchapter.Add(latex[i]);
                subchapter.AddRange(A3Presentation.LatexMap[A3Presentation.LatexLines.NOTEEND].Replace("_SLIDE_INDEX_", s.Index.ToString())
                                                                                              .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                                                                                              .ToList());
            });
            subchapter.Add(@"\clearpage");
            File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter, @"\", Title, @".tex"), subchapter);
        }
    }
}
