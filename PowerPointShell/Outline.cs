using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointShell
{
    public class Outline
    {
        public string Course { get; set; }
        public List<Chapter> Chapters { get; set; }
        public List<Lab> Labs { get; set; }

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
            foreach (var chapter in this.Chapters)
            {
                chapter.Generate(presentation, chapterCount);
                chapterCount += 1;
            }
        }
        private void GenerateCourseSlide(PowerPoint.Presentation presentation)
        {
            presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 1, 1);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
            activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = this.Course;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
        }
        private void GenerateEndOfDeckSlide(PowerPoint.Presentation presentation)
        {
            presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 3, 3);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
            activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = "End of Deck";
            string scrubber = this.Course + ": End Of Deck";
            activeslide.Shapes.Range("SCRUBBER").TextFrame.TextRange.Text = scrubber;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
        }
        private void GenerateTOCSlide(PowerPoint.Presentation presentation)
        {
            presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 4, 4);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
            activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = "Table of Contents";
            string scrubber = this.Course + ": TOC";
            activeslide.Shapes.Range("SCRUBBER").TextFrame.TextRange.Text = scrubber;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
        }
        private void GenerateQuizSlide(PowerPoint.Presentation presentation)
        {
            presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 6, 6);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
            activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = "Knowledge Check";
            presentation.SectionProperties.AddBeforeSlide(presentation.Slides.Count, "Knowledge Check");
            System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
        }
    }
}
