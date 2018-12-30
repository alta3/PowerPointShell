using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Chapter
    {
        public string Title { get; set; }
        public List<A3Subchapter> Subchapters { get; set; }
        public List<A3Vocab> Vocab { get; set; }
        public string Day { get; set; }
        public string Guid { get; set; }
        public List<string> HistoricGuids { get; set; }

        public void Generate(PowerPoint.Presentation presentation, int chapterCount)
        {
            this.GenerateChapterSlide(presentation, chapterCount);
            int count = presentation.Slides.Count;
            presentation.SectionProperties.AddBeforeSlide(count, this.Title);
            this.GenerateSubChapters(presentation);
            this.GenerateVocab(chapterCount);
        }

        private void GenerateChapterSlide(PowerPoint.Presentation presentation, int chapterCount)
        {
            // Open the appropriate slide and set it to the active slide in the presentation
            presentation.Slides[2].Duplicate().MoveTo(presentation.Slides.Count);

            // Change the title of the slide and the chap:sub to accurately reflect the outline
            A3Slide a3Slide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Vocabluary",
                ChapSub = this.Title,
                Guid = System.Guid.NewGuid().ToString()
            };
            a3Slide.WriteFromMemory();

            // Ensure the slide TOC button is linked to the second slide in the presentation
            PowerPoint.Shape activeShape;
            foreach (PowerPoint.Shape shape in a3Slide.Slide.Shapes)
            {
                if (shape.Title == "TOC" || shape.Name == "TOC")
                {
                    activeShape = shape;
                    activeShape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
                    activeShape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = null;
                    activeShape.TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
                    activeShape.TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = presentation.Slides[2].SlideID + "," + presentation.Slides[2].SlideIndex + "," + presentation.Slides[2].Name;
                }
            }

            // Write the Chapter VBA to the slide itself
            string index = String.Concat("Slide ", presentation.Slides.Count);
            foreach (Microsoft.Vbe.Interop.VBComponent component in presentation.VBProject.VBComponents)
            {
                if (component.Name.ToLower().StartsWith("slide"))
                {
                    component.CodeModule.AddFromString(A3Globals.CHAPTER_VBA);
                }
            }

        }
        private void GenerateSubChapters(PowerPoint.Presentation presentation)
        {
            if (this.Subchapters != null)
            {
                foreach (var subchapter in this.Subchapters)
                {
                    subchapter.Generate(presentation, String.Concat(this.Title, ": ", subchapter.Title));
                }
            }
        }
        private void GenerateVocab(int chapterCount)
        {
            if (this.Vocab != null)
            {
                foreach (var word in this.Vocab)
                {
                    word.Generate(chapterCount);
                }
            }
        }
    }
}
