using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Chapter
    {
        public string Title { get; set; }
        public List<A3Subchapter> Subchapters { get; set; }
        public List<A3Vocab> Vocab { get; set; }
        public string Guid { get; set; }
        public List<string> HistoricGuids { get; set; }

        public void Generate(PowerPoint.Presentation presentation, int chapterCount)
        {
            GenerateChapterSlide(presentation, chapterCount);
            int count = presentation.Slides.Count;
            presentation.SectionProperties.AddBeforeSlide(count, this.Title);
            GenerateSubChapters(presentation);
            GenerateVocab(chapterCount);
        }

        private void GenerateChapterSlide(PowerPoint.Presentation presentation, int chapterCount)
        {
            // Open the appropriate slide and set it to the active slide in the presentation
            presentation.Slides[2].Duplicate().MoveTo(presentation.Slides.Count);

            // Change the title of the slide and the scrubber to accurately reflect the outline
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Vocabluary",
                ChapSub = this.Title,
                Guid = System.Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();

            // Ensure the slide TOC button is linked to the second slide in the presentation
            PowerPoint.Shape activeShape;
            foreach (PowerPoint.Shape shape in a3ActiveSlide.Slide.Shapes)
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
            try
            {
                foreach (Microsoft.Vbe.Interop.VBComponent component in presentation.VBProject.VBComponents)
                {
                    if (component.Name.ToLower().StartsWith("slide"))
                    {
                        component.CodeModule.AddFromString(A3Environment.CHAPTER_VBA);
                    }
                }
            }
            catch
            {
                if (!A3Environment.QUIT_FROM_CURRENT_LOOP)
                {
                    MessageBox.Show("You must give access to the VBA Object Model for this plugin to work: \r\n File -> Options -> Trust Center -> Trust Center Setttings -> Macro Settings -> Trust Access to the VBA Project object model. This build will fail.", "Security Setting Problem", MessageBoxButtons.OK);
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

        public void Validate(int chapterCount, A3Log log, int process)
        {
            this.ValidateTitle(chapterCount, log);
            // If process dicatates different checks make a switch statement here locally but pass the process varaible to the chapter and labs so that it can independently handle those
            this.ValidateSubchapters(chapterCount, log);
            this.ValidateVocab(chapterCount, log);
        }

        private void ValidateTitle(int chapterCount, A3Log log)
        {
            if (this.Title == null || this.Title.Count(c => !Char.IsWhiteSpace(c)) == 0)
            {
                log.Write(A3Log.Level.Error, String.Concat("Chapter ", chapterCount.ToString(), ": No Title Found"));
            }
        }
        private void ValidateSubchapters(int chapterCount, A3Log log)
        {
            foreach (A3Subchapter subchapter in this.Subchapters)
            {
                //subchapter.Valid();
            }
        }
        private void ValidateVocab(int chapterCount, A3Log log)
        {
            foreach (A3Vocab vocab in this.Vocab)
            {
                //vocab.Valid();
            }
        }
    }
}
