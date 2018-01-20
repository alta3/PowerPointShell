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
    public class A3Chapter
    {
        public string Title { get; set; }
        public List<A3Subchapter> Subchapters { get; set; }
        public List<A3Vocab> Vocab { get; set; }
        public string Day { get; set; }
        public string ActiveGuid { get; set; }
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
            presentation.Slides.InsertFromFile(A3Globals.MODEL_POWERPOINT, presentation.Slides.Count + 1, 2);

            // Change the title of the slide and the scrubber to accurately reflect the outline
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Vocabluary",
                ChapSub = this.Title,
                ActiveGuid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();

            // Ensure the slide TOC button is linked to the second slide in the presentation
            a3ActiveSlide.Slide.Shapes.Range("TOC").ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
            a3ActiveSlide.Slide.Shapes.Range("TOC").ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = null;
            a3ActiveSlide.Slide.Shapes.Range("TOC").TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
            a3ActiveSlide.Slide.Shapes.Range("TOC").TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = presentation.Slides[2].SlideID + "," + presentation.Slides[2].SlideIndex + "," + presentation.Slides[2].Name;

            // Write the Chapter VBA to the slide itself
            string index = String.Concat("Slide", presentation.Slides.Count);
            Microsoft.Vbe.Interop.VBComponent component = presentation.VBProject.VBComponents.Item(index);
            component.CodeModule.AddFromString(A3Globals.CHAPTER_VBA);
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

        public void Validate(int chapterCount, A3LogFile logFile, int process)
        {
            this.ValidateTitle(chapterCount, logFile);
            // If process dicatates different checks make a switch statement here locally but pass the process varaible to the chapter and labs so that it can independently handle those
            this.ValidateSubchapters(chapterCount, logFile);
            this.ValidateVocab(chapterCount, logFile);
        }

        private void ValidateTitle(int chapterCount, A3LogFile logFile)
        {
            if (this.Title == null || this.Title.Count(c => !Char.IsWhiteSpace(c)) == 0)
            {
                logFile.WriteError(String.Concat("Chapter ", chapterCount.ToString(), ": No Title Found"));
            }
        }
        private void ValidateSubchapters(int chapterCount, A3LogFile logFile)
        {
            foreach (A3Subchapter subchapter in this.Subchapters)
            {
                //subchapter.Valid();
            }
        }
        private void ValidateVocab(int chapterCount, A3LogFile logFile)
        {
            foreach (A3Vocab vocab in this.Vocab)
            {
                //vocab.Valid();
            }
        }
    }
}
