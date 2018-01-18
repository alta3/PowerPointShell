using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointShell
{
    public class Chapter
    {
        public string Title { get; set; }
        public List<Subchapter> Subchapters { get; set; }
        public List<Vocab> Vocab { get; set; }
        public int Day { get; set; }

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
            presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 2, 2);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);

            // Change the title of the slide and the scrubber to accurately reflect the outline
            activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = "Vocabulary";
            activeslide.Shapes.Range("SCRUBBER").TextFrame.TextRange.Text = this.Title;

            // Ensure the slide TOC button is linked to the second slide in the presentation
            activeslide.Shapes.Range("TOC").ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
            activeslide.Shapes.Range("TOC").ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = null;
            activeslide.Shapes.Range("TOC").TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = null;
            activeslide.Shapes.Range("TOC").TextFrame.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = presentation.Slides.Range(2).SlideID + "," + presentation.Slides.Range(2).SlideIndex + "," + presentation.Slides.Range(2).Name;

            // TODO: ADD VBA CODE HERE -- THIS WORKS FOR ADDING SO NOW JUST NEED TO ACTUALLY IMPLEMENT!
            string index = string.Concat("Slide",presentation.Slides.Count.ToString());
            Microsoft.Vbe.Interop.VBComponent component = presentation.VBProject.VBComponents.Item(index);
            component.CodeModule.AddFromString("THIS IS ONLY A TEST!");

            // Release to ComObject in order to avoid HRESULT E_FAIL errors
            System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
        }
        private void GenerateSubChapters(PowerPoint.Presentation presentation)
        {
            if (this.Subchapters != null)
            {
                foreach (var subchapter in this.Subchapters)
                {
                    string scrubber = this.Title + ": " + subchapter.Title;
                    subchapter.Generate(presentation, scrubber);
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
