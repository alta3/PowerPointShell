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
    public class Subchapter
    {
        public string Title { get; set; }
        public List<BodySlide> Slides { get; set; }
        public List<Question> Questions { get; set; }

        public void Generate(PowerPoint.Presentation presentation, string scrubber)
        {
            this.GenerateBodySlides(presentation, scrubber);
            this.GenerateQuestions(scrubber);
        }

        private void GenerateBodySlides(PowerPoint.Presentation presentation, string scrubber)
        {
            if (this.Slides != null)
            {
                foreach (var slide in this.Slides)
                {
                    slide.Generate(presentation, scrubber);
                }
            }
            else
            {
                presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 3, 3);
                PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
                activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = "SLIDE TITLE";
                activeslide.Shapes.Range("SCRUBBER").TextFrame.TextRange.Text = scrubber;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
            }
            // Add Question Triangle Here To the Last Slide Of The Subchapter
        }
        private void GenerateQuestions(string scrubber)
        {
            if (this.Questions != null)
            {
                foreach (var question in this.Questions)
                {
                    question.Generate(scrubber);
                }
            }
        }
    }
}
