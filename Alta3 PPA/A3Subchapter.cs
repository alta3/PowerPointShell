using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Subchapter
    {
        public string Title { get; set; }
        public List<A3Content> Slides { get; set; }
        public List<A3Question> Questions { get; set; }

        public A3Subchapter()
        {
            this.Title = null;
            this.Slides = new List<A3Content>();
            this.Questions = new List<A3Question>();
        }

        public A3Subchapter(String subchapter)
        {
            this.Title = subchapter;
            this.Slides = new List<A3Content>();
            this.Questions = new List<A3Question>();
        }

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
                A3Content slide = new A3Content
                {
                    Title = "SLIDE TITLE"
                };
                slide.Generate(presentation, scrubber);
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
