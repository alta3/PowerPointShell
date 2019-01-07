using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Subchapter
    {
        public string Title { get; set; }
        public List<A3Content> Slides { get; set; }
        public List<A3Question> Questions { get; set; }

        public A3Subchapter(A3Slide slide)
        {
            Title       = slide.Subchapter;
            Slides      = new List<A3Content>();
            Questions   = new List<A3Question>();
        }

        public void WriteToPresentation(Presentation presentation, string chapter)
        {
            Slides?.ForEach(s => s.WriteToPresentation(presentation));
            Questions?.ForEach(q => q.Generate(chapter, Title));
        }

    }
}
