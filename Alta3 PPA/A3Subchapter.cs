using System.Collections.Generic;
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
    }
}
