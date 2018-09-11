using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Content
    {
        public int Index { get; set; }
        public string Title { get; set; }
        public string Chapter { get; set; }
        public string Subchapter { get; set; }
        public string Day { get; set; }
        public string Notes { get; set; }
        public string Type { get; set; }
        public string Guid { get; set; }
        public List<string> HistoricGuids { get; set; }

        public void Generate(PowerPoint.Presentation presentation, string chapSub)
        {
            presentation.Slides[3].Duplicate().MoveTo(presentation.Slides.Count);
            A3Slide a3Slide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = this.Title,
                Type = "CONTENT",
                ChapSub = chapSub,
                Guid = System.Guid.NewGuid().ToString(),
                Notes = this.Notes
            };
            a3Slide.WriteFromMemory();
        }
    }
}
