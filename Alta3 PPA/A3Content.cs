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
        public string Title { get; set; }
        public string Day { get; set; }
        public string Notes { get; set; }
        public string ActiveGuid { get; set; }
        public List<string> HistoricGuids { get; set; }

        public void Generate(PowerPoint.Presentation presentation, string chapSub)
        {
            presentation.Slides.AddSlide(presentation.Slides.Count + 1, presentation.SlideMaster.CustomLayouts[3]);
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = this.Title,
                ChapSub = chapSub,
                ActiveGuid = Guid.NewGuid().ToString(),
                Notes = this.Notes
            };
            a3ActiveSlide.WriteFromMemory();
        }
    }
}
