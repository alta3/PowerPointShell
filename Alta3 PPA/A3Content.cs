using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Content
    {
        public int Index { get; set; }
        public string Title { get; set; }
        public string Chapter { get; set; }
        public string Subchapter { get; set; }
        public string Notes { get; set; }
        public string Type { get; set; }
        public string Guid { get; set; }
        public List<string> HGuids { get; set; }

        public A3Content(A3Slide slide) 
        {
            Guid        = slide.Guid;
            HGuids      = slide.HGuids;
            Title       = slide.Title;
            Chapter     = slide.Chapter;
            Subchapter  = slide.Subchapter;
            Type        = slide.Type.ToString();
            Notes       = slide.Notes;
            Index       = slide.Slide.SlideIndex;
        }

        public void WriteToPresentation(PowerPoint.Presentation presentation)
        {
            presentation.Slides[3].Duplicate().MoveTo(presentation.Slides.Count);
            A3Slide slide = new A3Slide(presentation.Slides[presentation.Slides.Count]) {
                Title =         Title,
                Type =          Enum.TryParse(Type.ToUpper(), out A3Slide.Types type) ? type : A3Slide.Types.NULL,
                Chapter =       Chapter,
                Subchapter =    Subchapter,
                Guid =          Guid is null ? System.Guid.NewGuid().ToString() : Guid,
                HGuids =        HGuids is null || HGuids.Count < 1 ? null : HGuids,
                Notes =         Notes
            };
            slide.WriteFromMemory();
        }
    }
}
