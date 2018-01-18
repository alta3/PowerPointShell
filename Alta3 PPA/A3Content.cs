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

        public void Generate(PowerPoint.Presentation presentation, string scrubber)
        {
            presentation.Slides.AddSlide(-1, A3Globals.TITLE_LAYOUT);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
            activeslide.Shapes.Range("TITLE").TextFrame.TextRange.Text = this.Title;
            activeslide.Shapes.Range("CHAP:SUB").TextFrame.TextRange.Text = scrubber;
            activeslide.Shapes.Range("ACTIVE_GUID").TextFrame.TextRange.Text = Guid.NewGuid().ToString();
            activeslide.NotesPage.Shapes.Range("notes").TextFrame.TextRange.Text = this.Notes;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
        }
    }
}
