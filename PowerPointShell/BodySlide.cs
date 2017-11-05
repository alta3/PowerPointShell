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
    public class BodySlide
    {
        public string Title { get; set; }
        public string Type { get; set; }
        public string Visable { get; set; }
        public string Notes { get; set; }

        public void Generate(PowerPoint.Presentation presentation, string scrubber)
        {
            string type = this.Type.ToLower();

            if (type == "title")
            {
                presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 3, 3);
                PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
                activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = this.Title;
                activeslide.Shapes.Range("SCRUBBER").TextFrame.TextRange.Text = scrubber;
                activeslide.NotesPage.Shapes.Range("notes").TextFrame.TextRange.Text = this.Notes;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
            }
            else if (type == "split")
            {
                presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 4, 4);
                PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
                activeslide.Shapes.Range("Title").TextFrame.TextRange.Text = this.Title;
                activeslide.Shapes.Range("SCRUBBER").TextFrame.TextRange.Text = scrubber;
                activeslide.NotesPage.Shapes.Range("notes").TextFrame.TextRange.Text = this.Notes;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
            }
            else if (type == "blank")
            {
                presentation.Slides.InsertFromFile(GlobalVars.MODEL_POWERPOINT, presentation.Slides.Count, 5, 5);
                PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);
                activeslide.NotesPage.Shapes.Range("notes").TextFrame.TextRange.Text = this.Notes;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(activeslide);
            }
        }
    }
}
