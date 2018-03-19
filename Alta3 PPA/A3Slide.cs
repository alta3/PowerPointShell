using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Slide
    {
        public string ActiveGuid { get; set; }
        public List<string> HistoricGuids { get; set; }
        public string Day { get; set; }
        public string Type { get; set; }
        public string ChapSub { get; set; }
        public string Chapter { get; set; }
        public string Subchapter { get; set; }
        public string Title { get; set; }
        public string Notes { get; set; }
        public List<string> ShapeNames { get; set; }
        public PowerPoint.Slide Slide { get; set; }
        
        public A3Slide(PowerPoint.Slide slide)
        {
            this.Slide = slide;
            this.ReadActiveGuid();
            this.ReadHistoricGuid();
            this.ReadType();
            this.ReadChapSub();
            this.ReadChapter();
            this.ReadSubchapter();
            this.ReadTitle();
            this.ReadDay();
            this.ReadNotes();
        }

        public static void SetActiveSlide(PowerPoint.Slide slide)
        {
            A3Globals.A3SLIDE = new A3Slide(slide);
        }
        public static void ShowMetadataForm()
        {
            A3Globals.A3SLIDE.Slide.Select();
            SlideMetadata slideMetadata = new SlideMetadata()
            {
                StartPosition = FormStartPosition.CenterScreen
            };
            slideMetadata.DrawSlideInfo();
            slideMetadata.ShowDialog();
        }
        public static void FixNullMetadata(bool firstCheck, A3LogFile logFile)
        {
            A3Globals.A3SLIDE.Slide.Select();
            string msg = null;
            if (A3Globals.A3SLIDE.Type == null)
            {
                msg = String.Concat("A Type Must Be Specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
            }
            else if (A3Globals.A3SLIDE.ActiveGuid == null)
            {
                msg = String.Concat("An ActiveGuid Must Be Specified For Every Slide -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
            }
            else if (A3Globals.A3SLIDE.Type.ToUpper() == "CONTENT")
            {
                if (A3Globals.A3SLIDE.Title == null ||
                    A3Globals.A3SLIDE.ChapSub == null ||
                    A3Globals.A3SLIDE.Chapter == null ||
                    A3Globals.A3SLIDE.Subchapter == null)
                {
                    msg = String.Concat("A Title, ActiveGuid, ChapSub, Chapter, and Subchapter must be specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
                }
            }
            else if (A3Globals.A3SLIDE.Type.ToUpper() == "CHAPTER")
            {
                if (A3Globals.A3SLIDE.Title == null ||
                    A3Globals.A3SLIDE.ActiveGuid == null ||
                    A3Globals.A3SLIDE.ChapSub == null)
                {
                    msg = String.Concat("A Ttitle, ActiveGuid, and ChapSub must be specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
                }
            }
            else if (A3Globals.A3SLIDE.Type.ToUpper() == "COURSE")
            {
                if (A3Globals.A3SLIDE.Title == null)
                {
                    msg = String.Concat("A Title And AtiveGuid Must Be Specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
                }
            }
            else { }

            if (firstCheck)
            {
                if (msg != null)
                {
                    A3Globals.A3SLIDE.ReadShapes();
                    A3Slide.ShowMetadataForm();
                    A3Slide.FixNullMetadata(false, logFile);
                }
            }
            else
            {
                if (msg != null)
                {
                    logFile.WriteError(msg);
                    DialogResult dialogResult = MessageBox.Show(msg, "Properties Still Contain A Null", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        A3Globals.A3SLIDE.ReadShapes();
                        A3Slide.ShowMetadataForm();
                        A3Slide.FixNullMetadata(false, logFile);
                    }
                }
            }     
        }

        public object TypeConversion()
        {
            switch (this.Type.ToLower())
            {
                case "course":
                    A3Outline a3Outline = new A3Outline()
                    {
                        Course = this.Title,
                        Chapters = new List<A3Chapter>()
                    };
                    return a3Outline;
                case "chapter":
                    A3Chapter a3Chapter = new A3Chapter()
                    {
                        ActiveGuid = this.ActiveGuid,
                        HistoricGuids = this.HistoricGuids,
                        Day = this.Day,
                        Title = this.Title,
                        Subchapters = new List<A3Subchapter>()
                    };
                    return a3Chapter;
                default:
                    A3Content a3Content = new A3Content()
                    {
                        ActiveGuid = this.ActiveGuid,
                        HistoricGuids = this.HistoricGuids,
                        Day = this.Day,
                        Title = this.Title,
                        Type = this.Type,
                        Notes = this.Notes,
                        Index = this.Slide.SlideIndex
                    };
                    return a3Content;
            } 
        }

        public void ReadFromSlide()
        {
            this.ReadActiveGuid();
            this.ReadHistoricGuid();
            this.ReadType();
            this.ReadChapSub();
            this.ReadChapter();
            this.ReadSubchapter();
            this.ReadTitle();
            this.ReadDay();
        }
        public void ReadActiveGuid()
        {
            try { this.ActiveGuid = this.Slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text; }
            catch { this.ActiveGuid = null; }
        }
        public void ReadHistoricGuid()
        {
            try
            {
                string guids = this.Slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text;
                this.HistoricGuids = new List<string>();

                if (guids.Contains(';')) { this.HistoricGuids.AddRange(guids.Split(';')); }
                else { this.HistoricGuids.Add(guids); }
            }
            catch { this.HistoricGuids = null; }
        }
        public void ReadType()
        {
            try { this.Type = this.Slide.Shapes["TYPE"].TextFrame.TextRange.Text; }
            catch { this.Type = null; }
        }
        public void ReadChapSub()
        {
            try { this.ChapSub = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text; }
            catch { this.ChapSub = null; }
        }
        public void ReadChapter()
        {
            try { this.Chapter = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text.Split(':')[0].Trim(); }
            catch { this.Chapter = null; }
        }
        public void ReadSubchapter()
        {
            try { this.Subchapter = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text.Split(':')[1].Trim(); }
            catch { this.Subchapter = null; }
        }
        public void ReadTitle()
        {
            try { this.Title = this.Slide.Shapes["TITLE"].TextFrame.TextRange.Text; }
            catch { this.Title = null; }
        }
        public void ReadDay()
        {
            try { this.Day = this.Slide.Shapes["DAY"].TextFrame.TextRange.Text; }
            catch { this.Day = null; }
        }
        public void ReadNotes()
        {
            try
            {
                foreach (PowerPoint.Shape shape in this.Slide.NotesPage.Shapes)
                {
                    if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.TextRange.Text != "")
                        {
                            this.Notes = shape.TextFrame.TextRange.Text;
                            break;
                        }
                    }
                }
            }
            catch { this.Notes = null; }
        }
        public void ReadShapes()
        {
            try
            {
                this.ShapeNames = new List<string>();
                foreach (PowerPoint.Shape shape in this.Slide.Shapes)
                {
                    try
                    {
                        if (shape.TextFrame.TextRange.Text != null)
                        {
                            this.ShapeNames.Add(shape.Name);
                        }
                    }
                    catch { }
                }
            }
            catch 
            {
                this.ShapeNames = null;
            }
        }

        public string InferType()
        {
            switch (this.Slide.CustomLayout.Name)
            {
                case "Course Title":
                    return "COURSE";
                case "Chapter Title":
                    return "CHAPTER";
                case "Review Questions":
                    return "QUESTION";
                default:
                    return "CONTENT";
            }
        }
        public string InferChapSub()
        {
            foreach (string shapeName in A3Globals.A3SLIDE.ShapeNames)
            {
                if (A3Globals.A3SLIDE.Slide.Shapes[shapeName].Height >= 20
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Height <= 33
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Width >= 700
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Width <= 1000
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Top >= 0
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Top <= 20)
                {
                    return shapeName;
                }
            }
            return null;
        }
        public string InferTitle()
        {
            foreach (string shapeName in A3Globals.A3SLIDE.ShapeNames)
            {
                if (A3Globals.A3SLIDE.Slide.Shapes[shapeName].Height >= 30
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Height <= 60
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Width >= 700
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Width <= 900
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Top >= 21
                    && A3Globals.A3SLIDE.Slide.Shapes[shapeName].Top <= 50)
                {
                    return shapeName;
                }
            }
            return null;
        }
        public string InferDay()
        {
            int slideIndex = this.Slide.SlideIndex;
            PowerPoint.Slide previousSlide = this.Slide.Application.ActivePresentation.Slides[slideIndex - 1];
            string previousDay = "1";
            try { previousDay = previousSlide.Shapes["DAY"].TextFrame.TextRange.Text; } catch { }
            return previousDay;
        }

        public void WriteFromMemory()
        {
            this.WriteType();
            this.WriteActiveGuid();
            this.WriteHistoricGuid();
            this.WriteChapSub();
            this.WriteTitle();
            this.WriteDay();
            this.WriteNotes();
        }
        public void WriteActiveGuid()
        {
            PowerPoint.Shape aguid;
            try { aguid = this.Slide.Shapes["ACTIVE_GUID"]; } catch { aguid = this.MakeActiveGuid(); }
            aguid.TextFrame.TextRange.Text = this.ActiveGuid;
            aguid.Name = "ACTIVE_GUID";
            aguid.Title = "ACTIVE_GUID";
        }
        public void WriteHistoricGuid()
        {
            PowerPoint.Shape hguid;
            try { hguid = this.Slide.Shapes["HISTORIC_GUID"]; } catch { hguid = this.MakeHistoricGuid(); }
            string hguidText = "";
            try
            {
                foreach (string guid in this.HistoricGuids)
                {
                    hguidText += guid;
                }
            }
            catch 
            {

            }
            hguid.TextFrame.TextRange.Text = hguidText;
            hguid.Name = "HISTORIC_GUID";
            hguid.Title = "HISTORIC_GUID";
        }
        public void WriteType()
        {
            PowerPoint.Shape type;
            try { type = this.Slide.Shapes["TYPE"]; } catch { type = this.MakeSlideType(); }
            type.TextFrame.TextRange.Text = this.Type.ToUpper();
            if (this.Type.ToUpper() == "COURSE" || this.Type.ToUpper() == "CHAPTER" || this.Type.ToUpper() == "QUESTION")
            {
                this.Slide.CustomLayout.Name = this.Type.ToUpper();
            }
            else
            {
                this.Slide.CustomLayout.Name = "CONTENT";
            }
            type.Name = "TYPE";
            type.Title = "TYPE";
        }
        public void WriteChapSub()
        {
            PowerPoint.Shape chapsub;
            try { chapsub = this.Slide.Shapes["CHAP:SUB"]; } catch { chapsub = this.MakeChapSub(); }
            chapsub.TextFrame.TextRange.Text = this.ChapSub;
            chapsub.Name = "CHAP:SUB";
            chapsub.Title = "CHAP:SUB";
        }
        public void WriteTitle()
        {
            PowerPoint.Shape title;
            try { title = this.Slide.Shapes["TITLE"]; } catch { title = this.MakeTitle(); }
            title.TextFrame.TextRange.Text = this.Title;
            title.Name = "TITLE";
            title.Title = "TITLE";
        }
        public void WriteDay()
        {
            PowerPoint.Shape day;
            try { day = this.Slide.Shapes["DAY"]; } catch { day = this.MakeDay(); }
            day.TextFrame.TextRange.Text = this.Day;
            day.Name = "DAY";
            day.Title = "DAY";
        }
        public void WriteNotes()
        {

        }

        public PowerPoint.Shape MakeActiveGuid()
        {
            PowerPoint.Shape aguid = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 400, 500, 30);
            aguid.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            aguid.Name = "ACTIVE_GUID";
            aguid.Title = "ACTIVE_GUID";
            return aguid;
        }
        public PowerPoint.Shape MakeHistoricGuid()
        {
            PowerPoint.Shape hguid = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 430, 500, 30);
            hguid.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            hguid.Name = "HISTORIC_GUID";
            hguid.Title = "HISTORIC_GUID";
            return hguid;
        }
        public PowerPoint.Shape MakeSlideType()
        {
            PowerPoint.Shape type = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 500, 400, 500, 30);
            type.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            type.Name = "TYPE";
            type.Title = "TYPE";
            return type;
        }
        public PowerPoint.Shape MakeChapSub()
        {
            PowerPoint.Shape chapsub = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 12, 1, 720, 28);
            chapsub.TextFrame.TextRange.Characters().Font.Size = 16;
            chapsub.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
            chapsub.Name = "CHAP:SUB";
            chapsub.Title = "CHAP:SUB";
            if (this.Type == "COURSE")
            {
                chapsub.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            return chapsub;
        }
        private PowerPoint.Shape MakeTitle()
        {
            PowerPoint.Shape title;
            switch (this.Type)
            {
                case "COURSE":
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 312, 192, 358, 106);
                    break;
                case "CHAPTER":
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 18, 81, 934, 50);
                    break;
                case "QUESTION":
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 30, 960, 50, 28);
                    break;
                default:
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 12, 30, 936, 50);
                    break;
            }
            title.Name = "TITLE";
            title.Title = "TITLE";
            return title;
        }
        private PowerPoint.Shape MakeDay()
        {
            PowerPoint.Shape day = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 500, 430, 1000, 30);
            day.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            day.Name = "DAY";
            day.Title = "DAY";
            return day;
        }
    }
}
