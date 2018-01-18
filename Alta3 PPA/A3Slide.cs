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
        public string Title { get; set; }
        public PowerPoint.Slide Slide { get; set; }

        public A3Slide(PowerPoint.Slide slide)
        {
            this.Slide = slide;
            this.ReadActiveGuid();
            this.ReadHistoricGuid();
            this.ReadType();
            this.ReadChapSub();
            this.ReadTitle();
            this.ReadDay();
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
            slideMetadata.ShowDialog();
        }
        public static void FixNullMetadata(bool firstCheck)
        {
            A3Globals.A3SLIDE.Slide.Select();
            bool anyNull = A3Globals.A3SLIDE.GetType().GetProperties().Any(p => p.GetValue(A3Globals.A3SLIDE) == null);
            if (firstCheck)
            {
                if (anyNull)
                {
                    A3Slide.ShowMetadataForm();
                    A3Slide.FixNullMetadata(false);
                }
            }
            else
            {
                if (anyNull)
                {
                    string msg = "There are still null fields for this slide would you like to fix these errors?";
                    DialogResult dialogResult = MessageBox.Show(msg, "Properties Still Contain A Null", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        A3Slide.ShowMetadataForm();
                        A3Slide.FixNullMetadata(false);
                    }
                }
            }
        }

        public object TypeConversion()
        {
            switch (this.Type)
            {
                case "COURSE":
                    A3Outline a3Outline = new A3Outline()
                    {
                        Course = this.Title,
                        Chapters = new List<A3Chapter>()
                    };
                    return a3Outline;
                case "CHAPTER":
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
                        // Figure out how to do the Notes page correctly
                        // Notes = this.Slide.NotesPage.Shapes[1].TextFrame.TextRange.Text
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
            try { this.ChapSub = this.Slide.Shapes["CHAP_SUB"].TextFrame.TextRange.Text; }
            catch { this.ChapSub = null; }
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
            foreach (string shapeName in A3Globals.SHAPE_NAMES)
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
            foreach (string shapeName in A3Globals.SHAPE_NAMES)
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

        }
        public void WriteActiveGuid()
        {
            PowerPoint.Shape aguid;
            try { aguid = this.Slide.Shapes["ACTIVE_GUID"]; } catch { aguid = this.MakeActiveGuid(); }
            aguid.TextFrame.TextRange.Text = this.ActiveGuid;
            aguid.Left = 0;
            aguid.Top = 400;
            aguid.Height = 30;
            aguid.Width = 500;
            aguid.Name = "ACTIVE_GUID";
            aguid.Title = "ACTIVE_GUID";
        }
        public void WriteHistoricGuid()
        {
            PowerPoint.Shape hguid;
            try { hguid = this.Slide.Shapes["HISTORIC_GUID"]; } catch { hguid = this.MakeHistoricGuid(); }
            string hguidText = null;
            foreach (string guid in this.HistoricGuids)
            {
                hguidText += guid;
            }
            hguid.TextFrame.TextRange.Text = hguidText;
            hguid.Left = 0;
            hguid.Top = 430;
            hguid.Height = 30;
            hguid.Width = 500;
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
            type.Left = 500;
            type.Top = 400;
            type.Height = 30;
            type.Width = 500;
            type.Name = "TYPE";
            type.Title = "TYPE";
        }
        public void WriteChapSub()
        {
            PowerPoint.Shape chapsub;
            try { chapsub = this.Slide.Shapes["CHAP:SUB"]; } catch { chapsub = this.MakeChapSub(); }
            chapsub.TextFrame.TextRange.Text = this.ChapSub;
            chapsub.Left = 12;
            chapsub.Height = 28;
            chapsub.Top = 1;
            chapsub.Width = 720;
            chapsub.TextFrame.TextRange.Characters().Font.Size = 16;
            chapsub.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
            chapsub.Name = "CHAP:SUB";
            chapsub.Title = "CHAP:SUB";
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
        private PowerPoint.Shape MakeSlideType()
        {
            PowerPoint.Shape type = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 500, 400, 500, 30);
            type.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            type.Name = "TYPE";
            type.Title = "TYPE";
            return type;
        }
        private PowerPoint.Shape MakeChapSub()
        {
            PowerPoint.Shape chapsub = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 12, 1, 720, 28);
            chapsub.TextFrame.TextRange.Characters().Font.Size = 16;
            chapsub.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
            chapsub.Name = "CHAP:SUB";
            chapsub.Title = "CHAP:SUB";
            return chapsub;
        }

        #region TODO: IMPLEMENT
        public void WriteTitle(string SlideType)
        {
            PowerPoint.Shape title;
            try { title = this.Slide.Shapes["TITLE"]; } catch { title = this.MakeTitle(SlideType); }
            title.TextFrame.TextRange.Text = this.Title;
            title.Name = "TITLE";
            title.Title = "TITLE";
            switch (SlideType)
            {
                case "COURSE":
                    title.Left = 12;
                    title.Height = 28;
                    title.Top = 1;
                    title.Width = 720;
                    title.TextFrame.TextRange.Characters().Font.Size = 16;
                    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
                    break;
                case "CHAPTER":
                    title.Left = 12;
                    title.Height = 28;
                    title.Top = 1;
                    title.Width = 720;
                    title.TextFrame.TextRange.Characters().Font.Size = 16;
                    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
                    break;
                case "QUESTION":
                    title.Left = 12;
                    title.Height = 28;
                    title.Top = 1;
                    title.Width = 720;
                    title.TextFrame.TextRange.Characters().Font.Size = 16;
                    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
                    break;
                default:
                    title.Left = 12;
                    title.Height = 28;
                    title.Top = 1;
                    title.Width = 720;
                    title.TextFrame.TextRange.Characters().Font.Size = 16;
                    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
                    break;
            }
        }
        public void WriteDay()
        {
            
        }

        private PowerPoint.Shape MakeTitle(string SlideType)
        {
            PowerPoint.Shape title;
            switch (SlideType)
            {
                case "COURSE":
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 10, 10, 30);
                    break;
                case "CHAPTER":
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 10, 10, 30);
                    break;
                case "QUESTION":
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 10, 10, 30);
                    break;
                default:
                    title = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 10, 10, 30);
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

        public void ValidateMetadata(A3LogFile logFile)
        {

        }
        #endregion
    }
}
