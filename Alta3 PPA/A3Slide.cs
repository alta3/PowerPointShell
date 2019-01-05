using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Slide
    {
        #region Properties
        public enum Types
        {
            COURSE,
            TOC,
            CHAPTER,
            CONTENT,
            NOPUB,
            QUESTION,
            NULL
        };
        public enum Tags
        {
            GUID,
            HGUID,
            TYPE,
            CHAPSUB,
            TITLE
        };

        public string Guid { get; set; }
        public List<string> HGuids { get; set; }
        public Types Type { get; set; }
        public string Chapter { get; set; }
        public string Subchapter { get; set; }
        public string Title { get; set; }
        public string Notes { get; set; }
        public List<Shape> Shapes { get; set; }
        public Slide Slide { get; set; }
        #endregion

        public A3Slide(Slide slide)
        {
            // Set default properties
            Guid = null;
            HGuids = new List<string>();
            Type = Types.NULL;
            Chapter = null;
            Subchapter = null;
            Title = null;
            Notes = null;
            Shapes = new List<Shape>();
            Slide = slide;

            // Read in the current information from the slide
            ReadFromSlide();
        }
        public object TypeConversion()
        {
            switch (Type)
            {
                case Types.COURSE:
                    A3Outline outline = new A3Outline()
                    {
                        Course = Title,
                        Chapters = new List<A3Chapter>()
                    };
                    return outline;
                case Types.CHAPTER:
                    A3Chapter chapter = new A3Chapter()
                    {
                        Guid = Guid,
                        HistoricGuids = HGuids,
                        Title = Title,
                        Subchapters = new List<A3Subchapter>()
                    };
                    return chapter;
                default:
                    A3Content content = new A3Content()
                    {
                        Guid = Guid,
                        HistoricGuids = HGuids,
                        Title = Title,
                        Type = Type.ToString(),
                        Notes = Notes,
                        Index = Slide.SlideIndex
                    };
                    return content;
            }
        }
        public void ShowMetadataForm()
        {
            Slide.Select();
            SlideMetadata slideMetadata = new SlideMetadata()
            {
                StartPosition = FormStartPosition.CenterScreen
            };
            slideMetadata.DrawSlideInfo();
            slideMetadata.Show();
        }

        // TODO: Clean this code to not have to be a static solution. 
        public static void FixNullMetadata(bool firstCheck, A3Log log)
        {
            A3Environment.A3SLIDE.Slide.Select();
            string msg = null;

            List<string> typesAllowed = new List<string> {
                "course",
                "toc",
                "chapter",
                "content",
                "no-pub",
                "question"
            };
            A3Slide.ScrubMetadata(A3Environment.A3SLIDE);

            if (A3Environment.A3SLIDE.Type == null)
            {
                msg = String.Concat("A Type Must Be Specified -- please check slide number: ", A3Environment.A3SLIDE.Slide.SlideIndex);
            }
            else if (!typesAllowed.Contains(A3Environment.A3SLIDE.Type.ToLower()))
            {
                msg = String.Concat("A Proper Type Must Be Specified -- please check slide number: ", A3Environment.A3SLIDE.Slide.SlideIndex);
                if (A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE == true)
                {
                    A3Environment.A3SLIDE.Type = "CONTENT";
                    A3Environment.A3SLIDE.WriteType();
                    msg = null;
                }
            }
            else if (A3Environment.A3SLIDE.Guid == null)
            {
                msg = String.Concat("An ActiveGuid Must Be Specified For Every Slide -- please check slide number: ", A3Environment.A3SLIDE.Slide.SlideIndex);
            }
            else if (A3Environment.A3SLIDE.Type.ToUpper() == "CONTENT")
            {
                if (A3Environment.A3SLIDE.Title == null ||
                    A3Environment.A3SLIDE.ChapSub == null ||
                    (A3Environment.A3SLIDE.Chapter == null && A3Environment.ENFORCE_CHAP_SUB_SPLITTING == true) ||
                    (A3Environment.A3SLIDE.Subchapter == null && A3Environment.ENFORCE_CHAP_SUB_SPLITTING == true))
                {
                    msg = String.Concat("A Title, ActiveGuid, and ChapSub must be specified. Chapter and Subchapter must be split by the \":\" character -- please check slide number: ", A3Environment.A3SLIDE.Slide.SlideIndex);
                }
            }
            else if (A3Environment.A3SLIDE.Type.ToUpper() == "CHAPTER")
            {
                if (A3Environment.A3SLIDE.Title == null)
                {
                    msg = String.Concat("A Title and ActiveGuid must be specified -- please check slide number: ", A3Environment.A3SLIDE.Slide.SlideIndex);
                }
            }
            else if (A3Environment.A3SLIDE.Type.ToUpper() == "COURSE")
            {
                if (A3Environment.A3SLIDE.Title == null)
                {
                    msg = String.Concat("A Title And AtiveGuid Must Be Specified -- please check slide number: ", A3Environment.A3SLIDE.Slide.SlideIndex);
                }
            }

            if (firstCheck)
            {
                if (msg != null)
                {
                    A3Slide.ShowMetadataForm();
                    A3Slide.FixNullMetadata(false, log);
                }
            }
            else
            {
                if (msg != null)
                {
                    log.Write(A3Log.Level.Error, msg);
                    DialogResult dialogResult = MessageBox.Show(msg, "Properties Still Contain A Null", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //A3Environment.A3SLIDE.ReadShapes();
                        A3Slide.ShowMetadataForm();
                        A3Slide.FixNullMetadata(false, log);
                    }
                }
            }
        }

        // TODO: Move this method to a more appropriate place perhaps A3Presentation? 
        public static void ScrubMetadata(A3Slide a3Slide)
        {
            a3Slide.ReadShapes();
            if (a3Slide.ShapeNames.Contains("SCRUBBER"))
            {
                if (a3Slide.Type.ToUpper() == "COURSE" || a3Slide.Type.ToUpper() == "CHAPTER")
                {
                    if (a3Slide.ShapeNames.Contains("TITLE"))
                    {
                        a3Slide.Slide.Shapes["TITLE"].Delete();
                    }
                    PowerPoint.Shape shape = a3Slide.Slide.Shapes["SCRUBBER"];
                    shape.Name = "TITLE";
                    shape.Title = "TITLE";
                }
                else
                {
                    if (a3Slide.ShapeNames.Contains("CHAPSUB"))
                    {
                        a3Slide.Slide.Shapes["CHAPSUB"].Delete();
                    }
                    PowerPoint.Shape shape = a3Slide.Slide.Shapes["SCRUBBER"];
                    shape.Name = "CHAPSUB";
                    shape.Title = "CHAPSUB";
                }
            }
        }

        #region Read Functions
        public void ReadFromSlide()
        {
            ReadShapes();
            ReadTag(Tags.GUID);
            ReadTag(Tags.HGUID);
            ReadTag(Tags.TYPE);
            ReadTag(Tags.TITLE);
            ReadTag(Tags.CHAPSUB);
            ReadNotes();
        }
        public void ReadShapes()
        {
            Shapes.Clear();
            var eShapes = Slide.Shapes.GetEnumerator();
            while (eShapes.MoveNext())
            {
                Shape shape = (Shape)eShapes.Current;
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    Shapes.Add(shape);
                }
            }
        }
        public void ReadTag(Tags tag)
        {
            Shape shape = Shapes?.FirstOrDefault(s => string.Equals(s.Name, tag.ToString(), StringComparison.OrdinalIgnoreCase));
            if (shape is null)
            {
                if (A3Environment.ALLOW_INFER_FROM_SLIDE)
                {
                    switch (tag)
                    {
                        case Tags.TYPE:
                            InferType();
                            return;
                        case Tags.CHAPSUB:
                            InferFromTypeTag(Type, Tags.CHAPSUB);
                            return;
                        case Tags.TITLE:
                            InferFromTypeTag(Type, Tags.TITLE);
                            return;
                        default:
                            break;
                    }
                }
                switch (tag)
                {
                    case Tags.GUID:
                        Guid = null;
                        return;
                    case Tags.HGUID:
                        HGuids = null;
                        return;
                    case Tags.TYPE:
                        Type = Types.NULL;
                        return;
                    case Tags.CHAPSUB:
                        Chapter = null;
                        Subchapter = null;
                        return;
                    case Tags.TITLE:
                        Title = null;
                        return;
                    default:
                        break;
                }
            }
            switch (tag)
            {
                case Tags.GUID:
                    Guid = shape.TextFrame.TextRange.Text;
                    break;
                case Tags.HGUID:
                    HGuids = shape.TextFrame.TextRange.Text.Split(';').ToList();
                    break;
                case Tags.TYPE:
                    Type = Enum.TryParse(shape.TextFrame.TextRange.Text, true, out Types t) ? t : Types.NULL;
                    break;
                case Tags.CHAPSUB:
                    SplitChapSub(shape);
                    break;
                case Tags.TITLE:
                    Title = shape.TextFrame.TextRange.Text;
                    break;
                default:
                    break;
            }

        }
        public void ReadNotes()
        {
            var eShapes = Slide.NotesPage.Shapes.GetEnumerator();
            while (eShapes.MoveNext())
            {
                Shape shape = (Shape)eShapes.Current;
                Notes = shape?.TextFrame?.TextRange?.Text;
            }
        }
        private void SplitChapSub(Shape chapsub)
        {
            string[] chapSubArr = chapsub.TextFrame?.TextRange.Text?.Split(':');
            Chapter = chapSubArr[0].Trim();
            Subchapter = chapSubArr.Length > 1 ? string.Join(":", chapSubArr.Skip(1)).Trim() : null;
        }
        #endregion

        #region Write Functions
        public void WriteFromMemory()
        {
            WriteType();
            WriteGuid();
            WriteHistoricGuid();
            WriteChapSub();
            WriteTitle();
            WriteNotes();
        }
        public void WriteTag(Tags tag)
        {
            Shape shape = GetShapeByTag(tag);
            shape = shape is null ? MakeTag(Type, tag) : shape;
            switch (tag)
            {
                case Tags.GUID:
                    shape.TextFrame.TextRange.Text = Guid;
                    break;
                case Tags.HGUID:
                    shape.TextFrame.TextRange.Text = string.Join(";", HGuids);
                    break;
                case Tags.TYPE:
                    Slide.CustomLayout.Name = Type.ToString();
                    shape.TextFrame.TextRange.Text = Type.ToString();
                    break;
                case Tags.CHAPSUB:
                    shape.TextFrame.TextRange.Text = Subchapter is null ? Chapter : string.Concat(Chapter, ": ", Subchapter);
                    break;
                case Tags.TITLE:
                    shape.TextFrame.TextRange.Text = Title;
                    break;
                default:
                    break;
            }
        }
        public void WriteNotes()
        {
            var eShapes = Slide.NotesPage.Shapes.GetEnumerator();
            while (eShapes.MoveNext())
            {
                Shape shape = (Shape)eShapes.Current;
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Text = Notes;
                }
            }
        }
        public Shape GetShapeByTag(Tags tag)
        {
            var eShapes = Slide.Shapes.GetEnumerator();
            while (eShapes.MoveNext())
            {
                Shape shape = (Shape)eShapes.Current;
                if (string.Equals(shape.Name, tag.ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    return shape;
                }
            }
            return null;
        }
        private Shape MakeTag(Types type, Tags tag)
        {
            List<int> sDim = GetDimensions(type, tag);
            Shape shape = Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, sDim[0], sDim[1], sDim[2], sDim[3]);
            shape.Name = tag.ToString();
            if (tag is Tags.GUID || tag is Tags.HGUID || tag is Tags.TYPE)
            {
                shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            else if (true)
            {
                shape.TextFrame.TextRange.Characters().Font.Size = 16;
                shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
                shape.Visible = Type is Types.COURSE ? Microsoft.Office.Core.MsoTriState.msoFalse : Microsoft.Office.Core.MsoTriState.msoTrue;
            }
            return shape;
        }
        #endregion

        #region Infer Functions
        private void InferType()
        {
            // Check For Course Slide Indications
            if (Slide.SlideNumber == 1)
            {
                Type = Types.COURSE;
                WriteType();
                return;
            }

            // Check for shape names indications both chapter and questions slides
            List<string> chapterShapeNames = new List<string>
            {
                "wordquan",
                "wordcounter",
                "wordsufer",
                "chapternumber",
                "vocabwordbox",
                "vocabbox"
            };
            List<string> questionShapeNames = new List<string>
            {
                "addquestion",
                "editquestion",
                "forward",
                "back",
                "qindex",
                "q",
                "bluetest",
                "greentest",
                "repair",
                "returnslide",
                "whyabox",
                "whybbox",
                "whycbox",
                "whydbox",
                "questionbox"
            };

            // Check for chapter slide size indications
            bool chapterChapSub = false;
            bool chapterTitle = false;

            Shapes?.ForEach(s =>
            {
                if (chapterShapeNames.Contains(s.Name.ToLower()))
                {
                    Type = Types.CHAPTER;
                    WriteType();
                    return;
                }
                if (questionShapeNames.Contains(s.Name.ToLower()))
                {
                    Type = Types.QUESTION;
                    WriteType();
                    return;
                }
                chapterTitle = SatisfiesDimensions(s, Types.CHAPTER, Tags.TITLE) || chapterTitle ? true : false;
                chapterChapSub = SatisfiesDimensions(s, Types.CHAPTER, Tags.CHAPSUB) || chapterChapSub ? true : false;
                if (chapterTitle && chapterChapSub)
                {
                    Type = Types.CHAPTER;
                    WriteType();
                    return;
                }
            });

            Type = A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE ? Types.CONTENT : Types.NULL;
            WriteType();
        }
        private void InferFromTypeTag(Types type, Tags tag)
        {
            Shape shape = Shapes?.FirstOrDefault(s => SatisfiesDimensions(s, type, tag));
            if (shape is null)
            {
                switch (tag)
                {
                    case Tags.CHAPSUB:
                        Chapter = null;
                        Subchapter = null;
                        return;
                    case Tags.TITLE:
                        return;
                    default:
                        return;
                }
            }
            shape.Name = tag.ToString();
            switch (tag)
            {
                case Tags.CHAPSUB:
                    SplitChapSub(shape);
                    break;
                case Tags.TITLE:
                    Title = shape.TextFrame?.TextRange.Text;
                    break;
                default:
                    break;
            }
            
        }
        #endregion

        #region Dimension Helper Functions
        private bool SatisfiesDimensions(Shape s, Types type, Tags tag)
        {
            // sDim in order are hMin, hMax, wMin, wMax, tMin, tMax
            List<int> sDim = GetMinMaxDimensions(type, tag);
            bool height = IsInclusiveBetween(Convert.ToInt32(s.Height), sDim[0], sDim[1]);
            bool width = IsInclusiveBetween(Convert.ToInt32(s.Width), sDim[2], sDim[3]);
            bool top = IsInclusiveBetween(Convert.ToInt32(s.Top), sDim[4], sDim[5]);
            bool match = height && width && top ? true : false;
            return match;
        }
        private List<int> GetDimensions(Types type, Tags tag)
        {
            List<int> dimensions;
            switch (tag)
            {
                case Tags.GUID:
                    dimensions = new List<int>() { 0, 400, 500, 30 };
                    break;
                case Tags.HGUID:
                    dimensions = new List<int>() { 0, 430, 500, 30 };
                    break;
                case Tags.TYPE:
                    dimensions = new List<int>() { 500, 400, 500, 30 };
                    break;
                case Tags.CHAPSUB:
                    dimensions = type is Types.CHAPTER ? new List<int>() { 0, 1, 2, 3 } : new List<int>() { 0, 1, 2, 3 };
                    break;
                case Tags.TITLE:
                    switch (type)
                    {
                        case Types.COURSE:
                            dimensions = new List<int>() { 312, 192, 358, 106 };
                            break;
                        case Types.CHAPTER:
                            dimensions = new List<int>() { 18, 81, 934, 50 };
                            break;
                        case Types.QUESTION:
                            dimensions = new List<int>() { 30, 960, 50, 28 };
                            break;
                        default:
                            dimensions = new List<int>() { 12, 30, 936, 50 };
                            break;
                    }
                    break;
                default:
                    dimensions = new List<int>() { 0, 1, 2, 3 };
                    break;
            }
            return dimensions;
        }
        private List<int> GetMinMaxDimensions(Types type, Tags tag)
        {
            // dimensions in order are hMin, hMax, wMin, wMax, tMin, tMax
            List<int> dimensions = new List<int>();
            dimensions.AddRange(Enumerable.Repeat(0, 6));
            if (tag is Tags.TITLE)
            {
                dimensions = type is Types.CHAPTER ? new List<int>() { 45, 55, 850, 1100, 75, 85 } : new List<int>() { 30, 60, 600, 1100, 15, 50 };
            }
            else if (tag is Tags.CHAPSUB)
            {
                dimensions = type is Types.CHAPTER ? new List<int>() { 47, 55, 900, 1100, 5, 15 } : new List<int>() { 20, 33, 700, 1000, 0, 20 };
            }
            return dimensions;
        }
        private bool IsInclusiveBetween(int value, int Min, int Max)
        {
            bool inclusiveBetween = value >= Min && value <= Max ? true : false;
            return inclusiveBetween;
        }
        #endregion
    }
}
