using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Alta3_PPA
{
    public class A3Slide {
        #region Properties
        public enum Alerts {
            MessageInfo,
            SlideInfo,
            TypeIsNull,
            TypeDefaultInfered,
            TitleIsNull,
            ChapSubContainsNull,
            Continue,
            Success,
        }
        public enum Types {
            COURSE,
            TOC,
            CHAPTER,
            CONTENT,
            NOPUB,
            QUESTION,
            NULL
        };
        public enum Tags {
            GUID,
            HGUID,
            TYPE,
            CHAPSUB,
            TITLE,
            TOC
        };

        public static Dictionary<Alerts, string> AlertMessages = new Dictionary<Alerts, string>()
        {
            { Alerts.MessageInfo, "Properties Warning Message"},
            { Alerts.SlideInfo, "Slide Number: {} (Guid: {GUID}) ::" },
            { Alerts.TypeIsNull, "The TYPE is null and was not default infered based on the calling functions settings. It will either need to be default infered to content or manually fixed before the slide can be published."},
            { Alerts.TypeDefaultInfered, "The TYPE was infered to the default type of CONTENT. This may not be what was desired and could potentially cause parsing problems for the entire slide deck." },
            { Alerts.TitleIsNull, "The TITLE is null. Slides that have a type of COURSE, CHAPTER, or CONTENT require a TITLE in order to be properly parsed. This will have to be manually fixed." },
            { Alerts.ChapSubContainsNull, "This slide was identified as a CONTENT slide type and the CHAPSUB field contains a null field. CONTENT slides require that at least the CHAPTER field exist. The also potentially require a SUBCHAPTER to exist depending on the global settings. If the SUBCHAPTER is null this will be required to be fixed before publishing can occur." },
            { Alerts.Continue, "Click Yes if you would like to fix the issue(s) now or click No to exit without fixing the issue(s)."}
        };
        public static Dictionary<string, List<Tags>> OldToNewMap = new Dictionary<string, List<Tags>>()
        {
            { "scrubber", new List<Tags>() { Tags.CHAPSUB, Tags.TITLE } },
            { "chap:sub", new List<Tags>() { Tags.CHAPSUB } },
            { "historic_guid", new List<Tags>() { Tags.HGUID } },
            { "historic_guids", new List<Tags>() { Tags.HGUID } }
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

        public void FixMetadata(A3Log log, bool alert)
        {
            List<Alerts> alerts = AlertOrDefaultMetadataValues();
            while (alerts.Count > 0 && A3Environment.QUIT_FROM_CURRENT_LOOP is false)
            {
                alerts.Insert(0, Alerts.SlideInfo);
                string message = string.Join(" ", alerts.Select(a => AlertMessages[a].Replace("{SN}", Slide.SlideNumber.ToString())
                                                                                     .Replace("{GUID}", Guid))
                                                                                     .ToList());
                log.Write(A3Log.Level.Warn, message);
                if (alerts.All(a => a is Alerts.TypeDefaultInfered)) return;
                else if (alert)
                {
                    Slide.Select();
                    message = string.Concat(message, AlertMessages[Alerts.Continue]);
                    DialogResult dialogResult = MessageBox.Show(message, AlertMessages[Alerts.MessageInfo], MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.No) return;
                }
                ShowMetadataForm();
                FixMetadata(log, true);
            }
        }
        public List<Alerts> AlertOrDefaultMetadataValues()
        {
            List<Alerts> alerts = new List<Alerts>();

            if (Type is Types.NULL)
            {
                Type = A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE ? Types.CONTENT : Types.NULL;
                alerts.Add(Type is Types.NULL ? Alerts.TypeIsNull : Alerts.TypeDefaultInfered);
                if (Type is Types.CONTENT) WriteTag(Tags.TYPE);
            }

            if (Guid is null)
            {
                Guid = System.Guid.NewGuid().ToString();
                WriteTag(Tags.GUID);
            }

            alerts.Add((Chapter is null || (Subchapter is null && A3Environment.ENFORCE_CHAP_SUB_SPLITTING)) && Type is Types.CONTENT ? Alerts.ChapSubContainsNull : Alerts.Success);
            alerts.Add(Title is null && (Type is Types.COURSE || Type is Types.CHAPTER || Type is Types.CONTENT) ? Alerts.TitleIsNull : Alerts.Success);
            alerts.RemoveAll(a => a is Alerts.Success);

            return alerts;
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
                    if (tag is Tags.TYPE) InferType();
                    else if (tag is Tags.CHAPSUB || tag is Tags.TITLE) InferFromTypeTag(Type, tag);
                    return;
                }
                if (tag is Tags.GUID) Guid = null;
                if (tag is Tags.HGUID) HGuids = null;
                if (tag is Tags.TYPE) Type = Types.NULL;
                if (tag is Tags.TITLE) Title = null;
                if (tag is Tags.CHAPSUB)
                {
                    Chapter = null;
                    Subchapter = null;
                }
                return;
            }
            if (tag is Tags.GUID) Guid = shape.TextFrame.TextRange.Text;
            if (tag is Tags.HGUID) HGuids = shape.TextFrame.TextRange.Text.Split(';').ToList();
            if (tag is Tags.TYPE) Type = Enum.TryParse(shape.TextFrame.TextRange.Text, true, out Types t) ? t : Types.NULL;
            if (tag is Tags.TITLE) Title = shape.TextFrame.TextRange.Text;
            if (tag is Tags.CHAPSUB) SplitChapSub(shape);
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
            WriteTag(Tags.TYPE);
            WriteTag(Tags.GUID);
            WriteTag(Tags.HGUID);
            WriteTag(Tags.TITLE);
            WriteTag(Tags.CHAPSUB);
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
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue) shape.TextFrame.TextRange.Text = Notes;
            }
        }
        public Shape GetShapeByTag(Tags tag)
        {
            var eShapes = Slide.Shapes.GetEnumerator();
            while (eShapes.MoveNext())
            {
                Shape shape = (Shape)eShapes.Current;
                if (string.Equals(shape.Name, tag.ToString(), StringComparison.OrdinalIgnoreCase)) return shape;
            }
            return null;
        }
        public Shape MakeTag(Types type, Tags tag)
        {
            List<int> sDim = GetDimensions(type, tag);
            Shape shape = Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, sDim[0], sDim[1], sDim[2], sDim[3]);
            shape.Name = tag.ToString();
            if (tag is Tags.GUID || tag is Tags.HGUID || tag is Tags.TYPE) shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            if (tag is Tags.CHAPSUB)
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
                WriteTag(Tags.TYPE);
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
                    WriteTag(Tags.TYPE);
                    return;
                }
                if (questionShapeNames.Contains(s.Name.ToLower()))
                {
                    Type = Types.QUESTION;
                    WriteTag(Tags.TYPE);
                    return;
                }
                chapterTitle = SatisfiesDimensions(s, Types.CHAPTER, Tags.TITLE) || chapterTitle ? true : false;
                chapterChapSub = SatisfiesDimensions(s, Types.CHAPTER, Tags.CHAPSUB) || chapterChapSub ? true : false;
                if (chapterTitle && chapterChapSub)
                {
                    Type = Types.CHAPTER;
                    WriteTag(Tags.TYPE);
                    return;
                }
            });

            Type = A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE ? Types.CONTENT : Types.NULL;
            WriteTag(Tags.TYPE);
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

        #region Scrub Metadata
        public void ScrubMetadata(string search, Tags tag)
        {
            ReadFromSlide();
            search = search.ToLower().Trim();
            if (search is null)
            {
                Shapes?.ForEach(s =>
                {
                    string name = s.Name.ToLower().Trim();
                    if (OldToNewMap.TryGetValue(name, out List<Tags> tags))
                    {
                        int index = string.Equals(name, "scrubber") && (Type is Types.COURSE || Type is Types.CHAPTER) ? 1 : 0;
                        tag = tags[index];
                        ScrubTag(s, tag);
                    }
                });
                return;
            }
            Shapes?.ForEach(s =>
            {
                if (string.Equals(search, s.Name.ToLower().Trim()))
                {
                    ScrubTag(s, tag);
                    return;
                }
            });
        }
        private void ScrubTag(Shape shape, Tags tag)
        {
            Shape tagShape = GetShapeByTag(tag);
            string value = tagShape?.TextFrame?.TextRange?.Text;
            tagShape?.Delete();
            shape.TextFrame.TextRange.Text = value;
            shape.Name = tag.ToString();
        }
        #endregion

        #region FillSubchapter
        public string FillSubchapter(A3Log log, A3Slide slide, string subchapter, int count)
        {
            switch (slide.Type)
            {
                case Types.CHAPTER:
                    log.Write(A3Log.Level.Info, "Slide number {} was identified as a Chapter slide.".Replace("{}", count.ToString()));
                    subchapter = "Contents";
                    A3Environment.AFTER_CHAPTER = true;
                    break;
                case Types.CONTENT:
                    if (slide.Subchapter != subchapter && A3Environment.AFTER_CHAPTER)
                    {
                        if (string.Equals(slide.Subchapter, "Contents"))
                        {
                            slide.Subchapter = subchapter;
                            slide.WriteTag(Tags.CHAPSUB);
                            log.Write(A3Log.Level.Info, "Slide number {N} was identified as a Content slide which has a unique subchapter name: {SC}, which has overwritten the current \"Contents\" subchapter name.".Replace("{N}", count.ToString()).Replace("{SC}", subchapter));
                        }
                        else
                        {
                            subchapter = slide.Subchapter;
                            log.Write(A3Log.Level.Info, "Slide number {N} was identified as a Content slide which has a new subchapter name: {SC}.".Replace("{N}", count.ToString()).Replace("{SC}", subchapter));
                        }
                    }
                    else
                    {
                        log.Write(A3Log.Level.Info, "Slide number {N} was identified as a Content slide which matched the prvious subchapter: {SC}.".Replace("{N}", count.ToString()).Replace("{SC}", subchapter));
                    }
                    break;
                case Types.QUESTION:
                    A3Environment.Clean();
                    log.Write(A3Log.Level.Info, "Slide number {} was identified as a Question slide, no more slides will be parsed.".Replace("{}", count.ToString()));
                    break;
            }
            return subchapter;
        }
        #endregion

        public void WriteMarkdown()
        {
            Encoding utf8 = Encoding.UTF8;
            Encoding ascii = Encoding.ASCII;

            string asciiMarkdown = Notes is null ? "" : ascii.GetString(Encoding.Convert(utf8, ascii, utf8.GetBytes(Notes)));
            string markdownPath = string.Concat(A3Environment.A3_MARKDOWN, @"\", Guid, @".md");
            File.WriteAllText(markdownPath, asciiMarkdown);
        }
        public List<string> GetLatex()
        {
            string markdownPath = string.Concat(A3Environment.A3_MARKDOWN, @"\", Guid, @".md");
            if (File.Exists(markdownPath) is false) WriteMarkdown();
            ProcessStartInfo pandoc = new ProcessStartInfo()
            {
                CreateNoWindow = false,
                UseShellExecute = true,
                FileName = "pandoc.exe",
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = string.Concat(@"-f html -t latex -o ", "\"", A3Environment.A3_LATEX, @"\", "out.tex\" \"", markdownPath)
            };
            using (Process process = Process.Start(pandoc))
            {
                process.WaitForExit();
            }

            string[] latex = File.ReadAllLines(string.Concat(A3Environment.A3_LATEX, @"\out.tex"));
            File.Delete(string.Concat(A3Environment.A3_LATEX, @"\out.tex"));
            List<string> newtex = latex.ToList();

            return newtex;
        }

        // TODO: NEED TO GET THE DIMENSIONS FOR CHAPSUB && TITLE ON CHAPTERS VS CONTENT. ALSO NEED TO DECIDE HOW THIS INFO WILL BE RECORDED. I AM THINKING TITLE MAKES THE MOST SENSE.
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
