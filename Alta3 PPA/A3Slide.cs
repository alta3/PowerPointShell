using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA {
    public class A3Slide{

        public string Guid { get; set; }
        public List<string> HistoricGuids { get; set; }
        public string Type { get; set; }
        public string ChapSub { get; set; }
        public string Chapter { get; set; }
        public string Subchapter { get; set; }
        public string Title { get; set; }
        public string Notes { get; set; }
        public List<string> ShapeNames { get; set; }
        public PowerPoint.Slide Slide { get; set; }
        public enum SlideType {
            COURSE,
            TOC,
            CHAPTER,
            CONTENT,
            NOPUB,
            QUESTION
        }
        public static List<string> TypeStrings = new List<string> {
            "COURSE",
            "TOC",
            "CHAPTER",
            "CONTENT",
            "NOPUB",
            "QUESTION"
        };
        private enum MetadataError {
            Type,
            Title,
            ChapSub,
            Parse
        }
        private static readonly List<string> MetadataErrorStrings = new List<string> {
            "A valid type is expected for every slide. ",
            "A title is expected for every course, chapter, & content slide. ",
            "A Chap:Sub field is expected for every content slide. ",
            "The Chap:Sub field failed to parse into a chapter and subchapter field. "
        };

        public A3Slide(PowerPoint.Slide slide) {
            this.Slide = slide;
            this.ReadFromSlide();
        }

        public static void SetA3SlideFromPPTSlide(PowerPoint.Slide slide) {
            A3Globals.A3SLIDE = new A3Slide(slide);
        }
        public static void ShowMetadataForm() {
            A3Globals.A3SLIDE.Slide.Select();
            SlideMetadata slideMetadata = new SlideMetadata() {
                StartPosition = FormStartPosition.CenterScreen
            };
            slideMetadata.DrawSlideInfo();
            slideMetadata.ShowDialog();
        }
        
        public static void FixMetadataErrors(bool reviewed, A3LogFile logFile) {
            A3Globals.A3SLIDE.Slide.Select();
            A3Globals.A3SLIDE.ReadFromSlide();
            string msg = FindMetadataErrors();

            if (msg != null) {
                if (reviewed) {
                    DialogResult dialogResult = MessageBox.Show(msg, "Still Contains Error! Would You Like To Fix It Now?", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.No) {
                        logFile.WriteError(msg);
                        return;
                    }
                }
                ShowMetadataForm();
                FixMetadataErrors(true, logFile);
            }
        }
        public static string FindMetadataErrors() {
            A3Globals.A3SLIDE.Slide.Select();
            string msg = "";
            if (A3Globals.A3SLIDE.Guid == null) {
                A3Globals.A3SLIDE.Guid = System.Guid.NewGuid().ToString();
                A3Globals.A3SLIDE.WriteActiveGuid();
            }
            if (A3Globals.A3SLIDE.Type == null) {
                msg = String.Concat(MetadataErrorStrings[(int)MetadataError.Type], "Please Check slide number: ", A3Globals.A3SLIDE.Slide.SlideNumber.ToString());
                return msg;
            }
            if (A3Globals.A3SLIDE.Type == TypeStrings[(int)SlideType.COURSE] || A3Globals.A3SLIDE.Type == TypeStrings[(int)SlideType.CHAPTER] || A3Globals.A3SLIDE.Type == TypeStrings[(int)SlideType.CONTENT]) {
                if (A3Globals.A3SLIDE.Title == null || A3Globals.A3SLIDE.Title == "") {
                    msg = String.Concat(msg, MetadataErrorStrings[(int)MetadataError.Title]);
                }
            }
            if (A3Globals.A3SLIDE.ChapSub == null || A3Globals.A3SLIDE.ChapSub == "") {
                msg = String.Concat(msg, MetadataErrorStrings[(int)MetadataError.ChapSub]);
            }
            if (A3Globals.A3SLIDE.Chapter == null || A3Globals.A3SLIDE.Subchapter == null || A3Globals.A3SLIDE.Chapter == "" || A3Globals.A3SLIDE.Subchapter == "") {
                msg = String.Concat(msg, MetadataErrorStrings[(int)MetadataError.Parse]);
            }
            if (msg != "") {
                String.Concat(msg, "Please check slide number: ", A3Globals.A3SLIDE.Slide.SlideNumber.ToString());
            }
            else {
                msg = null;
            }
            return msg;
        }

        public static void NewBaseline(PowerPoint.Slide slide, A3LogFile logFile) {
            // Read slide and set it as current global slide
            SetA3SlideFromPPTSlide(slide);

            // Set new guid
            A3Globals.A3SLIDE.Guid = System.Guid.NewGuid().ToString();
            A3Globals.A3SLIDE.WriteActiveGuid();

            // Log metadata errors
            logFile.WriteError(FindMetadataErrors());

            // Reconstruct the chapter line and write it to the slide
            if (A3Globals.SLIDE_ITTERATION_AFTER_CHAPTER && !A3Globals.SLIDE_ITTERATION_AFTER_QUESTION && A3Globals.A3SLIDE.Type == TypeStrings[(int)SlideType.CONTENT]) {
                try { A3Globals.SLIDE_ITTERATION_CURRENT_SUBCHAPTER = A3Globals.A3SLIDE.Subchapter; }
                catch { A3Globals.SLIDE_ITTERATION_CURRENT_SUBCHAPTER = "Contents"; }
                A3Globals.A3SLIDE.ChapSub = String.Concat(A3Globals.SLIDE_ITTERATION_CURRENT_CHAPTER, @": ", A3Globals.SLIDE_ITTERATION_CURRENT_SUBCHAPTER);
            }
            A3Globals.A3SLIDE.WriteFromMemory();
        }

        // Utilized to change an A3Slide into an A3Outline Object for creating YAML files during the publishing process
        public object ObjectConversion() {
            switch (this.Type.ToLower()) {
                case "course":
                    A3Outline a3Outline = new A3Outline() {
                        Name = this.Title,
                        Chapters = new List<A3Chapter>()
                    };
                    return a3Outline;
                case "chapter":
                    A3Chapter a3Chapter = new A3Chapter() {
                        Guid = this.Guid,
                        HistoricGuids = this.HistoricGuids,
                        Title = this.Title,
                        Subchapters = new List<A3Subchapter>()
                    };
                    return a3Chapter;
                default:
                    A3Content a3Content = new A3Content() {
                        Guid = this.Guid,
                        HistoricGuids = this.HistoricGuids,
                        Title = this.Title,
                        Type = this.Type,
                        Notes = this.Notes,
                        Index = this.Slide.SlideIndex
                    };
                    return a3Content;
            }
        }

        public void ReadFromSlide() {
            this.ReadShapes();
            this.ReadActiveGuid();
            this.ReadHistoricGuid();
            this.ReadType();
            this.ReadTitle();
            this.ReadChapSub();
            this.ReadChapter();
            this.ReadSubchapter();
        }
        public void ReadActiveGuid() {
            try { this.Guid = this.Slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text; }
            catch { this.Guid = null; }
        }
        public void ReadHistoricGuid() {
            try {
                string guids = this.Slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text;
                this.HistoricGuids = new List<string>();
                this.HistoricGuids.AddRange(guids.Split(';'));
            }
            catch {
                this.HistoricGuids = null;
            }
        }
        public void ReadType() {
            try {
                this.Type = this.Slide.Shapes["TYPE"].TextFrame.TextRange.Text;
            }
            catch {
                if (A3Globals.ALLOW_INFER_FROM_SLIDE) {
                    this.InferType();
                    this.ReadFromSlide();
                }
                else {
                    this.Type = null;
                }
            }
        }
        public void ReadChapSub() {
            try { this.ChapSub = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text; }
            catch {
                if (A3Globals.ALLOW_INFER_FROM_SLIDE) {
                    this.InferChapSub(this.Type);
                }
            }
        }
        public void ReadChapter() {
            try { this.Chapter = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text.Split(':')[0].Trim(); }
            catch { this.Chapter = null; }
        }
        public void ReadSubchapter() {
            try { this.Subchapter = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text.Split(':')[1].Trim(); }
            catch { this.Subchapter = null; }
        }
        public void ReadTitle() {
            try { this.Title = this.Slide.Shapes["TITLE"].TextFrame.TextRange.Text; }
            catch {
                if (A3Globals.ALLOW_INFER_FROM_SLIDE == true) {
                    this.InferTitle(this.Type);
                }
            }
        }
        public void ReadNotes() {
            foreach (PowerPoint.Shape shape in this.Slide.NotesPage.Shapes) {
                if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue) {
                    if (shape.TextFrame.TextRange.Text != "") {
                        this.Notes = shape.TextFrame.TextRange.Text;
                        return;
                    }
                }
            }
        }
        public void ReadShapes() {
            this.ShapeNames = new List<string>();
            foreach (PowerPoint.Shape shape in this.Slide.Shapes) {
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue) {
                    if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue) {
                        if (shape.Name != null && shape.Name != "") {
                            this.ShapeNames.Add(shape.Name);
                        }
                    }
                }
            }
        }

        public void InferType() {
            PowerPoint.Shape type = this.MakeSlideType();
            // Check for shape names indications both chapter and questions slides
            List<string> chapShapeNames = new List<string> {
                "wordquan",
                "wordcounter",
                "wordsufer",
                "chapternumber",
                "vocabwordbox",
                "vocabbox"
            };
            List<string> questionShapeNames = new List<string> {
                "addquestion",
                "editquestion",
                "qindex",
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
            foreach (string shapeName in this.ShapeNames) {
                if (chapShapeNames.Contains(shapeName.ToLower())) {
                    this.Type = "CHAPTER";
                    this.WriteType();
                    return;
                }
                if (questionShapeNames.Contains(shapeName.ToLower())) {
                    this.Type = "QUESTION";
                    this.WriteType();
                    return;
                }
            }

            // Check for chapter slide size indications
            bool chapChapSub = false;
            bool chapTitle = false;
            foreach (string shapeName in this.ShapeNames) {
                PowerPoint.Shape shape = this.Slide.Shapes[shapeName];
                if (shape.Height >= 47 && shape.Height <= 55 && shape.Width >= 900 && shape.Width <= 1100 && shape.Top >= 5 && shape.Top <= 15) {
                    chapChapSub = true;
                }
                else if (shape.Height >= 45 && shape.Height <= 55 && shape.Width >= 850 && shape.Width <= 1100 && shape.Top >= 75 && shape.Top <= 85) {
                    chapTitle = true;
                }
            }
            if (chapTitle && chapChapSub) {
                this.Type = "CHAPTER";
                this.WriteType();
                return;
            }

            // Check For Course Slide Indications
            if (this.Slide.SlideNumber == 1) {
                this.Type = "COURSE";
                this.WriteType();
                return;
            }

            // Default To Type of Content If Allow Default Infer Is Set to True.
            if (A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE) {
                this.Type = "CONTENT";
                this.WriteType();
            }
        }
        public void InferChapSub(string type) {
            if (type == TypeStrings[(int)SlideType.CHAPTER]) {
                return;
            }

            foreach (string shapeName in this.ShapeNames) {
                PowerPoint.Shape shape;
                try { shape = this.Slide.Shapes[shapeName]; }
                catch { continue; }
                if (shape.Height >= 20 && shape.Height <= 33 && shape.Width >= 700 && shape.Width <= 1000 && shape.Top >= 0 && shape.Top <= 20) {
                    shape.Name = "CHAP:SUB";
                    shape.Title = "CHAP:SUB";
		    this.ReadChapSub();
                    return;
                }
            }

	    if (A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE) {
	        PowerPoint.Shape chapSub = this.Shapes[0];
                foreach (PowerPoint.Shape shape in this.Shapes) {
		    if (shape.HasTextFrame) {
		    	if (shape.TextFrame.HasText) {
			    if (shape.Left <= chapSub.Left && shape.Top < chapSub.Top && shape.Left >= this.Shapes["TITLE"].Left && shape.Top > this.Shapes["TITLE"].Top ) {
				chapSub = shape;
			    }
			}
		    }
		}
		chapSub.Name = "CHAP:SUB";
		chapSub.Title = "CHAP:SUB";
		this.ReadChapSub();
	    }
        }
        public void InferTitle(string type) {
            List<int> checks;
            if (A3Globals.A3SLIDE.Type == TypeStrings[(int)SlideType.COURSE] || A3Globals.A3SLIDE.Type == TypeStrings[(int)SlideType.CONTENT]) {
                checks = new List<int> { 45, 55, 850, 1100, 75, 85 }; 
            }
            else if (A3Globals.A3SLIDE.Type == TypeStrings[(int)SlideType.CHAPTER]) {
                checks = new List<int> { 30, 55, 850, 1100, 75, 85 }; //CHECK THESE NUMBERS PROBABLY NOT ACCURATE FOR TITLE ARE PROBABLY FOR VOCAB
            }
            else {
                return;
            }

            foreach (string shapeName in this.ShapeNames) {
                PowerPoint.Shape shape;
                try { shape = this.Slide.Shapes[shapeName]; }
                catch { continue; }
                if (shape.Height >= checks[0] && shape.Height <= checks[1] && shape.Width >= checks[2] && shape.Width <= checks[3] && shape.Top >= checks[4] && shape.Top <= checks[5]) {
                    shape.Name = "TITLE";
                    shape.Title = "TITLE";
                    this.ReadTitle();
                    return;
                }
            }
	    
	    if (A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE) {
	        PowerPoint.Shape title = this.Shapes[0];
                foreach (PowerPoint.Shape shape in this.Shapes) {
		    if (shape.HasTextFrame) {
		    	if (shape.TextFrame.HasText) {
			    if (shape.Left <= title.Left && shape.Top < title.Top) {
				title = shape;
			    }
			}
		    }
		}
		title.Name = "TITLE";
		title.Title = "TITLE";
		this.ReadTitle();
	    }
        }

        public void WriteFromMemory() {
            this.WriteType();
            this.WriteActiveGuid();
            this.WriteHistoricGuid();
            this.WriteTitle();
            this.WriteChapSub();
            this.WriteNotes();
        }
        public void WriteActiveGuid() {
            PowerPoint.Shape aguid;
            try {
                aguid = this.Slide.Shapes["ACTIVE_GUID"];
            }
            catch {
                aguid = this.MakeActiveGuid();
            }
            aguid.TextFrame.TextRange.Text = this.Guid;
            aguid.Name = "ACTIVE_GUID";
            aguid.Title = "ACTIVE_GUID";
        }
        public void WriteHistoricGuid() {
            PowerPoint.Shape hguid;
            try { hguid = this.Slide.Shapes["HISTORIC_GUID"]; } 
	    catch { hguid = this.MakeHistoricGuid(); }
            string hguidText = "";
            try {
                foreach (string guid in this.HistoricGuids {
                    hguidText += guid;
                }
            }
            catch {

            }
            hguid.TextFrame.TextRange.Text = hguidText;
            hguid.Name = "HISTORIC_GUID";
            hguid.Title = "HISTORIC_GUID";
        }
        public void WriteType() {
            PowerPoint.Shape type;
            try { type = this.Slide.Shapes["TYPE"]; }
            catch { type = this.MakeSlideType(); }

            try { type.TextFrame.TextRange.Text = this.Type.ToUpper(); }
            catch {  type.TextFrame.TextRange.Text = ""; }

            if (this.Type.ToUpper() == "COURSE" || this.Type.ToUpper() == "CHAPTER" || this.Type.ToUpper() == "QUESTION") {
                this.Slide.CustomLayout.Name = this.Type.ToUpper();
            }
            else {
                this.Slide.CustomLayout.Name = "CONTENT";
            }
            type.Name = "TYPE";
            type.Title = "TYPE";
        }
        public void WriteChapSub() {
            PowerPoint.Shape chapsub;
            try { chapsub = this.Slide.Shapes["CHAP:SUB"]; }
            catch { chapsub = this.MakeChapSub(); }
            chapsub.TextFrame.TextRange.Text = this.ChapSub;
            chapsub.Name = "CHAP:SUB";
            chapsub.Title = "CHAP:SUB";
        }
        public void WriteTitle() {
            PowerPoint.Shape title;
            try { title = this.Slide.Shapes["TITLE"]; }
            catch { title = this.MakeTitle(); }
            title.TextFrame.TextRange.Text = this.Title;
            title.Name = "TITLE";
            title.Title = "TITLE";
        }
        public void WriteNotes() {
            try {
                foreach (PowerPoint.Shape shape in this.Slide.NotesPage.Shapes) {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue) {
                        shape.TextFrame.TextRange.Text = this.Notes;
                    }
                }
            }
            catch { }
        }

        public PowerPoint.Shape MakeActiveGuid() {
            PowerPoint.Shape aguid = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 400, 500, 30);
            aguid.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            aguid.Name = "ACTIVE_GUID";
            aguid.Title = "ACTIVE_GUID";
            return aguid;
        }
        public PowerPoint.Shape MakeHistoricGuid() {
            PowerPoint.Shape hguid = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 430, 500, 30);
            hguid.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            hguid.Name = "HISTORIC_GUID";
            hguid.Title = "HISTORIC_GUID";
            return hguid;
        }
        public PowerPoint.Shape MakeSlideType() {
            PowerPoint.Shape type = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 500, 400, 500, 30);
            type.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            type.Name = "TYPE";
            type.Title = "TYPE";
            return type;
        }
        public PowerPoint.Shape MakeChapSub() {
            PowerPoint.Shape chapsub = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 12, 1, 720, 28);
            chapsub.TextFrame.TextRange.Characters().Font.Size = 16;
            chapsub.TextFrame.TextRange.Font.Color.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
            chapsub.Name = "CHAP:SUB";
            chapsub.Title = "CHAP:SUB";
            if (this.Type == "COURSE") {
                chapsub.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            return chapsub;
        }
        public PowerPoint.Shape MakeTitle() {
            PowerPoint.Shape title;
            switch (this.Type) {
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
    }
}
