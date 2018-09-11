using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA {
    public class A3Slide {

        public string Guid { get; set; }
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

        public A3Slide(PowerPoint.Slide slide) {
            this.Slide = slide;
            this.ReadFromSlide();
        }

        // TODO: Further test the functionality of the REST API that I setup... and expand/enhance its capability
        // TODO: Create the helper script to build the pdf for latex documents
        // TODO: Enhance the user interface dramatically to improve functionality, make the settings persist in memory and have a place to set them seperate from everything else.
        // TODO: Clean the code and make it more consistent throughout
        // TODO: Get rid of references to SCRUBBER throughout the code and make it more unified. 
        // TODO: Determine whether or not this is a useful function call... I doubt it so get rid of SetA3Slide and Most Likely ShowMetadataForm calls. 
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
        // TODO: Clean this code to not have to be a static solution. 
        public static void FixNullMetadata(bool firstCheck, A3LogFile logFile) {
            A3Globals.A3SLIDE.Slide.Select();
            string msg = null;

            List<string> typesAllowed = new List<string> {
                "course",
                "toc",
                "chapter",
                "content",
                "no-pub",
                "question"
            };
            A3Slide.ScrubMetadata(A3Globals.A3SLIDE);

            if (A3Globals.A3SLIDE.Type == null) {
                msg = String.Concat("A Type Must Be Specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
            }
            else if (!typesAllowed.Contains(A3Globals.A3SLIDE.Type.ToLower())) {
                msg = String.Concat("A Proper Type Must Be Specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
                if (A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE == true) {
                    A3Globals.A3SLIDE.Type = "CONTENT";
                    A3Globals.A3SLIDE.WriteType();
                    msg = null;
                }
            }
            else if (A3Globals.A3SLIDE.Guid == null) {
                A3Globals.A3SLIDE.WriteActiveGuid();
            }
            else if (A3Globals.A3SLIDE.Type.ToUpper() == "CONTENT") {
                if (A3Globals.A3SLIDE.Title == null || A3Globals.A3SLIDE.ChapSub == null || (A3Globals.A3SLIDE.Chapter == null && A3Globals.ENFORCE_CHAP_SUB_SPLITTING == true) || (A3Globals.A3SLIDE.Subchapter == null && A3Globals.ENFORCE_CHAP_SUB_SPLITTING == true)) {
                    msg = String.Concat("A Title, ActiveGuid, and ChapSub must be specified. Chapter and Subchapter must be split by the \":\" character -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
                }
            }
            else if (A3Globals.A3SLIDE.Type.ToUpper() == "CHAPTER") {
                if (A3Globals.A3SLIDE.Title == null) {
                    msg = String.Concat("A Title and ActiveGuid must be specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
                }
            }
            else if (A3Globals.A3SLIDE.Type.ToUpper() == "COURSE") {
                if (A3Globals.A3SLIDE.Title == null) {
                    msg = String.Concat("A Title And AtiveGuid Must Be Specified -- please check slide number: ", A3Globals.A3SLIDE.Slide.SlideIndex);
                }
            }

            if (firstCheck) {
                if (msg != null) {
                    A3Slide.ShowMetadataForm();
                    A3Slide.FixNullMetadata(false, logFile);
                }
            }
            else {
                if (msg != null) {
                    logFile.WriteError(msg);
                    DialogResult dialogResult = MessageBox.Show(msg, "Properties Still Contain A Null", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes) {
                        A3Globals.A3SLIDE.ReadShapes();
                        A3Slide.ShowMetadataForm();
                    }
                    A3Slide.FixNullMetadata(false, logFile);
                }
            }     
        }
        
        // TODO: Move this method to a more appropriate place perhaps A3Presentation? 
        public static void NewBaseline(PowerPoint.Slide slide, string chapterName, bool before_chap, bool after_question, A3LogFile logFile) {
            // Set current slide
            A3Slide.SetA3SlideFromPPTSlide(slide);
            
            // Set new guid
            A3Globals.A3SLIDE.Guid = System.Guid.NewGuid().ToString();
            A3Globals.A3SLIDE.WriteActiveGuid();

            // Fix unacceptable null metadata fields
            A3Slide.FixNullMetadata(true, logFile);

            // Reconstruct the chapter line and write it to the slide
            if (!before_chap && !after_question && A3Globals.A3SLIDE.Type != "CHAPTER" && A3Globals.A3SLIDE.Type != "COURSE" && A3Globals.A3SLIDE.Type != "QUESTION") {
                A3Globals.A3SLIDE.ChapSub = String.Concat(chapterName, @": Contents");
            }
            A3Globals.A3SLIDE.WriteChapSub();
        }
        public static void ScrubMetadata(A3Slide a3Slide) {
            a3Slide.ReadShapes();
            if (a3Slide.ShapeNames.Contains("SCRUBBER")) {
                if (a3Slide.Type.ToUpper() == "COURSE" || a3Slide.Type.ToUpper() == "CHAPTER") {
                    if (a3Slide.ShapeNames.Contains("TITLE")) {
                        a3Slide.Slide.Shapes["TITLE"].Delete();
                    }
                    PowerPoint.Shape shape = a3Slide.Slide.Shapes["SCRUBBER"];
                    shape.Name = "TITLE";
                    shape.Title = "TITLE";
                }
                else {
                    if (a3Slide.ShapeNames.Contains("CHAP:SUB")) {
                        a3Slide.Slide.Shapes["CHAP:SUB"].Delete();
                    }
                    PowerPoint.Shape shape = a3Slide.Slide.Shapes["SCRUBBER"];
                    shape.Name = "CHAP:SUB";
                    shape.Title = "CHAP:SUB";
                }
            }
        }
        
        // TODO: Create a file for several types of enums that can be used throughout the code for more useful functions calls. 
        // TODO: Implement the following functions and consolidate and clean code to better fit within these contexts.
        public void CheckType() { }
        public void CheckActiveGuid() { }
        public void CheckHistoricGuid() { }
        public void CheckTitle() { }
        public void CheckChapter() { }
        public void CheckSubchapter() { }
        public void CheckPreviousTerms() { }
        public void CheckMetadata() { }

        // TODO: Document the following function to include its purpose; implementation; and how it works. 
        public object TypeConversion() {
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
                        Day = this.Day,
                        Title = this.Title,
                        Subchapters = new List<A3Subchapter>()
                    };
                    return a3Chapter;
                default:
                    A3Content a3Content = new A3Content() {
                        Guid = this.Guid,
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

        // TODO: Document the following functions to include their purpose; implementations; and basic understanding of how it gathers information.
        // TODO: Determine if any of this metadata should ever be let to fall into a null state, it is causing problems with YAML ingestion, probably better to find sane defaults: For now enabling INFER_FROM _SLIDE for YAML generation. 
        public void ReadFromSlide() {
            this.ReadShapes();
            this.ReadActiveGuid();
            this.ReadHistoricGuid();
            this.ReadType();
            this.ReadChapSub();
            this.ReadChapter();
            this.ReadSubchapter();
            this.ReadTitle();
            this.ReadDay();
        }
        public void ReadActiveGuid() {
            try { this.Guid = this.Slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text; }
            catch { this.Guid = null; }
        }
        public void ReadHistoricGuid() {
            try {
                string guids = this.Slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text;
                this.HistoricGuids = new List<string>();

                if (guids.Contains(';')) {
                    this.HistoricGuids.AddRange(guids.Split(';'));
                }
                else {
                    this.HistoricGuids.Add(guids);
                }
            }
            catch {
                this.HistoricGuids = null;
            }
        }
        public void ReadType()
        {
            try {
                this.Type = this.Slide.Shapes["TYPE"].TextFrame.TextRange.Text;
            }
            catch {
                if (A3Globals.ALLOW_INFER_FROM_SLIDE == true) {
                    this.InferType();
                    this.ReadShapes();
                }
                else {
                    this.Type = null;
                }
            }
        }
        public void ReadChapSub() {
            try {
                this.ChapSub = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text;
            }
            catch {
                if (A3Globals.ALLOW_INFER_FROM_SLIDE == true) {
                    string chapSub = this.InferChapSub(this.Type);
                    try {
                        this.ChapSub = this.Slide.Shapes[chapSub].TextFrame.TextRange.Text;
                        PowerPoint.Shape shape = this.Slide.Shapes[chapSub];
                        shape.Name = "CHAP:SUB";
                        shape.Title = "CHAP:SUB";
                        this.ReadShapes();
                    }
                    catch { 
                        this.ChapSub = null;
                    }
                }
                else {
                    this.ChapSub = null;
                }
                
            }
        }
        public void ReadChapter() {
            try {
                this.Chapter = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text.Split(':')[0].Trim();
            }
            catch {
                this.Chapter = null;
            }
        }
        public void ReadSubchapter() {
            try {
                this.Subchapter = this.Slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text.Split(':')[1].Trim();
            }
            catch {
                this.Subchapter = null;
            }
        }
        public void ReadTitle() {
            try {
                this.Title = this.Slide.Shapes["TITLE"].TextFrame.TextRange.Text;
            }
            catch {
                if (A3Globals.ALLOW_INFER_FROM_SLIDE == true) {
                    string title = this.InferTitle(this.Type);
                    try {
                        this.Title = this.Slide.Shapes[title].TextFrame.TextRange.Text;
                        PowerPoint.Shape shape = this.Slide.Shapes[title];
                        shape.Name = "TITLE";
                        shape.Title = "TITLE";
                        this.ReadShapes();
                    }
                    catch {
                        this.Title = null;
                    }
                }
                else {
                    this.Title = null;
                }
            }
        }
        public void ReadDay() {
            try {
                this.Day = this.Slide.Shapes["DAY"].TextFrame.TextRange.Text;
            }
            catch {
                this.Day = null;
            }
        }
        public void ReadNotes() {
            try {
                foreach (PowerPoint.Shape shape in this.Slide.NotesPage.Shapes) {
                    if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue) {
                        if (shape.TextFrame.TextRange.Text != "") {
                            this.Notes = shape.TextFrame.TextRange.Text;
                            break;
                        }
                    }
                }
            }
            catch {
                this.Notes = null;
            }
        }
        public void ReadShapes() {
            try {
                this.ShapeNames = new List<string>();
                foreach (PowerPoint.Shape shape in this.Slide.Shapes) {
                    try {
                        if (shape.TextFrame.TextRange.Text != null) {
                            this.ShapeNames.Add(shape.Name);
                        }
                    }
                    catch { }
                }
            }
            catch {
                this.ShapeNames = null;
            }
        }

        // TODO: Document the following functions to incldue their purpose; implementations; and basic understanding of how they preform infrences
        public void InferType() {
            // Check For Course Slide Indications
            if (this.Slide.SlideNumber == 1) {
                DialogResult dialogResult = MessageBox.Show("Is the first slide of this deck the Course Title Slide?", "Infering First Slides Type", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes) {
                    this.MakeSlideType();
                    this.Type = "COURSE";
                    this.Slide.Shapes["TYPE"].TextFrame.TextRange.Text = "COURSE";
                    return;
                }
            }

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
            try {
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
            }
            catch { }

            // Check for chapter slide size indications
            bool chapChapSub = false;
            bool chapTitle = false;
            foreach (string shapeName in this.ShapeNames) {
                if (this.Slide.Shapes[shapeName].Height >= 47
                    && this.Slide.Shapes[shapeName].Height <= 55
                    && this.Slide.Shapes[shapeName].Width >= 900
                    && this.Slide.Shapes[shapeName].Width <= 1100
                    && this.Slide.Shapes[shapeName].Top >= 5
                    && this.Slide.Shapes[shapeName].Top <= 15) {
                        chapChapSub = true;
                }
                else if (this.Slide.Shapes[shapeName].Height >= 45
                    && this.Slide.Shapes[shapeName].Height <= 55
                    && this.Slide.Shapes[shapeName].Width >= 850
                    && this.Slide.Shapes[shapeName].Width <= 1100
                    && this.Slide.Shapes[shapeName].Top >= 75
                    && this.Slide.Shapes[shapeName].Top <= 85) {
                        chapTitle = true;
                }
            }
            if (chapTitle && chapChapSub) {
                this.Type = "CHAPTER";
                this.WriteType();
                return;
            }

            // Default To Type of Content If Allow Default Infer Is Set to True.
            if (A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE == true) {
                this.MakeSlideType();
                this.Type = "CONTENT";
                this.Slide.Shapes["TYPE"].TextFrame.TextRange.Text = "CONTENT";
            }
        }
        public string InferChapSub(string type) {
            List<int> checks = new List<int>();
            switch (type) {
                case "CHAPTER":
                    checks.Add(47);
                    checks.Add(55);
                    checks.Add(900);
                    checks.Add(1100);
                    checks.Add(5);
                    checks.Add(15);
                    break;
                default:
                    checks.Add(20);
                    checks.Add(33);
                    checks.Add(700);
                    checks.Add(1000);
                    checks.Add(0);
                    checks.Add(20);
                    break;
            }

            if (this.ShapeNames.Count >= 1) {
                foreach (string shapeName in this.ShapeNames) {
                    try {
                        if (this.Slide.Shapes[shapeName].Height >= checks[0]
                            && this.Slide.Shapes[shapeName].Height <= checks[1]
                            && this.Slide.Shapes[shapeName].Width >= checks[2]
                            && this.Slide.Shapes[shapeName].Width <= checks[3]
                            && this.Slide.Shapes[shapeName].Top >= checks[4]
                            && this.Slide.Shapes[shapeName].Top <= checks[5]) {
                                return shapeName;
                        }
                    }
                    catch { }
                }
            }
            return null;
        }
        public string InferTitle(string type) {
            List<int> checks = new List<int>();
            switch (type) {
                case "CHAPTER":
                    checks.Add(45);
                    checks.Add(55);
                    checks.Add(850);
                    checks.Add(1100);
                    checks.Add(75);
                    checks.Add(85);
                    break;
                default:
                    checks.Add(30);
                    checks.Add(60);
                    checks.Add(600);
                    checks.Add(1100);
                    checks.Add(15);
                    checks.Add(50);
                    break;
            }

            if (this.ShapeNames.Count >= 1) {
                foreach (string shapeName in this.ShapeNames) {
                    if (this.Slide.Shapes[shapeName].Height >= checks[0]
                        && this.Slide.Shapes[shapeName].Height <= checks[1]
                        && this.Slide.Shapes[shapeName].Width >= checks[2]
                        && this.Slide.Shapes[shapeName].Width <= checks[3]
                        && this.Slide.Shapes[shapeName].Top >= checks[4]
                        && this.Slide.Shapes[shapeName].Top <= checks[5]) {
                            return shapeName;
                    }
                }
            }
            return null;
        }
        /* public string InferDay()
        {
            int slideIndex = this.Slide.SlideIndex;
            PowerPoint.Slide previousSlide = this.Slide.Application.ActivePresentation.Slides[slideIndex - 1];
            string previousDay = "1";
            try { previousDay = previousSlide.Shapes["DAY"].TextFrame.TextRange.Text; } catch { }
            return previousDay;
        }
        */

        // TODO: Document the following functiosn to include their purpose; implementations; and basic understading of how they write information to the slide.
        public void WriteFromMemory() {
            this.WriteType();
            this.WriteActiveGuid();
            this.WriteHistoricGuid();
            this.WriteChapSub();
            this.WriteTitle();
            this.WriteDay();
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
        public void WriteType() {
            PowerPoint.Shape type;
            try {
                type = this.Slide.Shapes["TYPE"];
            }
            catch {
                type = this.MakeSlideType();
            }
            try {
                type.TextFrame.TextRange.Text = this.Type.ToUpper();
            }
            catch {
                type.TextFrame.TextRange.Text = "";
            }
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
            try {
                chapsub = this.Slide.Shapes["CHAP:SUB"];
            }
            catch {
                chapsub = this.MakeChapSub();
            }
            chapsub.TextFrame.TextRange.Text = this.ChapSub;
            chapsub.Name = "CHAP:SUB";
            chapsub.Title = "CHAP:SUB";
        }
        public void WriteTitle() {
            PowerPoint.Shape title;
            try {
                title = this.Slide.Shapes["TITLE"];
            }
            catch {
                title = this.MakeTitle();
            }
            title.TextFrame.TextRange.Text = this.Title;
            title.Name = "TITLE";
            title.Title = "TITLE";
        }
        public void WriteDay() {
            PowerPoint.Shape day;
            try {
                day = this.Slide.Shapes["DAY"];
            }
            catch {
                day = this.MakeDay();
            }
            day.TextFrame.TextRange.Text = this.Day;
            day.Name = "DAY";
            day.Title = "DAY";
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

        // TODO: Document the following functiosn to include their purpose; implemenations; invocations; and basic understanding of how and when they are utilized. 
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
        private PowerPoint.Shape MakeTitle() {
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
        private PowerPoint.Shape MakeDay() {
            PowerPoint.Shape day = this.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 500, 430, 1000, 30);
            day.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            day.Name = "DAY";
            day.Title = "DAY";
            return day;
        }
    }
}
