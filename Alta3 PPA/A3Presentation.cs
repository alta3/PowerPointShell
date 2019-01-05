using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    class A3Presentation
    {
        #region Properties
        public string Path { get; set; }
        public PowerPoint.Presentation Presentation { get; set; }
        public List<A3Slide> Slides { get; set; }
        #endregion

        public A3Presentation(PowerPoint.Presentation presentation)
        {
            Presentation = presentation;
            Slides = new List<A3Slide>();
            foreach (PowerPoint.Slide slide in Presentation.Slides)
            {
                Slides.Add(new A3Slide(slide));
            }
            Path = presentation.Path;
        }
        public void UpdateSlidesFromPresentation()
        {
            List<A3Slide> updatedSlides = new List<A3Slide>();
            foreach (PowerPoint.Slide slide in Presentation.Slides)
            {
                Slides.Add(new A3Slide(slide));
            }
            Slides = updatedSlides;
        }
        public void SavePresentationAs(string name)
        {
            string saveDir = String.Concat(A3Environment.A3_WORKING, "\\", name);
            try { Directory.CreateDirectory(saveDir); } catch { }
            string savePath = String.Concat(saveDir, "\\", name);
            int version = 0;
            while (File.Exists(String.Concat(savePath, ".pptm")))
            {
                version++;
                savePath = string.Concat(saveDir, "\\", name, version.ToString());
            }
            Presentation.SaveAs(String.Concat(savePath, ".pptm"));
            Path = savePath;
        }

        #region Generate Outline
        public A3Outline GenerateOutline(A3Log log)
        {
            // Set Enviornment
            A3Environment.Clean();

            // Create new blank outline
            A3Outline outline = new A3Outline();

            // Get the course info
            (outline.Course, outline.Filename, outline.HasLabs, outline.HasSlides, outline.HasVideos, outline.Weburl) = GetCourseInfo(log);

            // Retrieve each of the chapters contents. Each chapter will recurse its own internal tree to collect all the related content ie(subchapters and their content as well. ). 
            outline.Chapters = GetChapters(log);

            // Return the outline
            return outline;
        }
        private (string name, string filename, bool haslabs, bool hasslides, bool hasvideos, string weburl) GetCourseInfo(A3Log log)
        {
            // Set the default values for the course info 
            (string name, string filename, bool haslabs, bool hasslides, bool hasvideos, string weburl) = (null, null, false, false, false, null);

            // Find the course slide and log errors 
            A3Slide a3CourseSlide = GetCourse(log);

            // Split the notes section by the lines and then look for the specified metadata keys
            List<string> noteLines = new List<string>(a3CourseSlide.Notes.Split(new string[] { Environment.NewLine }, StringSplitOptions.None));
            foreach (string line in noteLines)
            {
                List<string> map = new List<string>(line.Trim().Split(':'));
                if (map.Count > 1)
                {
                    if (Enum.TryParse(map[0].Remove('-').Trim().ToUpper(), out A3Outline.Metadata enumValue))
                    {
                        switch (enumValue)
                        {
                            case A3Outline.Metadata.NAME:
                                name = map[1];
                                break;
                            case A3Outline.Metadata.FILENAME:
                                filename = map[1];
                                break;
                            case A3Outline.Metadata.HASLABS:
                                try { haslabs = Convert.ToBoolean(map[1].ToLower()); }
                                catch { log.Write(A3Log.Level.Warn, "Failed to convert has-labs value to a boolean. -- Defaulting to false."); }
                                break;
                            case A3Outline.Metadata.HASSLIDES:
                                try { hasslides = Convert.ToBoolean(map[1].ToLower()); }
                                catch { log.Write(A3Log.Level.Warn, "Failed to convert has-slides value to a boolean. -- Defaulting to false."); }
                                break;
                            case A3Outline.Metadata.HASVIDEOS:
                                try { hasvideos = Convert.ToBoolean(map[1].ToLower()); }
                                catch { log.Write(A3Log.Level.Warn, "Failed to convert has-videos value to a boolean. -- Defaulting to false."); }
                                break;
                            case A3Outline.Metadata.WEBURL:
                                weburl = map[1];
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            return (name, filename, haslabs, hasslides, hasvideos, weburl);
        }
        private A3Slide GetCourse(A3Log log)
        {
            return Slides?.FirstOrDefault(slide => slide.Type == A3Slide.Types.COURSE);
        }
        private List<A3Chapter> GetChapters(A3Log log)
        {
            return Slides?.FindAll(slide => slide.Type == A3Slide.Types.CHAPTER)
                             .Select(chap => {
                                 A3Chapter chapter = (A3Chapter)chap.TypeConversion();
                                 chapter.Subchapters.AddRange(GetSubChapters(log, chapter.Title));
                                 return chapter;
                             }).ToList();
        }
        private List<A3Subchapter> GetSubChapters(A3Log log, string chapterTitle)
        {
            return Slides?.FindAll(slide => (slide.Type == A3Slide.Types.CONTENT || slide.Type == A3Slide.Types.NOPUB) && string.Equals(slide.Chapter, chapterTitle, StringComparison.OrdinalIgnoreCase))
                                        .GroupBy(slide => slide.Subchapter).Select(sc => sc.ToList()).ToList()
                                        .Select(sub => {
                                            A3Subchapter subchapter = (A3Subchapter)sub[0].TypeConversion();
                                            subchapter.Slides = sub.Select(slide => (A3Content)slide.TypeConversion()).ToList();
                                            return subchapter;
                                        }).ToList();
        }
        #endregion

        #region Generate From Yaml
        public void GenerateFromYaml(string yamlPath)
        {
            // Set global variables after starting with a clean slate
            A3Environment.Clean();
            A3Environment.ALLOW_INFER_FROM_SLIDE = true;
            A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE = true;

            // Setup logging
            A3Log log = new A3Log(A3Log.Operations.GenerateFromYaml);

            // Ingest the yaml file
            A3Yaml a3Yaml = new A3Yaml(yamlPath);

            // Lint the YAML file before attempting to deserialize the outline and exit early if the user cancels the operation
            a3Yaml.Lint(log);
            if (A3Environment.QUIT_FROM_CURRENT_LOOP)
            {
                A3Environment.Clean();
                return;
            }

            // Create the outline from the YAML file and exit early if the user cancels the operation
            A3Outline outline = a3Yaml.Deserialize(log);
            if (A3Environment.QUIT_FROM_CURRENT_LOOP)
            {
                A3Environment.Clean();
                return;
            }

            // Open a copy of the model PowerPoint in the current PowerPoint context
            A3Presentation a3Presentation = new A3Presentation(Globals.ThisAddIn.Application.Presentations.Open(A3Environment.MODEL_POWERPOINT, 0, 0, Microsoft.Office.Core.MsoTriState.msoTrue));

            // Save the presentation to a unqiue location
            a3Presentation.SavePresentationAs(outline.Course);

            // Generate the Presentation
            a3Presentation.GenerateFromOutline(outline);
            // outline.GeneratePresentation(a3Presentation.Presentation);

            // Cleanup the initial slides
            for (int i = 0; i < 6; i++)
            {
                a3Presentation.Presentation.Slides[1].Delete();
            }

            // Save the generated presentation and handoff control back to the user
            a3Presentation.Presentation.Save();
            MessageBox.Show(A3Yaml.AlertDescriptions[A3Yaml.Alerts.YamlGenSuccess].Replace("{}", Path), "POWERPOINT GENERATION COMPLETE!", MessageBoxButtons.OK);
            A3Environment.Clean();
        }
        #endregion

        #region Generate From Outline
        public void GenerateFromOutline(A3Outline outline)
        {
            GenerateCourseSlide(outline.Course, outline.Filename, outline.HasLabs, outline.HasSlides, outline.HasVideos, outline.Weburl);
            GenerateTOCSlide(outline.Course);
            Presentation.SectionProperties.AddBeforeSlide(1, outline.Course);
            GenerateChapters(outline.Chapters);
            GenerateEndOfDeckSlide(outline.Course);
            GenerateQuizSlide();
        }
        private void GenerateCourseSlide(string course, string filename, bool haslabs, bool hasslides, bool hasvideos, string weburl)
        {
            // Insert the course slide from the model PowerPoint
            Presentation.Slides[1].Duplicate().MoveTo(Presentation.Slides.Count);

            // Change the title to the course title given in the yaml file
            A3Slide a3ActiveSlide = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title = course,
                Type = A3Slide.Types.COURSE,
                Guid = Guid.NewGuid().ToString(),
                Notes = string.Concat("name: ", course,
                                      "\r\nfilename: ", filename,
                                      "\r\nhas-labs: ", haslabs.ToString(),
                                      "\r\nhas-slides: ", hasslides.ToString(),
                                      "\r\nhas-videos: ", hasvideos.ToString(),
                                      "\r\nweburl: ", weburl)
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateTOCSlide(string course)
        {
            // Insert a split slide from the model PowerPoint
            Presentation.Slides[4].Duplicate().MoveTo(Presentation.Slides.Count);

            // Populate the appropriate values of the slide deck here
            A3Slide a3ActiveSlide = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title = "Table of Contents",
                ChapSub = string.Concat(course, ": TOC"),
                Type = A3Slide.Types.TOC,
                Guid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();

            // TODO: Create a linked list to the first chapter of each day and colorize the results
        }
        private void GenerateChapters(List<A3Chapter> chapters)
        {
            int chapterCount = 1;
            foreach (A3Chapter chapter in chapters)
            {
                chapter.Generate(Presentation, chapterCount);
                chapterCount += 1;
            }
        }
        private void GenerateEndOfDeckSlide(string course)
        {
            // Insert a title slide from the model PowerPoint
            Presentation.Slides[3].Duplicate().MoveTo(Presentation.Slides.Count);

            // Change the title, chapsub, type, and active guid to accurately reflect what is happening
            A3Slide a3ActiveSlide = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title = "End of Deck",
                ChapSub = string.Concat(course, ": End Of Deck"),
                Guid = Guid.NewGuid().ToString(),
                Type = A3Slide.Types.CONTENT
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateQuizSlide()
        {
            // Insert a question slide from the model PowerPoint
            Presentation.Slides[6].Duplicate().MoveTo(Presentation.Slides.Count);

            // Ensure the title is Knowledge Check and move on 
            A3Slide a3ActiveSlide = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title = "Knowledge Check",
                Type = A3Slide.Types.QUESTION,
                Guid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();
            Presentation.SectionProperties.AddBeforeSlide(Presentation.Slides.Count, "Knowledge Check");
        }
        #endregion

        public void FillSubChapters()
        {
            // Clean the global variables
            A3Environment.Clean();

            // Setup logging
            A3Log log = new A3Log(A3Log.Operations.FillSubChapters);

            // Initialize variables
            string subChapName = "Contents";
            int slideCounter = 1;

            foreach (A3Slide a3Slide in Slides)
            {
                switch (a3Slide.Type)
                {
                    case A3Slide.Types.CHAPTER:
                        log.Write(A3Log.Level.Info, "Slide number {} was identified as a Chapter slide.".Replace("{}", slideCounter.ToString()));
                        subChapName = "Contents";
                        A3Environment.AFTER_CHAPTER = true;
                        break;

                    case A3Slide.Types.CONTENT:
                        if (a3Slide.Subchapter != subChapName && A3Environment.AFTER_CHAPTER)
                        {
                            if (a3Slide.Subchapter == "Contents")
                            {
                                a3Slide.Subchapter = subChapName;
                                a3Slide.WriteSubchapter();
                                log.Write(A3Log.Level.Info, "Slide number {N} was identified as a Content slide which has a unique subchapter name: {SC}, which has overwritten the current \"Contents\" subchapter name.".Replace("{N}", slideCounter.ToString()).Replace("{SC}", subChapName));
                            }
                            else
                            {
                                subChapName = a3Slide.Subchapter;
                                log.Write(A3Log.Level.Info, "Slide number {N} was identified as a Content slide which has a new subchapter name: {SC}.".Replace("{N}", slideCounter.ToString()).Replace("{SC}", subChapName));
                            }
                        }
                        else
                        {
                            log.Write(A3Log.Level.Info, "Slide number {N} was identified as a Content slide which matched the prvious subchapter: {SC}.".Replace("{N}", slideCounter.ToString()).Replace("{SC}", subChapName));
                        }
                        break;

                    case A3Slide.Types.QUESTION:
                        A3Environment.Clean();
                        log.Write(A3Log.Level.Info, "Slide number {} was identified as a Question slide, no more slides will be parsed.".Replace("{}", slideCounter.ToString()));
                        return;
                }
                slideCounter++;
            }

            // Clean up the global variables state
            A3Environment.Clean();
        }
        public void NewBaseLine()
        {
            // Set Environment
            A3Environment.Clean();
            A3Environment.ALLOW_INFER_FROM_SLIDE = true;
            A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE = true;
            A3Environment.ENFORCE_CHAP_SUB_SPLITTING = false;

            // Setup logging
            A3Log log = new A3Log(A3Log.Operations.NewBaseline);

            SavePresentationAs("new-baseline");

            string chapterName = null;
            foreach (A3Slide a3Slide in Slides)
            {
                if (!A3Environment.QUIT_FROM_CURRENT_LOOP)
                {
                    a3Slide.HGuids.Add(a3Slide.Guid);
                    a3Slide.Guid = Guid.NewGuid().ToString();
                    a3Slide.FixNullMetadata(true, log);
                    if (a3Slide.Type == A3Slide.Types.CHAPTER)
                    {
                        A3Environment.AFTER_CHAPTER = true;
                        chapterName = a3Slide.Chapter;
                    }
                    if (a3Slide.Type == A3Slide.Types.QUESTION)
                    {
                        break;
                    }
                    if (A3Environment.AFTER_CHAPTER && a3Slide.Type == A3Slide.Types.CONTENT)
                    {
                        a3Slide.Chapter = chapterName;
                        a3Slide.WriteChapter();
                        a3Slide.Subchapter = "Contents";
                        a3Slide.WriteSubchapter();
                    }
                }
                else
                {
                    break;
                }
            }

            Presentation.Save();

            // Cleanup environemnt
            A3Environment.Clean();
        }
        public void FixAllMetadata(bool allowInfer, bool allowDefault)
        {
            //Set Enviornment
            A3Environment.Clean();
            A3Environment.ALLOW_INFER_FROM_SLIDE = allowInfer;
            A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE = allowDefault;

            // Setup logging
            A3Log log = new A3Log(A3Log.Operations.FixMetadata);

            foreach (A3Slide a3Slide in Slides)
            {
                if (!A3Environment.QUIT_FROM_CURRENT_LOOP)
                {
                    a3Slide.FixSlideMetadata(true, log);
                }
            }

            A3Environment.Clean();
        }
    }
}
