using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using YamlDotNet.Serialization;

namespace Alta3_PPA
{
    class A3Presentation
    {
        public enum LatexLines {
            MAINCHAPTER,
            CHAPTER,
            CHAPTERSUBCHAPTER,
            SECTION,
            FIGURE,
            NOTESTART,
            NOTEEND
        }
        readonly public static Dictionary<LatexLines, string> LatexMap = new Dictionary<LatexLines, string>()
        {
            { LatexLines.MAINCHAPTER, "\\input{\"_LATEX_PATH_/chapters/_CHAPTER_TITLE_.tex\"}" },
            { LatexLines.CHAPTER, "\\chapter{_CHAPTER_TITLE_}\r\n\\newpage\r\n\r\n" },
            { LatexLines.CHAPTERSUBCHAPTER, "\\input{\"_LATEX_PATH_/chapters/_CHAPTER_TITLE_/_SUBCHAPTER_TITLE_.tex\"}" },
            { LatexLines.SECTION, "\\section{_SUBCHAPTER_TITLE_}" },
            { LatexLines.FIGURE,  "\\begin{figure}[H]\r\n\\includegraphics*[width=1\\linewidth, height=.425\\textheight, trim= 0 0 0 0, clip]{\"_BOOK_PNGS_/_GUID_}\r\n\\end{figure}" },
            { LatexLines.NOTESTART, "%SLIDE_INDEX_OF_ABOVE_FIGURE: _SLIDE_INDEX_\r\n\\begin{flushleft}" },
            { LatexLines.NOTEEND, "\\end{flushleft}\r\n%SLIDE_INDEX_OF_ABOVE_FIGURE: _SLIDE_INDEX_"}
};
        
        #region Properties
        public string Path { get; set; }
        public Presentation Presentation { get; set; }
        public A3Outline Outline { get; set; }
        public List<A3Slide> Slides { get; set; }
        #endregion

        public A3Presentation(Presentation presentation)
        {
            Presentation = presentation;
            Slides = new List<A3Slide>();
            foreach (Slide s in Presentation.Slides)
            {
                Slides.Add(new A3Slide(s));
            }
            Outline = GenerateOutline(new A3Log(A3Log.Operations.ToOutline));
            Path = presentation.Path;
        }
        public void UpdateSlidesFromPresentation()
        {
            List<A3Slide> updatedSlides = new List<A3Slide>();
            foreach (Slide s in Presentation.Slides)
            {
                Slides.Add(new A3Slide(s));
            }
            Slides = updatedSlides;
        }
        public void SavePresentationAs(string name)
        {
            string saveDir = string.Concat(A3Environment.A3_WORKING, "\\", name);
            Directory.CreateDirectory(saveDir);
            string savePath = string.Concat(saveDir, "\\", name);

            int version = 0;
            while (File.Exists(string.Concat(savePath, ".pptm")))
            {
                version++;
                savePath = string.Concat(saveDir, "\\", name, version.ToString());
            }

            Presentation.SaveAs(string.Concat(savePath, ".pptm"));
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
            A3Slide course = GetCourse(log);

            // Split the notes section by the lines and then look for the specified metadata keys
            List<string> noteLines = new List<string>(course.Notes.Split(new string[] { Environment.NewLine }, StringSplitOptions.None));
            foreach (string line in noteLines)
            {
                List<string> map = new List<string>(line.Trim().Split(':'));
                if (Enum.TryParse(map[0].Remove('-').Trim().ToUpper(), out A3Outline.Metadata enumValue) && map.Count > 1)
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
                    }
                }
            }
            return (name, filename, haslabs, hasslides, hasvideos, weburl);
        }
        private A3Slide GetCourse(A3Log log)
        {
            return Slides?.FirstOrDefault(s => s.Type is A3Slide.Types.COURSE);
        }
        private List<A3Chapter> GetChapters(A3Log log)
        {
            return Slides?.Where(s => s.Type is A3Slide.Types.CHAPTER)
                          .Select(c => {
                              A3Chapter chapter = new A3Chapter(c);
                              chapter.Subchapters.AddRange(GetSubChapters(log, chapter.Title));
                              return chapter;
                          }).ToList();
        }
        private List<A3Subchapter> GetSubChapters(A3Log log, string chapterTitle)
        {
            return Slides?.Where(s => (s.Type is A3Slide.Types.CONTENT) && string.Equals(s.Chapter, chapterTitle, StringComparison.OrdinalIgnoreCase))
                           .GroupBy(s => s.Subchapter).Select(sc => sc.ToList()).ToList()
                           .Select(sub => {
                               A3Subchapter subchapter = new A3Subchapter(sub[0]);
                               subchapter.Slides = sub.Select(s => new A3Content(s)).ToList();
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
            A3Yaml yaml = new A3Yaml(yamlPath);

            // Lint the YAML file before attempting to deserialize the outline and exit early if the user cancels the operation
            yaml.Lint(log);
            if (A3Environment.QUIT_FROM_CURRENT_LOOP)
            {
                A3Environment.Clean();
                return;
            }

            // Create the outline from the YAML file and exit early if the user cancels the operation
            A3Outline outline = yaml.Deserialize(log);
            if (A3Environment.QUIT_FROM_CURRENT_LOOP)
            {
                A3Environment.Clean();
                return;
            }

            // Open a copy of the model PowerPoint in the current PowerPoint context
            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.Presentations.Open(A3Environment.MODEL_POWERPOINT, 0, 0, Microsoft.Office.Core.MsoTriState.msoTrue));

            // Save the presentation to a unqiue location
            presentation.SavePresentationAs(outline.Course);

            // Generate the Presentation
            presentation.WriteFromOutline(outline);

            // Cleanup the initial slides
            for (int i = 0; i < 6; i++) presentation.Presentation.Slides[1].Delete();

            // Save the generated presentation and handoff control back to the user
            presentation.Presentation.Save();
            MessageBox.Show(A3Yaml.AlertMessages[A3Yaml.Alerts.YamlGenSuccess].Replace("{}", Path), "POWERPOINT GENERATION COMPLETE!", MessageBoxButtons.OK);
            A3Environment.Clean();
        }
        #endregion

        #region Write From Outline
        public void WriteFromOutline(A3Outline outline)
        {
            GenerateCourseSlide(outline.Course, outline.Filename, outline.HasLabs, outline.HasSlides, outline.HasVideos, outline.Weburl);
            GenerateTOCSlide(outline.Course);
            Presentation.SectionProperties.AddBeforeSlide(1, outline.Course);
            GenerateChapterSlides(outline.Chapters);
            GenerateEndOfDeckSlide(outline.Course);
            GenerateQuizSlide();
        }
        private void GenerateCourseSlide(string title, string filename, bool haslabs, bool hasslides, bool hasvideos, string weburl)
        {
            // Insert the course slide from the model PowerPoint
            Presentation.Slides[1].Duplicate().MoveTo(Presentation.Slides.Count);

            // Change the title to the course title given in the yaml file
            A3Slide course = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title = title,
                Type = A3Slide.Types.COURSE,
                Guid = Guid.NewGuid().ToString(),
                Notes = string.Concat("name: ",             title,
                                      "\r\nfilename: ",     filename,
                                      "\r\nhas-labs: ",     haslabs.ToString(),
                                      "\r\nhas-slides: ",   hasslides.ToString(),
                                      "\r\nhas-videos: ",   hasvideos.ToString(),
                                      "\r\nweburl: ",       weburl)
            };
            course.WriteFromMemory();
        }
        private void GenerateTOCSlide(string course)
        {
            // Insert a split slide from the model PowerPoint
            Presentation.Slides[4].Duplicate().MoveTo(Presentation.Slides.Count);

            // Populate the appropriate values of the slide deck here
            A3Slide slide = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title =         "Table of Contents",
                Chapter =       course,
                Subchapter =    "TOC",
                Type =          A3Slide.Types.TOC,
                Guid =          Guid.NewGuid().ToString()
            };
            slide.WriteFromMemory();
        }
        private void GenerateChapterSlides(List<A3Chapter> chapters)
        {
            int chapter = 1;
            chapters.ForEach(c =>
            {
                c.WriteToPresentation(Presentation, chapter);
                chapter++;
            });
        }
        private void GenerateEndOfDeckSlide(string course)
        {
            // Insert a title slide from the model PowerPoint
            Presentation.Slides[3].Duplicate().MoveTo(Presentation.Slides.Count);

            // Change the title, chapsub, type, and active guid to accurately reflect what is happening
            A3Slide slide = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title =         "End of Deck",
                Chapter =       course, 
                Subchapter =    "End of Deck",
                Guid =          Guid.NewGuid().ToString(),
                Type =          A3Slide.Types.CONTENT
            };
            slide.WriteFromMemory();
        }
        private void GenerateQuizSlide()
        {
            // Insert a question slide from the model PowerPoint
            Presentation.Slides[5].Duplicate().MoveTo(Presentation.Slides.Count);

            // Ensure the title is Knowledge Check and move on 
            A3Slide slide = new A3Slide(Presentation.Slides[Presentation.Slides.Count])
            {
                Title =     "Knowledge Check",
                Type =      A3Slide.Types.QUESTION,
                Guid =      Guid.NewGuid().ToString()
            };
            slide.WriteFromMemory();
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
            string subchapter = "Contents";
            int count= 1;

            Slides.ForEach(s =>
            {
                subchapter = s.FillSubchapter(log, s, subchapter, count);
                count++;
            });

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

            SavePresentationAs("new_baseline");

            string chapterName = null;
            foreach (A3Slide s in Slides)
            {
                if (A3Environment.QUIT_FROM_CURRENT_LOOP) return;
                if (!(s.Guid is null)) s.HGuids.Add(s.Guid);
                s.Guid = Guid.NewGuid().ToString();
                s.FixMetadata(log, false);
                if (s.Type == A3Slide.Types.CHAPTER)
                {
                    A3Environment.AFTER_CHAPTER = true;
                    chapterName = s.Chapter;
                    continue;
                }
                if (A3Environment.AFTER_CHAPTER && s.Type is A3Slide.Types.CONTENT)
                {
                    s.Chapter = chapterName;
                    s.Subchapter = "Contents";
                    s.WriteTag(A3Slide.Tags.CHAPSUB);
                    continue;
                }
                if (s.Type == A3Slide.Types.QUESTION) break;
            }

            Presentation.Save();

            // Cleanup environemnt
            A3Environment.Clean();
        }
        public void FixMetadata(bool allowInfer, bool allowDefault)
        {
            //Set Enviornment
            A3Environment.Clean();
            A3Environment.ALLOW_INFER_FROM_SLIDE = allowInfer;
            A3Environment.ALLOW_DEFAULT_INFER_FROM_SLIDE = allowDefault;

            // Setup logging
            A3Log log = new A3Log(A3Log.Operations.FixMetadata);

            // Fix Metadata
            Slides?.ForEach(s => { if (A3Environment.QUIT_FROM_CURRENT_LOOP is false) s.FixMetadata(log, false); });

            // Cleanup
            A3Environment.Clean();
        }
        public void ShowGuids()
        {
            Slides.ForEach(s =>
            {
                Shape guid = s.GetShapeByTag(A3Slide.Tags.GUID);
                if (guid is null)
                {
                    guid = s.MakeTag(A3Slide.Types.NULL, A3Slide.Tags.GUID);
                    s.Guid = Guid.NewGuid().ToString();
                    s.WriteTag(A3Slide.Tags.GUID);
                }
                guid.Visible = A3Environment.SHOW_GUID ? Microsoft.Office.Core.MsoTriState.msoFalse : Microsoft.Office.Core.MsoTriState.msoTrue;
                guid.Fill.ForeColor.RGB = 763355;
            });
            A3Environment.SHOW_GUID = A3Environment.SHOW_GUID ? false : true;
        }

        public void ScrubMetadata(string search, A3Slide.Tags tag)
        {
            Slides?.ForEach(s => s.ScrubMetadata(search, tag));
        }

        #region Publish Products
        public void PublishPNGs()
        {
            Slides.ForEach(s => s.Slide.Export(string.Concat(A3Environment.A3_PRES_PNGS, "\\", s.Guid, ".png"), "png", 1920, 1080));
            Parallel.ForEach(Directory.EnumerateFiles(A3Environment.A3_PRES_PNGS), picture =>
                {
                    Bitmap bmp = new Bitmap(picture);

                    int width = bmp.Width;
                    int height = bmp.Height;

                    Color p;

                    // make grayscale
                    for (int y = 0; y < height; y++)
                    {
                        for (int x = 0; x < width; x++)
                        {
                            p = bmp.GetPixel(x, y);
                            p = Color.FromArgb(255, (255 - p.R), (255 - p.G), (255 - p.B));
                            int avg = (p.R + p.G + p.B) / 3;
                            bmp.SetPixel(x, y, Color.FromArgb(p.A, avg, avg, avg));
                        }
                    }
                    bmp.Save(picture.Replace("pres_pngs", "book_pngs"));
                });
        }
        public void PublishMarkdown()
        {
            Slides.ForEach(s => s.WriteMarkdown());
        }
        public void PublishYaml()
        {
            // ORDER MATTERS: Make sure the YAML publishin always takes place last in order to ensure the integrity of the Outline for other portions of the publishing process
            // Remove nopub and null slides before publishing. Set the other metadata to null. May be configurable in the future for different levels of detail.
            Outline.Chapters.ForEach(c => {
                c.HGuids = null;
                c.Subchapters.ForEach(sub => {
                    sub.Slides.ForEach(s => {
                        s.Type = null;
                        s.Chapter = null;
                        s.Subchapter = null;
                        s.HGuids = null;
                    });
                });
            });

            // Build the serializer and create the YAML from the outline
            ISerializer serializer = new SerializerBuilder().Build();
            string yaml = serializer.Serialize(Outline);

            // Write the YAML to the proper location as indicated by A3Environment.A3_PUBLISH
            File.WriteAllText(string.Concat(A3Environment.A3_PUBLISH, @"\yaml.yml"), yaml);
        }
        public void PublishPDF()
        {
            if (Directory.EnumerateFiles(A3Environment.A3_BOOK_PNGS)?.Any() is false) PublishPNGs();
            if (Directory.EnumerateFiles(A3Environment.A3_MARKDOWN)?.Any() is false) PublishMarkdown();
            PublishLatex();
            ProcessStartInfo build = new ProcessStartInfo()
            {
                UseShellExecute = true,
                CreateNoWindow = true,
                FileName = "pdflatex.exe",
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = string.Concat(@"-job-name=", Outline.Course, @" -output-directory=", A3Environment.A3_PUBLISH, @" -aux-directory=", A3Environment.A3_LATEX, @"main.tex")
            };
            using (Process process = Process.Start(build))
            {
                process.WaitForExit();
            }
        }
        public void PublishLatex()
        {
            string latexPath = A3Environment.A3_LATEX.Replace('\\', '/');
            string resourcePath = A3Environment.A3_RESOURCE.Replace('\\', '/');

            List<string> main = new List<string>(File.ReadAllLines(A3Environment.MAIN_LATEX));
            main.ForEach(l => l.Replace("_RESOURCE_LOCATION_", resourcePath)
                               .Replace("_COURSE_TITLE_", Outline.Course));

            Outline.Chapters.ForEach(c => {
                Directory.CreateDirectory(string.Concat(A3Environment.A3_LATEX, @"\chapters\", c.Title));
                main.Add(LatexMap[LatexLines.MAINCHAPTER].Replace("_LATEX_PATH_", latexPath)
                                                         .Replace("_CHAPTER_TITLE", c.Title));
                c.PublishLatex();
            });

            main.AddRange(File.ReadAllLines(A3Environment.END_LATEX));
            File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"\main.tex"), main);
        }
        public void WriteLatex()
        {
            List<string> latex = new List<string>();
            Slides.ForEach(s => {
                latex.Add(s.Guid);
                latex.AddRange(s.GetLatex());
                latex.Add(s.Guid);
            });
            File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"\raw.tex"), latex);
        }
        #endregion


    }
}
