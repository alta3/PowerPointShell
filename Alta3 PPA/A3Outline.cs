using System.Collections.Generic;
using System.IO;
using Markdig;
using YamlDotNet.Serialization;

namespace Alta3_PPA {
    public class A3Outline {
        #region Outline Properites
        public enum Metadata {
            NAME,
            FILENAME,
            HASLABS,
            HASSLIDES,
            HASVIDEOS,
            WEBURL
        }
        public enum LatexLines {
            MAINCHAPTERLINE
        }
        readonly public Dictionary<LatexLines, string> LatexMap = new Dictionary<LatexLines, string>()
        {
            { LatexLines.MAINCHAPTERLINE, string.Concat(@"\input{", "\"_LATEX_PATH_", @"/chapters/_CHAPTER_TITLE_.tex}") }
        };

        public string Course { get; set; }
        public string Filename { get; set; }
        public bool HasLabs { get; set; }
        public bool HasSlides { get; set; }
        public bool HasVideos { get; set; }
        public string Weburl { get; set; }
        public List<A3Chapter> Chapters { get; set; }
        #endregion

        public A3Outline()
        {
            Course = null;
            Filename = null;
            HasLabs = false;
            HasSlides = false;
            HasVideos = false;
            Weburl = null;
            Chapters = new List<A3Chapter>();
        }

        public void Publish()
        {
            
        }

        #region Publish LaTex
        public void GenerateLaTex()
        {
            GenerateLaTexMain();
            GenerateLaTexChapters();
            GenerateLaTexSubchapters();
        }

        private void GenerateLaTexMain()
        {
            string latexPath = A3Environment.A3_LATEX.Replace('\\', '/');
            string resourcePath = A3Environment.A3_RESOURCE.Replace('\\', '/');

            List<string> main = new List<string>(File.ReadAllLines(A3Environment.MAIN_LATEX));
            main.ForEach(l => l.Replace("_RESOURCE_LOCATION_", resourcePath)
                               .Replace("_COURSE_TITLE_", Course));

            Chapters.ForEach(c => {
                Directory.CreateDirectory(string.Concat(A3Environment.A3_LATEX, @"\chapters\", c.Title));
                main.Add(LatexMap[LatexLines.MAINCHAPTERLINE].Replace("_LATEX_PATH_", latexPath)
                                                         .Replace("_CHAPTER_TITLE", c.Title));
            });

            main.AddRange(File.ReadAllLines(A3Environment.END_LATEX));
            File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"\", "main.tex"), main);
        }
        private void GenerateLaTexChapters()
        {
            foreach (A3Chapter chapter in Chapters)
            {
                List<string> chap = new List<string>
                {
                    string.Concat(@"\chapter{", chapter.Title, @"}"),
                    @"\newpage",
                    ""
                };
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    Directory.CreateDirectory(string.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter.Title));
                    chap.Add(string.Concat(@"\input{", "\"", A3Environment.A3_LATEX.Replace('\\', '/'), @"/chapters/", chapter.Title, @"/", subchapter.Title, ".tex\"}"));
                }
                File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter.Title, @".tex"), chap);
            }
        }
        private void GenerateLaTexSubchapters()
        {
            string[] mdFiles = Directory.GetFiles(string.Concat(A3Environment.A3_MARKDOWN));
            List<string> htmlNotes = new List<string>();
            foreach (string filePath in mdFiles)
            {
                string aguid = filePath.Split('.')[0];
                string note = File.ReadAllText(filePath);
                if (note != null)
                {
                    if (note != "")
                    {
                        htmlNotes.Add(aguid);
                        htmlNotes.Add(Markdown.ToHtml(note));
                        htmlNotes.Add(aguid);
                    }
                }
            }
            File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"notes.html"), htmlNotes);
            List<string> notes = A3Note.GetLatex(outline, string.Concat(A3Environment.A3_LATEX, @"notes.html"));

            foreach (A3Chapter chapter in Chapters)
            {
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    List<string> sub = new List<string>
                    {
                        string.Concat(@"\section{", subchapter.Title, @"}")
                    };
                    foreach (A3Content a3Content in subchapter.Slides)
                    {
                        sub.Add(@"\begin{figure}[H]");
                        sub.Add(string.Concat(@"\includegraphics*[width=1\linewidth, height=.425\textheight, trim= 0 0 0 0, clip]{", "\"", A3Environment.A3_BOOK_PNGS.Replace('\\','/'),a3Content.Guid, "\"}"));
                        sub.Add(@"\end{figure}");
                        if (a3Content.Notes != null)
                        {
                            if (a3Content.Notes != "")
                            {
                                int startIndex = notes.FindIndex(s => s.Contains(a3Content.Guid));
                                int endIndex = notes.FindLastIndex(s => s.Contains(a3Content.Guid));
                                startIndex++;
                                endIndex--;

                                sub.Add(string.Concat(@"%SLIDE_INDEX_OF_ABOVE_FIGURE: ", a3Content.Index));
                                sub.Add(@"\begin{flushleft}");
                                for (int i = startIndex; i < endIndex; i++)
                                {
                                    sub.Add(notes[i]);
                                }
                                sub.Add(@"\end{flushleft}");
                                sub.Add(string.Concat(@"%SLIDE_INDEX_OF_ABOVE_TEXT: ", a3Content.Index));
                            }
                        }

                    }
                    sub.Add(@"\clearpage");
                    File.WriteAllLines(string.Concat(A3Environment.A3_LATEX, @"\chapters\", chapter.Title, @"\", subchapter.Title, @".tex"), sub);
                }
            }
        }
        #endregion

        #region Publish YAML
        public void PublishYaml() {
            // Remove nopub and null slides before publishing. Set the other metadata to null... eventaully make this configurable to the level of detail. 
            Chapters.ForEach(c => {
                c.HGuids                = null;
                c.Subchapters.ForEach(sub => {
                    sub.Slides.ForEach(s => {
                        s.Type          = null;
                        s.Chapter       = null;
                        s.Subchapter    = null;
                        s.HGuids        = null;
                    });
                });
            });

            // Build the serializer and create the YAML from the outline
            ISerializer serializer = new SerializerBuilder().Build();
            string yaml = serializer.Serialize(this);

            // Write the YAML to the proper location as indicated by A3Environment.A3_PUBLISH
            File.WriteAllText(string.Concat(A3Environment.A3_PUBLISH, @"\yaml.yml"), yaml);
        }
        #endregion
    }
}
