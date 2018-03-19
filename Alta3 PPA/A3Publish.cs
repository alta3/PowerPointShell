using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    class A3Publish
    {
        public static void PublishPowerPoint()
        {

        }
        public static void PublishPresentationPNGs(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                string guid = slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text;
                string path = String.Concat(A3Globals.A3_PRES_PNGS, "\\", guid, ".png");
                slide.Export(path, "png", 1920, 1080);
            }
        }
        public static void PublishBookPNGs(PowerPoint.Presentation presentation)
        {
            Parallel.ForEach(Directory.EnumerateFiles(A3Globals.A3_PRES_PNGS), picture =>
            {
                //read image
                Bitmap bmp = new Bitmap(picture);
                
                //get image dimension
                int width = bmp.Width;
                int height = bmp.Height;

                //color of pixel
                Color p;

                //grayscale
                for (int y = 0; y < height; y++)
                {
                    for (int x = 0; x < width; x++)
                    {
                        //get pixel value
                        p = bmp.GetPixel(x, y);
                        p = Color.FromArgb(255, (255 - p.R), (255 - p.G), (255 - p.B));

                        //extract pixel component ARGB
                        int a = p.A;
                        int r = p.R;
                        int g = p.G;
                        int b = p.B;

                        //find average
                        int avg = (r + g + b) / 3;

                        //set new pixel value
                        bmp.SetPixel(x, y, Color.FromArgb(a, avg, avg, avg));
                    }
                }
                bmp.Save(picture.Replace("pres_pngs", "book_pngs"));
            });
        }
        public static void PublishMarkdown(A3Outline outline)
        {
            foreach (A3Chapter chapter in outline.Chapters)
            {
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    foreach (A3Content content in subchapter.Slides)
                    {
                        A3Notes.ToMarkdown(content.Notes, content.ActiveGuid);
                    }
                }
            }
        }
        public static void PublishLaTex(PowerPoint.Presentation presentation, A3Outline outline)
        {
            if (!Directory.EnumerateFiles(A3Globals.A3_MARKDOWN).Any())
            {
                A3Publish.PublishMarkdown(outline);
            }
            outline.GenerateLaTex(presentation, outline);
        }
        public static void PublishPDF(PowerPoint.Presentation presentation, A3Outline outline)
        {
            if (!Directory.EnumerateFiles(A3Globals.A3_BOOK_PNGS).Any())
            {
                A3Publish.PublishBookPNGs(presentation);
            }
            if (!Directory.EnumerateFiles(A3Globals.A3_LATEX).Any())
            {
                A3Publish.PublishLaTex(presentation, outline);
            }

            ProcessStartInfo build = new ProcessStartInfo()
            {
                UseShellExecute = true,
                CreateNoWindow = false,
                FileName = "pdflatex.exe",
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = String.Concat(@"-job-name=", outline.Course, @" -output-directory=", A3Globals.A3_PUBLISH, @" -aux-directory=", A3Globals.A3_LATEX, @"main.tex")
            };
            try
            {
                using (Process process = Process.Start(build))
                {
                    process.WaitForExit();
                }
            }
            catch 
            {
            
            }
        }
        public static void PublishQuestions()
        { }
        public static void PublishVocabulary()
        { }
        public static void PublishYAML(A3LogFile logFile, A3Outline outline)
        {
            A3Yaml.ProduceYaml(logFile, outline);
        }
    }
}
