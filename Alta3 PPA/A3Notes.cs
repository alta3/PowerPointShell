using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Alta3_PPA
{
    class A3Notes
    {
        public static string ToMarkdown(string notes, string activeGuid)
        {
            Encoding utf8 = Encoding.UTF8;
            Encoding ascii = Encoding.ASCII;

            string asciiMarkdown = "";

            if (notes != null)
            {
                asciiMarkdown = ascii.GetString(Encoding.Convert(utf8, ascii, utf8.GetBytes(notes)));
            }

            File.WriteAllText(string.Concat(A3Environment.A3_MARKDOWN, @"\", activeGuid, @".md"), asciiMarkdown);

            return asciiMarkdown;
        }
        public static List<string> ToLatex(A3Outline outline, string path)
        {
            if (!Directory.EnumerateFiles(A3Environment.A3_MARKDOWN).Any())
            {
                A3Publish.PublishMarkdown(outline);
            }

            ProcessStartInfo pandoc = new ProcessStartInfo()
            {
                CreateNoWindow = false,
                UseShellExecute = true,
                FileName = "pandoc.exe",
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = string.Concat(@"-f html -t latex -o ", "\"", A3Environment.A3_LATEX, @"\", "out.tex\" \"", path)
            };
            try
            {
                using (Process process = Process.Start(pandoc))
                {
                    process.WaitForExit();
                }
            }
            catch
            {

            }

            string[] latex = File.ReadAllLines(string.Concat(A3Environment.A3_LATEX, @"\out.tex"));
            File.Delete(string.Concat(A3Environment.A3_LATEX, @"\out.tex"));
            List<string> newtex = latex.ToList();

            return newtex;
        }
    }
}
