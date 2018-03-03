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

            string asciiMarkdown = ascii.GetString(Encoding.Convert(utf8, ascii, utf8.GetBytes(notes)));

            File.WriteAllText(String.Concat(A3Globals.A3_MARKDOWN, @"\", activeGuid, @".md"), asciiMarkdown);

            return asciiMarkdown;
        }
        public static string ToLatex(A3Outline outline, string activeGuid)
        {
            if (!Directory.EnumerateFiles(A3Globals.A3_MARKDOWN).Any())
            {
                A3Publish.PublishMarkdown(outline);
            }

            ProcessStartInfo build = new ProcessStartInfo(String.Concat(@"powershell.exe -file ", A3Globals.A3_RESOURCE, @"\pandoc.ps1 -markdownPath ", A3Globals.A3_MARKDOWN, @" -latexPath ", A3Globals.A3_LATEX, @" -activeGuid ", activeGuid, @".md"))
            {
                UseShellExecute = true,
            };
            Process.Start(build);

            string latex = File.ReadAllText(String.Concat(A3Globals.A3_LATEX, @"\out.tex"));

            File.Delete(String.Concat(A3Globals.A3_LATEX, @"\out.txt"));

            return latex;
        }
    }
}
