using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Vocab
    {
        public string Word { get; set; }
        public string Def { get; set; }

        public void Generate(int chapterCount)
        {
            string line = String.Concat(this.Word, " [", chapterCount.ToString(), "] ", this.Def, "\r\n");
            string path = String.Concat(A3Globals.A3_WORKING, @"\Vocab.txt");
            File.AppendAllText(path, line);
        }
    }
}
