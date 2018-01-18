using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointShell
{
    public class Vocab
    {
        public string Word { get; set; }
        public string Def { get; set; }

        public void Generate(int chapterCount)
        {
            string line = this.Word + " [" + chapterCount.ToString() + "] " + this.Def + "\r\n";
            string path = GlobalVars.WORKING_PATH + "\\Vocab.txt";
            File.AppendAllText(path, line);
        }
    }
}
