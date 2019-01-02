using System;

using System.IO;


namespace Alta3_PPA
{
    public class A3Vocab
    {
        public string Word { get; set; }
        public string Def { get; set; }

        public void Generate(int chapterCount)
        {
            string line = String.Concat(Word, " [", chapterCount.ToString(), "] ", Def, Environment.NewLine);
            string path = String.Concat(A3Environment.A3_WORKING, @"\Vocab.txt");
            File.AppendAllText(path, line);
        }
    }
}
