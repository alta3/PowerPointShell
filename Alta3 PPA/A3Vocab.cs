using System;

using System.IO;


namespace Alta3_PPA
{
    public class A3Vocab
    {
        public string Word { get; set; }
        public string Def { get; set; }

        public void Generate(int chapter)
        {
            string line = string.Concat(Word, " [", chapter.ToString(), "] ", Def, Environment.NewLine);
            string path = string.Concat(A3Environment.A3_WORKING, @"\Vocab.txt");
            File.AppendAllText(path, line);
        }
    }
}
