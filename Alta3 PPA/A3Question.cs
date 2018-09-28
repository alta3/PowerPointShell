using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Question
    {
        [YamlMember(Alias = "question", ApplyNamingConventions = false)]
        public string Text { get; set; }
        public string Correct { get; set; }
        public List<A3Incorrect> Incorrect { get; set; }
        public int Value { get; set; }

        private static Random random = new Random();

        public void Generate(string chapSub)
        {
            //TODO: CONVERT TO STRING BUILDER 
            string lines = null;
            int correct = random.Next(1, 5);
            string path = A3Globals.A3_PATH + "\\quiz.txt";

            lines += "id: " + this.ID(16) + "\r\n";
            lines += "chapsubchap: " + chapSub + "\r\n";
            lines += "MediaURL: \r\n";
            lines += "Points: " + this.Value + "\r\n";
            lines += "Question: " + this.Text + "\r\n";
            lines += this.GenerateOrder(correct);
            lines += "--------\r\n";
            File.AppendAllText(path, lines);
        }

        private string ID(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZacbdefghijklmnopqrstuvwxyz0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private string GenerateOrder(int correct)
        {
            int line = 1;
            int incorrect = 0;

            Dictionary<int, string> numDict = new Dictionary<int, string>();
            numDict.Add(1, "A");
            numDict.Add(2, "B");
            numDict.Add(3, "C");
            numDict.Add(4, "D");

            string choiceLines = null;
            string answerBoolLines = null;
            string answerExplanationLines = null;

            while (line <= 4)
            {
                //TODO: CONVERT TO STRING BUILDERS
                string choiceText = null;
                string answerBool = null;
                string answerExplanation = null;
                if (line == correct)
                {
                    choiceText = this.Correct;
                    answerBool = "1";
                    answerExplanation = "Correct!";
                }
                else
                {
                    choiceText = this.Incorrect[incorrect].Text;
                    answerBool = "0";
                    answerExplanation = this.Incorrect[incorrect].Explanation;
                    incorrect += 1;
                }
                choiceLines += "Choice" + numDict[line] + ": " + choiceText.Trim() + "\r\n";
                answerBoolLines += "Correct" + numDict[line] + ": " + answerBool.Trim() + "\r\n";
                answerExplanationLines += "Why" + numDict[line] + ": " + answerExplanation.Trim() + "\r\n";
                line += 1;
            }
            string lines = string.Concat(choiceLines, answerBoolLines, answerExplanationLines);
            return lines;
        }

        public class A3Incorrect
        {
            public string Text { get; set; }
            public string Explanation { get; set; }
        }
    }
}
