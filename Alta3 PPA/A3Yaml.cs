using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Alta3_PPA
{
    class A3Yaml
    {
        public static void Lint(A3LogFile logFile, string yamlText)
        {
            List<string> yamlLines = new List<string>(Regex.Split(yamlText, Environment.NewLine));
            List<string> outlineKeys = new List<string> { "outline", "chapters", "labs", "meta"};
            List<string> chapterKeys = new List<string> { "subchapters" };
            List<string> subchapterKeys = new List<string> { "slides" };
            List<string> slideKeys = new List<string> { "notes" };

            #region Course
            List<string> courses = yamlLines.FindAll(s => s.Split(':')[0].ToLower() == "course");
            A3Yaml.LogNotPresent(logFile, courses, 2, "Course");
            A3Yaml.ErrorNullCheck(logFile, "course", courses[0], false);
            A3Yaml.ErrorDuplicateMapping(logFile, yamlLines, courses);
            #endregion

            #region Chapters
            List<string> chapsMap = yamlLines.FindAll(s => s.Split(':')[0].ToLower() == "chapters");
            A3Yaml.LogNotPresent(logFile, chapterKeys, 1, "Chapters");
            A3Yaml.ErrorNullCheck(logFile, "chapters", chapsMap[0], true);
            A3Yaml.ErrorDuplicateMapping(logFile, yamlLines, chapsMap);

            List<string> chapters = A3Yaml.GetValidBlock(logFile, yamlLines, chapsMap, true, 0);
            #endregion

            #region TODO:
            List<string> metas = yamlLines.FindAll(str => str.Split(':')[0].ToLower() == "meta");
            
            List<string> labs = yamlLines.FindAll(str => str.Split(':')[0].ToLower() == "labs");
            #endregion
        }

        private static List<string> GetValidBlock(A3LogFile logFile, List<string> yamlLines, List<string> map, bool unique, int indent)
        {
            List<string> list = new List<string>();

            int lastIndex = 0;
            foreach (string mapping in map)
            {
                if (lastIndex == 0)
                {
                    int startIndex = yamlLines.FindIndex(s => s == mapping);
                    int endIndex = yamlLines.FindIndex(startIndex + 1, s => char.IsWhiteSpace(s, 0));
                }
                else
                {

                }
            }

            return list;
        }
        private static void ErrorDuplicateBlocks(A3LogFile logFile, List<string> yamlLines, int startIndex, int indent)
        {
            int endIndex = yamlLines.FindIndex(startIndex, s => !char.IsWhiteSpace(s, indent));
            for (int i = startIndex; i < endIndex; i++)
            {
                logFile.WriteError(String.Concat(""));
            }
        }
        private static void ErrorDuplicateMapping(A3LogFile logFile, List<string> yamlLines, List<string> list)
        {
            var duplicates = list.GroupBy(c => c).Where(g => g.Count() > 1).Select(g => g.Key);
            foreach (var duplicate in duplicates)
            {
                List<string> duplicated = yamlLines.FindAll(s => s == duplicate);
                int lineIndex = yamlLines.FindIndex(s => s == duplicate);
                for (int i = 1; i < duplicated.Count; i++)
                {
                    lineIndex = yamlLines.FindIndex(lineIndex, s => s == duplicate);
                    logFile.WriteError(String.Concat("Dupicate, ", "\"", duplicate, "\" found at line number: ", lineIndex));
                    lineIndex++;
                }
            }
        }
        private static void ErrorNullCheck(A3LogFile logFile, string key, string line, bool beNull)
        {
            bool isNull;
            try { isNull = line.Split(':')[1].Trim().Length > 0 ? false : true; } catch { isNull = true; }
            if (beNull != isNull)
            {
                logFile.WriteError(String.Concat(""));
            }
        }
        private static void ErrorUnkownKey(A3LogFile logFile, List<string> keys, List<string> lines, int indent)
        {
            foreach (string line in lines)
            {
                if (keys.Contains(line.Split(':')[0].Trim().ToLower()))
                {
                    logFile.WriteError(String.Concat(""));
                }
            }
        }
        private static void ErrorImproperIndentation(A3LogFile logFile, List<string> yamlLines)
        {
            foreach (string line in yamlLines)
            {
                int i = 0;
                while (char.IsWhiteSpace(line, i) || i < 15)
                {
                    i++;
                }
                if (i % 2 != 0 || i > 14)
                {
                    logFile.WriteError(String.Concat(""));
                }
            }
        }
        private static void LogNotPresent(A3LogFile logFile, List<string> list, int logType, string item)
        {
            if (list.Count < 1)
            {
                switch (logType)
                {
                    case 0:
                        logFile.WriteInfo(string.Concat(""));
                        break;
                    case 1:
                        logFile.WriteWarn(string.Concat(""));
                        break;
                    case 2:
                        logFile.WriteError(string.Concat(""));
                        break;
                    default:
                        logFile.WriteError(string.Concat(""));
                        break;
                }
            }
        }
    }
}
