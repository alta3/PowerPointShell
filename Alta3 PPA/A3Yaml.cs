using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Alta3_PPA
{
    class A3Yaml
    {
        public static void GenerateFromYaml(A3LogFile logFile, string yamlPath)
        {
            // Read the file into a string for processing
            string text = File.ReadAllText(yamlPath);

            // TODO: Normailze the text for now make it convert easy, but in the future change it to be consiste with everything else. 
            // Lint the YAML file before attempting to deserialize the outline
            // A3Yaml.Lint(logFile, text);

            // Log that we are about to try and desearilize this will help to see if our linting is effective or not
            logFile.WriteInfo("YAML lint complete. About to desearilize outline.");

            // Create the outline from the YAML file
            Deserializer deserializer = new DeserializerBuilder().WithNamingConvention(new CamelCaseNamingConvention()).Build();
            A3Outline outline = new A3Outline();
            try { outline = deserializer.Deserialize<A3Outline>(text); }
            catch (Exception ex) { logFile.WriteError(ex.Message); }

            // Open a copy of the blank PowerPoint in the current PowerPoint context
            Microsoft.Office.Interop.PowerPoint.Presentation ppt = Globals.ThisAddIn.Application.Presentations.Open(A3Globals.MODEL_POWERPOINT, 0, 0, Microsoft.Office.Core.MsoTriState.msoTrue);

            // Save the powerpoint presentation to the working directory so that changes do not affect the model presentation
            string saveDir = String.Concat(A3Globals.A3_WORKING, "\\", outline.Name);
            A3Presentation.SavePresentation(ppt, saveDir, outline.Name);

            // Infer Defaults
            A3Environment.DefaultInfer();


            // Generate the Presentation
            outline.GeneratePresentation(ppt);

            // Clean the environment
            A3Environment.Clean();

            for (int i = 0; i < 6; i++)
            {
                ppt.Slides[1].Delete();
            }
            
            // Save the newly generated Presentation
            ppt.Save();

            // Alert the user the operation has concluded
            string message = String.Concat("The PowerPoint has been successfully built and saved.");
            MessageBox.Show(message, "Build Success", MessageBoxButtons.OK);
        }
        public static void Lint(A3LogFile logFile, string yamlText)
        {
            List<string> yamlLines = new List<string>(Regex.Split(yamlText, Environment.NewLine));
            List<string> outlineKeys = new List<string> { "outline", "chapters", "labs", "meta"};
            List<string> titleKey = new List<string> { "- title"};
            List<string> chapterKeys = new List<string> { "subchapters" };
            List<string> subchapterKeys = new List<string> { "slides" };
            List<string> slideKeys = new List<string> { "notes" };

            // TODO: GET RID OF THIS ENTIRE SECTION BY MAKING IT REPORT CORRECTLY FROM THE GROUND UP THIS IS ONLY A TEMPORARY SOLUTION. 
            /*string newYamlText = null;
            foreach (string yamlLine in yamlLines)
            {
                string newYamlLine = null;
                List<string> keys = new List<string>();
                keys.Add("name");
                keys.AddRange(outlineKeys);
                keys.AddRange(titleKey);
                keys.AddRange(chapterKeys);
                keys.AddRange(subchapterKeys);
                keys.AddRange(slideKeys);
                foreach (string key in keys)
                {
                    if (yamlLine.Split(':')[0].ToLower() == key)
                    {
                        if (key == "name")
                        {
                            newYamlLine = String.Concat("course:", yamlLine.Split(':')[1]);
                        }
                        newYamlLine = String.Concat(yamlLine.Split(':')[0].ToLower(), ":", yamlLine.Split(':')[1]);
                    }
                }
                newYamlText = String.Concat(newYamlText, newYamlLine, Environment.NewLine);
            }

            List<string> newYamlLines = new List<string>(Regex.Split(newYamlText, Environment.NewLine));
            */

            #region Course
            List<string> courses = yamlLines.FindAll(s => s.Split(':')[0].ToLower() == "course");
            A3Yaml.LogNotPresent(logFile, courses, 2, "Course");
            try { A3Yaml.ErrorNullCheck(logFile, "course", courses[0], false); } catch { }
            A3Yaml.ErrorDuplicateMapping(logFile, yamlLines, courses);
            #endregion

            #region Chapters
            List<string> chapsMap = yamlLines.FindAll(s => s.Split(':')[0].ToLower() == "chapters");
            A3Yaml.LogNotPresent(logFile, chapsMap, 1, "Chapters");
            A3Yaml.ErrorNullCheck(logFile, "chapters", chapsMap[0], true);
            A3Yaml.ErrorDuplicateMapping(logFile, yamlLines, chapsMap);

            List<string> chapters = A3Yaml.GetValidBlock(logFile, yamlLines, chapsMap, 0);
            #endregion

            #region TODO:
            List<string> metas = yamlLines.FindAll(str => str.Split(':')[0].ToLower() == "meta");
            
            List<string> labs = yamlLines.FindAll(str => str.Split(':')[0].ToLower() == "labs");
            #endregion
            // return newYamlText;
        }
        public static void ProduceYaml(A3LogFile logFile, A3Outline _outline)
        {
            A3Outline outline = new A3Outline();
            outline = _outline;
            // Check for NO-PUB slides and remove them from the outline
            foreach (A3Chapter chapter in outline.Chapters)
            {
                chapter.Day = null;
                chapter.Vocab = null;
                chapter.HistoricGuids = null;
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    foreach (A3Content slide in subchapter.Slides)
                    {
                        slide.Day = null;
                        slide.Type = null;
                        slide.Chapter = null;
                        slide.Subchapter = null;
                        slide.HistoricGuids = null;                       
                        if (slide.Type == "NOPUB" || slide.Type == "BLANK")
                        {
                            subchapter.Slides.Remove(slide);
                        }
                    }
                }
            }

            // Build the serializer and create the YAML from the outline
            var serializer = new SerializerBuilder().Build();
            var yaml = serializer.Serialize(outline);

            // Write the YAML to the proper location as indicated by A3Globals.A3_PUBLISH
            File.WriteAllText(String.Concat(A3Globals.A3_PUBLISH, @"\yaml.yml"), yaml);
        }

        private static List<string> GetValidBlock(A3LogFile logFile, List<string> yamlLines, List<string> map, int indent)
        {
            List<string> list = new List<string>();

            int lastIndex = 0;
            foreach (string mapping in map)
            {
                if (lastIndex == 0)
                {
                    int startIndex = yamlLines.FindIndex(s => s == mapping);
                    int endIndex = yamlLines.FindIndex(startIndex + 1, s => char.IsWhiteSpace(s, indent));
                    lastIndex = endIndex;
                    for (int i = startIndex; i < endIndex; i++)
                    {
                        list.Add(yamlLines[i]);
                    }
                }
                else
                {
                    int startIndex = yamlLines.FindIndex(s => s == mapping);
                    int endIndex = yamlLines.FindIndex(startIndex + 1, s => char.IsWhiteSpace(s, indent));
                    for (int i = startIndex; i < endIndex; i++)
                    {
                        logFile.WriteError("This section of YAML is part of a duplicate block - please check the scope of each block and ensure only one mapping exists.");
                    }
                    lastIndex = endIndex;
                }
                return list;
            }

            return list;
        }

        private static void ErrorDuplicateBlocks(A3LogFile logFile, List<string> yamlLines, int startIndex, int indent)
        {
            int endIndex = yamlLines.FindIndex(startIndex, s => !char.IsWhiteSpace(s, indent));
            for (int i = startIndex; i < endIndex; i++)
            {
                logFile.WriteError(String.Concat("Duplicates problem"));
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
                logFile.WriteError(String.Concat("Null Key Problem"));
            }
        }
        private static void ErrorUnkownKey(A3LogFile logFile, List<string> keys, List<string> lines, int indent)
        {
            foreach (string line in lines)
            {
                if (keys.Contains(line.Split(':')[0].Trim().ToLower()))
                {
                    logFile.WriteError(String.Concat("Unkown Key Problem"));
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
                    logFile.WriteError(String.Concat("Indentation Problem"));
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
                        logFile.WriteInfo(string.Concat("Not Present"));
                        break;
                    case 1:
                        logFile.WriteWarn(string.Concat("Not Present"));
                        break;
                    case 2:
                        logFile.WriteError(string.Concat("Not Present"));
                        break;
                    default:
                        logFile.WriteError(string.Concat("Not Present"));
                        break;
                }
            }
        }
    }
}
