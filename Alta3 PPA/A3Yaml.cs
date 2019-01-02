using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    class A3Yaml
    {
        public enum Alerts
        {
            SyntaxError,
            DeserializationError,
            YamlIncomingKeyMapWarn,
            YamlGenSuccess
        }
        public static Dictionary<Alerts, string> AlertDescriptions = new Dictionary<Alerts, string>
        {
            { Alerts.SyntaxError, "The source YAML file contains at least one syntax error. Opening the log file reference and the source YAML file for editing. Once done editing the source file save it and click Retry to attempt the parsing action again. If you would like to exit without fixing the issue click Cancel."},
            { Alerts.DeserializationError, "The program failed to deserialize the YAML tree into an A3Outline object with the following error: \r\n {}.  Opening the log file reference and the source YAML file for editing. Once done editing the source file save it and click Retry to attempt the deserialization action again. If you would like to exit without fixing the issue click Cancel." },
            { Alerts.YamlIncomingKeyMapWarn, "No valid YAML key mapping found on line: {}.  If the YAML file fails to be deserialized or the generated PowerPoint is not as you expected, please ensure this line does not require a valid yaml key."},
            { Alerts.YamlGenSuccess, "The PowerPoint has been generated and saved to the following location:\r\n {}"}
        };

        public static void GenerateFromYaml(A3Log log, string yamlPath)
        {
            // Lint the YAML file before attempting to deserialize the outline and exit early if the user cancels the operation
            (bool lintProceed, string lintedText) = Lint(log, yamlPath);
            if (!lintProceed)
            {
                return;
            }

            // Create the outline from the YAML file and exit early if the user cancels the operation
            (bool deserializeProceed, A3Outline outline) = Deserialize(log, lintedText, yamlPath);
            if (!deserializeProceed)
            {
                return;
            }

            // Open a copy of the blank PowerPoint in the current PowerPoint context
            Presentation ppt = Globals.ThisAddIn.Application.Presentations.Open(A3Environment.MODEL_POWERPOINT, 0, 0, Microsoft.Office.Core.MsoTriState.msoTrue);

            // Save the presentation to a unqiue location
            string savePath = SavePresentationAs(outline.Course, ppt);

            // Generate the Presentation
            outline.GeneratePresentation(ppt);

            // Cleanup the initial slides
            for (int i = 0; i < 6; i++)
            {
                ppt.Slides[1].Delete();
            }

            // Save the generated presentation and handoff control back to the user
            ppt.Save();
            MessageBox.Show(AlertDescriptions[Alerts.YamlGenSuccess].Replace("{}", savePath), "POWERPOINT GENERATION COMPLETE!", MessageBoxButtons.OK);
        }
        public static void ProduceYaml(A3Log log, A3Outline _outline)
        {
            A3Outline outline = new A3Outline();
            outline = _outline;
            // Check for NO-PUB slides and remove them from the outline
            foreach (A3Chapter chapter in outline.Chapters)
            {
                chapter.Vocab = null;
                chapter.HistoricGuids = null;
                foreach (A3Subchapter subchapter in chapter.Subchapters)
                {
                    foreach (A3Content slide in subchapter.Slides)
                    {
                        slide.Type = null;
                        slide.Chapter = null;
                        slide.Subchapter = null;
                        slide.HistoricGuids = null;
                        if (slide.Type == "NO-PUB" || slide.Type == "BLANK")
                        {
                            subchapter.Slides.Remove(slide);
                        }
                    }
                }
            }

            // Build the serializer and create the YAML from the outline
            var serializer = new SerializerBuilder().Build();
            var yaml = serializer.Serialize(outline);

            // Write the YAML to the proper location as indicated by A3Environment.A3_PUBLISH
            File.WriteAllText(String.Concat(A3Environment.A3_PUBLISH, @"\yaml.yml"), yaml);
        }

        private static (bool, string) Lint(A3Log log, string yamlPath)
        {
            List<string> lines = new List<string>(File.ReadAllLines(yamlPath));

            string lintedText = ConvertIncomingYamlKeys(log, lines);
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, lintedText);

            bool returnedError = RunSyntaxLinter(log, tempFile);

            if (returnedError)
            {
                Process.Start(log.Path);
                Process.Start(yamlPath);

                // check to see if the user wants to retry
                DialogResult dialogResult = MessageBox.Show(AlertDescriptions[Alerts.SyntaxError], "YAML SYNTAX ERROR!", MessageBoxButtons.RetryCancel);
                if (dialogResult == DialogResult.Retry)
                {
                    bool proceed = true;
                    (proceed, lintedText) = Lint(log, yamlPath);
                    return (proceed, lintedText);
                }
                return (false, "");
            }
            return (true, lintedText);
        }
        private static string ConvertIncomingYamlKeys(A3Log log, List<string> lines)
        {
            List<string> convertedLines = new List<string>() { };
            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];
                string word = line.Trim().TrimStart('-').Split(' ')[0];
                if (A3Environment.A3_YAML_KEY_MAPPING.TryGetValue(word.ToLower(), out string replace))
                {
                    convertedLines.Add(ReplaceFirstOccurance(line, word, replace));
                }
                else
                {
                    convertedLines.Add(line);
                    log.Write(A3Log.Level.Warn, AlertDescriptions[Alerts.YamlIncomingKeyMapWarn].Replace("{}", i.ToString()));
                }
            }
            return String.Join("\r\n",convertedLines);
        }
        private static string ReplaceFirstOccurance(string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            text = pos < 0 ? text : String.Concat(text.Substring(0, pos), replace, text.Substring(pos + search.Length));
            return text;
        }
        private static bool RunSyntaxLinter(A3Log log, string path)
        {
            Process process = new Process();

            process.StartInfo.FileName = "yamllint";
            process.StartInfo.Arguments = String.Concat("-c \"", A3Environment.YAML_LINT_CONFIG, "\" -f parsable \"", path.Trim().Replace("\"", ""), "\"");
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;

            process.Start();
            string text = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            List<string> lines = new List<string>(text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None));
            foreach (string line in lines)
            {
                if (line.Contains("[error]"))
                {
                    log.Write(A3Log.Level.Error, line);
                    return true;
                }
                else if (line.Contains("[warning]"))
                {
                    log.Write(A3Log.Level.Warn, line);
                }
            }
            return false;
        }

        private static (bool, A3Outline) Deserialize(A3Log log, string lintedText, string yamlPath) 
        {
            IDeserializer deserializer = new DeserializerBuilder().WithNamingConvention(new CamelCaseNamingConvention()).Build();
            A3Outline outline = new A3Outline();
            try
            {
                outline = deserializer.Deserialize<A3Outline>(lintedText);
                return (true, outline);
            }
            catch (Exception ex)
            {
                log.Write(A3Log.Level.Error, ex.Message);
                Process.Start(log.Path);
                Process.Start(yamlPath);
                DialogResult dialogResult = MessageBox.Show(AlertDescriptions[Alerts.DeserializationError].Replace("{}", ex.Message), "DESERIALIZATION ERROR!", MessageBoxButtons.RetryCancel);
                if (dialogResult == DialogResult.Retry)
                {
                    bool proceed = true;
                    (proceed, outline) = Deserialize(log, lintedText, yamlPath);
                    return (proceed, outline);
                }
                return (false, outline);
            }
        }

        private static string SavePresentationAs(string course, Presentation ppt)
        {
            string saveDir = String.Concat(A3Environment.A3_WORKING, "\\", course);
            try { Directory.CreateDirectory(saveDir); } catch { }
            string savePath = String.Concat(saveDir, "\\", course);
            int version = 0;
            while (File.Exists(String.Concat(savePath, ".pptm")))
            {
                version++;
                savePath = string.Concat(saveDir, "\\", course, version.ToString());
            }
            ppt.SaveAs(String.Concat(savePath, ".pptm"));
            return savePath;
        }
    }
}
