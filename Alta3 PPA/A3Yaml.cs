using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Alta3_PPA
{
    class A3Yaml
    {
        public static void GenerateFromYaml(A3LogFile logFile, string yamlPath)
        {
            // Lint the YAML file before attempting to deserialize the outline
            bool lintSuccess = false;
            string lintText = "";
            while (!lintSuccess)
            {
                Tuple<string, bool> linter = Lint(logFile, yamlPath);
                lintText = linter.Item1;
                lintSuccess = linter.Item2;
                if (!lintSuccess) {
                    DialogResult dialogResult = MessageBox.Show(String.Concat("The source YAML file contains at least one syntax error. Opening the log and the YAML file for reference and editing. Once done editing the source file save it and click Retry to attempt to parse again. If you would like to exit without fixing the issue click Cancel."), "YAML SYNTAX ERROR!", MessageBoxButtons.RetryCancel);
                    Process.Start(logFile.Path);
                    Process.Start(yamlPath);
                    if (dialogResult == DialogResult.Retry)
                    {
                        continue;
                    }
                    else if (dialogResult == DialogResult.Cancel)
                    {
                        return;
                    }
                    else
                    {
                        return;
                    }
                }
            }

            // Create the outline from the YAML file and exit early if unable to deserialize and point user to the the logs.
            IDeserializer deserializer = new DeserializerBuilder().WithNamingConvention(new CamelCaseNamingConvention()).Build();
            A3Outline outline = new A3Outline();
            try { outline = deserializer.Deserialize<A3Outline>(lintText); }
            catch (Exception ex) {
                logFile.WriteError(ex.Message);
                MessageBox.Show(String.Concat("The program failed to deserialize the YAML tree into an A3Outline object with the following error: \r\n", ex.Message, "\r\n Opening the logfile and the source YAML file for you."), "DESERIALIZATION FAILURE!", MessageBoxButtons.OK);
                Process.Start(logFile.Path);
                Process.Start(yamlPath);
                return; }

            // Open a copy of the blank PowerPoint in the current PowerPoint context
            Microsoft.Office.Interop.PowerPoint.Presentation ppt = Globals.ThisAddIn.Application.Presentations.Open(A3Globals.MODEL_POWERPOINT, 0, 0, Microsoft.Office.Core.MsoTriState.msoTrue);

            // Save the powerpoint presentation to the working directory so that changes do not affect the model presentation
            string saveDir = String.Concat(A3Globals.A3_WORKING, "\\", outline.Course);
            try { Directory.CreateDirectory(saveDir); } catch { }
            string savePath = String.Concat(saveDir, "\\", outline.Course);
            int version = 0;
            while (File.Exists(String.Concat(savePath, ".pptm")))
            {
                version += 1;
                savePath = string.Concat(saveDir, "\\", outline.Course, version.ToString());
            }
            ppt.SaveAs(String.Concat(savePath, ".pptm"));

            // Grab the current global infer states and then switch them to true while generating the powerpoint from yaml
            bool inferState = A3Globals.ALLOW_INFER_FROM_SLIDE;
            bool defaultInferState = A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE;
            A3Globals.ALLOW_INFER_FROM_SLIDE = true;
            A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE = true;

            // Generate the Presentation
            outline.GeneratePresentation(ppt);

            // Return to the original state of the infer global states 
            A3Globals.ALLOW_INFER_FROM_SLIDE = inferState;
            A3Globals.ALLOW_DEFAULT_INFER_FROM_SLIDE = defaultInferState;

            for (int i = 0; i < 6; i++)
            {
                ppt.Slides[1].Delete();
            }
            
            // Save the newly generated Presentation
            ppt.Save();

            // Alert the user the operation has concluded
            string message = String.Concat("The PowerPoint has been successfully built and saved to the following location:\r\n", savePath);
            MessageBox.Show(message, "Build Success", MessageBoxButtons.OK);
        }
        public static Tuple<string, bool> Lint(A3LogFile logFile, string yamlPath)
        {
            bool success = true;

            logFile.WriteInfo("Starting: YAML Lint Process.");

            logFile.WriteInfo("Starting: Read YAML file.");
            string[] lines = File.ReadAllLines(yamlPath);
            logFile.WriteInfo("Finished: Read YAML file.");

            logFile.WriteInfo("Starting: Fix YAML Keys");
            string lintText = FixYamlKeys(logFile, lines);
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, lintText);
            logFile.WriteInfo("Finished: Fix YAML Keys");

            logFile.WriteInfo("Starting: YAML Syntax Linter");
            string syntaxLinterOutput = RunSyntaxLinter(tempFile);
            string[] lintLines = syntaxLinterOutput.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            foreach (string line in lintLines)
            {
                if (line.Contains("[error]"))
                {
                    logFile.WriteError(line);
                    success = false;
                }
                else if (line.Contains("[warning]"))
                {
                    logFile.WriteWarn(line);
                }
                else
                {
                    logFile.WriteInfo(line);
                }
            }
            logFile.WriteInfo("Finished: YAML Syntax Linter");

            logFile.WriteInfo("Finished: YAML Lint Process.");
            return Tuple.Create(lintText, success);
        }
        public static void ProduceYaml(A3LogFile logFile, A3Outline _outline)
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

            // Write the YAML to the proper location as indicated by A3Globals.A3_PUBLISH
            File.WriteAllText(String.Concat(A3Globals.A3_PUBLISH, @"\yaml.yml"), yaml);
        }

        private static string FixYamlKeys(A3LogFile logFile, string[] lines)
        {
            string text = "";
            int lineCounter = 1; 
            foreach (string line in lines)
            { 
                string[] words = line.Trim().Split(' ');
                string word = "";
                string replace = "";
                if (words.Length > 0)
                {
                    word = words[0] == "-" ? words[1] : words[0];
                    if (A3Globals.A3_YAML_KEY_MAPPING.TryGetValue(word.ToLower(), out string value))
                    {
                        replace = value;
                    }
                    else
                    {
                        replace = word.ToLower();
                        logFile.WriteWarn(String.Concat("No valid yaml mapping found on line: ", lineCounter.ToString(), ". If the YAML file fails to parse properly, please ensure this line does not require a valid yaml key."));
                    } 
                }
                text = String.Concat(text, ReplaceFirst(line, word, replace));
                lineCounter++;
            }
            return text;
        }
        private static string ReplaceFirst(string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
        private static string RunSyntaxLinter(string path)
        {
            ProcessStartInfo start = new ProcessStartInfo
            {
                FileName = "yamllint",
                Arguments = String.Concat("-c \"", A3Globals.YAML_LINT_CONFIG, "\" -f parsable \"", path.Trim().Replace("\"",""), "\""),
                UseShellExecute = true,
                CreateNoWindow = true, 
                RedirectStandardOutput = true,
                RedirectStandardError = true 
            };
            using (Process process = Process.Start(start))
            {
                using (StreamReader reader = process.StandardOutput)
                {
                    string stderr = process.StandardError.ReadToEnd();
                    string result = reader.ReadToEnd();
                    result = String.Concat(result, stderr);
                    return result;
                }
            }
        }
    }
}
