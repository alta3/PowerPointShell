using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Alta3_PPA
{
    class A3Yaml
    {
        public string Path;
        public List<string> Lines;
        public string Text;

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
            { Alerts.YamlIncomingKeyMapWarn, "A potential Key mapping issue is located on line: {}. No valid YAML key map was found on this line. This could be a block line in which case there will be no impact, but if the YAML file fails to be deserialized or the generated PowerPoint is not as you expected, please ensure this line does not require a valid yaml key."},
            { Alerts.YamlGenSuccess, "The PowerPoint has been generated and saved to the following location:\r\n {}"}
        };

        public static Dictionary<string, string>  KeyMappings = new Dictionary<string, string>
            {
                { "name:", "course:" },
                { "chapters:", "chapters:" },
                { "questions:", "questions:" },
                { "type:", "type:" },
                { "title:", "title:" },
                { "index:", "index:" },
                { "guid:", "guid:" },
                { "historicguids:", "historicguids:" },
                { "filename:", "filename:" },
                { "notes:", "notes:" },
                { "subchapters:", "subchapters:" },
                { "has-labs:", "haslabs:" },
                { "has-slides:", "hasslides:" },
                { "has-videos:", "hasvideos:" },
                { "weburl:", "weburl:" },
                { "vocab:", "vocab:" },
                { "slides:", "slides:"}
            };

        public A3Yaml(string path)
        {
            Path = path;
            Lines = new List<string>(File.ReadAllLines(path));
            Text = string.Join("\r\n", Lines);
        }
        public void UpdateTextFromLines()
        {
            Text = string.Join("\r\n", Lines);
        }
        public void UpdateLinesFromText()
        {
            Lines = new List<string>(Text.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None));
        }

        public void Lint(A3Log log)
        {
            ConvertIncomingYamlKeys(log);
            string tempFile = System.IO.Path.GetTempFileName();
            File.WriteAllText(tempFile, Text);
            A3Yaml a3Yaml = new A3Yaml(tempFile);

            bool returnedError = a3Yaml.RunSyntaxLinter(log);

            if (returnedError)
            {
                Process.Start(log.Path);
                Process.Start(Path);

                // check to see if the user wants to retry
                DialogResult dialogResult = MessageBox.Show(AlertDescriptions[Alerts.SyntaxError], "YAML SYNTAX ERROR!", MessageBoxButtons.RetryCancel);
                if (dialogResult == DialogResult.Retry)
                {
                    Lint(log);
                }
                A3Environment.QUIT_FROM_CURRENT_LOOP = true;
                return;
            }
            return;
        }
        private void ConvertIncomingYamlKeys(A3Log log)
        {
            List<string> convertedLines = new List<string>() { };
            for (int i = 0; i < Lines.Count; i++)
            {
                string line = Lines[i];
                string word = line.Trim().TrimStart('-').Split(' ')[0];
                if (KeyMappings.TryGetValue(word.ToLower(), out string replace))
                {
                    convertedLines.Add(ReplaceFirstOccurance(line, word, replace));
                }
                else if (line.Split(':').Length > 1)
                {
                    log.Write(A3Log.Level.Warn, AlertDescriptions[Alerts.YamlIncomingKeyMapWarn].Replace("{}", i.ToString()));
                }
                convertedLines.Add(line);
            }
            Lines = convertedLines;
            UpdateTextFromLines();
        }
        private static string ReplaceFirstOccurance(string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            text = pos < 0 ? text : String.Concat(text.Substring(0, pos), replace, text.Substring(pos + search.Length));
            return text;
        }
        private bool RunSyntaxLinter(A3Log log)
        {
            Process process = new Process();

            process.StartInfo.FileName = "yamllint";
            process.StartInfo.Arguments = String.Concat("-c \"", A3Environment.YAML_LINT_CONFIG, "\" -f parsable \"", Path.Trim().Replace("\"", ""), "\"");
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

        public A3Outline Deserialize(A3Log log)
        {
            IDeserializer deserializer = new DeserializerBuilder().WithNamingConvention(new CamelCaseNamingConvention()).Build();
            A3Outline outline = new A3Outline();
            try
            {
                outline = deserializer.Deserialize<A3Outline>(Text);
                return outline;
            }
            catch (Exception ex)
            {
                log.Write(A3Log.Level.Error, ex.Message);
                Process.Start(log.Path);
                Process.Start(Path);
                DialogResult dialogResult = MessageBox.Show(AlertDescriptions[Alerts.DeserializationError].Replace("{}", ex.Message), "DESERIALIZATION ERROR!", MessageBoxButtons.RetryCancel);
                if (dialogResult == DialogResult.Retry)
                {
                    A3Yaml a3Yaml = new A3Yaml(Path);
                    a3Yaml.Lint(log);
                    return Deserialize(log);
                }
                A3Environment.QUIT_FROM_CURRENT_LOOP = true;
                return outline;
            }
        }

        // move to outline
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
            ISerializer serializer = new SerializerBuilder().Build();
            string yaml = serializer.Serialize(outline);

            // Write the YAML to the proper location as indicated by A3Environment.A3_PUBLISH
            File.WriteAllText(String.Concat(A3Environment.A3_PUBLISH, @"\yaml.yml"), yaml);
        }

        public static string ConvertOutgoingYamlKeys(string yaml)
        {
            List<string> lines = new List<string>(yaml.Split(new string[] { Environment.NewLine }, StringSplitOptions.None));
            List<string> convertedLines = new List<string>() { };
            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];
                string word = line.Trim().TrimStart('-').Split(' ')[0];
                if (KeyMappings.ContainsValue(word.ToLower()))
                {
                    string replace = KeyMappings.Values.Where(v => v == word.ToLower()).First();
                    convertedLines.Add(ReplaceFirstOccurance(line, word, replace));
                }
                convertedLines.Add(line);
            }
            return String.Join("\r\n", convertedLines);
        }
    }
}
