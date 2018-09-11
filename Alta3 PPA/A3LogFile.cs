using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Alta3_PPA {
    public class A3LogFile {
        public string Path { get; set; }
        public string Timestamp { get; set; }
        public List<A3LogEntry> Entries { get; set; }

        public A3LogFile() {
            this.Timestamp = DateTimeOffset.Now.DateTime.ToString().Replace('/','.').Replace(':','.').Replace(' ','-');
            this.Path = string.Concat(A3Globals.A3_LOG, @"\", this.Timestamp, ".log.txt");
            this.Entries = new List<A3LogEntry>();
        }

        public A3LogFile(string path) {
            if (File.Exists(path)) {
                this.Timestamp = path.Remove(0, A3Globals.A3_LOG.Length).Split('.')[0];
                this.Path = path;
                this.Entries = new List<A3LogEntry>();
                string[] entries = File.ReadAllLines(path);
                foreach (string entry in entries) {
                    this.Entries.Add(new A3LogEntry(entry));
                }
            }
            else {
                MessageBox.Show("The requested log does not exist", "Log File Not Found", MessageBoxButtons.OK);
            }
        }

        public bool HasError() {
            foreach (A3LogEntry a3Entry in this.Entries) {
                if (a3Entry.Type == "ERROR") {
                    return true;
                }
            }
            return false;
        }

        public void WriteEntry(A3LogEntry entry) {
            File.AppendAllText(this.Path, entry.Entry);
        }
        public void WriteError(string message) {
            this.Entries.Add(new A3LogEntry("ERROR", message));
            File.AppendAllText(this.Path, String.Concat(this.Entries.Last().Entry.ToString(), Environment.NewLine));
        }
        public void WriteWarn(string message) {
            this.Entries.Add(new A3LogEntry("WARN", message));
            File.AppendAllText(this.Path, String.Concat(this.Entries.Last().Entry.ToString(), Environment.NewLine));
        }
        public void WriteInfo(string message) {
            this.Entries.Add(new A3LogEntry("INFO", message));
            File.AppendAllText(this.Path, String.Concat(this.Entries.Last().Entry.ToString(), Environment.NewLine));
        }
    }
}
