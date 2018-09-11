using System;

namespace Alta3_PPA {
    public class A3LogEntry {
        public string Timestamp { get; set; }
        public string Type { get; set; }
        public string Message { get; set; }
        public string Entry { get; set; }

        public A3LogEntry(string type, string message) {
            this.Timestamp = DateTimeOffset.Now.ToUnixTimeSeconds().ToString();
            this.Type = type;
            this.Message = message;
            this.Entry = String.Concat(this.Timestamp, " ", this.Type, " ", this.Message);
        }
        public A3LogEntry(string timestamp, string type, string message) {
            this.Timestamp = timestamp;
            this.Type = type;
            this.Message = message;
        }
        public A3LogEntry(string entry) {
            this.Entry = entry;
            string[] fields = entry.Split(new char[0], 3);
            this.Timestamp = fields[0];
            this.Type = fields[1];
            this.Message = fields[2];
        }
    }
}
