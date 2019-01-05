using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Alta3_PPA
{
    public class A3Log
    {
        public enum Level
        {
            Info,
            Warn,
            Error
        }
        public enum Operations
        {
            GenerateFromYaml,
            NewBaseline,
            FillSubChapters,
            FixMetadata,
            PrePublish,
            Publish,
            ToOutline,
            FillSubchapters
        }

        public string Timestamp { get; set; }
        public Operations Operation { get; set; }
        public string Path { get; set; }
        public List<A3LogEntry> Entries { get; set; }

        public A3Log(Operations operation)
        {
            Timestamp = DateTimeOffset.Now.DateTime.ToString().Replace('/', '.').Replace(':', '.').Replace(' ', '-');
            Operation = operation;
            Path = string.Concat(A3Environment.A3_LOG, @"\", Operation.ToString("g"), "--", Timestamp, ".log.txt");
            Entries = new List<A3LogEntry>();
        }

        public bool HasType(Level type)
        {
            return Entries.Any(entry => entry.Type == type);
        }

        public void Write(Level type, string message)
        {
            A3LogEntry entry = new A3LogEntry(type, message);
            Entries.Add(entry);
            File.AppendAllText(Path, entry.Entry.ToString());
        }
    }

    public class A3LogEntry
    {
        public string Timestamp { get; set; }
        public A3Log.Level Type { get; set; }
        public string Message { get; set; }
        public string Entry { get; set; }

        public A3LogEntry(A3Log.Level type, string message)
        {
            Timestamp = DateTimeOffset.Now.ToUnixTimeSeconds().ToString();
            Type = type;
            Message = message;
            Entry = String.Concat(Timestamp, " ", Type.ToString("g"), " ", Message, Environment.NewLine);
        }
    }
}
