using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    class A3Globals
    {
        // Global bools to help with loops and environment settings
        public static bool QUIT_FROM_CURRENT_LOOP = false;
        public static bool DEV_INITIALIZED = false;
        public static bool SHOW_ACTIVE_GUID = false;

        // The Alta3 directory structure variables
        public static string A3_PATH = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"\Documents\Alta3 PowerPoints");
        public static string A3_WORKING = String.Concat(A3_PATH, @"\working");
        public static string A3_PUBLISH = String.Concat(A3_PATH, @"\publish");
        public static string A3_PNGS = String.Concat(A3_PUBLISH, @"\pngs");
        public static string A3_LOG = String.Concat(A3_PATH, @"\log");

        // Alta3 Resoures location
        public static string A3_RESOURCE = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"\Alta 3\A3PPA\resources");

        // Global resoures including the model powerpoint and the model vba
        public static string MODEL_POWERPOINT = String.Concat(A3_RESOURCE, @"\mod.pptm");
        public static string BLANK_POWERPOINT = String.Concat(A3_RESOURCE, @"\blank.pptm");
        public static string CHAPTER_VBA = File.ReadAllText(String.Concat(A3_RESOURCE, @"\chapter_vba.txt"));
        public static string QUESTION_VBA = File.ReadAllText(String.Concat(A3_RESOURCE, @"\question_vba.txt"));

        // References to active/current presentation and slide
        public static A3Presentation A3PRESENTATION;
        public static A3Slide A3SLIDE;
        public static List<string> SHAPE_NAMES = new List<string>();
    }
}
