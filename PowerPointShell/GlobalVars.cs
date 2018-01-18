using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointShell
{
    public static class GlobalVars
    {
        public static string MODEL_POWERPOINT = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\documents\\Alta3 PowerPoints\\resource\\mod.pptm";
        public static string BLANK_POWERPOINT = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\documents\\Alta3 PowerPoints\\resource\\blank.pptm";
        public static string WORKING_PATH = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\documents\\Alta3 PowerPoints\\working";
    }
}
