using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;

namespace Alta3_PPA
{
    class A3Environment
    {
        // Ensure environment is ready for next run
        public static void Clean()
        {
            A3Globals.QUIT_FROM_CURRENT_LOOP = false;
        }

        public static void StartUp()
        {
            // Create the A3 directory structure if it does not exist. 
            try { Directory.CreateDirectory(A3Globals.A3_PATH); } catch { }
            try { Directory.CreateDirectory(A3Globals.A3_WORKING); } catch { }
            try { Directory.CreateDirectory(A3Globals.A3_PUBLISH); } catch { }
            try { Directory.CreateDirectory(A3Globals.A3_LOG); } catch { }
            try { Directory.CreateDirectory(A3Globals.A3_PRES_PNGS); } catch { }
            try { Directory.CreateDirectory(A3Globals.A3_BOOK_PNGS); } catch { }
            try { Directory.CreateDirectory(A3Globals.A3_MARKDOWN); } catch { }
            try { Directory.CreateDirectory(A3Globals.A3_LATEX); } catch { }
            if (!A3Environment.IsApplictionInstalled("pandoc") || !A3Environment.IsApplictionInstalled("texstudio") || !A3Environment.IsApplictionInstalled("miktex"))
            {
                // RUN THE UPDATE CHOCO SCRIPT
            }
        }
        private static bool IsApplictionInstalled(string p_name)
        {
            string displayName;
            RegistryKey key;

            // search in: CurrentUser
            key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            foreach (String keyName in key.GetSubKeyNames())
            {
                RegistryKey subkey = key.OpenSubKey(keyName);
                displayName = subkey.GetValue("DisplayName") as string;
                if (p_name.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return true;
                }
            }

            // search in: LocalMachine_32
            key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            foreach (String keyName in key.GetSubKeyNames())
            {
                RegistryKey subkey = key.OpenSubKey(keyName);
                displayName = subkey.GetValue("DisplayName") as string;
                if (p_name.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return true;
                }
            }

            // search in: LocalMachine_64
            key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall");
            foreach (String keyName in key.GetSubKeyNames())
            {
                RegistryKey subkey = key.OpenSubKey(keyName);
                displayName = subkey.GetValue("DisplayName") as string;
                if (p_name.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return true;
                }
            }

            // NOT FOUND
            return false;
        }
    }
}
