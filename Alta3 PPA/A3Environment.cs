using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

namespace Alta3_PPA
{
    class A3Environment
    {
        // Global bools to help with loops and environment settings
        public static bool QUIT_FROM_CURRENT_LOOP = false;
        public static bool AFTER_CHAPTER = false;
        public static bool BEFORE_QUESTION = true;
        public static bool SHOW_GUID = false;
        public static bool ALLOW_INFER_FROM_SLIDE = false;
        public static bool ALLOW_DEFAULT_INFER_FROM_SLIDE = false;
        public static bool ENFORCE_CHAP_SUB_SPLITTING = true;

        // The Alta3 directory structure variables
        public static string A3_PATH = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"\Documents\Alta3 PowerPoints");
        public static string A3_WORKING = String.Concat(A3_PATH, @"\working");
        public static string A3_PUBLISH = String.Concat(A3_PATH, @"\publish");
        public static string A3_PRES_PNGS = String.Concat(A3_PUBLISH, @"\pres_pngs");
        public static string A3_BOOK_PNGS = String.Concat(A3_PUBLISH, @"\book_pngs");
        public static string A3_LATEX = String.Concat(A3_PUBLISH, @"\latex");
        public static string A3_MARKDOWN = String.Concat(A3_PUBLISH, @"\markdown");

        public static string A3_LOG = String.Concat(A3_PATH, @"\log");

        // Alta3 Resoures location
        public static string A3_RESOURCE = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"\Alta3\A3PPA\resources");

        // Global resoures including the model powerpoint and the model vba
        public static string MODEL_POWERPOINT = String.Concat(A3_RESOURCE, @"\mod.pptm");
        public static string BLANK_POWERPOINT = String.Concat(A3_RESOURCE, @"\blank.pptm");
        public static string CHAPTER_VBA = File.ReadAllText(String.Concat(A3_RESOURCE, @"\chapter_vba.txt"));
        public static string QUESTION_VBA = File.ReadAllText(String.Concat(A3_RESOURCE, @"\question_vba.txt"));
        public static string YAML_LINT_CONFIG = String.Concat(A3_RESOURCE, @"\yamllint_config.yml");

        // References to active/current presentation and slide
        public static A3Slide A3SLIDE;

        // Ensure environment is ready for next run
        public static void Clean()
        {
            // Global bools to help with loops and environment settings
            QUIT_FROM_CURRENT_LOOP = false;
            AFTER_CHAPTER = false;
            BEFORE_QUESTION = true;
            SHOW_GUID = false;
            ALLOW_INFER_FROM_SLIDE = false;
            ALLOW_DEFAULT_INFER_FROM_SLIDE = false;
            ENFORCE_CHAP_SUB_SPLITTING = true;

            // The Alta3 directory structure variables
            A3_PATH = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"\Documents\Alta3 PowerPoints");
            A3_WORKING = String.Concat(A3_PATH, @"\working");
            A3_PUBLISH = String.Concat(A3_PATH, @"\publish");
            A3_PRES_PNGS = String.Concat(A3_PUBLISH, @"\pres_pngs");
            A3_BOOK_PNGS = String.Concat(A3_PUBLISH, @"\book_pngs");
            A3_LATEX = String.Concat(A3_PUBLISH, @"\latex");
            A3_MARKDOWN = String.Concat(A3_PUBLISH, @"\markdown");

            A3_LOG = String.Concat(A3_PATH, @"\log");

            // Alta3 Resoures location
            A3_RESOURCE = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"\Alta3\A3PPA\resources");

            // Global resoures including the model powerpoint and the model vba
            MODEL_POWERPOINT = String.Concat(A3_RESOURCE, @"\mod.pptm");
            BLANK_POWERPOINT = String.Concat(A3_RESOURCE, @"\blank.pptm");
            CHAPTER_VBA = File.ReadAllText(String.Concat(A3_RESOURCE, @"\chapter_vba.txt"));
            QUESTION_VBA = File.ReadAllText(String.Concat(A3_RESOURCE, @"\question_vba.txt"));
            YAML_LINT_CONFIG = String.Concat(A3_RESOURCE, @"\yamllint_config.yml");
    }

        public static void StartUp()
        {
            // Create the A3 directory structure if it does not exist. 
            try { Directory.CreateDirectory(A3Environment.A3_PATH); } catch { }
            try { Directory.CreateDirectory(A3Environment.A3_WORKING); } catch { }
            try { Directory.CreateDirectory(A3Environment.A3_PUBLISH); } catch { }
            try { Directory.CreateDirectory(A3Environment.A3_LOG); } catch { }
            try { Directory.CreateDirectory(A3Environment.A3_PRES_PNGS); } catch { }
            try { Directory.CreateDirectory(A3Environment.A3_BOOK_PNGS); } catch { }
            try { Directory.CreateDirectory(A3Environment.A3_MARKDOWN); } catch { }
            try { Directory.CreateDirectory(A3Environment.A3_LATEX); } catch { }
            if (!ApplictionInstalled("pandoc") || !ApplictionInstalled("texstudio") || !ApplictionInstalled("miktex"))
            {
                // RUN THE UPDATE CHOCO SCRIPT
            }
        }
        private static bool ApplictionInstalled(string program)
        {
            string displayName;
            RegistryKey key;

            // search in: CurrentUser
            key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            foreach (String keyName in key.GetSubKeyNames())
            {
                RegistryKey subkey = key.OpenSubKey(keyName);
                displayName = subkey.GetValue("DisplayName") as string;
                if (program.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
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
                if (program.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
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
                if (program.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return true;
                }
            }

            // NOT FOUND
            return false;
        }
    }
}
