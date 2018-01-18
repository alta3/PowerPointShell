using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
            try { Directory.CreateDirectory(A3Globals.A3_PNGS); } catch { }
        }
    }
}
