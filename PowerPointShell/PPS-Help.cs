using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;

namespace PowerPointShell
{
    [Cmdlet("A3", "Help")]
    public class PPSHelpCmdlet : Cmdlet
    {
        [Parameter(Position = 0)]
        public string Verbose { get; set; }

        protected override void ProcessRecord()
        {

        }
    }
}
