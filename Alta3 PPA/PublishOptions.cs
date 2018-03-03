using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public partial class PublishOptions : Form
    {
        public PublishOptions()
        {
            InitializeComponent();
        }

        private void btnPublish_Click(object sender, EventArgs e)
        {
            A3LogFile a3LogFile = new A3LogFile();
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

        }
    }
}
