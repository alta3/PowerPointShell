using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using YamlDotNet.Serialization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public partial class FirstChapter : Form
    {
        public FirstChapter()
        {
            InitializeComponent();
        }

        private void BtnGo_Click(object sender, EventArgs e)
        {   
            // Find all the slides marked as Question slides
            List<A3Slide> a3SlideQuestions = A3Globals.A3PRESENTATION.Slides.FindAll(a3Slide => a3Slide.Type == "QUESTION");
            // If there are no slides marked as Question report this error and do not attempt to continue
            if (a3SlideQuestions == null || a3SlideQuestions.Count == 0)
            {
                a3LogFile.WriteError("No slide could be found marked as a question type. Ensure that the question slide in the deck has the type metadata field set to question");
                goto REPORT;
            }
            // If there are more than one slide marked as question report this error and do not attempt to continue
            else if (a3SlideQuestions.Count > 1)
            {
                string slideNum = null;
                foreach (A3Slide a3Question in a3SlideQuestions)
                {
                    String.Concat(slideNum, a3Question.Slide.SlideNumber.ToString().Trim(), ",");
                };
                slideNum.TrimEnd(',');
                a3LogFile.WriteError(String.Concat("More than one slide is marked question. The following slides are marked as question slides: ", slideNum));
                goto REPORT;
            }
            else
            {
                chapStart.Add(a3SlideQuestions[0].Slide.SlideNumber);
            }
        }
    }
}

