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
            // Spawn a new log file to track everything that happens during the publishing process
            A3LogFile a3LogFile = new A3LogFile();

            // Consume the presentation into the a3Slide formats to be worked with
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            A3Globals.A3PRESENTATION = new A3Presentation(presentation);
            // One day all the validation of structure and layout should be handled in the function below
            // A3Globals.A3PRESENTATION.ValidateStructure();

            // Find all the slides marked Course
            List<A3Slide> a3SlideCourses = A3Globals.A3PRESENTATION.Slides.FindAll(a3Slide => a3Slide.Type == "COURSE");
            // If there are no slides marked as Course report this error and do not attempt to continue
            if (a3SlideCourses == null || a3SlideCourses.Count == 0)
            {
                a3LogFile.WriteError("No slide could be found marked as a course type. Ensure that the first slide in the deck has the type metadata field set to Course");
                goto REPORT;
            }
            // If there are more than one slide marked as course report this error and do not attempt to continue
            else if (a3SlideCourses.Count > 1)
            {
                string slideNum = null;
                foreach (A3Slide a3Course in a3SlideCourses)
                {
                    String.Concat(slideNum, a3Course.Slide.SlideNumber.ToString().Trim(), ",");
                };
                slideNum.TrimEnd(',');
                a3LogFile.WriteError(String.Concat("More than one slide is marked course. The following slides are marked as course slides: ", slideNum));
                goto REPORT;
            }

            // Create the outline that will be used for yaml presentation and the json output as well
            A3Outline outline = new A3Outline
            {
                Course = a3SlideCourses[0].Title,
                Chapters = new List<A3Chapter>()
            };

            // Find all the slides marked as Chapters
            List<A3Slide> a3SlideChapters = A3Globals.A3PRESENTATION.Slides.FindAll(a3slide => a3slide.Type == "CHAPTER");

            // Track the start of each chapter -- the last number is the start of the question slide
            List<int> chapStart = new List<int>();

            // Convert the A3Slide to an A3Chapter and add it to the outline
            foreach (A3Slide a3Chapter in a3SlideChapters)
            {
                outline.Chapters.Add((A3Chapter)a3Chapter.TypeConversion());
                chapStart.Add(a3Chapter.Slide.SlideNumber);
            }

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

            // For each chapter found build the subchapters & fill them with content slides
            for (int i = 0; i - 1 < chapStart.Count; i++)
            {
                // Get the slide range of the current chpater and then create a list of a3slides that represent the chapter. 
                int chapstart = chapStart[i];
                int chapend = chapStart[i + 1];
                List<A3Slide> a3SlideChapterSlides = A3Globals.A3PRESENTATION.Slides.FindAll(a3Slide => (a3Slide.Slide.SlideNumber > chapstart) && (a3Slide.Slide.SlideNumber < chapend));

                // Track the subchapters in a list
                List<string> subchapters = new List<string>();

                // Track the start and end of each subchapter in a list
                List<int> subchaptersStart = new List<int>();

                // For each slide in the chapter split the chap:sub into its respective parts, and add different subchapters to the list of subchapters
                foreach (A3Slide a3Slide in a3SlideChapterSlides)
                {
                    string[] chapSubArray = a3Slide.ChapSub.Split(':');
                    string chapter = null;
                    string subchapter = null;
                    try { chapter = chapSubArray[0]; } catch { a3LogFile.WriteError(String.Concat("Could not split on ':' for the chap:sub metadata field at slide number: ", a3Slide.Slide.SlideNumber.ToString().Trim(), ". Continuing with 'ERROR' as the chapter title.")); chapter = "ERROR"; }
                    try { subchapter = chapSubArray[1]; } catch { a3LogFile.WriteError(String.Concat("Could not split on ':' for the chap:sub metadata field at slide number: ", a3Slide.Slide.SlideNumber.ToString().Trim(), ". Continuing with 'ERROR' as the subchapter title.")); subchapter = "ERROR"; }
                    if (subchapters.Count == 0)
                    {
                        subchapters.Add(subchapter);
                        subchaptersStart.Add(a3Slide.Slide.SlideNumber);
                    }
                    else
                    {
                        if (subchapters.Last() != subchapter)
                        {
                            subchapters.Add(subchapter);
                            subchaptersStart.Add(a3Slide.Slide.SlideNumber);
                        }
                    }
                }

                List<A3Subchapter> a3Subchapters = new List<A3Subchapter>();

                foreach (string subchap in subchapters)
                {
                    a3Subchapters.Add(new A3Subchapter(subchap));
                }

                for (int z = 0; z < subchaptersStart.Count; z++)
                {
                    int substart = subchaptersStart[z];
                    int subend = (subchaptersStart.Count == z - 1) ? chapend : subchaptersStart[z + 1];
                    List<A3Slide> a3ContentSlides = A3Globals.A3PRESENTATION.Slides.FindAll(a3slide => (a3slide.Slide.SlideNumber >= substart) && (a3slide.Slide.SlideNumber < subend));

                    foreach (A3Slide a3slide in a3ContentSlides)
                    {
                        a3Subchapters[z].Slides.Add((A3Content)a3slide.TypeConversion());
                    }
                }

            }
            A3Yaml.ProduceYaml(a3LogFile, outline);

            REPORT:
            if (a3LogFile.HasError())
            {
                MessageBox.Show(String.Concat("There were errors while producing this output. The specifics can be seen in the following location:\r\n", a3LogFile.Path));
            }
            else
            {
                MessageBox.Show("Successfully Produced Outline & PNGs");
            }
        }
    }
}

