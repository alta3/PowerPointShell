using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using YamlDotNet.Core;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NodeDeserializers;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA
{
    public class A3Outline
    {
        #region Outline Properites
        public string Id { get; set; }
        public string Course { get; set; }
        public List<A3Content> TOC { get; set; }
        public List<A3Chapter> Chapters { get; set; }
        public List<A3Lab> Labs { get; set; }
        #endregion

        #region Generate From YAML Methods
        public void GeneratePresentation(PowerPoint.Presentation presentation)
        {
            this.GenerateCourseSlide(presentation);
            this.GenerateTOCSlide(presentation);
            presentation.SectionProperties.AddBeforeSlide(1, this.Course);
            this.GenerateChapters(presentation);
            this.GenerateEndOfDeckSlide(presentation);
            this.GenerateQuizSlide(presentation);
        }

        private void GenerateChapters(PowerPoint.Presentation presentation)
        {
            int chapterCount = 1;
            foreach (A3Chapter chapter in this.Chapters)
            {
                chapter.Generate(presentation, chapterCount);
                chapterCount += 1;
            }
        }
        private void GenerateCourseSlide(PowerPoint.Presentation presentation)
        {
            // Insert the course slide from the model PowerPoint
            presentation.Slides.AddSlide(presentation.Slides.Count + 1, presentation.SlideMaster.CustomLayouts[1]);

            // Change the title to the course title given in the yaml file
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = this.Course,
                Type = "COURSE",
                ActiveGuid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateEndOfDeckSlide(PowerPoint.Presentation presentation)
        {
            // Insert a title slide from the model PowerPoint
            presentation.Slides.AddSlide(presentation.Slides.Count + 1, presentation.SlideMaster.CustomLayouts[4]);
            PowerPoint.SlideRange activeslide = presentation.Slides.Range(presentation.Slides.Count);

            // Change the title, chapsub, type, and active guid to accurately reflect what is happening
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "End of Deck",
                ChapSub = String.Concat(this.Course, ": End Of Deck"),
                ActiveGuid = Guid.NewGuid().ToString(),
                Type = "CONTENT"
            };
            a3ActiveSlide.WriteFromMemory();
        }
        private void GenerateTOCSlide(PowerPoint.Presentation presentation)
        {
            // Insert a split slide from the model PowerPoint
            presentation.Slides.AddSlide(presentation.Slides.Count + 1, presentation.SlideMaster.CustomLayouts[4]);

            // Populate the appropriate values of the slide deck here
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Table of Contents",
                ChapSub = String.Concat(this.Course, ": TOC"),
                Type = "TOC",
                ActiveGuid = Guid.NewGuid().ToString()
            };
            a3ActiveSlide.WriteFromMemory();

            // TODO: Create a linked list to the first chapter of each day and colorize the results
        }
        private void GenerateQuizSlide(PowerPoint.Presentation presentation)
        {
            // Insert a question slide from the model PowerPoint
            presentation.Slides.InsertFromFile(A3Globals.MODEL_POWERPOINT, presentation.Slides.Count + 1, 6);       

            // Ensure the title is Knowledge Check and move on 
            A3Slide a3ActiveSlide = new A3Slide(presentation.Slides[presentation.Slides.Count])
            {
                Title = "Knowledge Check",
                Type = "QUESTION",
                ActiveGuid = Guid.NewGuid().ToString(),
            };
            a3ActiveSlide.WriteFromMemory();
            presentation.SectionProperties.AddBeforeSlide(presentation.Slides.Count, "Knowledge Check");

        }
        #endregion

        #region Validation Methods
        public void ValidateYAMLStructure(A3LogFile logFile, string yamlText)
        {
            A3Yaml.Lint(logFile, yamlText);
        }

        public void Validate(A3LogFile logFile, int processingLevel)
        {
            this.ValidateTitle(logFile);
            // If process dicatates different checks make a switch statement here locally but pass the process varaible to the chapter and labs so that it can independently handle those
            this.ValidateChapters(logFile, processingLevel);
            this.ValidateLabs(logFile, processingLevel);
        }

        private void ValidateTitle(A3LogFile logFile)
        {
            if (this.Course == null || this.Course.Count(c => !Char.IsWhiteSpace(c)) == 0)
            {
                logFile.WriteEntry(new A3LogEntry("ERROR", "No Course Title Found"));
            }
        }
        private void ValidateChapters(A3LogFile logFile, int processingLevel)
        {
            int chapterCount = 1;
            foreach (A3Chapter chapter in this.Chapters)
            {
                chapter.Validate(chapterCount, logFile, processingLevel);
                chapterCount++;
            }
        }
        private void ValidateLabs(A3LogFile logFile, int processingLevel)
        {
            int labCount = 1;
            foreach (A3Lab lab in this.Labs)
            {
                // lab.Validate(logFile, process);
                labCount++;
            }
        }
        #endregion
    }
}
