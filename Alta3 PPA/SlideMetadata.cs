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
    public partial class SlideMetadata : Form
    {
        public SlideMetadata()
        {
            // Generated code
            InitializeComponent();
            
            // Fix the frame of the form so that it cannot be resized
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Set Quit from current loop to true so that the red X works without looping through the checks
            A3Globals.QUIT_FROM_CURRENT_LOOP = true;

            // Draw Current Slide Information To Screen
            this.DrawSlideInfo();
        }

        private void DrawSlideInfo()
        {
            // Read all the current values from the slide
            this.GetValuesFromSlide();

            // Initialize the controls
            this.InitializeType();
        }

        private void GetValuesFromSlide()
        {
            // Read the current values from the slide f
            A3Globals.A3SLIDE.ReadDay();
            A3Globals.A3SLIDE.ReadType();
            A3Globals.A3SLIDE.ReadTitle();
            A3Globals.A3SLIDE.ReadChapSub();
            A3Globals.A3SLIDE.ReadActiveGuid();
            A3Globals.A3SLIDE.ReadHistoricGuid();

            // Get a list of all the shape names that have text and are not tracked metadata fields
            foreach (PowerPoint.Shape shape in A3Globals.A3SLIDE.Slide.Shapes)
            {
                if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue
                    && shape.Name != "ACTIVE_GUID"
                    && shape.Name != "HISTORIC_GUID"
                    && shape.Name != "TYPE"
                    && shape.Name != "TITLE"
                    && shape.Name != "DAY"
                    && shape.Name != "CHAP:SUB")
                {
                    A3Globals.SHAPE_NAMES.Add(shape.Name);
                }
            }
        }

        // Type Functions
        private void InitializeType()
        {
            // Add all possible options to the type combobox
            CBType.Items.Add("Course Title Card");
            CBType.Items.Add("Table Of Contents Slide");
            CBType.Items.Add("Chapter Title Card");
            CBType.Items.Add("Content Slide");
            CBType.Items.Add("Do Not Publish Slide");
            CBType.Items.Add("Question Slide");

            // Select the proper index to display based on the current type indicated by the slide
            switch (A3Globals.A3SLIDE.Type.ToUpper())
            {
                case "COURSE":
                    CBType.SelectedIndex = 1;
                    break;
                case "TOC":
                    CBType.SelectedIndex = 2;
                    break;
                case "CHAPTER":
                    CBType.SelectedIndex = 3;
                    break;
                case "CONTENT":
                    CBType.SelectedIndex = 4;
                    break;
                case "NO-PUB":
                    CBType.SelectedIndex = 5;
                    break;
                case "QUESTION":
                    CBType.SelectedIndex = 6;
                    break;
                default:
                    // TODO: Log whatever happened here... cause its wrong... yup wrong.
                    break;
            }
        }
        private void CBType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.CBType.SelectedText == "Course Title Card" || this.CBType.SelectedText == "Question Slide")
            {
                this.CBScrubberKey.Enabled = false;
                this.TBScrubberValue.Enabled = false;
                this.TBScrubberValue.Clear();
                this.BtnNewScrubber.Enabled = false;
                this.BtnSwapScrubber.Enabled = false;
            }
            else
            {
                this.CBScrubberKey.Enabled = true;
                this.TBScrubberValue.Enabled = true;
                this.BtnNewScrubber.Enabled = true;
                this.BtnSwapScrubber.Enabled = true;
            }
        }
        private void SaveType()
        {
            switch (this.CBType.SelectedText)
            {
                case "Course Title Card":
                    A3Globals.A3SLIDE.Type = "COURSE";
                    break;
                case "Table Of Contents Slide":
                    A3Globals.A3SLIDE.Type = "TOC";
                    break;
                case "Chapter Title Card":
                    A3Globals.A3SLIDE.Type = "CHAPTER";
                    break;
                case "Content Slide":
                    A3Globals.A3SLIDE.Type = "CONTENT";
                    break;
                case "Do Not Publish Slide":
                    A3Globals.A3SLIDE.Type = "NO-PUB";
                    break;
                case "Question Slide":
                    A3Globals.A3SLIDE.Type = "QUESTION";
                    break;
                default:
                    A3Globals.A3SLIDE.Type = null;
                    break;
            }
        }

        // Active Guid Functions
        private void InitializeActiveGuid()
        {
            if (A3Globals.A3SLIDE.ActiveGuid != null)
            {
                TBActiveGuid.Text = A3Globals.A3SLIDE.ActiveGuid;
                TBActiveGuid.Enabled = false;
            }
            else
            {

            }
        }

        private void BtnShowActiveGuids_Click(object sender, EventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            if (!A3Globals.SHOW_ACTIVE_GUID)
            {
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    slide.Shapes["ACTIVE_GUID"].Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    slide.Shapes["ACTIVE_GUID"].Fill.ForeColor.RGB = 763355;
                    A3Globals.SHOW_ACTIVE_GUID = true;
                }
            }
            else
            {
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    slide.Shapes["ACTIVE_GUID"].Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    A3Globals.SHOW_ACTIVE_GUID = false;
                }
            }
        }
        private void BtnCopyActiveGuid_Click(object sender, EventArgs e)
        {
            // TODO: Read and update the form before pulling this information
            Clipboard.SetText(this.TBActiveGuid.Text);
        }

        private void CBScrubberKey_SelectedIndexChanged(object sender, EventArgs e)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            TBScrubberValue.Text = slide.Shapes[CBScrubberKey.SelectedItem].TextFrame.TextRange.Text;
        }
        private void CBTitleKey_SelectedIndexChanged(object sender, EventArgs e)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            TBTitleValue.Text = slide.Shapes[CBTitleKey.SelectedItem].TextFrame.TextRange.Text;
        }

        private void UpdateFields(A3Slide a3Slide)
        {
            PowerPoint.Slide slide = a3Slide.Slide;
            #region Type Setup
            if (a3Slide.Type == null)
            {
                CBType.SelectedIndex = 3;
                PowerPoint.Shape type = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 30);
                type.Name = "TYPE";
                type.TextFrame.TextRange.Text = "";
                type.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                CBType.Enabled = true;
            }
            else
            {
                CBType.SelectedItem = slide.Shapes["TYPE"].TextFrame.TextRange.Text;
            }
            #endregion

            #region Active Guid Setup
            TBActiveGuid.Enabled = false;
            if (a3Slide.ActiveGuid == null)
            {
                PowerPoint.Shape aguid = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 500, 400, 30);
                aguid.Name = "ACTIVE_GUID";
                aguid.TextFrame.TextRange.Text = Guid.NewGuid().ToString();
                aguid.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                TBActiveGuid.Text = aguid.TextFrame.TextRange.Text;
            }
            else
            {
                TBActiveGuid.Text = slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text;
            }
            #endregion

            #region Historic Guid Setup
            if (a3Slide.HistoricGuids == null)
            {
                PowerPoint.Shape hguid = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 450, 400, 30);
                hguid.Name = "HISTORIC_GUID";
                hguid.TextFrame.TextRange.Text = "";
                hguid.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                CBHistoricGuid.Text = hguid.TextFrame.TextRange.Text;
            }
            else
            {
                CBHistoricGuid.Text = slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text;
            }
            #endregion

            #region Scrubber Setup
            if (a3Slide.ChapSub == null)
            {
                string msg = "Could not find a field named CHAP:SUB\r\nIs one of the textboxes on the slide currently a fields that should be labeled scrubber?";
                DialogResult dialogResult = MessageBox.Show(msg, "Scrubber Check", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    CBScrubberKey.Items.AddRange(A3Globals.SHAPE_NAMES.ToArray());
                }
                else
                {
                    PowerPoint.Shape scrubber = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 450, 400, 30);
                    scrubber.Name = "CHAP:SUB";
                    CBScrubberKey.Items.Add("CHAP:SUB");
                    CBScrubberKey.SelectedIndex = 0;
                    CBScrubberKey.Enabled = false;
                    TBScrubberValue.Text = "";
                    TBScrubberValue.Enabled = true;
                }
            }
            else
            {
                CBScrubberKey.Items.Add("CHAP:SUB");
                CBScrubberKey.SelectedIndex = 0;
                CBScrubberKey.Enabled = false;
                TBScrubberValue.Text = slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text;
            }
            #endregion

            #region Title Setup
            if (a3Slide.Title == null)
            {
                string msg = "Could not find a field named TITLE\r\nIs one of the textboxes on the slide currently a fields that should be labeled title?";
                DialogResult dialogResult = MessageBox.Show(msg, "Title Check", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    CBTitleKey.Items.AddRange(A3Globals.SHAPE_NAMES.ToArray());
                }
                else
                {
                    PowerPoint.Shape title = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 450, 400, 30);
                    title.Name = "TITLE";
                    CBTitleKey.Items.Add("TITLE");
                    CBTitleKey.SelectedIndex = 0;
                    CBTitleKey.Enabled = false;
                    TBTitleValue.Text = "";
                    TBTitleValue.Enabled = true;
                }
            }
            else
            {
                CBTitleKey.Items.Add("TITLE");
                CBTitleKey.SelectedIndex = 0;
                CBTitleKey.Enabled = false;
                TBTitleValue.Text = slide.Shapes["TITLE"].TextFrame.TextRange.Text;
            }
            #endregion

            #region Day Setup
            if (a3Slide.Day == null)
            {
                PowerPoint.Shape day = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 450, 400, 30);
                day.Name = "DAY";
                day.TextFrame.TextRange.Text = "";
                day.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                TBDay.Text = day.TextFrame.TextRange.Text;
            }
            else
            {
                TBDay.Text = slide.Shapes["DAY"].TextFrame.TextRange.Text;
            }
            #endregion
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (CBType.Text == "")
            {
                MessageBox.Show("YOU MUST SELECT A TYPE!", "NO TYPE!", MessageBoxButtons.OK);
            }
            else
            {
                if (CBTitleKey.Text == CBScrubberKey.Text)
                {
                    MessageBox.Show("CHAP:SUB AND TITLE MAY NOT BE THE SAME OBJECT", "DUPLICATE ITEM", MessageBoxButtons.OK);
                }
                else
                {
                    if (CBTitleKey.Text == null || CBTitleKey.Text == "" || CBScrubberKey.Text == null || CBScrubberKey.Text == "")
                    {
                        MessageBox.Show("MUST SELECT A TITLE AND CHAP:SUB", "NULL TITLE OR CHAP:SUB", MessageBoxButtons.OK);
                    }
                    else
                    {
                        PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                        slide.Shapes[CBScrubberKey.Text].Name = "CHAP:SUB";
                        slide.Shapes[CBTitleKey.Text].Name = "TITLE";

                        slide.Shapes["TYPE"].TextFrame.TextRange.Text = this.CBType.Text;

                        if (TBActiveGuid.Text == null)
                        {
                            slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text = this.TBActiveGuid.Text;
                        }

                        if (CBHistoricGuid.Text == null)
                        {
                            slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text = this.CBHistoricGuid.Text;
                        }

                        if (TBScrubberValue.Text == null)
                        {
                            slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text = this.TBScrubberValue.Text;
                        }

                        if (TBTitleValue.Text == null)
                        {
                            slide.Shapes["TITLE"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["TITLE"].TextFrame.TextRange.Text = this.TBTitleValue.Text;
                        }

                        if (TBDay.Text == null)
                        {
                            slide.Shapes["DAY"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["DAY"].TextFrame.TextRange.Text = this.TBDay.Text;
                        }

                        A3Slide a3Slide = new A3Slide(slide);
                        Point point = this.Location;
                        this.Close();
                        this.Dispose();
                        SlideMetadata slideMetadata = new SlideMetadata()
                        {
                            StartPosition = FormStartPosition.Manual,
                            Location = point
                        };
                        slideMetadata.ShowDialog();
                    }
                }
            }
        }
        private void BtnSaveAndProceed_Click(object sender, EventArgs e)
        {
            if (CBType.Text == "")
            {
                MessageBox.Show("YOU MUST SELECT A TYPE!", "NO TYPE!", MessageBoxButtons.OK);
            }
            else
            {
                if (CBTitleKey.Text == CBScrubberKey.Text)
                {
                    MessageBox.Show("CHAP:SUB AND TITLE MAY NOT BE THE SAME OBJECT", "DUPLICATE ITEM", MessageBoxButtons.OK);
                }
                else
                {
                    if (CBTitleKey.Text == null || CBTitleKey.Text == "" || CBScrubberKey.Text == null || CBScrubberKey.Text == "")
                    {
                        MessageBox.Show("MUST SELECT A TITLE AND CHAP:SUB", "NULL TITLE OR CHAP:SUB", MessageBoxButtons.OK);
                    }
                    else
                    {
                        PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                        slide.Shapes[CBScrubberKey.Text].Name = "CHAP:SUB";
                        slide.Shapes[CBTitleKey.Text].Name = "TITLE";

                        slide.Shapes["TYPE"].TextFrame.TextRange.Text = this.CBType.Text;

                        if (TBActiveGuid.Text == null)
                        {
                            slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["ACTIVE_GUID"].TextFrame.TextRange.Text = this.TBActiveGuid.Text;
                        }

                        if (CBHistoricGuid.Text == null)
                        {
                            slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["HISTORIC_GUID"].TextFrame.TextRange.Text = this.CBHistoricGuid.Text;
                        }

                        if (TBScrubberValue.Text == null)
                        {
                            slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["CHAP:SUB"].TextFrame.TextRange.Text = this.TBScrubberValue.Text;
                        }

                        if (TBTitleValue.Text == null)
                        {
                            slide.Shapes["TITLE"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["TITLE"].TextFrame.TextRange.Text = this.TBTitleValue.Text;
                        }

                        if (TBDay.Text == null)
                        {
                            slide.Shapes["DAY"].TextFrame.TextRange.Text = "";
                        }
                        else
                        {
                            slide.Shapes["DAY"].TextFrame.TextRange.Text = this.TBDay.Text;
                        }
                    }
                }
            }
            this.Close();
            A3Globals.QUIT_FROM_CURRENT_LOOP = false;
        }
        private void SaveTitle()
        {

        }
        private void SaveChapSubchap()
        {

        }
        private void SaveDay()
        {

        }
        
        #region TODO: IMPLEMENT
        private void BtnPreviousSlide_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnNextSlide_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }

        private void BtnNewScrubber_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnSwapScrubber_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }

        private void BtnNewTitle_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnSwapTitle_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }

        private void BtnNextResult_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnPreviousResult_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }

        private void BtnHistoricToWorking_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnCommitGuid_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnNewGuid_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        #endregion


    }
}
