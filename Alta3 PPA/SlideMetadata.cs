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
        }

        public void DrawSlideInfo()
        {
            // Initialize the controls
            A3Globals.A3SLIDE.Slide.Select();
            this.InitializeTitle();
            this.InitializeChapSub();
            this.InitializeActiveGuid();
            this.InitializeType();
        }

        // Type Functions
        private void InitializeType()
        {
            // Add all possible options to the type combobox
            CBType.Items.Clear();
            CBType.Items.Add("Course Title Card");
            CBType.Items.Add("Table Of Contents Slide");
            CBType.Items.Add("Chapter Title Card");
            CBType.Items.Add("Content Slide");
            CBType.Items.Add("Do Not Publish Slide");
            CBType.Items.Add("Question Slide");

            CBType.SelectedIndex = 3;

            // Select the proper index to display based on the current type indicated by the slide
            try
            {
                switch (A3Globals.A3SLIDE.Type.ToUpper())
                {
                    case "COURSE":
                        CBType.SelectedIndex = 0;
                        break;
                    case "TOC":
                        CBType.SelectedIndex = 1;
                        break;
                    case "CHAPTER":
                        CBType.SelectedIndex = 2;
                        break;
                    case "CONTENT":
                        CBType.SelectedIndex = 3;
                        break;
                    case "NO-PUB":
                        CBType.SelectedIndex = 4;
                        break;
                    case "QUESTION":
                        CBType.SelectedIndex = 5;
                        break;
                    default:
                        CBType.SelectedIndex = 3;
                        break;
                }
            }
            catch 
            {
                
            }
        }
        private void CBType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.CBType.SelectedIndex == 0 || this.CBType.SelectedIndex == 2 || this.CBType.SelectedIndex == 5)
            {
                this.CBScrubberKey.Enabled = false;
                this.TBScrubberValue.Enabled = false;
                this.TBScrubberValue.Clear();
                this.CBScrubberKey.Items.Clear();
                this.CBScrubberKey.Text = "";
                this.BtnNewScrubber.Enabled = false;
                this.BtnSwapScrubber.Enabled = false;
            }
            else
            {
                this.CBScrubberKey.Enabled = true;
                this.TBScrubberValue.Enabled = true;
                this.InitializeChapSub();
                this.BtnNewScrubber.Enabled = true;
                this.BtnSwapScrubber.Enabled = true;
            }
        }
        private void SaveType()
        {
            switch (this.CBType.Text)
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

        // Guid Functions
        private void InitializeActiveGuid()
        {
            TBActiveGuid.Clear();
            TBActiveGuid.Enabled = false;
            if (A3Globals.A3SLIDE.ActiveGuid != null)
            {
                TBActiveGuid.Text = A3Globals.A3SLIDE.ActiveGuid;
            }
            else
            {
                TBActiveGuid.Text = Guid.NewGuid().ToString();
            }
        }
        private void BtnShowActiveGuids_Click(object sender, EventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            if (!A3Globals.SHOW_ACTIVE_GUID)
            {
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    try
                    {
                        slide.Shapes["ACTIVE_GUID"].Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        slide.Shapes["ACTIVE_GUID"].Fill.ForeColor.RGB = 763355;
                    }
                    catch { }
                    A3Globals.SHOW_ACTIVE_GUID = true;
                }
            }
            else
            {
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    try { slide.Shapes["ACTIVE_GUID"].Visible = Microsoft.Office.Core.MsoTriState.msoFalse; } catch { }
                    A3Globals.SHOW_ACTIVE_GUID = false;
                }
            }
        }
        private void BtnCopyActiveGuid_Click(object sender, EventArgs e)
        {
            // TODO: Read and update the form before pulling this information
            Clipboard.SetText(this.TBActiveGuid.Text);
        }
        private void SaveGuids()
        {
            A3Globals.A3SLIDE.ActiveGuid = TBActiveGuid.Text;
            foreach (string hguid in CBHistoricGuid.Items)
            {
                A3Globals.A3SLIDE.HistoricGuids.Add(hguid);
            }
        }

        // Title Functions
        private void InitializeTitle()
        {
            CBTitleKey.SelectedIndex = -1;
            CBTitleKey.Items.Clear();
            TBTitleValue.Clear();

            try
            {
                foreach (string shapeName in A3Globals.A3SLIDE.ShapeNames)
                {
                    CBTitleKey.Items.Add(shapeName);
                }

                int index = A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "TITLE") >= 0 ? A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "TITLE") : 0;
                CBTitleKey.SelectedIndex = index;

                try { TBTitleValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBTitleKey.SelectedItem].TextFrame.TextRange.Text; } catch { }
            }
            catch { }
            CBTitleKey.Update();
            CBTitleKey.Show();
        }
        private void CBTitleKey_SelectedIndexChanged(object sender, EventArgs e)
        {
            try { TBTitleValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBTitleKey.SelectedItem].TextFrame.TextRange.Text; } catch { }
        }
        private void SaveTitle()
        {
            A3Globals.A3SLIDE.Title = TBTitleValue.Text;
            PowerPoint.Shape shape = A3Globals.A3SLIDE.Slide.Shapes[CBTitleKey.SelectedItem];
            shape.Name = "TITLE";
            shape.Title = "TITLE";
        }

        // CHAP:SUB Functions
        private void InitializeChapSub()
        {
            try
            {
                CBScrubberKey.Items.Clear();
                TBScrubberValue.Clear();

                foreach (string shapeName in A3Globals.A3SLIDE.ShapeNames)
                {
                    CBScrubberKey.Items.Add(shapeName);
                }

                int index = A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "CHAP:SUB") > 0 ? A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "CHAP:SUB") : 0;
                CBScrubberKey.SelectedIndex = index;

                try { TBScrubberValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBScrubberKey.SelectedItem].TextFrame.TextRange.Text; } catch { }
            }
            catch
            {

            }

        }
        private void CBScrubberKey_SelectedIndexChanged(object sender, EventArgs e)
        {
            TBScrubberValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBScrubberKey.SelectedItem].TextFrame.TextRange.Text;
        }
        private void SaveChapSub()
        {
            A3Globals.A3SLIDE.ChapSub = TBScrubberValue.Text;
        }

        private void Save()
        {
            if (CBType.Text == "" || !CBType.Items.Contains(CBType.Text))
            {
                MessageBox.Show("YOU MUST SELECT A VALID TYPE!", "SELECT VALID TYPE!", MessageBoxButtons.OK);
                return;
            }
            if (CBTitleKey.Text == CBScrubberKey.Text)
            {
                MessageBox.Show("CHAP:SUB AND TITLE MAY NOT BE THE SAME OBJECT", "DUPLICATE ITEM", MessageBoxButtons.OK);
                return;
            }
            if (CBTitleKey.Text == null || CBTitleKey.Text == "")
            {
                MessageBox.Show("MUST SELECT A TITLE", "NULL TITLE", MessageBoxButtons.OK);
                return;
            }
            this.SaveType();
            this.SaveTitle();
            this.SaveChapSub();
            this.SaveGuids();
            A3Globals.A3SLIDE.WriteFromMemory();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            this.Save();
            A3Globals.A3SLIDE.ReadFromSlide();
            this.DrawSlideInfo();
        }
        private void BtnSaveAndProceed_Click(object sender, EventArgs e)
        {
            this.Save();
            this.Close();
            A3Globals.QUIT_FROM_CURRENT_LOOP = false;
        }
        
        private void BtnPreviousSlide_Click(object sender, EventArgs e)
        {
            int slideIndex = A3Globals.A3SLIDE.Slide.SlideIndex - 1;
            try { A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]); A3Globals.A3SLIDE.ReadShapes(); this.DrawSlideInfo(); }
            catch { MessageBox.Show("BEGINING OF SLIDE SHOW", "ERROR", MessageBoxButtons.OK); }
        }
        private void BtnNextSlide_Click(object sender, EventArgs e)
        {
            int slideIndex = A3Globals.A3SLIDE.Slide.SlideIndex + 1;
            try { A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]); A3Globals.A3SLIDE.ReadShapes(); this.DrawSlideInfo(); }
            catch { MessageBox.Show("END OF SLIDE SHOW", "ERROR", MessageBoxButtons.OK); }
        }

        #region TODO: IMPLEMENT
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
