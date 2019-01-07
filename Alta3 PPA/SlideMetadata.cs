using System;
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
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;

            // Set Quit from current loop to true so that the red X works without looping through the checks
            A3Environment.QUIT_FROM_CURRENT_LOOP = true;
        }

        public void DrawSlideInfo()
        {
            // Initialize the controls
            A3Environment.A3SLIDE.Slide.Select();
            InitializeTitle();
            InitializeChapSub();
            InitializeGuid();
            InitializeType();
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

            CBType.SelectedIndex = (int)A3Environment.A3SLIDE.Type < 6 ? (int)A3Environment.A3SLIDE.Type : 3;
        }
        private void CBType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBType.SelectedIndex == 0 || CBType.SelectedIndex == 2 || CBType.SelectedIndex == 5)
            {
                CBTitle.Enabled = false;
                TBTitle.Enabled = false;
                TBTitle.Clear();
                CBTitle.Items.Clear();
                return;
            }
            CBTitle.Enabled = true;
            CBTitle.Enabled = true;
            InitializeChapSub();
        }
        private void SaveType()
        {
            A3Environment.A3SLIDE.Type = (A3Slide.Types)CBType.SelectedIndex;
        }

        // Guid Functions
        private void InitializeGuid()
        {
            TBGuid.Clear();
            TBGuid.Enabled = false;
            TBGuid.Text = A3Environment.A3SLIDE.Guid is null ? Guid.NewGuid().ToString() : A3Environment.A3SLIDE.Guid;
        }
        private void BtnShowGuids_Click(object sender, EventArgs e)
        {
            A3Presentation presentation = new A3Presentation(Globals.ThisAddIn.Application.ActivePresentation);
            presentation.ShowGuids();
        }
        private void BtnCopyGuid_Click(object sender, EventArgs e)
        {
            // TODO: Read and update the form before pulling this information
            Clipboard.SetText(this.TBGuid.Text);
        }
        private void SaveGuids()
        {
            A3Environment.A3SLIDE.Guid = TBGuid.Text;
            foreach (string hguid in CBHistoricGuid.Items)
            {
                A3Environment.A3SLIDE.HGuids.Add(hguid);
            }
        }

        // Title Functions
        private void InitializeTitle()
        {
            CBChapSub.SelectedIndex = -1;
            CBChapSub.Items.Clear();
            TBChapter.Clear();

            try
            {
                foreach (string shapeName in A3Environment.A3SLIDE.ShapeNames)
                {
                    CBChapSub.Items.Add(shapeName);
                }

                int index = A3Environment.A3SLIDE.ShapeNames.FindIndex(s => s == "TITLE") >= 0 ? A3Environment.A3SLIDE.ShapeNames.FindIndex(s => s == "TITLE") : 0;
                CBChapSub.SelectedIndex = index;

                try { TBChapter.Text = A3Environment.A3SLIDE.Slide.Shapes[CBChapSub.SelectedItem].TextFrame.TextRange.Text; } catch { }
            }
            catch { }
            CBChapSub.Update();
            CBChapSub.Show();
        }
        private void CBTitleKey_SelectedIndexChanged(object sender, EventArgs e)
        {
            try { TBChapter.Text = A3Environment.A3SLIDE.Slide.Shapes[CBChapSub.SelectedItem].TextFrame.TextRange.Text; } catch { }
        }
        private void SaveTitle()
        {
            A3Environment.A3SLIDE.Title = TBChapter.Text;
            PowerPoint.Shape shape = A3Environment.A3SLIDE.Slide.Shapes[CBChapSub.SelectedItem];
            shape.Name = "TITLE";
            shape.Title = "TITLE";
        }

        // CHAP:SUB Functions
        private void InitializeChapSub()
        {
            try
            {
                CBTitle.Items.Clear();
                TBTitle.Clear();

                foreach (string shapeName in A3Environment.A3SLIDE.ShapeNames)
                {
                    CBTitle.Items.Add(shapeName);
                }

                int index = A3Environment.A3SLIDE.ShapeNames.FindIndex(s => s == "CHAP:SUB") > 0 ? A3Environment.A3SLIDE.ShapeNames.FindIndex(s => s == "CHAP:SUB") : 0;
                CBTitle.SelectedIndex = index;

                try { TBTitle.Text = A3Environment.A3SLIDE.Slide.Shapes[CBTitle.SelectedItem].TextFrame.TextRange.Text; } catch { }
            }
            catch
            {

            }

        }
        private void CBScrubberKey_SelectedIndexChanged(object sender, EventArgs e)
        {
            TBTitle.Text = A3Environment.A3SLIDE.Slide.Shapes[CBTitle.SelectedItem].TextFrame.TextRange.Text;
        }
        private void SaveChapSub()
        {
            A3Environment.A3SLIDE.ChapSub = TBTitle.Text;
        }

        private void Save()
        {
            if (CBChapSub.Text == CBTitle.Text)
            {
                MessageBox.Show("TITLE AND CHAPSUB MAY NOT BE THE SAME OBJECT", "DUPLICATE ITEM", MessageBoxButtons.OK);
                return;
            }
            if (CBChapSub.Text == null || CBChapSub.Text == "")
            {
                MessageBox.Show("MUST SELECT A TITLE", "NULL TITLE", MessageBoxButtons.OK);
                return;
            }
            SaveType();
            SaveTitle();
            SaveChapSub();
            SaveGuids();
            A3Environment.A3SLIDE.WriteFromMemory();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            Save();
            A3Environment.A3SLIDE.ReadFromSlide();
            DrawSlideInfo();
        }
        private void BtnSaveAndProceed_Click(object sender, EventArgs e)
        {
            Save();
            Close();
            A3Environment.QUIT_FROM_CURRENT_LOOP = false;
        }
        
        private void BtnPreviousSlide_Click(object sender, EventArgs e)
        {
            int slideIndex = A3Environment.A3SLIDE.Slide.SlideIndex - 1;
            try { A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]); A3Environment.A3SLIDE.ReadShapes(); this.DrawSlideInfo(); }
            catch { MessageBox.Show("BEGINING OF SLIDE SHOW", "ERROR", MessageBoxButtons.OK); }
        }
        private void BtnNextSlide_Click(object sender, EventArgs e)
        {
            int slideIndex = A3Environment.A3SLIDE.Slide.SlideIndex + 1;
            try { A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]); A3Environment.A3SLIDE.ReadShapes(); this.DrawSlideInfo(); }
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
