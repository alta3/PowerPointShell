﻿using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Alta3_PPA {
    public partial class SlideMetadata : Form {
        // Initialize the Slide Metadata Components
        public SlideMetadata() {
            // Generated code
            InitializeComponent();
            
            // Fix the frame of the form so that it cannot be resized
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Set Quit from current loop to true so that the red X works without looping through the checks
            A3Globals.QUIT_FROM_CURRENT_LOOP = true;
        }

        public void DrawSlideInfo() {
            // Initialize the controls
            A3Globals.A3SLIDE.Slide.Select();
            this.InitializeTitle();
            this.InitializeChapSub();
            this.InitializeActiveGuid();
            this.InitializeType();
        }

        // Type Functions
        private void InitializeType() {
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
            A3Slide.SlideType slideType;
            if (Enum.TryParse<A3Slide.SlideType>(A3Globals.A3SLIDE.Type.ToUpper(), out slideType)) {
                CBType.SelectedIndex = (int)slideType;
            }
            else {
                CBType.SelectedIndex = 3;
            }
        }
        private void CBType_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.CBType.SelectedIndex == (int)A3Slide.SlideType.COURSE || this.CBType.SelectedIndex == (int)A3Slide.SlideType.CHAPTER || this.CBType.SelectedIndex == (int)A3Slide.SlideType.QUESTION) {
                this.CBChapSubKey.Enabled = false;
                this.TBChapSubValue.Enabled = false;
                this.TBChapSubValue.Clear();
                this.CBChapSubKey.Items.Clear();
                this.CBChapSubKey.Text = "";
                this.BtnNewChapSub.Enabled = false;
                this.BtnSwapChapSub.Enabled = false;
            }
            else {
                this.CBChapSubKey.Enabled = true;
                this.TBChapSubValue.Enabled = true;
                this.InitializeChapSub();
                this.BtnNewChapSub.Enabled = true;
                this.BtnSwapChapSub.Enabled = true;
            }
        }
        private void SaveType() {
            A3Globals.A3SLIDE.Type = A3Slide.TypeStrings[this.CBType.SelectedIndex];
        }

        // Guid Functions
        private void InitializeActiveGuid() {
            TBActiveGuid.Clear();
            TBActiveGuid.Enabled = false;
            if (A3Globals.A3SLIDE.Guid != null) {
                TBActiveGuid.Text = A3Globals.A3SLIDE.Guid;
            }
            else {
                TBActiveGuid.Text = Guid.NewGuid().ToString();
            }
        }
        private void BtnShowActiveGuids_Click(object sender, EventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            if (!A3Globals.SHOW_ACTIVE_GUID) {
                foreach (PowerPoint.Slide slide in presentation.Slides) {
                    try {
                        slide.Shapes["ACTIVE_GUID"].Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        slide.Shapes["ACTIVE_GUID"].Fill.ForeColor.RGB = 763355;
                    }
                    catch { }
                    A3Globals.SHOW_ACTIVE_GUID = true;
                }
            }
            else {
                foreach (PowerPoint.Slide slide in presentation.Slides) {
                    try { slide.Shapes["ACTIVE_GUID"].Visible = Microsoft.Office.Core.MsoTriState.msoFalse; } catch { }
                    A3Globals.SHOW_ACTIVE_GUID = false;
                }
            }
        }
        private void BtnCopyActiveGuid_Click(object sender, EventArgs e) {
            // TODO: Read and update the form before pulling this information
            Clipboard.SetText(this.TBActiveGuid.Text);
        }
        private void SaveGuids() {
            A3Globals.A3SLIDE.Guid = TBActiveGuid.Text;
            foreach (string hGuid in CBHistoricGuid.Items) {
                A3Globals.A3SLIDE.HistoricGuids.Add(hGuid);
            }
        }

        // Title Functions
        private void InitializeTitle() {
            CBTitleKey.SelectedIndex = -1;
            CBTitleKey.Items.Clear();
            TBTitleValue.Clear();

            try {
                foreach (string shapeName in A3Globals.A3SLIDE.ShapeNames) {
                    CBTitleKey.Items.Add(shapeName);
                }

                int index = A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "TITLE") >= 0 ? A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "TITLE") : 0;
                CBTitleKey.SelectedIndex = index;

                try { TBTitleValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBTitleKey.SelectedItem].TextFrame.TextRange.Text; }
                catch { }
            }
            catch { }
            CBTitleKey.Update();
            CBTitleKey.Show();
        }
        private void CBTitleKey_SelectedIndexChanged(object sender, EventArgs e) {
            try {
                TBTitleValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBTitleKey.SelectedItem].TextFrame.TextRange.Text;
            }
            catch { }
        }
        private void SaveTitle() {
            A3Globals.A3SLIDE.Title = TBTitleValue.Text;
            PowerPoint.Shape shape = A3Globals.A3SLIDE.Slide.Shapes[CBTitleKey.SelectedItem];
            shape.Name = "TITLE";
            shape.Title = "TITLE";
        }

        // CHAP:SUB Functions
        private void InitializeChapSub() {
            try
            {
                CBChapSubKey.Items.Clear();
                TBChapSubValue.Clear();

                foreach (string shapeName in A3Globals.A3SLIDE.ShapeNames) {
                    CBChapSubKey.Items.Add(shapeName);
                }

                int index = A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "CHAP:SUB") > 0 ? A3Globals.A3SLIDE.ShapeNames.FindIndex(s => s == "CHAP:SUB") : 0;
                CBChapSubKey.SelectedIndex = index;

                try {
                    TBChapSubValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBChapSubKey.SelectedItem].TextFrame.TextRange.Text;
                }
                catch { }
            }
            catch
            {

            }

        }
        private void CBChapSubKey_SelectedIndexChanged(object sender, EventArgs e) {
            TBChapSubValue.Text = A3Globals.A3SLIDE.Slide.Shapes[CBChapSubKey.SelectedItem].TextFrame.TextRange.Text;
        }
        private void SaveChapSub() {
            A3Globals.A3SLIDE.ChapSub = TBChapSubValue.Text;
            PowerPoint.Shape shape = A3Globals.A3SLIDE.Slide.Shapes[CBChapSubKey.SelectedItem];
            shape.Name = "CHAP:SUB";
            shape.Title = "CHAP:SUB";
        }

        private void Save() {
            if (CBType.Text == "" || !CBType.Items.Contains(CBType.Text)) {
                MessageBox.Show("YOU MUST SELECT A VALID TYPE!", "SELECT VALID TYPE!", MessageBoxButtons.OK);
                return;
            }
            if (CBTitleKey.Text == CBChapSubKey.Text) {
                MessageBox.Show("CHAP:SUB AND TITLE MAY NOT BE THE SAME OBJECT", "DUPLICATE ITEM", MessageBoxButtons.OK);
                return;
            }
            if (CBTitleKey.Text == null || CBTitleKey.Text == "") {
                MessageBox.Show("MUST SELECT A TITLE", "NULL TITLE", MessageBoxButtons.OK);
                return;
            }
            this.SaveType();
            this.SaveTitle();
            if (CBType.Text == "Course Title Card" || CBType.Text == "Chapter Title Card")
            { }
            else {
                this.SaveChapSub();
            }
            this.SaveGuids();
            A3Globals.A3SLIDE.WriteFromMemory();
        }

        private void BtnSave_Click(object sender, EventArgs e) {
            this.Save();
            A3Globals.A3SLIDE.ReadFromSlide();
            this.DrawSlideInfo();
        }
        private void BtnSaveAndProceed_Click(object sender, EventArgs e) {
            this.Save();
            this.Close();
            A3Globals.QUIT_FROM_CURRENT_LOOP = false;
        }
        
        private void BtnPreviousSlide_Click(object sender, EventArgs e) {
            int slideIndex = A3Globals.A3SLIDE.Slide.SlideIndex - 1;
            try {
                A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]);
                A3Globals.A3SLIDE.ReadShapes();
                this.DrawSlideInfo();
            }
            catch {
                MessageBox.Show("BEGINING OF SLIDE SHOW", "ERROR", MessageBoxButtons.OK);
            }
        }
        private void BtnNextSlide_Click(object sender, EventArgs e)
        {
            int slideIndex = A3Globals.A3SLIDE.Slide.SlideIndex + 1;
            try {
                A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]);
                A3Globals.A3SLIDE.ReadShapes();
                this.DrawSlideInfo();
            }
            catch {
                MessageBox.Show("END OF SLIDE SHOW", "ERROR", MessageBoxButtons.OK);
            }
        }

        #region TODO: IMPLEMENT
        private void BtnNewChapSub_Click(object sender, EventArgs e)
        {
            MessageBox.Show("NOT IMPLEMENTED AT THIS TIME", "ERROR", MessageBoxButtons.OK);
        }
        private void BtnSwapChapSub_Click(object sender, EventArgs e)
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
