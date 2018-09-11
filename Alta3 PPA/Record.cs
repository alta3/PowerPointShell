using System;
using System.Windows.Forms;

namespace Alta3_PPA {
    public partial class Record : Form {
        int count = 0;
        public Record() {
            InitializeComponent();
        }

        public void DrawSlideInfo() {
            A3Globals.A3SLIDE.Slide.Select();
            TBTitle.Text = A3Globals.A3SLIDE.Title;
            TBType.Text = A3Globals.A3SLIDE.Type;
            TBChapter.Text = A3Globals.A3SLIDE.Chapter;
            TBSubchapter.Text = A3Globals.A3SLIDE.Subchapter;
            TBIndex.Text = A3Globals.A3SLIDE.Slide.SlideIndex.ToString();
            TBGuid.Text = A3Globals.A3SLIDE.Guid;
        }

        private void SendJson() {
            string json = String.Concat(@"{",
                                        "\"Title\": \"", TBTitle.Text, "\"",
                                        "\"Type\": \"", TBType.Text, "\"",
                                        "\"Chapter\": \"", TBChapter.Text, "\"",
                                        "\"Subchapter\": \"", TBSubchapter.Text, "\"",
                                        "\"Index\": \"", TBIndex.Text, "\"",
                                        "\"GUID\": \"", TBGuid.Text, "\"",
                                        "\"Location\": \"", TBLocation.Text, "\"",
                                        @"}");
            string address = String.Concat(@"http:\\", TBFqdn.Text, @":", TBPort.Text);
            Uri uri = A3Record.ConvertToUri(address);
            A3Record.PostIt(uri, json);
        }

        private void btnBrowser_Click(object sender, EventArgs e) {
            if (fldBrowser.ShowDialog() == DialogResult.OK) {
                A3Globals.A3_PUBLISH = fldBrowser.SelectedPath;
                TBLocation.Text = A3Globals.A3_PUBLISH;
            }
        }

        private void btnNextSlide_Click(object sender, EventArgs e) {
            int slideIndex = A3Globals.A3SLIDE.Slide.SlideIndex + 1;
            try {
                A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]);
                this.DrawSlideInfo();
            }
            catch {
                MessageBox.Show("END OF SLIDE SHOW", "ERROR", MessageBoxButtons.OK);
            }
        }
        private void btnPrevious_Click(object sender, EventArgs e) {
            int slideIndex = A3Globals.A3SLIDE.Slide.SlideIndex - 1;
            try {
                A3Slide.SetA3SlideFromPPTSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIndex]);
                this.DrawSlideInfo();
            }
            catch {
                MessageBox.Show("BEGINING OF SLIDE SHOW", "ERROR", MessageBoxButtons.OK);
            }
        }

        private void btnStartStop_Click(object sender, EventArgs e) {
            this.SendJson();
            if (this.count == 0) {
                recordTime.Start();
            }
            else {
                this.count = 0;
                recordTime.Stop();
            }
        }

        private void recordTime_Tick(object sender, EventArgs e) {
            TBTime.Text = String.Concat(count.ToString(), @" seconds");
            count++;
        }
    }
}
