namespace Alta3_PPA
{
    partial class A3_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public A3_Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.TabA3 = this.Factory.CreateRibbonTab();
            this.GrpInitialize = this.Factory.CreateRibbonGroup();
            this.BtnGenFromYaml = this.Factory.CreateRibbonButton();
            this.BtnInitialize = this.Factory.CreateRibbonButton();
            this.GrpActiveDev = this.Factory.CreateRibbonGroup();
            this.BtnEnvironmentSettings = this.Factory.CreateRibbonButton();
            this.BtnShowSlideMetadata = this.Factory.CreateRibbonButton();
            this.BtnFixAllMetadata = this.Factory.CreateRibbonButton();
            this.BtnNewBaseline = this.Factory.CreateRibbonButton();
            this.BtnFillSubChaps = this.Factory.CreateRibbonButton();
            this.GrpProduce = this.Factory.CreateRibbonGroup();
            this.BtnPublish = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.BtnRecord = this.Factory.CreateRibbonButton();
            this.OpenYamlForGen = new System.Windows.Forms.OpenFileDialog();
            this.TabA3.SuspendLayout();
            this.GrpInitialize.SuspendLayout();
            this.GrpActiveDev.SuspendLayout();
            this.GrpProduce.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabA3
            // 
            this.TabA3.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabA3.Groups.Add(this.GrpInitialize);
            this.TabA3.Groups.Add(this.GrpActiveDev);
            this.TabA3.Groups.Add(this.GrpProduce);
            this.TabA3.Label = "Alta3";
            this.TabA3.Name = "TabA3";
            // 
            // GrpInitialize
            // 
            this.GrpInitialize.Items.Add(this.BtnGenFromYaml);
            this.GrpInitialize.Items.Add(this.BtnInitialize);
            this.GrpInitialize.Label = "Initialize";
            this.GrpInitialize.Name = "GrpInitialize";
            // 
            // BtnGenFromYaml
            // 
            this.BtnGenFromYaml.Label = "Generate From YAML";
            this.BtnGenFromYaml.Name = "BtnGenFromYaml";
            this.BtnGenFromYaml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGenFromYaml_Click);
            // 
            // BtnInitialize
            // 
            this.BtnInitialize.Label = "";
            this.BtnInitialize.Name = "BtnInitialize";
            // 
            // GrpActiveDev
            // 
            this.GrpActiveDev.Items.Add(this.BtnEnvironmentSettings);
            this.GrpActiveDev.Items.Add(this.BtnShowSlideMetadata);
            this.GrpActiveDev.Items.Add(this.BtnFixAllMetadata);
            this.GrpActiveDev.Items.Add(this.BtnNewBaseline);
            this.GrpActiveDev.Items.Add(this.BtnFillSubChaps);
            this.GrpActiveDev.Label = "Active Development";
            this.GrpActiveDev.Name = "GrpActiveDev";
            // 
            // BtnEnvironmentSettings
            // 
            this.BtnEnvironmentSettings.Label = "Environment Settings";
            this.BtnEnvironmentSettings.Name = "BtnEnvironmentSettings";
            // 
            // BtnShowSlideMetadata
            // 
            this.BtnShowSlideMetadata.Label = "Show Slide Metadata";
            this.BtnShowSlideMetadata.Name = "BtnShowSlideMetadata";
            this.BtnShowSlideMetadata.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnShowSlideMetadata_Click);
            // 
            // BtnFixAllMetadata
            // 
            this.BtnFixAllMetadata.Label = "Fix Null Metadata";
            this.BtnFixAllMetadata.Name = "BtnFixAllMetadata";
            this.BtnFixAllMetadata.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFixAllMetadata_Click);
            // 
            // BtnNewBaseline
            // 
            this.BtnNewBaseline.Label = "New Baseline";
            this.BtnNewBaseline.Name = "BtnNewBaseline";
            this.BtnNewBaseline.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnNewBaseline_Click);
            // 
            // BtnFillSubChaps
            // 
            this.BtnFillSubChaps.Label = "Fill Subchapters";
            this.BtnFillSubChaps.Name = "BtnFillSubChaps";
            this.BtnFillSubChaps.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFillSubChaps_Click);
            // 
            // GrpProduce
            // 
            this.GrpProduce.Items.Add(this.BtnPublish);
            this.GrpProduce.Items.Add(this.button1);
            this.GrpProduce.Items.Add(this.BtnRecord);
            this.GrpProduce.Label = "Produce";
            this.GrpProduce.Name = "GrpProduce";
            // 
            // BtnPublish
            // 
            this.BtnPublish.Label = "Publish Products";
            this.BtnPublish.Name = "BtnPublish";
            this.BtnPublish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnPublish_Click);
            // 
            // button1
            // 
            this.button1.Label = "Record Slides";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // BtnRecord
            // 
            this.BtnRecord.Label = "";
            this.BtnRecord.Name = "BtnRecord";
            // 
            // OpenYamlForGen
            // 
            this.OpenYamlForGen.FileOk += new System.ComponentModel.CancelEventHandler(this.OpenYamlForGen_FileOk);
            // 
            // A3_Ribbon
            // 
            this.Name = "A3_Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.TabA3);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.A3_Ribbon_Load);
            this.TabA3.ResumeLayout(false);
            this.TabA3.PerformLayout();
            this.GrpInitialize.ResumeLayout(false);
            this.GrpInitialize.PerformLayout();
            this.GrpActiveDev.ResumeLayout(false);
            this.GrpActiveDev.PerformLayout();
            this.GrpProduce.ResumeLayout(false);
            this.GrpProduce.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabA3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GrpInitialize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnGenFromYaml;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GrpActiveDev;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GrpProduce;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnPublish;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRecord;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnEnvironmentSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnInitialize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFixAllMetadata;
        private System.Windows.Forms.OpenFileDialog OpenYamlForGen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnShowSlideMetadata;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnNewBaseline;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFillSubChaps;
    }

    partial class ThisRibbonCollection
    {
        internal A3_Ribbon A3_Ribbon
        {
            get { return this.GetRibbon<A3_Ribbon>(); }
        }
    }
}
