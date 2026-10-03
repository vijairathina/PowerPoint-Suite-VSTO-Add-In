namespace OneAssist
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.groupQuality = this.Factory.CreateRibbonGroup();
            this.btnDesignAssistant = this.Factory.CreateRibbonButton();
            this.btnQuickScan = this.Factory.CreateRibbonButton();
            this.groupCompare = this.Factory.CreateRibbonGroup();
            this.btnSlideCompare = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupQuality.SuspendLayout();
            this.groupCompare.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupQuality);
            this.tab1.Groups.Add(this.groupCompare);
            this.tab1.Label = "OneAssist";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Picture Tools";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Image Editor";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "PictureSharpenSoftenGallery";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // groupQuality
            // 
            this.groupQuality.Items.Add(this.btnDesignAssistant);
            this.groupQuality.Items.Add(this.btnQuickScan);
            this.groupQuality.Label = "Design & Quality";
            this.groupQuality.Name = "groupQuality";
            // 
            // btnDesignAssistant
            // 
            this.btnDesignAssistant.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDesignAssistant.Label = "Design Assistant";
            this.btnDesignAssistant.Name = "btnDesignAssistant";
            this.btnDesignAssistant.OfficeImageId = "SmartArtChangeColorsGallery";
            this.btnDesignAssistant.ShowImage = true;
            this.btnDesignAssistant.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDesignAssistant_Click);
            // 
            // btnQuickScan
            // 
            this.btnQuickScan.Label = "Quick Scan";
            this.btnQuickScan.Name = "btnQuickScan";
            this.btnQuickScan.OfficeImageId = "QueryShowTable";
            this.btnQuickScan.ShowImage = true;
            this.btnQuickScan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQuickScan_Click);
            // 
            // groupCompare
            // 
            this.groupCompare.Items.Add(this.btnSlideCompare);
            this.groupCompare.Label = "Compare";
            this.groupCompare.Name = "groupCompare";
            // 
            // btnSlideCompare
            // 
            this.btnSlideCompare.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSlideCompare.Label = "Slide Compare";
            this.btnSlideCompare.Name = "btnSlideCompare";
            this.btnSlideCompare.OfficeImageId = "CompareMergePresentations";
            this.btnSlideCompare.ShowImage = true;
            this.btnSlideCompare.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSlideCompare_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupQuality.ResumeLayout(false);
            this.groupQuality.PerformLayout();
            this.groupCompare.ResumeLayout(false);
            this.groupCompare.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupQuality;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDesignAssistant;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQuickScan;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCompare;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSlideCompare;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
