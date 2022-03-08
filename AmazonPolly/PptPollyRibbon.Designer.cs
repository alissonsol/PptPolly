namespace PptPolly
{
    partial class PptPollyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PptPollyRibbon()
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
            this.groupPptPolly = this.Factory.CreateRibbonGroup();
            this.buttonAddVoice = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupPptPolly.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupPptPolly);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupPptPolly
            // 
            this.groupPptPolly.Items.Add(this.buttonAddVoice);
            this.groupPptPolly.Label = "PptPolly";
            this.groupPptPolly.Name = "groupPptPolly";
            // 
            // buttonAddVoice
            // 
            this.buttonAddVoice.Label = "Add Voice";
            this.buttonAddVoice.Name = "buttonAddVoice";
            this.buttonAddVoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddVoice_Click);
            // 
            // PptPollyRibbon
            // 
            this.Name = "PptPollyRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PptPollyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupPptPolly.ResumeLayout(false);
            this.groupPptPolly.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPptPolly;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddVoice;
    }

    partial class ThisRibbonCollection
    {
        internal PptPollyRibbon PptPollyRibbon
        {
            get { return this.GetRibbon<PptPollyRibbon>(); }
        }
    }
}
