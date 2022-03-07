namespace PptTextToSpeechAzure
{
    partial class TextToSpeechAzureRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TextToSpeechAzureRibbon()
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
            this.groupTextToSpeechAzure = this.Factory.CreateRibbonGroup();
            this.buttonSlideNoteToSpeech = this.Factory.CreateRibbonButton();
            this.buttonUpdateAllSlides = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupTextToSpeechAzure.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupTextToSpeechAzure);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupTextToSpeechAzure
            // 
            this.groupTextToSpeechAzure.Items.Add(this.buttonSlideNoteToSpeech);
            this.groupTextToSpeechAzure.Items.Add(this.buttonUpdateAllSlides);
            this.groupTextToSpeechAzure.Label = "Text to Speech Azure";
            this.groupTextToSpeechAzure.Name = "groupTextToSpeechAzure";
            // 
            // buttonSlideNoteToSpeech
            // 
            this.buttonSlideNoteToSpeech.Label = "Slide Note to Speech";
            this.buttonSlideNoteToSpeech.Name = "buttonSlideNoteToSpeech";
            this.buttonSlideNoteToSpeech.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSlideNoteToSpeech_Click);
            // 
            // buttonUpdateAllSlides
            // 
            this.buttonUpdateAllSlides.Label = "Update All Slides";
            this.buttonUpdateAllSlides.Name = "buttonUpdateAllSlides";
            this.buttonUpdateAllSlides.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUpdateAllSlides_Click);
            // 
            // TextToSpeechAzureRibbon
            // 
            this.Name = "TextToSpeechAzureRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TextToSpeechAzureRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupTextToSpeechAzure.ResumeLayout(false);
            this.groupTextToSpeechAzure.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTextToSpeechAzure;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSlideNoteToSpeech;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUpdateAllSlides;
    }

    partial class ThisRibbonCollection
    {
        internal TextToSpeechAzureRibbon TextToSpeechAzureRibbon
        {
            get { return this.GetRibbon<TextToSpeechAzureRibbon>(); }
        }
    }
}
