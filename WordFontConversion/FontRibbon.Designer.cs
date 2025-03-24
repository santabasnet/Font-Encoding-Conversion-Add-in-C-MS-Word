namespace WordFontConversion
{
    partial class FontRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FontRibbon()
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
            this.nepaliFont = this.Factory.CreateRibbonGroup();
            this.fontConverterButton = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.nepaliFont.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.nepaliFont);
            this.tab1.Label = "सायक- Nepali Language Tools";
            this.tab1.Name = "tab1";
            // 
            // nepaliFont
            // 
            this.nepaliFont.Items.Add(this.fontConverterButton);
            this.nepaliFont.Label = "Font Action: OFF";
            this.nepaliFont.Name = "nepaliFont";
            // 
            // fontConverterButton
            // 
            this.fontConverterButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.fontConverterButton.Description = "Nepali Font Conversion State";
            this.fontConverterButton.Image = global::WordFontConversion.Properties.Resources.logo;
            this.fontConverterButton.Label = "नेपाली फन्ट रूपान्तरण";
            this.fontConverterButton.Name = "fontConverterButton";
            this.fontConverterButton.ShowImage = true;
            this.fontConverterButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fontConverterButton_Click);
            // 
            // FontRibbon
            // 
            this.Name = "FontRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.FontRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.nepaliFont.ResumeLayout(false);
            this.nepaliFont.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup nepaliFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton fontConverterButton;
    }

    partial class ThisRibbonCollection
    {
        internal FontRibbon FontRibbon
        {
            get { return this.GetRibbon<FontRibbon>(); }
        }
    }
}
