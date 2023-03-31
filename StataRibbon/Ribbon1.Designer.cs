namespace StataRibbon
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
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.makeDoFile = this.Factory.CreateRibbonButton();
            this.browseButton = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.editConfiguration = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.runStataLinesButton = this.Factory.CreateRibbonButton();
            this.runStata = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Stata";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.editBox1);
            this.group1.Items.Add(this.editBox2);
            this.group1.Items.Add(this.makeDoFile);
            this.group1.Items.Add(this.browseButton);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.editConfiguration);
            this.group1.Label = "Do File";
            this.group1.Name = "group1";
            // 
            // editBox1
            // 
            this.editBox1.Label = ".do Folder Path";
            this.editBox1.Name = "editBox1";
            this.editBox1.SizeString = "WWWWWWWWWWWWW";
            this.editBox1.Text = null;
            // 
            // editBox2
            // 
            this.editBox2.Label = ".do File Name";
            this.editBox2.Name = "editBox2";
            this.editBox2.SizeString = "WWWWWWWWWWWWW";
            this.editBox2.Text = null;
            // 
            // makeDoFile
            // 
            this.makeDoFile.Label = "Make .do File";
            this.makeDoFile.Name = "makeDoFile";
            this.makeDoFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.makeDoFile_Click);
            // 
            // browseButton
            // 
            this.browseButton.Label = "Browse";
            this.browseButton.Name = "browseButton";
            this.browseButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.browseButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // editConfiguration
            // 
            this.editConfiguration.Label = "Edit Configuration";
            this.editConfiguration.Name = "editConfiguration";
            this.editConfiguration.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editConfiguration_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.runStataLinesButton);
            this.group2.Items.Add(this.runStata);
            this.group2.Label = "Stata";
            this.group2.Name = "group2";
            // 
            // runStataLinesButton
            // 
            this.runStataLinesButton.Label = "Run Stata Lines";
            this.runStataLinesButton.Name = "runStataLinesButton";
            this.runStataLinesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runStataLinesButton_Click);
            // 
            // runStata
            // 
            this.runStata.Label = "Run Stata";
            this.runStata.Name = "runStata";
            this.runStata.SuperTip = "Runs Column C starting from row 2";
            this.runStata.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runStata_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.label1);
            this.group3.Label = "Error Messages";
            this.group3.Name = "group3";
            // 
            // label1
            // 
            this.label1.Label = "Errors: ";
            this.label1.Name = "label1";
            // 
            // button1
            // 
            this.button1.Label = "Make GLD.do File";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.makeGLDFile);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton editConfiguration;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton browseButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runStataLinesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runStata;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton makeDoFile;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        protected internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
