namespace XLToolbox
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.TabXlToolbox = this.Factory.CreateRibbonTab();
            this.GroupXlToolbox = this.Factory.CreateRibbonGroup();
            this.ButtonCheckForUpdate = this.Factory.CreateRibbonButton();
            this.ButtonTestError = this.Factory.CreateRibbonButton();
            this.TabXlToolbox.SuspendLayout();
            this.GroupXlToolbox.SuspendLayout();
            // 
            // TabXlToolbox
            // 
            this.TabXlToolbox.Groups.Add(this.GroupXlToolbox);
            resources.ApplyResources(this.TabXlToolbox, "TabXlToolbox");
            this.TabXlToolbox.Name = "TabXlToolbox";
            // 
            // GroupXlToolbox
            // 
            this.GroupXlToolbox.Items.Add(this.ButtonCheckForUpdate);
            this.GroupXlToolbox.Items.Add(this.ButtonTestError);
            resources.ApplyResources(this.GroupXlToolbox, "GroupXlToolbox");
            this.GroupXlToolbox.Name = "GroupXlToolbox";
            // 
            // ButtonCheckForUpdate
            // 
            resources.ApplyResources(this.ButtonCheckForUpdate, "ButtonCheckForUpdate");
            this.ButtonCheckForUpdate.Name = "ButtonCheckForUpdate";
            this.ButtonCheckForUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonCheckForUpdate_Click);
            // 
            // ButtonTestError
            // 
            resources.ApplyResources(this.ButtonTestError, "ButtonTestError");
            this.ButtonTestError.Name = "ButtonTestError";
            this.ButtonTestError.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonTestError_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TabXlToolbox);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.TabXlToolbox.ResumeLayout(false);
            this.TabXlToolbox.PerformLayout();
            this.GroupXlToolbox.ResumeLayout(false);
            this.GroupXlToolbox.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabXlToolbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupXlToolbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonCheckForUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonTestError;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
