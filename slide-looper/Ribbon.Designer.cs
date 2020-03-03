namespace slide_looper
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.sections = this.Factory.CreateRibbonButton();
            this.btn_class = this.Factory.CreateRibbonButton();
            this.btn_interface = this.Factory.CreateRibbonButton();
            this.btn_associ = this.Factory.CreateRibbonButton();
            this.btn_aggr = this.Factory.CreateRibbonButton();
            this.btn_comps = this.Factory.CreateRibbonButton();
            this.btn_gener = this.Factory.CreateRibbonButton();
            this.btn_impl = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.sections);
            this.group1.Label = "task-1";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_class);
            this.group2.Items.Add(this.btn_interface);
            this.group2.Items.Add(this.btn_associ);
            this.group2.Items.Add(this.btn_aggr);
            this.group2.Items.Add(this.btn_comps);
            this.group2.Items.Add(this.btn_gener);
            this.group2.Items.Add(this.btn_impl);
            this.group2.Label = "UML";
            this.group2.Name = "group2";
            // 
            // sections
            // 
            this.sections.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sections.Image = global::slide_looper.Properties.Resources.baseline_menu_black_48dp;
            this.sections.Label = "Get Sections";
            this.sections.Name = "sections";
            this.sections.ShowImage = true;
            this.sections.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sections_Click);
            // 
            // btn_class
            // 
            this.btn_class.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_class.Image = global::slide_looper.Properties.Resources.c;
            this.btn_class.Label = "Class";
            this.btn_class.Name = "btn_class";
            this.btn_class.ShowImage = true;
            this.btn_class.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_class_Click);
            // 
            // btn_interface
            // 
            this.btn_interface.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_interface.Image = global::slide_looper.Properties.Resources.i;
            this.btn_interface.Label = "Interface";
            this.btn_interface.Name = "btn_interface";
            this.btn_interface.ShowImage = true;
            this.btn_interface.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_interface_Click);
            // 
            // btn_associ
            // 
            this.btn_associ.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_associ.Image = global::slide_looper.Properties.Resources.associatoin;
            this.btn_associ.Label = "Association";
            this.btn_associ.Name = "btn_associ";
            this.btn_associ.ShowImage = true;
            this.btn_associ.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_associ_Click);
            // 
            // btn_aggr
            // 
            this.btn_aggr.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_aggr.Image = global::slide_looper.Properties.Resources.aggregation;
            this.btn_aggr.Label = "Aggregation";
            this.btn_aggr.Name = "btn_aggr";
            this.btn_aggr.ShowImage = true;
            this.btn_aggr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_aggr_Click);
            // 
            // btn_comps
            // 
            this.btn_comps.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_comps.Image = global::slide_looper.Properties.Resources.composition;
            this.btn_comps.Label = "Composition";
            this.btn_comps.Name = "btn_comps";
            this.btn_comps.ShowImage = true;
            this.btn_comps.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_comps_Click);
            // 
            // btn_gener
            // 
            this.btn_gener.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_gener.Image = global::slide_looper.Properties.Resources.generalization;
            this.btn_gener.Label = "Generaliztion";
            this.btn_gener.Name = "btn_gener";
            this.btn_gener.ShowImage = true;
            this.btn_gener.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_gener_Click);
            // 
            // btn_impl
            // 
            this.btn_impl.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_impl.Image = global::slide_looper.Properties.Resources.implementation;
            this.btn_impl.Label = "Implementation";
            this.btn_impl.Name = "btn_impl";
            this.btn_impl.ShowImage = true;
            this.btn_impl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_impl_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sections;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_class;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_interface;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_associ;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_aggr;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_comps;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_gener;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_impl;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
