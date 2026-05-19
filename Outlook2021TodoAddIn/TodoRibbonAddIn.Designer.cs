using System;
using Microsoft.Office.Tools.Ribbon;

namespace Outlook2021TodoAddIn
{
    partial class TodoRibbonAddIn : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public TodoRibbonAddIn() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tab1          = this.Factory.CreateRibbonTab();
            this.group1        = this.Factory.CreateRibbonGroup();
            this.btnToggleTodo = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();

            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name  = "tab1";

            this.group1.Items.Add(this.btnToggleTodo);
            this.group1.Label = "Group1";
            this.group1.Name  = "group1";

            this.btnToggleTodo.Label = "Task Toggle";
            this.btnToggleTodo.Name  = "btnToggleTodo";
            this.btnToggleTodo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleTodo_Click);

            this.Name       = "TodoRibbonAddIn";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load      += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TodoRibbonAddIn_Load);
            this.tab1.ResumeLayout(false);
            this.group1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab         tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup        group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnToggleTodo;
    }

    partial class ThisRibbonCollection
    {
        internal TodoRibbonAddIn TodoRibbonAddIn
        {
            get { return this.GetRibbon<TodoRibbonAddIn>(); }
        }
    }
}
