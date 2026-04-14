namespace Outlook2021TodoAddIn
{
    partial class AppointmentsControl
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();

            // ── Context Menu Termine ───────────────────────────────────────
            this.ctxMenuAppointments      = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuItemReplyAllEmail     = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuItemDeleteAppointment = new System.Windows.Forms.ToolStripMenuItem();

            // ── Context Menu Aufgaben ──────────────────────────────────────
            this.ctxMenuTasks        = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuItemMarkComplete = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuItemDeleteTask   = new System.Windows.Forms.ToolStripMenuItem();

            // ── Panels / Splitter ──────────────────────────────────────────
            this.panel1          = new System.Windows.Forms.Panel();
            this.panelCalendar   = new System.Windows.Forms.Panel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();

            // ── Termine-Panel (ersetzt ListView) ──────────────────────────
            this.pnlAppointments = new System.Windows.Forms.Panel();

            // ── Aufgaben-Panel (ersetzt ListView) ─────────────────────────
            this.pnlTasks = new System.Windows.Forms.Panel();

            this.ctxMenuAppointments.SuspendLayout();
            this.ctxMenuTasks.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();

            // ── ctxMenuAppointments ────────────────────────────────────────
            this.ctxMenuAppointments.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                this.mnuItemReplyAllEmail, this.mnuItemDeleteAppointment });
            this.ctxMenuAppointments.Name = "ctxMenuAppointments";
            this.ctxMenuAppointments.Size = new System.Drawing.Size(227, 52);

            this.mnuItemReplyAllEmail.Name   = "mnuItemReplyAllEmail";
            this.mnuItemReplyAllEmail.Size   = new System.Drawing.Size(226, 24);
            this.mnuItemReplyAllEmail.Text   = "Reply All With Email";
            this.mnuItemReplyAllEmail.Click += new System.EventHandler(this.mnuItemReplyAllEmail_Click);

            this.mnuItemDeleteAppointment.Name   = "mnuItemDeleteAppointment";
            this.mnuItemDeleteAppointment.Size   = new System.Drawing.Size(226, 24);
            this.mnuItemDeleteAppointment.Text   = "Delete Appointment";
            this.mnuItemDeleteAppointment.Click += new System.EventHandler(this.mnuItemDeleteAppointment_Click);

            // ── ctxMenuTasks ───────────────────────────────────────────────
            this.ctxMenuTasks.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                this.mnuItemMarkComplete, this.mnuItemDeleteTask });
            this.ctxMenuTasks.Name = "ctxMenuTasks";
            this.ctxMenuTasks.Size = new System.Drawing.Size(181, 52);

            this.mnuItemMarkComplete.Name   = "mnuItemMarkComplete";
            this.mnuItemMarkComplete.Size   = new System.Drawing.Size(180, 24);
            this.mnuItemMarkComplete.Text   = "Mark Complete";
            this.mnuItemMarkComplete.Click += new System.EventHandler(this.mnuItemMarkComplete_Click);

            this.mnuItemDeleteTask.Name   = "mnuItemDeleteTask";
            this.mnuItemDeleteTask.Size   = new System.Drawing.Size(180, 24);
            this.mnuItemDeleteTask.Text   = "Delete Task";
            this.mnuItemDeleteTask.Click += new System.EventHandler(this.mnuItemDeleteTask_Click);

            // ── panelCalendar (Höhe wird dynamisch gesetzt) ────────────────
            this.panelCalendar.BackColor = System.Drawing.SystemColors.Window;
            this.panelCalendar.Dock      = System.Windows.Forms.DockStyle.Top;
            this.panelCalendar.Name      = "panelCalendar";
            this.panelCalendar.TabIndex  = 7;

            // ── pnlAppointments ────────────────────────────────────────────
            this.pnlAppointments.AutoScroll  = false;
            this.pnlAppointments.BackColor   = System.Drawing.SystemColors.Window;
            this.pnlAppointments.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.pnlAppointments.Dock        = System.Windows.Forms.DockStyle.Fill;
            this.pnlAppointments.Name        = "pnlAppointments";
            this.pnlAppointments.TabIndex    = 4;

            // ── pnlTasks ───────────────────────────────────────────────────
            this.pnlTasks.AutoScroll  = false;
            this.pnlTasks.BackColor   = System.Drawing.SystemColors.Window;
            this.pnlTasks.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.pnlTasks.Dock        = System.Windows.Forms.DockStyle.Fill;
            this.pnlTasks.Name        = "pnlTasks";
            this.pnlTasks.TabIndex    = 5;

            // ── splitContainer1 ────────────────────────────────────────────
            this.splitContainer1.Dock             = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Name             = "splitContainer1";
            this.splitContainer1.Orientation      = System.Windows.Forms.Orientation.Horizontal;
            this.splitContainer1.Panel1.Controls.Add(this.pnlAppointments);
            this.splitContainer1.Panel2.Controls.Add(this.pnlTasks);
            this.splitContainer1.Panel2.Padding   = new System.Windows.Forms.Padding(0, 1, 0, 0);
            this.splitContainer1.Panel2.BackColor = System.Drawing.Color.LightGray;
            this.splitContainer1.Size             = new System.Drawing.Size(258, 567);
            this.splitContainer1.SplitterDistance = 300;
            this.splitContainer1.TabIndex         = 6;
            this.splitContainer1.SplitterMoved   += new System.Windows.Forms.SplitterEventHandler(this.splitContainer1_SplitterMoved);

            // ── panel1 ─────────────────────────────────────────────────────
            this.panel1.Controls.Add(this.splitContainer1);
            this.panel1.Controls.Add(this.panelCalendar);
            this.panel1.Dock     = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Name     = "panel1";
            this.panel1.TabIndex = 8;

            // ── AppointmentsControl ────────────────────────────────────────
            this.BackColor           = System.Drawing.SystemColors.Window;
            this.BorderStyle         = System.Windows.Forms.BorderStyle.None;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Name = "AppointmentsControl";
            this.Size = new System.Drawing.Size(258, 767);

            this.ctxMenuAppointments.ResumeLayout(false);
            this.ctxMenuTasks.ResumeLayout(false);
            this.panelCalendar.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip  ctxMenuAppointments;
        private System.Windows.Forms.ToolStripMenuItem mnuItemReplyAllEmail;
        private System.Windows.Forms.ToolStripMenuItem mnuItemDeleteAppointment;
        private System.Windows.Forms.ContextMenuStrip  ctxMenuTasks;
        private System.Windows.Forms.ToolStripMenuItem mnuItemMarkComplete;
        private System.Windows.Forms.ToolStripMenuItem mnuItemDeleteTask;
        private System.Windows.Forms.Panel             panel1;
        private System.Windows.Forms.Panel             panelCalendar;
        private System.Windows.Forms.SplitContainer    splitContainer1;
        private System.Windows.Forms.Panel             pnlAppointments;
        private System.Windows.Forms.Panel             pnlTasks;
    }
}
