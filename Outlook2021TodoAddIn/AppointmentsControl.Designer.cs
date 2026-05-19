namespace Outlook2021TodoAddIn
{
    partial class AppointmentsControl
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                DisposeCachedFonts();
            }
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

            // ── Panels ─────────────────────────────────────────────────────
            this.panel1          = new System.Windows.Forms.Panel();
            this.panelCalendar   = new System.Windows.Forms.Panel();
            this.pnlAppointments = new System.Windows.Forms.Panel();

            this.ctxMenuAppointments.SuspendLayout();
            this.panel1.SuspendLayout();
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

            // ── panelCalendar (Höhe wird dynamisch gesetzt) ────────────────
            this.panelCalendar.BackColor = System.Drawing.SystemColors.Window;
            this.panelCalendar.Dock      = System.Windows.Forms.DockStyle.Top;
            this.panelCalendar.Name      = "panelCalendar";
            this.panelCalendar.TabIndex  = 0;

            // ── pnlAppointments ────────────────────────────────────────────
            this.pnlAppointments.AutoScroll  = false;
            this.pnlAppointments.BackColor   = System.Drawing.SystemColors.Window;
            this.pnlAppointments.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.pnlAppointments.Dock        = System.Windows.Forms.DockStyle.Fill;
            this.pnlAppointments.Name        = "pnlAppointments";
            this.pnlAppointments.TabIndex    = 1;

            // ── panel1 ─────────────────────────────────────────────────────
            this.panel1.Controls.Add(this.pnlAppointments);   // Fill — zuerst hinzufügen
            this.panel1.Controls.Add(this.panelCalendar);     // Top  — danach (höhere Z-Order → zuerst gedockt)
            this.panel1.Dock     = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Name     = "panel1";
            this.panel1.TabIndex = 2;

            // ── AppointmentsControl ────────────────────────────────────────
            this.BackColor           = System.Drawing.SystemColors.Window;
            this.BorderStyle         = System.Windows.Forms.BorderStyle.None;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Name = "AppointmentsControl";
            this.Size = new System.Drawing.Size(258, 767);

            this.ctxMenuAppointments.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip  ctxMenuAppointments;
        private System.Windows.Forms.ToolStripMenuItem mnuItemReplyAllEmail;
        private System.Windows.Forms.ToolStripMenuItem mnuItemDeleteAppointment;
        private System.Windows.Forms.Panel             panel1;
        private System.Windows.Forms.Panel             panelCalendar;
        private System.Windows.Forms.Panel             pnlAppointments;
    }
}
