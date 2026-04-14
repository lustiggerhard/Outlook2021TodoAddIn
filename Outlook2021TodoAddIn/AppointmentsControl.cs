using Outlook2021TodoAddIn.Forms;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook2021TodoAddIn
{
    public partial class AppointmentsControl : UserControl
    {
        // ══════════════════════════════════════════════════════════════════
        // Hardcodierte Einstellungen (kein Config-Dialog mehr)
        // ══════════════════════════════════════════════════════════════════

        private const decimal  CFG_NUM_DAYS             = 90;
        private const bool     CFG_SHOW_COMPLETED_TASKS = false;
        private const bool     CFG_SHOW_PAST_APPTS      = false;
        private const bool     CFG_SHOW_DAY_NAMES       = true;
        private const bool     CFG_FRIENDLY_HEADERS     = true;
        private const bool     CFG_SHOW_WEEK_NUMBERS    = true;

        // ══════════════════════════════════════════════════════════════════
        // Felder
        // ══════════════════════════════════════════════════════════════════

        private DateTime                   _selectedDate   = DateTime.Today;
        private DateTime                   _calendarMonth;
        private HashSet<DateTime>          _boldedDates    = new HashSet<DateTime>();
        private FlowLayoutPanel            _flpAppointments;
        private FlowLayoutPanel            _flpTasks;
        private System.Windows.Forms.Timer _resizeTimer;
        private ToolTip                    _toolTip;
        private Outlook.AppointmentItem    _contextMenuAppt = null;
        private OLTaskItem                 _contextMenuTask = null;
        private bool                       _splitterLoaded  = false;
        private bool                       _splitterReady   = false;

        // Schriftgrößen (pt)
        private const float FS_SMALL  = 8.0f;
        private const float FS_NORMAL = 8.5f;
        private const float FS_BOLD   = 8.5f;

        // Spaltenbreiten
        private const int COL_TIME = 65;
        private const int COL_BAR  = 7;

        // ══════════════════════════════════════════════════════════════════
        // Properties (nur noch SelectedDate und Accounts von außen nutzbar)
        // ══════════════════════════════════════════════════════════════════

        public StringCollection Accounts { get; set; }

        /// <summary>Aufgabenliste ein/ausblenden (auch per Ribbon-Button togglebar)</summary>
        public bool ShowTasks
        {
            get { return !splitContainer1.Panel2Collapsed; }
            set
            {
                splitContainer1.Panel2Collapsed = !value;
                if (value) RetrieveTasks();
            }
        }

        public DateTime SelectedDate
        {
            get { return _selectedDate; }
            set { _selectedDate = value; _calendarMonth = new DateTime(value.Year, value.Month, 1); }
        }

        // ══════════════════════════════════════════════════════════════════
        // Konstruktor
        // ══════════════════════════════════════════════════════════════════

        public AppointmentsControl()
        {
            InitializeComponent();
            _calendarMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

            // FlowLayoutPanel Termine
            _flpAppointments = BuildFlowPanel();
            pnlAppointments.Controls.Add(_flpAppointments);

            // FlowLayoutPanel Aufgaben
            _flpTasks = BuildFlowPanel();
            pnlTasks.Controls.Add(_flpTasks);

            _toolTip = new ToolTip { AutoPopDelay = 8000 };

            // Resize-Debounce
            _resizeTimer       = new System.Windows.Forms.Timer { Interval = 300 };
            _resizeTimer.Tick += (s, e) => { _resizeTimer.Stop(); if (IsHandleCreated) RetrieveData(); };
            pnlAppointments.Resize += (s, e) => { _resizeTimer.Stop(); _resizeTimer.Start(); };
            pnlTasks.Resize        += (s, e) => { _resizeTimer.Stop(); _resizeTimer.Start(); };

            // Accounts aus Settings laden
            Accounts = Properties.Settings.Default.Accounts;

            try
            {
                int sd = Properties.Settings.Default.SplitterDistance;
                if (sd > splitContainer1.Panel1MinSize)
                    splitContainer1.SplitterDistance = sd;
            }
            catch { }
        }

        private static FlowLayoutPanel BuildFlowPanel()
        {
            return new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents  = false,
                AutoSize      = true,
                AutoSizeMode  = AutoSizeMode.GrowAndShrink,
                Padding       = new Padding(0),
                Margin        = new Padding(0),
                BackColor     = SystemColors.Window
            };
        }

        // ══════════════════════════════════════════════════════════════════
        // RetrieveData
        // ══════════════════════════════════════════════════════════════════

        public void RetrieveData()
        {
            if (!_splitterLoaded && splitContainer1.Height > 50)
            {
                _splitterLoaded = true;
                try
                {
                    int sd = Properties.Settings.Default.SplitterDistance;
                    if (sd > splitContainer1.Panel1MinSize &&
                        sd < splitContainer1.Height - splitContainer1.Panel2MinSize)
                        splitContainer1.SplitterDistance = sd;
                }
                catch { }
                _splitterReady = true;  // ab jetzt darf SplitterMoved speichern
            }

            RetrieveAppointments();

            if (ShowTasks)
            {
                splitContainer1.Panel2Collapsed = false;
                RetrieveTasks();
            }
            else
            {
                splitContainer1.Panel2Collapsed = true;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // KALENDER
        // ══════════════════════════════════════════════════════════════════

        private void BuildCalendar()
        {
            var fNav = new Font(Font.FontFamily, FS_BOLD,  FontStyle.Bold);
            var fHdr = new Font(Font.FontFamily, FS_SMALL, FontStyle.Bold);
            var fDay = new Font(Font.FontFamily, FS_NORMAL);
            var fKW  = new Font(Font.FontFamily, FS_SMALL - 0.5f);

            int rowH = fDay.Height + 5;
            panelCalendar.Height = 8 * rowH + 4;

            panelCalendar.SuspendLayout();
            foreach (Control c in panelCalendar.Controls) c.Dispose();
            panelCalendar.Controls.Clear();

            var tbl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 8, RowCount = 8,
                Padding = new Padding(2), Margin = new Padding(0),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30));
            for (int i = 0; i < 7; i++)
                tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / 7));
            for (int r = 0; r < 8; r++)
                tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, rowH));

            // Zeile 0: Navigation
            tbl.Controls.Add(CalNavBtn("<", fNav, () => { _calendarMonth = _calendarMonth.AddMonths(-1); BuildCalendar(); }), 0, 0);
            var lm = new Label { Text = _calendarMonth.ToString("MMMM yyyy"), Font = fNav,
                TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(0) };
            tbl.Controls.Add(lm, 1, 0);
            tbl.SetColumnSpan(lm, 6);
            tbl.Controls.Add(CalNavBtn(">", fNav, () => { _calendarMonth = _calendarMonth.AddMonths(1); BuildCalendar(); }), 7, 0);

            // Zeile 1: Tagesnamen
            tbl.Controls.Add(CalKWLbl("KW", fKW), 0, 1);
            string[] dn = { "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So" };
            for (int d = 0; d < 7; d++) tbl.Controls.Add(CalLbl(dn[d], fHdr, SystemColors.GrayText), d + 1, 1);

            // Zeilen 2–7: Wochen
            int dow = (int)_calendarMonth.DayOfWeek;
            DateTime ws = _calendarMonth.AddDays((dow == 0) ? -6 : -(dow - 1));

            for (int row = 2; row < 8; row++)
            {
                tbl.Controls.Add(CalKWLbl(GetWeekNumber(ws).ToString(), fKW), 0, row);
                for (int col = 1; col <= 7; col++)
                {
                    DateTime cd   = ws.AddDays(col - 1);
                    bool isCur    = cd.Month == _calendarMonth.Month;
                    bool isToday  = cd.Date == DateTime.Today;
                    bool isSel    = cd.Date == _selectedDate.Date;
                    bool isBold   = _boldedDates.Contains(cd.Date);

                    var lbl = new Label
                    {
                        Text = cd.Day.ToString(), TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill, Cursor = Cursors.Hand, Tag = cd, Margin = new Padding(0),
                        Font = isBold ? new Font(Font.FontFamily, FS_NORMAL, FontStyle.Bold) : fDay
                    };
                    if      (isSel)   { lbl.BackColor = Color.LightBlue; lbl.ForeColor = Color.DarkBlue; }
                    else if (isToday) { lbl.BackColor = Color.SteelBlue; lbl.ForeColor = Color.White; }
                    else if (!isCur)    lbl.ForeColor = Color.DarkGray;

                    lbl.Click       += CalDay_Click;
                    lbl.DoubleClick += CalDay_DblClick;
                    tbl.Controls.Add(lbl, col, row);
                }
                ws = ws.AddDays(7);
            }

            panelCalendar.Controls.Add(tbl);
            panelCalendar.ResumeLayout();
        }

        private Button CalNavBtn(string text, Font font, Action onClick)
        {
            var b = new Button { Text = text, FlatStyle = FlatStyle.Flat, Dock = DockStyle.Fill,
                Font = font, Cursor = Cursors.Hand, Margin = new Padding(0),
                BackColor = SystemColors.Window };
            b.FlatAppearance.BorderSize = 0;
            b.Click += (s, e) => onClick();
            return b;
        }

        private Label CalLbl(string text, Font font, Color fore)
            => new Label { Text = text, TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill, Font = font, ForeColor = fore, Margin = new Padding(0) };

        /// <summary>KW-Label mit rechtem Border-Strich</summary>
        private Label CalKWLbl(string text, Font font)
        {
            var lbl = new Label
            {
                Text      = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock      = DockStyle.Fill,
                Font      = font,
                ForeColor = SystemColors.GrayText,
                Margin    = new Padding(0)
            };
            lbl.Paint += (s, pe) =>
            {
                var c = (Control)s;
                using (var pen = new Pen(Color.LightGray))
                    pe.Graphics.DrawLine(pen, c.Width - 1, 2, c.Width - 1, c.Height - 3);
            };
            return lbl;
        }

        private static int GetWeekNumber(DateTime date)
            => CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
                   date.AddDays(3), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

        private void CalDay_Click(object sender, EventArgs e)
        {
            if (!(sender is Label l) || !(l.Tag is DateTime d)) return;
            _selectedDate = d;
            if (d.Month != _calendarMonth.Month || d.Year != _calendarMonth.Year)
                _calendarMonth = new DateTime(d.Year, d.Month, 1);
            RetrieveData();
        }

        private void CalDay_DblClick(object sender, EventArgs e)
        {
            if (!(sender is Label l) || !(l.Tag is DateTime d)) return;
            try
            {
                var f = Globals.ThisAddIn.Application.Session
                            .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
                Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder = f;
                var cv = (Outlook.CalendarView)Globals.ThisAddIn.Application.ActiveExplorer().CurrentView;
                cv.CalendarViewMode = Outlook.OlCalendarViewMode.olCalendarViewDay;
                cv.GoToDate(d);
            }
            catch { }
        }

        // ══════════════════════════════════════════════════════════════════
        // TERMINE
        // ══════════════════════════════════════════════════════════════════

        private void RetrieveAppointments()
        {
            var appts = new List<Outlook.AppointmentItem>();
            foreach (Outlook.Store store in Globals.ThisAddIn.Application.Session.Stores)
                if (Accounts == null || Accounts.Count == 0 || Accounts.Contains(store.DisplayName))
                    appts.AddRange(RetrieveAppointmentsForFolder(
                        store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder));

            appts.Sort(CompareAppointments);
            _boldedDates = new HashSet<DateTime>(appts.Select(a => a.Start.Date).Distinct());
            BuildCalendar();

            DateTime start = _selectedDate.Date;
            if (!CFG_SHOW_PAST_APPTS && start == DateTime.Today)
                start = start.Add(DateTime.Now.TimeOfDay);

            BuildItemsPanel(_flpAppointments, pnlAppointments,
                appts.Where(a => a.Start >= start && a.Start <= start.AddDays((double)CFG_NUM_DAYS))
                     .Cast<object>().ToList(),
                isTask: false);
        }

        // ══════════════════════════════════════════════════════════════════
        // AUFGABEN
        // ══════════════════════════════════════════════════════════════════

        private void RetrieveTasks()
        {
            var tasks = new List<OLTaskItem>();
            foreach (Outlook.Store store in Globals.ThisAddIn.Application.Session.Stores)
                if (Accounts == null || Accounts.Count == 0 || Accounts.Contains(store.DisplayName))
                    tasks.AddRange(RetrieveTasksForFolder(
                        store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderToDo) as Outlook.Folder));

            if (!CFG_SHOW_COMPLETED_TASKS) tasks = tasks.Where(t => !t.Completed).ToList();
            tasks.Sort(CompareTasks);

            BuildItemsPanel(_flpTasks, pnlTasks,
                tasks.Cast<object>().ToList(),
                isTask: true);
        }

        private List<OLTaskItem> RetrieveTasksForFolder(Outlook.Folder f)
        {
            var t = new List<OLTaskItem>();
            if (f == null) return t;
            foreach (object item in f.Items) try { t.Add(new OLTaskItem(item)); } catch { }
            return t.Where(x => x.ValidTaskItem).ToList();
        }

        // ══════════════════════════════════════════════════════════════════
        // GEMEINSAMES PANEL-BUILDER (Termine + Aufgaben)
        // ══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Baut FlowLayoutPanel mit Datum-Headern + Eintrags-Tabellen.
        /// items enthält entweder Outlook.AppointmentItem oder OLTaskItem.
        /// </summary>
        private void BuildItemsPanel(FlowLayoutPanel flp, Panel container,
            List<object> items, bool isTask)
        {
            flp.SuspendLayout();
            var old = flp.Controls.Cast<Control>().ToList();
            flp.Controls.Clear();
            foreach (var c in old) c.Dispose();

            if (items.Count == 0) { flp.ResumeLayout(); return; }

            int panelH = container.ClientSize.Height;
            if (panelH <= 0) panelH = 400;
            int w = Math.Max(container.ClientSize.Width - 2, 100);

            using (var fBold = new Font(Font.FontFamily, FS_BOLD, FontStyle.Bold))
            {
                int rowH    = fBold.Height + 6;
                int hdrH    = rowH + 2;
                int spacerH = 2;
                int usedH   = 0;
                panelH     -= rowH;   // -1: immer eine Zeile Puffer lassen
                int lastDay = -1, lastYear = -1;

                foreach (var item in items)
                {
                    DateTime itemDate;
                    bool hasSecondLine;

                    if (isTask)
                    {
                        var t = (OLTaskItem)item;
                        itemDate      = t.DueDate.Year == Constants.NullYear ? DateTime.MaxValue : t.DueDate.Date;
                        hasSecondLine = false;
                    }
                    else
                    {
                        var a = (Outlook.AppointmentItem)item;
                        itemDate      = a.Start.Date;
                        hasSecondLine = !string.IsNullOrEmpty(a.Location);
                    }

                    bool newDay = itemDate == DateTime.MaxValue
                        ? (lastDay != -2)
                        : (itemDate.Day != lastDay || itemDate.Year != lastYear);

                    int needed = (newDay ? hdrH : 0) + (hasSecondLine ? 2 : 1) * rowH + spacerH;
                    if (usedH + needed > panelH) break;

                    if (newDay)
                    {
                        if (itemDate == DateTime.MaxValue)
                        {
                            lastDay = -2; lastYear = -2;
                            flp.Controls.Add(BuildGroupHeader("Kein Fälligkeitsdatum", w, hdrH));
                        }
                        else
                        {
                            lastDay = itemDate.Day; lastYear = itemDate.Year;
                            flp.Controls.Add(BuildGroupHeader(FormatDateHeader(itemDate), w, hdrH));
                        }
                        usedH += hdrH;
                    }

                    if (isTask)
                        flp.Controls.Add(BuildTaskEntry((OLTaskItem)item, w, rowH));
                    else
                        flp.Controls.Add(BuildAppointmentEntry((Outlook.AppointmentItem)item, w, rowH));

                    usedH += (hasSecondLine ? 2 : 1) * rowH;

                    flp.Controls.Add(new Panel { Height = spacerH, Width = w, BackColor = SystemColors.Window });
                    usedH += spacerH;
                }
            }

            flp.ResumeLayout();
        }

        // ── Datum/Gruppen-Header ──────────────────────────────────────────

        private string FormatDateHeader(DateTime date)
        {
            int diff = (int)(date - DateTime.Today).TotalDays;
            string prefix = diff == -1 ? Constants.Yesterday + ":  " :
                            diff ==  0 ? Constants.Today     + ":  " :
                            diff ==  1 ? Constants.Tomorrow  + ":  " : "";
            string text = date.ToShortDateString();
            if (CFG_SHOW_DAY_NAMES) text += "  (" + date.ToString("dddd") + ")";
            return prefix + text;
        }

        private Label BuildGroupHeader(string text, int width, int height)
            => new Label
            {
                Text      = text,
                Font      = new Font(Font.FontFamily, FS_BOLD, FontStyle.Bold),
                BackColor = SystemColors.Window,
                ForeColor = Color.FromArgb(40, 60, 100),
                Width     = width,
                Height    = height,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding   = new Padding(4, 0, 0, 0),
                Margin    = new Padding(0)
            };

        // ── Termin-Eintrag ────────────────────────────────────────────────

        /// <summary>
        /// Type 1 (mit Ort): Zeile A = Zeit|Balken(rowspan2)|Titel  /  Zeile B = leer|–|Ort
        /// Type 2 (kein Ort): Zeile A = Zeit|Balken|Titel
        /// </summary>
        private TableLayoutPanel BuildAppointmentEntry(
            Outlook.AppointmentItem appt, int width, int rowH)
        {
            bool  hasLoc  = !string.IsNullOrEmpty(appt.Location);
            Color barColor = GetApptBarColor(appt);
            Color tint     = Color.FromArgb(30, barColor.R, barColor.G, barColor.B);
            var   tbl      = BuildEntryTable(width, rowH, hasLoc ? 2 : 1, tint);

            // Zeile 0
            tbl.Controls.Add(new Label
            {
                Text = appt.AllDayEvent ? "" : appt.Start.ToShortTimeString(),
                Font = new Font(Font.FontFamily, FS_SMALL),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill, AutoEllipsis = false,
                Padding = new Padding(15, 0, 2, 0), Margin = new Padding(0)
            }, 0, 0);

            var bar = new Panel { BackColor = barColor, Dock = DockStyle.Fill, Margin = new Padding(0) };
            tbl.Controls.Add(bar, 1, 0);
            if (hasLoc) tbl.SetRowSpan(bar, 2);

            tbl.Controls.Add(new Label
            {
                Text = appt.Subject ?? "", Font = new Font(Font.FontFamily, FS_BOLD, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, AutoEllipsis = true,
                Padding = new Padding(4, 0, 2, 0), Margin = new Padding(0)
            }, 2, 0);

            // Zeile 1 (nur Type 1)
            if (hasLoc)
            {
                tbl.Controls.Add(new Label { Margin = new Padding(0) }, 0, 1);
                tbl.Controls.Add(new Label
                {
                    Text = appt.Location, Font = new Font(Font.FontFamily, FS_SMALL),
                    TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, AutoEllipsis = true,
                    ForeColor = SystemColors.GrayText, Padding = new Padding(4, 0, 2, 0), Margin = new Padding(0)
                }, 2, 1);
            }

            TintChildren(tbl);
            AttachEvents(tbl, appt, null, BuildApptTooltip(appt));
            return tbl;
        }

        // ── Aufgaben-Eintrag (gleiches Template wie Termin Type 2) ────────

        private TableLayoutPanel BuildTaskEntry(OLTaskItem task, int width, int rowH)
        {
            Color barColor = GetTaskBarColor(task);
            Color tint     = Color.FromArgb(30, barColor.R, barColor.G, barColor.B);
            var   tbl      = BuildEntryTable(width, rowH, 1, tint);

            // Zeile 0: Fälligkeit | Balken | Betreff
            string timeText = task.DueDate.Year == Constants.NullYear
                ? "" : task.DueDate.ToShortDateString();

            tbl.Controls.Add(new Label
            {
                Text = timeText, Font = new Font(Font.FontFamily, FS_SMALL),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill, AutoEllipsis = false,
                Padding = new Padding(4, 0, 2, 0), Margin = new Padding(0),
                ForeColor = SystemColors.GrayText
            }, 0, 0);

            var bar = new Panel { BackColor = barColor, Dock = DockStyle.Fill, Margin = new Padding(0) };
            tbl.Controls.Add(bar, 1, 0);

            var subjectFont = task.Completed
                ? new Font(Font.FontFamily, FS_BOLD, FontStyle.Bold | FontStyle.Strikeout)
                : new Font(Font.FontFamily, FS_BOLD, FontStyle.Bold);

            tbl.Controls.Add(new Label
            {
                Text = task.TaskSubject ?? "", Font = subjectFont,
                TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, AutoEllipsis = true,
                Padding = new Padding(4, 0, 2, 0), Margin = new Padding(0)
            }, 2, 0);

            TintChildren(tbl);
            AttachEvents(tbl, null, task, BuildTaskTooltip(task));
            return tbl;
        }

        // ── Entry-Tabelle (shared) ────────────────────────────────────────

        private TableLayoutPanel BuildEntryTable(int width, int rowH, int rows, Color tint)
        {
            var tbl = new TableLayoutPanel
            {
                ColumnCount = 3, RowCount = rows, Width = width, Height = rows * rowH,
                Padding = new Padding(0), Margin = new Padding(0),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                BackColor = tint
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, COL_TIME));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, COL_BAR));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int r = 0; r < rows; r++)
                tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, rowH));
            return tbl;
        }

        /// <summary>Labels transparent färben außer Spalte 0 (Zeit/Datum) und Panel (Balken)</summary>
        private static void TintChildren(TableLayoutPanel tbl)
        {
            foreach (Control c in tbl.Controls)
            {
                if (c is Panel) continue;           // Farbbalken — nicht ändern
                if (tbl.GetColumn(c) == 0)
                    c.BackColor = SystemColors.Window;  // Zeit-/Datumzelle bleibt weiß
                else
                    c.BackColor = Color.Transparent;
            }
        }

        // ── Events anhängen ───────────────────────────────────────────────

        private void AttachEvents(Control ctrl,
            Outlook.AppointmentItem appt, OLTaskItem task, string tooltip)
        {
            _toolTip.SetToolTip(ctrl, tooltip);

            if (appt != null)
            {
                ctrl.DoubleClick += (s, e) => OpenAppt(appt);
                ctrl.MouseUp     += (s, e) =>
                {
                    if (((MouseEventArgs)e).Button == MouseButtons.Right)
                    { _contextMenuAppt = appt; ctxMenuAppointments.Show(ctrl, ((MouseEventArgs)e).Location); }
                };
            }
            else if (task != null)
            {
                ctrl.DoubleClick += (s, e) => OpenTask(task);
                ctrl.MouseUp     += (s, e) =>
                {
                    if (((MouseEventArgs)e).Button == MouseButtons.Right)
                    { _contextMenuTask = task; ctxMenuTasks.Show(ctrl, ((MouseEventArgs)e).Location); }
                };
            }

            foreach (Control child in ctrl.Controls)
                AttachEvents(child, appt, task, tooltip);
        }

        // ── Tooltips ─────────────────────────────────────────────────────

        private string BuildApptTooltip(Outlook.AppointmentItem appt)
        {
            string s = $"{appt.Start.ToShortTimeString()} – {appt.End.ToShortTimeString()}  {appt.Subject}";
            if (!string.IsNullOrEmpty(appt.Location)) s += "\nOrt: " + appt.Location;
            if (!string.IsNullOrEmpty(appt.Categories))
                foreach (string cat in appt.Categories.Split(','))
                {
                    var c = Globals.ThisAddIn.Application.Session.Categories[cat.Trim()] as Outlook.Category;
                    if (c != null) s += "\n – " + c.Name;
                }
            return s;
        }

        private string BuildTaskTooltip(OLTaskItem task)
        {
            string s = task.TaskSubject ?? "";
            if (task.DueDate.Year  != Constants.NullYear) s += "\nFällig: " + task.DueDate.ToShortDateString();
            if (task.StartDate.Year != Constants.NullYear) s += "\nStart: "  + task.StartDate.ToShortDateString();
            if (task.Reminder.Year  != Constants.NullYear) s += "\nErinnerung: " + task.Reminder;
            s += "\nOrdner: " + task.FolderName;
            return s;
        }

        // ── Farben ────────────────────────────────────────────────────────

        private Color GetApptBarColor(Outlook.AppointmentItem appt)
        {
            if (!string.IsNullOrEmpty(appt.Categories))
            {
                var c = Globals.ThisAddIn.Application.Session
                            .Categories[appt.Categories.Split(',')[0].Trim()] as Outlook.Category;
                if (c != null) return TranslateCategoryColor(c.Color);
            }
            switch (appt.BusyStatus)
            {
                case Outlook.OlBusyStatus.olBusy:             return Color.SteelBlue;
                case Outlook.OlBusyStatus.olOutOfOffice:      return Color.MediumPurple;
                case Outlook.OlBusyStatus.olTentative:        return Color.LightSteelBlue;
                case Outlook.OlBusyStatus.olWorkingElsewhere: return Color.LightSlateGray;
                default:                                      return Color.SteelBlue;
            }
        }

        private Color GetTaskBarColor(OLTaskItem task)
        {
            foreach (string cat in task.Categories)
            {
                var c = Globals.ThisAddIn.Application.Session.Categories[cat] as Outlook.Category;
                if (c != null) return TranslateCategoryColor(c.Color);
            }
            return Color.SteelBlue;
        }

        // ══════════════════════════════════════════════════════════════════
        // Termin öffnen / Context-Menu Termine
        // ══════════════════════════════════════════════════════════════════

        private void OpenAppt(Outlook.AppointmentItem appt)
        {
            if (appt == null) return;
            if (appt.IsRecurring)
            {
                var f = new FormRecurringOpen
                {
                    Title   = "Open Recurring Item",
                    Message = "This is one appointment in a series. What do you want to open?"
                };
                if (f.ShowDialog() == DialogResult.OK)
                { if (f.OpenRecurring) ((Outlook.AppointmentItem)appt.Parent).Display(true); else appt.Display(true); }
            }
            else appt.Display(true);
            RetrieveAppointments();
        }

        private void mnuItemReplyAllEmail_Click(object sender, EventArgs e)
        {
            if (_contextMenuAppt == null) return;
            var mail = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            string cur = OutlookHelper.GetEmailAddress(Globals.ThisAddIn.Application.Session.CurrentUser);
            foreach (Outlook.Recipient r in _contextMenuAppt.Recipients)
            { string a = OutlookHelper.GetEmailAddress(r); if (cur != a) mail.Recipients.Add(a); }
            mail.Body = "\n\n" + _contextMenuAppt.Body;
            mail.Subject = Constants.SubjectRE + ": " + _contextMenuAppt.Subject;
            mail.Display();
        }

        private void mnuItemDeleteAppointment_Click(object sender, EventArgs e)
        {
            if (_contextMenuAppt == null) return;
            if (_contextMenuAppt.IsRecurring)
            {
                var f = new FormRecurringOpen { Title = "Warning: Delete Recurring Item",
                    Message = "This is one appointment in a series. What do you want to delete?" };
                if (f.ShowDialog() == DialogResult.OK)
                { if (f.OpenRecurring) ((Outlook.AppointmentItem)_contextMenuAppt.Parent).Delete(); else _contextMenuAppt.Delete(); }
            }
            else if (MessageBox.Show("Termin wirklich löschen?",
                         "Termin löschen", MessageBoxButtons.YesNo) == DialogResult.Yes)
                _contextMenuAppt.Delete();
            RetrieveAppointments();
        }

        // ══════════════════════════════════════════════════════════════════
        // Aufgabe öffnen / Context-Menu Aufgaben
        // ══════════════════════════════════════════════════════════════════

        private void OpenTask(OLTaskItem task)
        {
            if (task == null) return;
            if      (task.OriginalItem is Outlook.MailItem m)    m.Display(true);
            else if (task.OriginalItem is Outlook.ContactItem co) co.Display(true);
            else if (task.OriginalItem is Outlook.TaskItem ti)   ti.Display(true);
            RetrieveTasks();
        }

        private void mnuItemMarkComplete_Click(object sender, EventArgs e)
        {
            if (_contextMenuTask == null || _contextMenuTask.Completed) return;
            if (MessageBox.Show("Aufgabe als erledigt markieren?",
                    "Erledigt", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            if      (_contextMenuTask.OriginalItem is Outlook.MailItem m)    m.ClearTaskFlag();
            else if (_contextMenuTask.OriginalItem is Outlook.ContactItem co) co.ClearTaskFlag();
            else if (_contextMenuTask.OriginalItem is Outlook.TaskItem ti)   ti.MarkComplete();
            RetrieveTasks();
        }

        private void mnuItemDeleteTask_Click(object sender, EventArgs e)
        {
            if (_contextMenuTask == null) return;
            if (MessageBox.Show("Aufgabe wirklich löschen?",
                    "Aufgabe löschen", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            if      (_contextMenuTask.OriginalItem is Outlook.MailItem m)    m.Delete();
            else if (_contextMenuTask.OriginalItem is Outlook.ContactItem co) co.Delete();
            else if (_contextMenuTask.OriginalItem is Outlook.TaskItem ti)   ti.Delete();
            RetrieveTasks();
        }

        // ══════════════════════════════════════════════════════════════════
        // Splitter
        // ══════════════════════════════════════════════════════════════════

        private void splitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {
            if (!_splitterReady) return;
            Properties.Settings.Default.SplitterDistance = splitContainer1.SplitterDistance;
            Properties.Settings.Default.Save();
        }

        // ══════════════════════════════════════════════════════════════════
        // Outlook-Datenabruf
        // ══════════════════════════════════════════════════════════════════

        private List<Outlook.AppointmentItem> RetrieveAppointmentsForFolder(Outlook.Folder cal)
        {
            var start = new DateTime(_selectedDate.Year, _selectedDate.Month, 1);
            var end   = start.AddMonths(1).AddDays(-1).AddDays((double)CFG_NUM_DAYS);
            var range = GetAppointmentsInRange(cal, start, end);
            var list  = new List<Outlook.AppointmentItem>();
            if (range != null) foreach (Outlook.AppointmentItem a in range) list.Add(a);
            return list;
        }

        private Outlook.Items GetAppointmentsInRange(Outlook.Folder folder, DateTime start, DateTime end)
        {
            string f = $"[Start] >= '{start:g}' AND [End] <= '{end:g}'";
            try { var i = folder.Items; i.IncludeRecurrences = true; i.Sort("[Start]", Type.Missing); var r = i.Restrict(f); return r.Count > 0 ? r : null; }
            catch { return null; }
        }

        private static int CompareAppointments(Outlook.AppointmentItem x, Outlook.AppointmentItem y)
            => x.Start.CompareTo(y.Start);
        private static int CompareTasks(OLTaskItem x, OLTaskItem y)
            => x.DueDate.CompareTo(y.DueDate);

        // ══════════════════════════════════════════════════════════════════
        // Kategorie-Farben
        // ══════════════════════════════════════════════════════════════════

        private Color TranslateCategoryColor(Outlook.OlCategoryColor col)
        {
            switch (col)
            {
                case Outlook.OlCategoryColor.olCategoryColorRed:        return Color.Red;
                case Outlook.OlCategoryColor.olCategoryColorOrange:     return Color.Orange;
                case Outlook.OlCategoryColor.olCategoryColorPeach:      return Color.PeachPuff;
                case Outlook.OlCategoryColor.olCategoryColorYellow:     return Color.Gold;
                case Outlook.OlCategoryColor.olCategoryColorGreen:      return Color.Green;
                case Outlook.OlCategoryColor.olCategoryColorTeal:       return Color.Teal;
                case Outlook.OlCategoryColor.olCategoryColorOlive:      return Color.Olive;
                case Outlook.OlCategoryColor.olCategoryColorBlue:       return Color.Blue;
                case Outlook.OlCategoryColor.olCategoryColorPurple:     return Color.Purple;
                case Outlook.OlCategoryColor.olCategoryColorMaroon:     return Color.Maroon;
                case Outlook.OlCategoryColor.olCategoryColorSteel:      return Color.LightSteelBlue;
                case Outlook.OlCategoryColor.olCategoryColorDarkSteel:  return Color.SteelBlue;
                case Outlook.OlCategoryColor.olCategoryColorGray:       return Color.Gray;
                case Outlook.OlCategoryColor.olCategoryColorDarkGray:   return Color.DarkGray;
                case Outlook.OlCategoryColor.olCategoryColorBlack:      return Color.Black;
                case Outlook.OlCategoryColor.olCategoryColorDarkRed:    return Color.DarkRed;
                case Outlook.OlCategoryColor.olCategoryColorDarkOrange: return Color.DarkOrange;
                case Outlook.OlCategoryColor.olCategoryColorDarkPeach:  return Color.DarkSalmon;
                case Outlook.OlCategoryColor.olCategoryColorDarkYellow: return Color.DarkGoldenrod;
                case Outlook.OlCategoryColor.olCategoryColorDarkGreen:  return Color.DarkGreen;
                case Outlook.OlCategoryColor.olCategoryColorDarkTeal:   return Color.DarkCyan;
                case Outlook.OlCategoryColor.olCategoryColorDarkOlive:  return Color.DarkOliveGreen;
                case Outlook.OlCategoryColor.olCategoryColorDarkBlue:   return Color.DarkBlue;
                case Outlook.OlCategoryColor.olCategoryColorDarkPurple: return Color.DarkViolet;
                case Outlook.OlCategoryColor.olCategoryColorDarkMaroon: return Color.DarkKhaki;
                default:                                                 return Color.SteelBlue;
            }
        }
    }
}
