/**
 * @file    AppointmentsControl.cs
 * @brief   UserControl: Monatskalender + Terminliste.
 * @author  Gerhard Lustig <gerhard@lustig.at>
 * @version 2.4.0
 * @date    2026-06-18
 * @history
 *   2.4.0  2026-06-18  Termin-Zeilen-Hintergrund = aufgehellte Kategoriefarbe
 *                      (Lighten-Helper, amount 0.72). Termine ohne Kategorie bleiben
 *                      neutral weiß. Balken behält volle Farbe.
 *   2.3.0  2026-06-18  Terminliste-Hintergrund über _listBg steuerbar (aktuell Darken 1.0 =
 *                      SystemColors.Window). Kategorie-Balken unverändert.
 *   2.2.0  2026-05-19  Kalender-Controls einmal erstellen (InitCalendarControls), danach
 *                      nur noch Properties updaten — kein GDI-Handle-Create/Destroy mehr
 *                      pro Rebuild (~55 Controls). BuildCalendar() ist jetzt reine
 *                      Property-Update-Schleife ohne Layout-Overhead.
 *   2.1.0  2026-05-19  Appointment-Cache: COM-Abruf nur bei Monatswechsel oder expliziter
 *                      Invalidierung — Tages-Klicks ohne Outlook-Roundtrip. BuildCalendar()
 *                      nur wenn Monat wechselt oder Cache neu geladen wurde. Neu:
 *                      InvalidateAndRefresh() für externe Trigger (ItemAdd/Change/Remove,
 *                      Sleep/Wake, Tageswechsel).
 *   2.0.0  2026-05-19  Task-Panel + alle CFG_-Konstanten entfernt; Font-Cache statt
 *                      GDI-Leak per Rebuild; DASL-Filter locale-fix (MM/dd/yyyy invariant);
 *                      toten tint-Parameter entfernt; BuildGroupHeader ohne §-Hack (Tuple);
 *                      BuildItemsPanel direkt auf List<AppointmentItem> getypt;
 *                      SplitContainer + Splitter-Persistenz entfernt.
 *   1.3.3  2026-05-19  CalDay_DblClick: neuen Termin für angeklickten Tag anlegen.
 *   1.3.2  2026-05-19  200ms-Timer nach CurrentFolder-Switch für DblClick.
 *   1.3.1  2026-05-19  Click-Timer verhindert dass Rebuild den DoubleClick verschluckt.
 *   1.3.0  2026-05-19  "Heute"-Button (⌂) in Navigationszeile.
 *   1.2.0  2026-05-05  COL_BAR von 7 auf 10 px erhöht.
 *   1.1.0  2026-05-03  Build-first-then-swap im Kalender.
 *   1.0.0  2026-04-24  Initial release.
 */
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
        // Felder
        // ══════════════════════════════════════════════════════════════════

        private DateTime                   _selectedDate    = DateTime.Today;
        private DateTime                   _calendarMonth;
        private HashSet<DateTime>          _boldedDates     = new HashSet<DateTime>();
        private FlowLayoutPanel            _flpAppointments;
        private System.Windows.Forms.Timer _resizeTimer;
        private System.Windows.Forms.Timer _dayChangeTimer;
        private DateTime                   _lastKnownDate   = DateTime.Today;
        private ToolTip                    _toolTip;
        private Outlook.AppointmentItem    _contextMenuAppt = null;
        private System.Windows.Forms.Timer _clickTimer;
        private DateTime                   _pendingClickDate;

        // Kalender-Control-Cache — einmal erstellt, danach nur Properties updaten
        private Label  _calMonthLbl;
        private Button _calNavHome;
        private Label[] _calKWLabels;   // [0]="KW"-Header, [1..6]=Wochennummer-Rows
        private Label[] _calDayCells;   // [0..41] = 6 Wochen × 7 Tage, row-major

        // Appointment-Cache — COM-Abruf nur bei Monatswechsel oder expliziter Invalidierung
        private List<Outlook.AppointmentItem> _cachedAppts   = null;
        private int                           _cacheYear     = -1;
        private int                           _cacheMonth    = -1;
        private DateTime                      _calendarBuilt = DateTime.MinValue;

        // Spaltenbreiten (px)
        private const int COL_TIME = 65;
        private const int COL_BAR  = 10;

        // Font-Cache — einmal erstellt, wiederverwendet; Dispose via DisposeCachedFonts()
        private Font _fontBold;    // 8.5pt Bold    — Nav-Buttons, Tages-Header, fette Kalendertage
        private Font _fontHdr;     // 8.0pt Bold    — Wochentag-Spaltenköpfe im Kalender
        private Font _fontDay;     // 8.5pt Regular — normale Kalendertage + Wochentag-Label
        private Font _fontKW;      // 7.5pt Regular — Kalenderwochen-Zahlen
        private Font _fontSmall;   // 8.0pt Regular — Terminuhrzeit
        private Font _fontItalic;  // 8.0pt Italic  — Terminort
        private Font _fontEmoji;   // Segoe UI Emoji 8.5pt Bold — Terminbetreff (Emoji-fähig)

        // Hintergrund der Terminliste — etwas dunkler als SystemColors.Window
        private static readonly Color _listBg = Darken(SystemColors.Window, 1.0f);

        // ══════════════════════════════════════════════════════════════════
        // Properties
        // ══════════════════════════════════════════════════════════════════

        public StringCollection Accounts { get; set; }

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

            InitCachedFonts();
            InitCalendarControls();

            _flpAppointments = BuildFlowPanel();
            pnlAppointments.BackColor = _listBg;              // Container-Hintergrund (Leerraum unter Terminen)
            pnlAppointments.Controls.Add(_flpAppointments);

            _toolTip = new ToolTip { AutoPopDelay = 8000 };

            // Resize-Debounce: 300 ms nach letztem Resize neu aufbauen
            _resizeTimer       = new System.Windows.Forms.Timer { Interval = 300 };
            _resizeTimer.Tick += (s, e) => { _resizeTimer.Stop(); if (IsHandleCreated) RetrieveData(); };
            pnlAppointments.Resize += (s, e) => { _resizeTimer.Stop(); _resizeTimer.Start(); };

            // Alle 60 s auf Tageswechsel prüfen
            _dayChangeTimer       = new System.Windows.Forms.Timer { Interval = 60000 };
            _dayChangeTimer.Tick += (s, e) =>
            {
                if (DateTime.Today == _lastKnownDate) return;
                _lastKnownDate = DateTime.Today;
                _selectedDate  = DateTime.Today;
                _calendarMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                if (IsHandleCreated) InvalidateAndRefresh();
            };
            _dayChangeTimer.Start();

            // Click-Timer: SingleClick erst nach DoubleClickTime ausführen damit DblClick nicht verschluckt wird
            _clickTimer = new System.Windows.Forms.Timer
                { Interval = SystemInformation.DoubleClickTime + 50 };
            _clickTimer.Tick += (s, e) =>
            {
                _clickTimer.Stop();
                _selectedDate = _pendingClickDate;
                if (_pendingClickDate.Month != _calendarMonth.Month ||
                    _pendingClickDate.Year  != _calendarMonth.Year)
                    _calendarMonth = new DateTime(_pendingClickDate.Year, _pendingClickDate.Month, 1);
                RetrieveData();
            };

            Accounts = Properties.Settings.Default.Accounts;
        }

        private void InitCachedFonts()
        {
            var ff    = Font.FontFamily;
            _fontBold   = new Font(ff, 8.5f, FontStyle.Bold);
            _fontHdr    = new Font(ff, 8.0f, FontStyle.Bold);
            _fontDay    = new Font(ff, 8.5f, FontStyle.Regular);
            _fontKW     = new Font(ff, 7.5f, FontStyle.Regular);
            _fontSmall  = new Font(ff, 8.0f, FontStyle.Regular);
            _fontItalic = new Font(ff, 8.0f, FontStyle.Italic);
            _fontEmoji  = new Font("Segoe UI Emoji", 8.5f, FontStyle.Bold);
        }

        internal void DisposeCachedFonts()
        {
            _fontBold?.Dispose();   _fontBold   = null;
            _fontHdr?.Dispose();    _fontHdr    = null;
            _fontDay?.Dispose();    _fontDay    = null;
            _fontKW?.Dispose();     _fontKW     = null;
            _fontSmall?.Dispose();  _fontSmall  = null;
            _fontItalic?.Dispose(); _fontItalic = null;
            _fontEmoji?.Dispose();  _fontEmoji  = null;
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
                BackColor     = _listBg
            };
        }

        // ══════════════════════════════════════════════════════════════════
        // RetrieveData / Cache-Steuerung
        // ══════════════════════════════════════════════════════════════════

        public void RetrieveData()
        {
            RetrieveAppointments();
        }

        // Für externe Trigger (ItemAdd/Change/Remove, Sleep/Wake, Tageswechsel):
        // Cache löschen und komplett neu laden.
        public void InvalidateAndRefresh()
        {
            _cachedAppts   = null;
            _cacheYear     = -1;
            _cacheMonth    = -1;
            _calendarBuilt = DateTime.MinValue;
            RetrieveData();
        }

        // ══════════════════════════════════════════════════════════════════
        // KALENDER — einmalige Initialisierung + schneller Property-Update
        // ══════════════════════════════════════════════════════════════════

        // Alle Controls einmal erstellen und in panelCalendar einhängen.
        // BuildCalendar() aktualisiert danach nur noch Properties — kein GDI-Overhead.
        private void InitCalendarControls()
        {
            int rowH = _fontDay.Height + 13;

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
            tbl.Controls.Add(CalNavBtn("<", () => { _calendarMonth = _calendarMonth.AddMonths(-1); BuildCalendar(); }), 0, 0);

            _calMonthLbl = new Label { Font = _fontBold, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(0) };
            tbl.Controls.Add(_calMonthLbl, 1, 0);
            tbl.SetColumnSpan(_calMonthLbl, 5);

            _calNavHome = CalNavBtn("⌂", () => { _selectedDate = DateTime.Today; _calendarMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1); RetrieveData(); });
            tbl.Controls.Add(_calNavHome, 6, 0);

            tbl.Controls.Add(CalNavBtn(">", () => { _calendarMonth = _calendarMonth.AddMonths(1); BuildCalendar(); }), 7, 0);

            // Zeile 1: KW-Header + Tagesnamen
            _calKWLabels = new Label[7];
            _calKWLabels[0] = CalKWLbl("KW");
            tbl.Controls.Add(_calKWLabels[0], 0, 1);
            string[] dn = { "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So" };
            for (int d = 0; d < 7; d++)
                tbl.Controls.Add(CalLbl(dn[d], _fontHdr, SystemColors.GrayText), d + 1, 1);

            // Zeilen 2–7: KW-Labels + Tageszellen (werden in BuildCalendar() nur noch upgedated)
            _calDayCells = new Label[42];
            for (int row = 0; row < 6; row++)
            {
                _calKWLabels[row + 1] = CalKWLbl("");
                tbl.Controls.Add(_calKWLabels[row + 1], 0, row + 2);
                for (int col = 0; col < 7; col++)
                {
                    var lbl = new Label
                    {
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill, Cursor = Cursors.Hand,
                        Margin = new Padding(0), Font = _fontDay
                    };
                    lbl.Click       += CalDay_Click;
                    lbl.DoubleClick += CalDay_DblClick;
                    _calDayCells[row * 7 + col] = lbl;
                    tbl.Controls.Add(lbl, col + 1, row + 2);
                }
            }

            panelCalendar.SuspendLayout();
            panelCalendar.Height = 8 * rowH + 4;
            panelCalendar.Controls.Add(tbl);
            panelCalendar.ResumeLayout();
        }

        // Nur Property-Updates — kein Control-Create, kein Dispose, kein Layout-Overhead.
        private void BuildCalendar()
        {
            _calMonthLbl.Text     = _calendarMonth.ToString("MMMM yyyy");
            _calNavHome.ForeColor = (_calendarMonth.Year  == DateTime.Today.Year &&
                                     _calendarMonth.Month == DateTime.Today.Month)
                                    ? Color.SteelBlue : SystemColors.ControlText;

            int dow = (int)_calendarMonth.DayOfWeek;
            DateTime ws = _calendarMonth.AddDays((dow == 0) ? -6 : -(dow - 1));

            for (int row = 0; row < 6; row++)
            {
                _calKWLabels[row + 1].Text = GetWeekNumber(ws).ToString();
                for (int col = 0; col < 7; col++)
                {
                    DateTime cd  = ws.AddDays(col);
                    bool isCur   = cd.Month == _calendarMonth.Month;
                    bool isToday = cd.Date  == DateTime.Today;
                    bool isSel   = cd.Date  == _selectedDate.Date;
                    bool isBold  = _boldedDates.Contains(cd.Date);

                    var lbl       = _calDayCells[row * 7 + col];
                    lbl.Text      = cd.Day.ToString();
                    lbl.Tag       = cd;
                    lbl.Font      = isBold ? _fontBold : _fontDay;
                    lbl.BackColor = isSel   ? Color.LightBlue :
                                    isToday ? Color.SteelBlue  : SystemColors.Window;
                    lbl.ForeColor = isSel   ? Color.DarkBlue  :
                                    isToday ? Color.White      :
                                    !isCur  ? Color.DarkGray   : SystemColors.WindowText;
                }
                ws = ws.AddDays(7);
            }
        }

        private Button CalNavBtn(string text, Action onClick, Color? foreColor = null)
        {
            var b = new Button
            {
                Text      = text,
                FlatStyle = FlatStyle.Flat,
                Dock      = DockStyle.Fill,
                Font      = _fontBold,
                Cursor    = Cursors.Hand,
                Margin    = new Padding(0),
                BackColor = SystemColors.Window,
                ForeColor = foreColor ?? SystemColors.ControlText
            };
            b.FlatAppearance.BorderSize = 0;
            b.Click += (s, e) => onClick();
            return b;
        }

        private static Label CalLbl(string text, Font font, Color fore)
            => new Label
            {
                Text      = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock      = DockStyle.Fill,
                Font      = font,
                ForeColor = fore,
                Margin    = new Padding(0)
            };

        private Label CalKWLbl(string text)
        {
            var lbl = new Label
            {
                Text      = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock      = DockStyle.Fill,
                Font      = _fontKW,
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
            _pendingClickDate = d;
            _clickTimer.Stop();
            _clickTimer.Start();
        }

        private void CalDay_DblClick(object sender, EventArgs e)
        {
            _clickTimer.Stop();
            if (!(sender is Label l) || !(l.Tag is DateTime d)) return;
            try
            {
                var appt = Globals.ThisAddIn.Application
                               .CreateItem(Outlook.OlItemType.olAppointmentItem)
                               as Outlook.AppointmentItem;
                appt.Start = d.Date.AddHours(DateTime.Now.Hour).AddMinutes(DateTime.Now.Minute);
                appt.End   = appt.Start.AddHours(1);
                appt.Display(true);
                InvalidateAndRefresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("DblClick: " + ex.Message, "AddIn-Fehler");
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // TERMINE — Datenabruf
        // ══════════════════════════════════════════════════════════════════

        private void RetrieveAppointments()
        {
            // COM-Roundtrip nur wenn Cache leer oder anderer Monat gewählt
            bool cacheStale = _cachedAppts == null
                              || _cacheYear  != _selectedDate.Year
                              || _cacheMonth != _selectedDate.Month;

            if (cacheStale)
            {
                var appts = new List<Outlook.AppointmentItem>();
                foreach (Outlook.Store store in Globals.ThisAddIn.Application.Session.Stores)
                    if (Accounts == null || Accounts.Count == 0 || Accounts.Contains(store.DisplayName))
                        appts.AddRange(RetrieveAppointmentsForFolder(
                            store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder));

                appts.Sort(CompareAppointments);
                _cachedAppts = appts;
                _cacheYear   = _selectedDate.Year;
                _cacheMonth  = _selectedDate.Month;
                _boldedDates = new HashSet<DateTime>(appts.Select(a => a.Start.Date).Distinct());
            }

            // Kalender nur neu aufbauen wenn Monat wechselt oder frische Daten vorliegen
            if (cacheStale || _calendarBuilt != _calendarMonth)
            {
                BuildCalendar();
                _calendarBuilt = _calendarMonth;
            }

            DateTime start  = _selectedDate.Date;
            // Heute: vergangene Termine ausblenden sobald End < jetzt; andere Tage: ab Tagesbeginn
            DateTime cutoff = (start == DateTime.Today) ? DateTime.Now : start;

            BuildItemsPanel(
                _cachedAppts.Where(a => a.End >= cutoff && a.Start <= start.AddDays(90))
                            .ToList());
        }

        private List<Outlook.AppointmentItem> RetrieveAppointmentsForFolder(Outlook.Folder cal)
        {
            var start = new DateTime(_selectedDate.Year, _selectedDate.Month, 1);
            var end   = start.AddMonths(1).AddDays(-1).AddDays(90);
            var range = GetAppointmentsInRange(cal, start, end);
            var list  = new List<Outlook.AppointmentItem>();
            if (range != null) foreach (Outlook.AppointmentItem a in range) list.Add(a);
            return list;
        }

        private Outlook.Items GetAppointmentsInRange(Outlook.Folder folder, DateTime start, DateTime end)
        {
            // Outlook DASL-Filter erwartet MM/dd/yyyy HH:mm (invariant) — Systemsprache ignorieren
            string f = "[Start] >= '" + start.ToString("MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture) + "'" +
                       " AND [End] <= '" + end.ToString("MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture) + "'";
            try
            {
                var i = folder.Items;
                i.IncludeRecurrences = true;
                i.Sort("[Start]", Type.Missing);
                var r = i.Restrict(f);
                return r.Count > 0 ? r : null;
            }
            catch { return null; }
        }

        private static int CompareAppointments(Outlook.AppointmentItem x, Outlook.AppointmentItem y)
            => x.Start.CompareTo(y.Start);

        // ══════════════════════════════════════════════════════════════════
        // TERMINLISTE — Darstellung
        // ══════════════════════════════════════════════════════════════════

        private void BuildItemsPanel(List<Outlook.AppointmentItem> appts)
        {
            _flpAppointments.SuspendLayout();
            var old = _flpAppointments.Controls.Cast<Control>().ToList();
            _flpAppointments.Controls.Clear();
            foreach (var c in old) c.Dispose();

            if (appts.Count == 0) { _flpAppointments.ResumeLayout(); return; }

            int panelH = pnlAppointments.ClientSize.Height - 49;   // -49: Outlook-Statusbalken
            if (panelH <= 0) panelH = 400;
            int w = Math.Max(pnlAppointments.ClientSize.Width - 2, 100);

            int rowH    = _fontBold.Height + 6;
            int hdrH    = rowH + 2;
            int spacerH = 2;
            int usedH   = 0;
            int lastDay = -1, lastYear = -1;

            foreach (var appt in appts)
            {
                DateTime itemDate = appt.Start.Date;
                bool     hasLoc   = !string.IsNullOrEmpty(appt.Location);

                bool newDay = itemDate.Day != lastDay || itemDate.Year != lastYear;
                int needed  = (newDay ? hdrH : 0) + (hasLoc ? 2 : 1) * rowH + spacerH;
                if (usedH + needed > panelH) break;

                if (newDay)
                {
                    lastDay = itemDate.Day; lastYear = itemDate.Year;
                    var (dateText, weekdayText) = FormatDateHeader(itemDate);
                    _flpAppointments.Controls.Add(BuildGroupHeader(dateText, weekdayText, w, hdrH));
                    usedH += hdrH;
                }

                _flpAppointments.Controls.Add(BuildAppointmentEntry(appt, w, rowH));
                usedH += (hasLoc ? 2 : 1) * rowH;

                _flpAppointments.Controls.Add(new Panel { Height = spacerH, Width = w, BackColor = _listBg });
                usedH += spacerH;
            }

            _flpAppointments.ResumeLayout();
        }

        // ── Datum/Gruppen-Header ──────────────────────────────────────────

        private (string dateText, string weekdayText) FormatDateHeader(DateTime date)
        {
            int diff = (int)(date - DateTime.Today).TotalDays;
            string prefix = diff == -1 ? Constants.Yesterday + ":  " :
                            diff ==  0 ? Constants.Today     + ":  " :
                            diff ==  1 ? Constants.Tomorrow  + ":  " : "";
            return (prefix + date.ToShortDateString(), date.ToString("dddd"));
        }

        private Control BuildGroupHeader(string dateText, string weekdayText, int width, int height)
        {
            var pnl = new Panel
            {
                Width     = width,
                Height    = height,
                BackColor = _listBg,
                Margin    = new Padding(0),
                Padding   = new Padding(0)
            };

            var lblDate = new Label
            {
                Text      = dateText,
                Font      = _fontBold,
                ForeColor = Color.FromArgb(40, 60, 100),
                BackColor = Color.Transparent,
                AutoSize  = true,
                Height    = height,
                TextAlign = ContentAlignment.MiddleLeft,
                Location  = new Point(4, 0)
            };
            var lblDay = new Label
            {
                Text      = weekdayText,
                Font      = _fontDay,
                ForeColor = Color.FromArgb(40, 60, 100),
                BackColor = Color.Transparent,
                AutoSize  = true,
                Height    = height,
                TextAlign = ContentAlignment.MiddleLeft
            };

            pnl.Controls.Add(lblDate);
            pnl.Controls.Add(lblDay);

            lblDate.Width   = TextRenderer.MeasureText(dateText, _fontBold).Width;
            lblDay.Location = new Point(lblDate.Left + lblDate.Width + 4, 0);

            return pnl;
        }

        // ── Termin-Eintrag ────────────────────────────────────────────────

        private TableLayoutPanel BuildAppointmentEntry(Outlook.AppointmentItem appt, int width, int rowH)
        {
            bool  hasLoc   = !string.IsNullOrEmpty(appt.Location);
            bool  hasCat   = !string.IsNullOrEmpty(appt.Categories);
            Color barColor = GetApptBarColor(appt);
            // mit Kategorie: aufgehellte Kategoriefarbe als Zeilen-Hintergrund; ohne: neutral weiß
            Color rowBg    = hasCat ? Lighten(barColor, 0.72f) : SystemColors.Window;
            var   tbl      = BuildEntryTable(width, rowH, hasLoc ? 2 : 1, rowBg);

            tbl.Controls.Add(new Label
            {
                Text      = appt.AllDayEvent ? "" : appt.Start.ToShortTimeString(),
                Font      = _fontSmall,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock      = DockStyle.Fill,
                AutoEllipsis = false,
                Padding   = new Padding(15, 0, 2, 0),
                Margin    = new Padding(0)
            }, 0, 0);

            var bar = new Panel { BackColor = barColor, Dock = DockStyle.Fill, Margin = new Padding(0) };
            tbl.Controls.Add(bar, 1, 0);
            if (hasLoc) tbl.SetRowSpan(bar, 2);

            tbl.Controls.Add(new Label
            {
                Text      = appt.Subject ?? "",
                Font      = _fontEmoji,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock      = DockStyle.Fill,
                AutoEllipsis = true,
                Padding   = new Padding(4, 0, 2, 0),
                Margin    = new Padding(0)
            }, 2, 0);

            if (hasLoc)
            {
                tbl.Controls.Add(new Label { Margin = new Padding(0) }, 0, 1);
                tbl.Controls.Add(new Label
                {
                    Text      = appt.Location,
                    Font      = _fontItalic,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Dock      = DockStyle.Fill,
                    AutoEllipsis = true,
                    ForeColor = SystemColors.GrayText,
                    Padding   = new Padding(4, 0, 2, 0),
                    Margin    = new Padding(0)
                }, 2, 1);
            }

            TintChildren(tbl, rowBg);
            AttachEvents(tbl, appt, BuildApptTooltip(appt));
            return tbl;
        }

        private TableLayoutPanel BuildEntryTable(int width, int rowH, int rows, Color bg)
        {
            var tbl = new TableLayoutPanel
            {
                ColumnCount     = 3,
                RowCount        = rows,
                Width           = width,
                Height          = rows * rowH,
                Padding         = new Padding(0),
                Margin          = new Padding(0),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                BackColor       = bg
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, COL_TIME));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, COL_BAR));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  100));
            for (int r = 0; r < rows; r++)
                tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, rowH));
            return tbl;
        }

        private static void TintChildren(TableLayoutPanel tbl, Color bg)
        {
            tbl.BackColor = bg;
            foreach (Control c in tbl.Controls)
            {
                if (c is Panel) continue;                               // Farbbalken — nicht anfassen
                // Spalte 0 (Uhrzeit) immer neutral weiß, restliche Zellen transparent (Tabellen-bg)
                c.BackColor = tbl.GetColumn(c) == 0 ? SystemColors.Window : Color.Transparent;
            }
        }

        // ── Events + Tooltip ─────────────────────────────────────────────

        private void AttachEvents(Control ctrl, Outlook.AppointmentItem appt, string tooltip)
        {
            _toolTip.SetToolTip(ctrl, tooltip);
            ctrl.DoubleClick += (s, e) => OpenAppt(appt);
            ctrl.MouseUp     += (s, e) =>
            {
                if (((MouseEventArgs)e).Button == MouseButtons.Right)
                    { _contextMenuAppt = appt; ctxMenuAppointments.Show(ctrl, ((MouseEventArgs)e).Location); }
            };
            foreach (Control child in ctrl.Controls)
                AttachEvents(child, appt, tooltip);
        }

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

        // ── Termin öffnen ─────────────────────────────────────────────────

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

        // ══════════════════════════════════════════════════════════════════
        // Context-Menu Termine
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
                {
                    if (f.OpenRecurring) ((Outlook.AppointmentItem)appt.Parent).Display(true);
                    else                 appt.Display(true);
                }
            }
            else appt.Display(true);
            InvalidateAndRefresh();
        }

        private void mnuItemReplyAllEmail_Click(object sender, EventArgs e)
        {
            if (_contextMenuAppt == null) return;
            var mail = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            string cur = OutlookHelper.GetEmailAddress(Globals.ThisAddIn.Application.Session.CurrentUser);
            foreach (Outlook.Recipient r in _contextMenuAppt.Recipients)
            {
                string a = OutlookHelper.GetEmailAddress(r);
                if (cur != a) mail.Recipients.Add(a);
            }
            mail.Body    = "\n\n" + _contextMenuAppt.Body;
            mail.Subject = Constants.SubjectRE + ": " + _contextMenuAppt.Subject;
            mail.Display();
        }

        private void mnuItemDeleteAppointment_Click(object sender, EventArgs e)
        {
            if (_contextMenuAppt == null) return;
            if (_contextMenuAppt.IsRecurring)
            {
                var f = new FormRecurringOpen
                {
                    Title   = "Warning: Delete Recurring Item",
                    Message = "This is one appointment in a series. What do you want to delete?"
                };
                if (f.ShowDialog() == DialogResult.OK)
                {
                    if (f.OpenRecurring) ((Outlook.AppointmentItem)_contextMenuAppt.Parent).Delete();
                    else                 _contextMenuAppt.Delete();
                }
            }
            else if (MessageBox.Show("Termin wirklich löschen?",
                         "Termin löschen", MessageBoxButtons.YesNo) == DialogResult.Yes)
                _contextMenuAppt.Delete();
            InvalidateAndRefresh();
        }

        // ══════════════════════════════════════════════════════════════════
        // Kategorie-Farben
        // ══════════════════════════════════════════════════════════════════

        // Kategoriefarbe abdunkeln (factor < 1 = dunkler)
        private static Color Darken(Color c, float factor)
        {
            return Color.FromArgb(
                c.A,
                (int)(c.R * factor),
                (int)(c.G * factor),
                (int)(c.B * factor));
        }

        // Kategoriefarbe aufhellen — Richtung Weiß mischen (amount 0 = Original, 1 = weiß)
        private static Color Lighten(Color c, float amount)
        {
            return Color.FromArgb(
                c.A,
                (int)(c.R + (255 - c.R) * amount),
                (int)(c.G + (255 - c.G) * amount),
                (int)(c.B + (255 - c.B) * amount));
        }

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
