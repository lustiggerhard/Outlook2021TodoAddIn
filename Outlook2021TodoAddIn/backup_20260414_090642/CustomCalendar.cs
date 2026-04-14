using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Globalization;

namespace Outlook2021TodoAddIn
{
    public partial class CustomCalendar : UserControl
    {
        #region "Variables"
        DateTime _selectedDate = DateTime.Today;
        #endregion "Variables"

        #region "Properties"
        public DateTime SelectedDate
        {
            get { return _selectedDate; }
            set
            {
                _selectedDate = value;
                this.OnSelectedDateChanged(EventArgs.Empty);
            }
        }
        public DayOfWeek FirstDayOfWeek { get; set; }
        public DateTime[] BoldedDates { get; set; }
        public Color CurrentMonthForeColor { get; set; }
        public Color OtherMonthForeColor { get; set; }
        public Color TodayForeColor { get; set; }
        public Color TodayBackColor { get; set; }
        public Color SelectedForeColor { get; set; }
        public Color SelectedBackColor { get; set; }
        public Color HoverForeColor { get; set; }
        public Color HoverBackColor { get; set; }
        public bool ShowWeekNumbers { get; set; }
        #endregion "Properties"

        #region "Methods"

        public CustomCalendar()
        {
            InitializeComponent();
            this.FirstDayOfWeek = DayOfWeek.Sunday;
            this.CurrentMonthForeColor = Color.Black;
            this.OtherMonthForeColor = Color.LightGray;
            this.TodayForeColor = Color.White;
            this.TodayBackColor = Color.Blue;
            this.SelectedForeColor = Color.Blue;
            this.SelectedBackColor = Color.LightBlue;
            this.HoverForeColor = Color.Black;
            this.HoverBackColor = Color.LightCyan;
            this.lnkToday.Visible = false;
            this.btnConfig.Visible = false;
            this.BackColor = SystemColors.Window;
            this.btnPrevious.FlatStyle = FlatStyle.Flat;
            this.btnPrevious.FlatAppearance.BorderSize = 0;
            this.btnPrevious.Font = new Font(this.btnPrevious.Font.FontFamily, 11f, FontStyle.Bold);
            this.btnNext.FlatStyle = FlatStyle.Flat;
            this.btnNext.FlatAppearance.BorderSize = 0;
            this.btnNext.Font = new Font(this.btnNext.Font.FontFamily, 11f, FontStyle.Bold);
            this.tableLayoutPanel1.BackColor = SystemColors.Window;
            this.tableLayoutPanel2.BackColor = SystemColors.Window;
            this.tableLayoutPanel1.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
            this.tableLayoutPanel2.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
            this.Paint += new PaintEventHandler(this.CustomCalendar_Paint);
            this.CreateTableControls();
        }

        private void CustomCalendar_Paint(object sender, PaintEventArgs e)
        {
            int x = this.tableLayoutPanel2.Right;
            using (var pen = new Pen(Color.LightGray, 1))
                e.Graphics.DrawLine(pen, x, this.tableLayoutPanel2.Top, x, this.tableLayoutPanel2.Bottom);
        }

        private void CustomCalendar_Load(object sender, EventArgs e) { }

        private void CreateTableControls()
        {
            this.tableLayoutPanel1.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
            for (int row = 0; row < this.tableLayoutPanel1.RowCount; row++)
            {
                for (int col = 0; col < this.tableLayoutPanel1.ColumnCount; col++)
                {
                    Label lblCtrl = new Label() { Text = "xx" };
                    lblCtrl.Name = String.Format("lbl_{0}_{1}", col.ToString(), row.ToString());
                    lblCtrl.Dock = DockStyle.Fill;
                    lblCtrl.TextAlign = ContentAlignment.MiddleCenter;
                    lblCtrl.Margin = Padding.Empty;
                    lblCtrl.Padding = Padding.Empty;
                    lblCtrl.FlatStyle = FlatStyle.Flat;
                    if (row == 0) { lblCtrl.AutoEllipsis = true; lblCtrl.AutoSize = false; }
                    if (row != 0)
                    {
                        lblCtrl.MouseEnter += lblCtrl_MouseEnter;
                        lblCtrl.MouseLeave += lblCtrl_MouseLeave;
                        lblCtrl.Click += lblCtrl_Click;
                        lblCtrl.DoubleClick += lblCtrl_DoubleClick;
                        lblCtrl.AllowDrop = true;
                        lblCtrl.DragEnter += lblCtrl_DragEnter;
                        lblCtrl.DragDrop += lblCtrl_DragDrop;
                    }
                    this.tableLayoutPanel1.Controls.Add(lblCtrl);
                    this.tableLayoutPanel1.SetCellPosition(lblCtrl, new TableLayoutPanelCellPosition(col, row));
                }
            }

            this.tableLayoutPanel2.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
            for (int row = 0; row < this.tableLayoutPanel2.RowCount; row++)
            {
                Label lblCtrl = new Label() { Text = "#" };
                lblCtrl.Name = String.Format("lbl_{0}", row.ToString());
                lblCtrl.Dock = DockStyle.Fill;
                lblCtrl.TextAlign = ContentAlignment.MiddleRight;
                lblCtrl.Margin = new Padding(0, 0, 4, 0);
                lblCtrl.Padding = Padding.Empty;
                lblCtrl.Font = new Font(this.Font.FontFamily, this.Font.Size - 2, FontStyle.Italic);
                lblCtrl.FlatStyle = FlatStyle.Flat;
                this.tableLayoutPanel2.Controls.Add(lblCtrl);
                this.tableLayoutPanel2.SetCellPosition(lblCtrl, new TableLayoutPanelCellPosition(0, row));
            }

            this.btnPrevious.FlatAppearance.MouseOverBackColor = this.HoverBackColor;
            this.btnNext.FlatAppearance.MouseOverBackColor = this.HoverBackColor;
            this.btnConfig.FlatAppearance.MouseOverBackColor = this.HoverBackColor;
        }

        private void lblCtrl_DragEnter(object sender, DragEventArgs e)
        {
            List<string> outlookRequiredFormats = new List<string>() { "RenPrivateSourceFolder", "RenPrivateMessages", "RenPrivateItem", "FileGroupDescriptor", "FileGroupDescriptorW", "FileContents", "Object Descriptor" };
            if (outlookRequiredFormats.All(r => e.Data.GetDataPresent(r)))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void lblCtrl_DragDrop(object sender, DragEventArgs e)
        {
            Label lblDay = sender as Label;
            if (sender != null)
            {
                Outlook.Explorer mailExpl = Globals.ThisAddIn.Application.ActiveExplorer();
                List<string> attendees = new List<string>();
                string curUserAddress = OutlookHelper.GetEmailAddress(Globals.ThisAddIn.Application.Session.CurrentUser);
                string body = String.Empty;
                string subject = String.Empty;
                foreach (object obj in mailExpl.Selection)
                {
                    Outlook.MailItem mail = obj as Outlook.MailItem;
                    if (mail != null)
                    {
                        subject = mail.Subject;
                        body = mail.Body;
                        if (mail.SenderEmailAddress != curUserAddress && !attendees.Contains(mail.SenderEmailAddress))
                            attendees.Add(mail.SenderEmailAddress);
                        attendees.AddRange(OutlookHelper.GetRecipentsEmailAddresses(mail.Recipients, curUserAddress));
                    }
                    else
                    {
                        Outlook.MeetingItem meeting = obj as Outlook.MeetingItem;
                        if (meeting != null)
                        {
                            subject = meeting.Subject;
                            body = meeting.Body;
                            if (meeting.SenderEmailAddress != curUserAddress && !attendees.Contains(meeting.SenderEmailAddress))
                                attendees.Add(meeting.SenderEmailAddress);
                            attendees.AddRange(OutlookHelper.GetRecipentsEmailAddresses(meeting.Recipients, curUserAddress));
                        }
                        else
                        {
                            Outlook.AppointmentItem appointment = obj as Outlook.AppointmentItem;
                            if (appointment != null)
                            {
                                subject = appointment.Subject;
                                body = appointment.Body;
                                if (appointment.Organizer != curUserAddress && !attendees.Contains(appointment.Organizer))
                                    attendees.Add(appointment.Organizer);
                                attendees.AddRange(OutlookHelper.GetRecipentsEmailAddresses(appointment.Recipients, curUserAddress));
                            }
                        }
                    }
                }
                Outlook.AppointmentItem appt = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                attendees.ForEach(a => appt.Recipients.Add(a));
                appt.Body = Environment.NewLine + Environment.NewLine + body;
                appt.Subject = subject;
                DateTime day = (DateTime)lblDay.Tag;
                DateTime now = DateTime.Now;
                appt.Start = OutlookHelper.RoundUp(new DateTime(day.Year, day.Month, day.Day, now.Hour, now.Minute, now.Second), TimeSpan.FromMinutes(15));
                appt.Display();
            }
        }

        public void UpdateCalendar()
        {
            this.lnkCurrentRange.Text = this.SelectedDate.ToString("MMM yyyy");
            this.lnkToday.Text = Constants.Today + ": " + DateTime.Today.ToShortDateString();

            string[] daysOfWeek = Enum.GetNames(typeof(DayOfWeek));
            string sFirstDayOfWeek = Enum.GetName(typeof(DayOfWeek), this.FirstDayOfWeek);
            List<string> sortedDays = new List<string>();
            sortedDays.AddRange(daysOfWeek.SkipWhile(ow => ow != sFirstDayOfWeek));
            sortedDays.AddRange(daysOfWeek.TakeWhile(ow => ow != sFirstDayOfWeek));

            int dayCurrent = 0;
            int firstIndex = 0;
            DateTime firstOfCurrentMonth = new DateTime(this.SelectedDate.Year, this.SelectedDate.Month, 1);
            DateTime previousMonth = firstOfCurrentMonth.AddMonths(-1);
            DateTime nextMonth = firstOfCurrentMonth.AddMonths(1);
            int daysInPreviousMonth = DateTime.DaysInMonth(previousMonth.Year, previousMonth.Month);
            int daysInCurrentMonth = DateTime.DaysInMonth(this.SelectedDate.Year, this.SelectedDate.Month);

            for (int col = 0; col < this.tableLayoutPanel1.ColumnCount; col++)
            {
                if (sortedDays[col] == Enum.GetName(typeof(DayOfWeek), firstOfCurrentMonth.DayOfWeek))
                    firstIndex = col;
                Label lblDay = this.tableLayoutPanel1.GetControlFromPosition(col, 0) as Label;
                lblDay.Text = sortedDays[col].Substring(0, 2).ToUpper();
            }

            dayCurrent = daysInPreviousMonth - firstIndex + 1;
            if (dayCurrent > daysInPreviousMonth)
                dayCurrent = daysInPreviousMonth - 6;
            bool previousMonthVisible = (dayCurrent != 1);
            bool nextMonthVisible = false;

            if (this.ShowWeekNumbers)
            {
                this.tableLayoutPanel1.Left = this.tableLayoutPanel2.Left + this.tableLayoutPanel2.Width + 2;
                this.tableLayoutPanel1.Width = this.btnNext.Left - this.btnPrevious.Left - this.tableLayoutPanel2.Width - 2 + this.btnNext.Width;
                this.tableLayoutPanel2.Visible = true;
            }
            else
            {
                this.tableLayoutPanel1.Left = this.btnPrevious.Left;
                this.tableLayoutPanel1.Width = this.btnNext.Left - this.btnPrevious.Left + this.btnNext.Width;
                this.tableLayoutPanel2.Visible = false;
            }

            for (int row = 1; row < this.tableLayoutPanel1.RowCount; row++)
            {
                if (this.ShowWeekNumbers)
                {
                    Label lblDoW = this.tableLayoutPanel2.GetControlFromPosition(0, row) as Label;
                    DateTime dateForWeek;
                    if (previousMonthVisible)
                        dateForWeek = new DateTime(previousMonth.Year, previousMonth.Month, dayCurrent);
                    else if (nextMonthVisible)
                        dateForWeek = new DateTime(nextMonth.Year, nextMonth.Month, dayCurrent);
                    else
                        dateForWeek = new DateTime(this.SelectedDate.Year, this.SelectedDate.Month, dayCurrent);
                    lblDoW.Text = GetWeekForDate(dateForWeek).ToString();
                }

                for (int col = 0; col < this.tableLayoutPanel1.ColumnCount; col++)
                {
                    Label lblCtrl = this.tableLayoutPanel1.GetControlFromPosition(col, row) as Label;
                    lblCtrl.Text = dayCurrent.ToString();

                    DateTime embeddedDate;
                    Font displayFont;
                    Color foreColor;
                    Color backColor = this.BackColor;
                    BorderStyle borderStyle = BorderStyle.None;

                    if (previousMonthVisible)
                    {
                        embeddedDate = new DateTime(previousMonth.Year, previousMonth.Month, dayCurrent);
                        displayFont = this.Font;
                        foreColor = this.OtherMonthForeColor;
                    }
                    else if (nextMonthVisible)
                    {
                        embeddedDate = new DateTime(nextMonth.Year, nextMonth.Month, dayCurrent);
                        displayFont = this.Font;
                        foreColor = this.OtherMonthForeColor;
                    }
                    else
                    {
                        embeddedDate = new DateTime(this.SelectedDate.Year, this.SelectedDate.Month, dayCurrent);
                        displayFont = this.Font;
                        foreColor = this.CurrentMonthForeColor;
                    }

                    if (this.BoldedDates != null && this.BoldedDates.Contains(embeddedDate))
                        displayFont = new Font(this.Font, FontStyle.Bold);

                    if (embeddedDate == DateTime.Today)
                    {
                        foreColor = this.TodayForeColor;
                        backColor = this.TodayBackColor;
                    }
                    else if (embeddedDate == this.SelectedDate)
                    {
                        borderStyle = BorderStyle.FixedSingle;
                        foreColor = this.SelectedForeColor;
                        backColor = this.SelectedBackColor;
                    }

                    lblCtrl.Tag = embeddedDate;
                    lblCtrl.Font = displayFont;
                    lblCtrl.ForeColor = foreColor;
                    lblCtrl.BackColor = backColor;
                    lblCtrl.BorderStyle = borderStyle;

                    dayCurrent++;

                    if (previousMonthVisible && dayCurrent > daysInPreviousMonth)
                    {
                        dayCurrent = 1;
                        previousMonthVisible = false;
                    }
                    if (!previousMonthVisible && dayCurrent > daysInCurrentMonth)
                    {
                        dayCurrent = 1;
                        nextMonthVisible = true;
                    }
                }
            }
        }

        private int GetWeekForDate(DateTime time)
        {
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time.AddDays(6), CalendarWeekRule.FirstDay, this.FirstDayOfWeek);
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            this.SelectedDate = this.SelectedDate.AddMonths(-1);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            this.SelectedDate = this.SelectedDate.AddMonths(1);
        }

        private void lblCtrl_MouseEnter(object sender, EventArgs e)
        {
            Label lblCtrl = sender as Label;
            DateTime curDate = (DateTime)lblCtrl.Tag;
            lblCtrl.ForeColor = (curDate.Month == this.SelectedDate.Month && curDate.Year == this.SelectedDate.Year)
                ? this.HoverForeColor : this.OtherMonthForeColor;
            lblCtrl.BackColor = this.HoverBackColor;
        }

        private void lblCtrl_MouseLeave(object sender, EventArgs e)
        {
            Label lblCtrl = sender as Label;
            DateTime curDate = (DateTime)lblCtrl.Tag;
            if (curDate == DateTime.Today)
            {
                lblCtrl.ForeColor = this.TodayForeColor;
                lblCtrl.BackColor = this.TodayBackColor;
            }
            else if (curDate == this.SelectedDate)
            {
                lblCtrl.ForeColor = this.SelectedForeColor;
                lblCtrl.BackColor = this.SelectedBackColor;
            }
            else
            {
                lblCtrl.ForeColor = (curDate.Month == this.SelectedDate.Month && curDate.Year == this.SelectedDate.Year)
                    ? this.CurrentMonthForeColor : this.OtherMonthForeColor;
                lblCtrl.BackColor = this.BackColor;
            }
        }

        private void lblCtrl_Click(object sender, EventArgs e)
        {
            this.SelectedDate = (DateTime)(sender as Label).Tag;
        }

        public event EventHandler CellDoubleClick;
        protected virtual void OnCellDoubleClick(EventArgs e)
        {
            CellDoubleClick?.Invoke(this, e);
        }

        private void lblCtrl_DoubleClick(object sender, EventArgs e)
        {
            this.SelectedDate = (DateTime)(sender as Label).Tag;
            OnCellDoubleClick(EventArgs.Empty);
        }

        public event EventHandler SelectedDateChanged;
        protected virtual void OnSelectedDateChanged(EventArgs e)
        {
            SelectedDateChanged?.Invoke(this, e);
        }

        public event EventHandler ConfigurationButtonClicked;
        private void btnConfig_Click(object sender, EventArgs e)
        {
            ConfigurationButtonClicked?.Invoke(this, e);
        }

        private void lnkToday_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.SelectedDate = DateTime.Today;
        }

        #endregion "Methods"
    }
}
