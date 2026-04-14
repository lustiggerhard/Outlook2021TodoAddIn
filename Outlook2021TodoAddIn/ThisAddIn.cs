using Microsoft.Win32;
using Outlook2021TodoAddIn.Forms;
using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace Outlook2021TodoAddIn
{
    public partial class ThisAddIn
    {
        public AppointmentsControl AppControl { get; set; }
        public Microsoft.Office.Tools.CustomTaskPane ToDoTaskPane { get; set; }
        private bool _taskPaneCreated = false;
        private System.Windows.Forms.Timer _refreshTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.AddRegistryNotification();
                var startupTimer = new System.Windows.Forms.Timer();
                startupTimer.Interval = 1000;
                startupTimer.Tick += (s, ev) =>
                {
                    startupTimer.Stop();
                    startupTimer.Dispose();
                    CreateTaskPane();
                };
                startupTimer.Start();
            }
            catch (Exception exc)
            {
                MessageBox.Show(String.Format("Error starting Calendar AddIn: {0}", exc.ToString()));
            }
        }

        private void CreateTaskPane()
        {
            try
            {
                this.AppControl = new AppointmentsControl();
                this.AppControl.MailAlertsEnabled        = Properties.Settings.Default.MailAlertsEnabled;
                this.AppControl.ShowPastAppointments     = Properties.Settings.Default.ShowPastAppointments;
                this.AppControl.ShowFriendlyGroupHeaders = Properties.Settings.Default.ShowFriendlyGroupHeaders;
                this.AppControl.ShowDayNames             = Properties.Settings.Default.ShowDayNames;
                this.AppControl.ShowWeekNumbers          = Properties.Settings.Default.ShowWeekNumbers;
                this.AppControl.ShowTasks                = Properties.Settings.Default.ShowTasks;
                this.AppControl.ShowCompletedTasks       = Properties.Settings.Default.ShowCompletedTasks;
                this.AppControl.FirstDayOfWeek           = Properties.Settings.Default.FirstDayOfWeek;
                this.AppControl.NumDays                  = Properties.Settings.Default.NumDays;

                ToDoTaskPane = this.CustomTaskPanes.Add(this.AppControl, " ");
            // Scrollbars deaktivieren für Outlook2021TodoAddIn (gerhard@lustig.at)
                ToDoTaskPane.Visible              = Properties.Settings.Default.Visible;
                ToDoTaskPane.Width                = Properties.Settings.Default.Width;
                ToDoTaskPane.DockPosition         = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                ToDoTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                ToDoTaskPane.VisibleChanged       += ToDoTaskPane_VisibleChanged;
                this.AppControl.SizeChanged       += appControl_SizeChanged;

                _taskPaneCreated = true;
                // Kalender-Änderungen überwachen
                var calFolder = this.Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar) as Microsoft.Office.Interop.Outlook.Folder;
                if (calFolder != null)
                {
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemAdd += (item) => { if (AppControl != null) AppControl.RetrieveData(); };
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemChange += (item) => { if (AppControl != null) AppControl.RetrieveData(); };
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemRemove += () => { if (AppControl != null) AppControl.RetrieveData(); };
                }
                this.AppControl.SelectedDate = DateTime.Today;

                ((Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)this.Application).Quit
                    += Application_Quit;
                this.Application.ActiveExplorer().Deactivate += ThisAddIn_Deactivate;

                _refreshTimer = new System.Windows.Forms.Timer();
                _refreshTimer.Interval = 30 * 60 * 1000;
                _refreshTimer.Tick += (s, e) => { if (AppControl != null) AppControl.RetrieveData(); };
                _refreshTimer.Start();
            }
            catch (Exception exc)
            {
                MessageBox.Show(String.Format("Error creating TaskPane: {0}", exc.ToString()));
            }
        }

        private void Application_Quit()
        {
            if (_taskPaneCreated && ToDoTaskPane != null)
                Properties.Settings.Default.Visible = ToDoTaskPane.Visible;
            Properties.Settings.Default.Save();
        }

        private void appControl_SizeChanged(object sender, EventArgs e)
        {
            if (ToDoTaskPane != null)
                Properties.Settings.Default.Width = ToDoTaskPane.Width;
        }

        private void ToDoTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (_taskPaneCreated && ToDoTaskPane != null)
                Properties.Settings.Default.Visible = ToDoTaskPane.Visible;
            TodoRibbonAddIn rbn = Globals.Ribbons.TodoRibbonAddIn;
            if (rbn != null)
                rbn.btnToggleTodo.Checked = ToDoTaskPane != null && ToDoTaskPane.Visible;
        }

        private void ThisAddIn_Deactivate()
        {
            if (_taskPaneCreated && ToDoTaskPane != null)
                Properties.Settings.Default.Visible = ToDoTaskPane.Visible;
            Properties.Settings.Default.Save();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_refreshTimer != null) { _refreshTimer.Stop(); _refreshTimer.Dispose(); }
            Properties.Settings.Default.Save();
        }

        private void AddRegistryNotification()
        {
            string subKey = @"Software\Microsoft\Office\Outlook\Addins\Outlook2021TodoAddIn";
            Microsoft.Win32.RegistryKey rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(subKey, true);
            if (rk == null) rk = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(subKey);
            if ((int)rk.GetValue("RequireShutdownNotification", 0) == 0)
                rk.SetValue("RequireShutdownNotification", 1, Microsoft.Win32.RegistryValueKind.DWord);
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup  += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
