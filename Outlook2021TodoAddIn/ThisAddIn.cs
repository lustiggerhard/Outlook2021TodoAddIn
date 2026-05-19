/**
 * @file    ThisAddIn.cs
 * @brief   VSTO Add-In Einstiegspunkt.
 * @author  Gerhard Lustig <gerhard@lustig.at>
 * @version 1.7.0
 * @date    2026-05-19
 * @history
 *   1.7.0  2026-05-19  ItemAdd/Change/Remove + TryResumeRefresh rufen jetzt
 *                      InvalidateAndRefresh() statt RetrieveData() — Cache wird
 *                      bei externen Datenänderungen korrekt geleert.
 *   1.6.0  2026-05-19  Task-Panel entfernt (ShowTasks weg); appControl_SizeChanged
 *                      debounced via _widthTimer (500 ms) statt direktem Save pro Pixel.
 *   1.5.0  2026-05-03  _isShuttingDown: verhindert dass VSTO-Cleanup den Visible-State
 *                      nach Application_Quit überschreibt.
 *   1.4.0  2026-05-03  OnPowerModeChanged: Retry-Logik nach Resume (max. 3 Versuche,
 *                      erster Versuch nach 5s, Folge-Versuche nach je 10s).
 *   1.3.0  2026-04-24  PowerModeChanged: bei Resume nach Sleep/Hibernate neu aufbauen.
 *   1.2.0  2026-04-23  Explorer_Activate: Visible-State wiederherstellen.
 *   1.1.0  2026-04-20  Initiale Version mit Refresh-Timer und Kalender-Events.
 */

using Microsoft.Win32;
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
        private bool _isShuttingDown  = false;
        private System.Windows.Forms.Timer _refreshTimer;
        private System.Windows.Forms.Timer _widthTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.AddRegistryNotification();
                SystemEvents.PowerModeChanged += OnPowerModeChanged;

                var startupTimer = new System.Windows.Forms.Timer { Interval = 1000 };
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
                MessageBox.Show(string.Format("Error starting Calendar AddIn: {0}", exc.ToString()));
            }
        }

        private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode != PowerModes.Resume) return;
            TryResumeRefresh(5000, 3);
        }

        // Nach Resume: erst nach 5s versuchen (MAPI/Exchange braucht Zeit).
        // Bei Fehler bis zu retriesLeft-mal nach 10s wiederholen.
        private void TryResumeRefresh(int delayMs, int retriesLeft)
        {
            var t = new System.Windows.Forms.Timer { Interval = delayMs };
            t.Tick += (s, ev) =>
            {
                t.Stop();
                t.Dispose();
                try
                {
                    if (AppControl == null || !AppControl.IsHandleCreated) return;
                    AppControl.Invalidate(true);
                    AppControl.Refresh();
                    AppControl.InvalidateAndRefresh();
                }
                catch
                {
                    if (retriesLeft > 0)
                        TryResumeRefresh(10000, retriesLeft - 1);
                }
            };
            t.Start();
        }

        private void CreateTaskPane()
        {
            try
            {
                this.AppControl = new AppointmentsControl();
                this.AppControl.Accounts = Properties.Settings.Default.Accounts;

                ToDoTaskPane = this.CustomTaskPanes.Add(this.AppControl, " ");
                ToDoTaskPane.Visible              = Properties.Settings.Default.Visible;
                ToDoTaskPane.Width                = Properties.Settings.Default.Width;
                ToDoTaskPane.DockPosition         = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                ToDoTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                ToDoTaskPane.VisibleChanged       += ToDoTaskPane_VisibleChanged;

                // Breite nur alle 500 ms speichern (nicht auf jeden Resize-Pixel)
                _widthTimer = new System.Windows.Forms.Timer { Interval = 500 };
                _widthTimer.Tick += (s, e) =>
                {
                    _widthTimer.Stop();
                    Properties.Settings.Default.Width = ToDoTaskPane.Width;
                };
                this.AppControl.SizeChanged += appControl_SizeChanged;

                _taskPaneCreated = true;

                // Kalender-Änderungen des Default-Stores überwachen
                var calFolder = this.Application.Session.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
                    as Microsoft.Office.Interop.Outlook.Folder;
                if (calFolder != null)
                {
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemAdd
                        += (item) => { if (AppControl != null) AppControl.InvalidateAndRefresh(); };
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemChange
                        += (item) => { if (AppControl != null) AppControl.InvalidateAndRefresh(); };
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemRemove
                        += () => { if (AppControl != null) AppControl.InvalidateAndRefresh(); };
                }

                this.AppControl.SelectedDate = DateTime.Today;

                ((Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)this.Application).Quit
                    += Application_Quit;

                var explorer = this.Application.ActiveExplorer();
                explorer.Deactivate += ThisAddIn_Deactivate;
                ((Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event)explorer).Activate
                    += Explorer_Activate;

                // Auto-Refresh alle 30 Minuten
                _refreshTimer = new System.Windows.Forms.Timer { Interval = 30 * 60 * 1000 };
                _refreshTimer.Tick += (s, e) => { if (AppControl != null) AppControl.InvalidateAndRefresh(); };
                _refreshTimer.Start();
            }
            catch (Exception exc)
            {
                MessageBox.Show(string.Format("Error creating TaskPane: {0}", exc.ToString()));
            }
        }

        private void Explorer_Activate()
        {
            try
            {
                if (!_taskPaneCreated || ToDoTaskPane == null) return;
                bool shouldBeVisible = Properties.Settings.Default.Visible;
                if (ToDoTaskPane.Visible != shouldBeVisible)
                    ToDoTaskPane.Visible = shouldBeVisible;
            }
            catch { }
        }

        private void Application_Quit()
        {
            // Visible-State einfrieren — VSTO-Cleanup löst danach VisibleChanged aus,
            // das darf den gespeicherten Wert nicht überschreiben.
            _isShuttingDown = true;
            if (_taskPaneCreated && ToDoTaskPane != null)
                Properties.Settings.Default.Visible = ToDoTaskPane.Visible;
            Properties.Settings.Default.Save();
        }

        private void appControl_SizeChanged(object sender, EventArgs e)
        {
            _widthTimer.Stop();
            _widthTimer.Start();
        }

        private void ToDoTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (_isShuttingDown) return;
            if (_taskPaneCreated && ToDoTaskPane != null)
                Properties.Settings.Default.Visible = ToDoTaskPane.Visible;
            TodoRibbonAddIn rbn = Globals.Ribbons.TodoRibbonAddIn;
            if (rbn != null)
                rbn.btnToggleTodo.Checked = ToDoTaskPane != null && ToDoTaskPane.Visible;
        }

        private void ThisAddIn_Deactivate()
        {
            if (_isShuttingDown) return;
            if (_taskPaneCreated && ToDoTaskPane != null)
                Properties.Settings.Default.Visible = ToDoTaskPane.Visible;
            Properties.Settings.Default.Save();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            SystemEvents.PowerModeChanged -= OnPowerModeChanged;
            if (_refreshTimer != null) { _refreshTimer.Stop(); _refreshTimer.Dispose(); }
            if (_widthTimer   != null) { _widthTimer.Stop();   _widthTimer.Dispose(); }
            // Application_Quit hat bereits den korrekten Visible-Wert gespeichert —
            // kein Save() mehr wenn über Quit gefahren, sonst überschreibt VSTO-State den echten Wert.
            if (!_isShuttingDown)
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
