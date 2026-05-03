/**
 * @file    ThisAddIn.cs
 * @brief   VSTO Add-In Einstiegspunkt.
 * @author  Gerhard Lustig <gerhard@lustig.at>
 * @version 1.5.0
 * @date    2026-05-03
 * @history
 *   1.5.0  2026-05-03  _isShuttingDown: verhindert dass VSTO-Cleanup den Visible-State
 *                      nach Application_Quit überschreibt (Panel war nach Neustart weg).
 *   1.4.0  2026-05-03  OnPowerModeChanged: Retry-Logik nach Resume (max. 3 Versuche,
 *                      erster Versuch nach 5s, Folge-Versuche nach je 10s).
 *   1.3.0  2026-04-24  PowerModeChanged: bei Resume nach Sleep/Hibernate
 *                      wird nach 2s neu gezeichnet und RetrieveData() aufgerufen.
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
        private bool _taskPaneCreated  = false;
        private bool _isShuttingDown   = false;
        private System.Windows.Forms.Timer _refreshTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.AddRegistryNotification();

                // PowerModeChanged: nach Sleep/Hibernate neu aufbauen
                SystemEvents.PowerModeChanged += OnPowerModeChanged;

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

        private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode != PowerModes.Resume) return;
            TryResumeRefresh(5000, 3);
        }

        // Nach Resume: erst nach 5s versuchen (MAPI/Exchange braucht Zeit).
        // Bei Fehler (z.B. Exchange noch nicht bereit) bis zu retriesLeft-mal nach 10s wiederholen.
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
                    AppControl.RetrieveData();
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

                this.AppControl.Accounts  = Properties.Settings.Default.Accounts;
                this.AppControl.ShowTasks = Properties.Settings.Default.ShowTasks;

                ToDoTaskPane = this.CustomTaskPanes.Add(this.AppControl, " ");
                ToDoTaskPane.Visible              = Properties.Settings.Default.Visible;
                ToDoTaskPane.Width                = Properties.Settings.Default.Width;
                ToDoTaskPane.DockPosition         = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                ToDoTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                ToDoTaskPane.VisibleChanged       += ToDoTaskPane_VisibleChanged;
                this.AppControl.SizeChanged       += appControl_SizeChanged;

                _taskPaneCreated = true;

                // Kalender-Änderungen überwachen
                var calFolder = this.Application.Session.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
                    as Microsoft.Office.Interop.Outlook.Folder;
                if (calFolder != null)
                {
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemAdd
                        += (item) => { if (AppControl != null) AppControl.RetrieveData(); };
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemChange
                        += (item) => { if (AppControl != null) AppControl.RetrieveData(); };
                    ((Microsoft.Office.Interop.Outlook.ItemsEvents_Event)calFolder.Items).ItemRemove
                        += () => { if (AppControl != null) AppControl.RetrieveData(); };
                }

                this.AppControl.SelectedDate = DateTime.Today;

                ((Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)this.Application).Quit
                    += Application_Quit;

                var explorer = this.Application.ActiveExplorer();
                explorer.Deactivate += ThisAddIn_Deactivate;

                ((Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event)explorer).Activate
                    += Explorer_Activate;

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
            // Ab hier: Visible-State einfrieren. VSTO versteckt den Pane
            // gleich danach (löst VisibleChanged aus) — das darf den echten Wert nicht überschreiben.
            _isShuttingDown = true;
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
            // Application_Quit hat bereits den korrekten Wert gespeichert —
            // kein Save() mehr, sonst überschreibt VSTO-Shutdown-State den echten Wert.
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
