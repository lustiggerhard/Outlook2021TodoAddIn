using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Outlook2021TodoAddIn
{
    public class NoLineListView : ListView
    {
        private const int LVM_FIRST        = 0x1000;
        private const int LVM_SETGROUPINFO = LVM_FIRST + 147;
        private const int LVGF_STATE       = 0x00000004;
        private const int LVGS_COLLAPSIBLE = 0x00000008;
        private const int WM_MOUSEHOVER    = 0x02A1;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct LVGROUP
        {
            public uint   cbSize;
            public uint   mask;
            public IntPtr pszHeader;
            public int    cchHeader;
            public IntPtr pszFooter;
            public int    cchFooter;
            public int    iGroupId;
            public uint   stateMask;
            public uint   state;
            public uint   uAlign;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref LVGROUP lParam);

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_MOUSEHOVER && Groups.Count > 0)
            {
                m.Result = IntPtr.Zero;
                return;
            }
            base.WndProc(ref m);
        }

        public void RemoveGroupLines()
        {
            if (!IsHandleCreated) return;
            for (int i = 0; i < Groups.Count; i++)
            {
                LVGROUP grp   = new LVGROUP();
                grp.cbSize    = (uint)Marshal.SizeOf(typeof(LVGROUP));
                grp.mask      = LVGF_STATE;
                grp.stateMask = LVGS_COLLAPSIBLE;
                grp.state     = 0;
                SendMessage(Handle, LVM_SETGROUPINFO, new IntPtr(i), ref grp);
            }
        }
    }
}

