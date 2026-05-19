using System;
using Microsoft.Office.Tools.Ribbon;

namespace Outlook2021TodoAddIn
{
    public partial class TodoRibbonAddIn
    {
        private void TodoRibbonAddIn_Load(object sender, RibbonUIEventArgs e)
        {
            if (Globals.ThisAddIn != null && Globals.ThisAddIn.ToDoTaskPane != null)
                btnToggleTodo.Checked = Globals.ThisAddIn.ToDoTaskPane.Visible;
        }

        private void btnToggleTodo_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn != null && Globals.ThisAddIn.ToDoTaskPane != null)
                Globals.ThisAddIn.ToDoTaskPane.Visible = btnToggleTodo.Checked;
        }
    }
}
