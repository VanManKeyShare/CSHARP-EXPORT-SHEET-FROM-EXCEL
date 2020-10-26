using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace EXPORT_SHEET_FROM_EXCEL
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Frm_Main());
        }
    }
}
