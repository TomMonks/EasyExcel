using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace TestEasyExcelForms
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var frm = new Form1();
            frm.ShowGrid();
            Application.Run(frm);
            
        }
    }
}
