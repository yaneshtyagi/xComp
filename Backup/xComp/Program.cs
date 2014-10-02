using System;
using System.Collections.Generic;
using System.Windows.Forms;
using yCompnents.OfficeTools.xComp;

namespace xComp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length == 2)
                Application.Run(new frmXComp(args[0], args[1]));
            else
                Application.Run(new frmXComp());
        }
    }
}