using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            TreeCalculator treeCalc = new TreeCalculator();
            treeCalc.LoadTreePricesAsync().Wait();
            Application.Run(new ReportGeneratorForm(treeCalc));
        }
    }
}
