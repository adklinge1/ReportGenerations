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
            // ExcelReader.ReadExcelFile("C:\\Users\\adklinge\\source\\repos\\GenerateReport\\WindowsFormsApp1\\ExcelTemplate.xlsx");
            Application.Run(new ReportGeneratorForm());
            //TreeCalculator t = new TreeCalculator();
            //Dictionary<string, double> x = t.FetchTreePricesAsync().Result;
            //
            //return;
        }
    }
}
