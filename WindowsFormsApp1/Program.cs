using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Windows.Forms;

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


            string filePath = "C:\\Users\\adklinge\\Downloads\\SekerFormula.xlsx";
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                string columnName = "מספר סידורי\r\n"; // Replace with the actual column name
                int rowNumber = 2; // Replace with the actual row number

                string cellValue = ExcelReader.GetCellValueByColumnAndRow(columnName, rowNumber, spreadsheetDocument);

                if (cellValue != null)
                {
                    Console.WriteLine($"Cell value at {columnName}{rowNumber}: {cellValue}");
                }
                else
                {
                    Console.WriteLine($"Cell not found at {columnName}{rowNumber}");
                }
            }


            Application.Run(new ReportGeneratorForm());
        }


    }

class ExcelReader
    {
        public static void ReadExcelTable(string filePath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First(); // Assuming there is only one worksheet

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                foreach (Row row in sheetData.Elements<Row>())
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(cell, workbookPart);
                        Console.Write(cellValue + "\t");
                    }
                    Console.WriteLine();
                }
            }
        }

        public static string GetCellValueByColumnAndRow(string columnName, int rowNumber, SpreadsheetDocument spreadsheetDocument)
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowNumber);

            if (row != null)
            {
                Cell cell = row.Elements<Cell>().FirstOrDefault(c => GetColumnName(c.CellReference.Value) == columnName);
                if (cell != null)
                {
                    return GetCellValue(cell, workbookPart);
                }
            }

            return null; // Cell not found
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
            string cellValue;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringIndex = int.Parse(cell.InnerText);
                cellValue = stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
            }
            else
            {
                cellValue = cell.InnerText;
            }

            return cellValue;
        }

        private static string GetColumnName(string cellReference)
        {
            // Extracts the column name from the cell reference (e.g., "A1" -> "A")
            return new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        }
    }
}
