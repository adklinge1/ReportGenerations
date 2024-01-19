using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WindowsFormsApp1
{
    public static class ExcelReader
    {

    // The DOM approach.
    // Note that the code below works only for cells that contain numeric values.
    // 
    public static void ReadExcelFile(string fileName, string startCol = "A", string endCol = "I")
    {
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            string text;

            // Get the shared string table
            SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

            foreach (Row r in sheetData.Elements<Row>().Skip(1))
            {
                foreach (Cell c in r.Elements<Cell>())
                {
                    text = GetCellValue(c, sharedStringTable);
                }
            }
        }
    }

    private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
    {
        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
        {
            int sharedStringIndex = int.Parse(cell.InnerText);
            return sharedStringTable.ChildElements[sharedStringIndex].InnerText;
        }
        else
        {
            return cell.InnerText;
        }
    }

    }
}
