using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1
{
    public static class ExcelReader
    {
        public static List<Tree> ReadExcelFile(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Get the shared string table
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                IEnumerable<Row> rows = sheetData.Elements<Row>();
                int numberOfTrees = rows?.Count() ?? 0;

                if (numberOfTrees <= 0)
                {
                        return new List<Tree>(numberOfTrees);
                }

                List<Tree> trees = new List<Tree>(numberOfTrees);

                // Read all rows and aocnvert to Trees
                // TODO: Stop when all cells are empty
                foreach (Row r in rows.Skip(1))
                {
                    if (r.Elements<Cell>().All(c => GetCellValue(c, sharedStringTable) == string.Empty))
                    {
                        break;
                    }

                    trees.Add(ConvertRowToTree(r, sharedStringTable));
                }

                return trees;
            }
        }

    private static Tree ConvertRowToTree(Row row, SharedStringTable sharedString)
    {
        int index = -1;
        string species = string.Empty;
        double height = -1.0;
        double diameter = -1.0;
        double canopy = -1.0;
        int health = -1;
        int location = -1;
        int speciesValue = -1;
        double price = -1;
        foreach (Cell cell in row.Elements<Cell>())
        {
            string cellValue = GetCellValue(cell, sharedString);
            var columneName = GetColumnName(cell.CellReference.Value);

                switch (columneName)
                {
                    case "A":
                        index = int.Parse(cellValue);
                        break;
                    case "B":
                        species= cellValue;
                        break;
                    case "C":
                       height = double.Parse(cellValue);
                        break;
                    case "D":
                        diameter = double.Parse(cellValue);
                        break;  
                    case "E":
                        canopy = double.Parse(cellValue);
                        break;
                    case "F": 
                        health = int.Parse(cellValue);
                        break;
                    
                    case "G":
                        location = int.Parse(cellValue);
                        break;

                    case "H":
                        speciesValue = int.Parse(cellValue);
                        break;

                    case "I":
                        price = double.Parse(cellValue);
                        break;

                    default:
                        // throw new ArgumentException($"Excel table shoudn't have a cell at column : {columneName}");
                        break;
                }
        }

        return new Tree(index, species, height, diameter, health, canopy, location, speciesValue, price);

    }

    static string GetColumnName(string cellReference)
    {
        // Extracts the column name from the cell reference (e.g., "A1" -> "A")
        return new String(cellReference.Where(Char.IsLetter).ToArray());
    }

        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                int sharedStringIndex = int.Parse(cell.InnerText);
                return sharedStringTable.ChildElements[sharedStringIndex].InnerText;
            }

            return cell.InnerText;
        }

    }
}
