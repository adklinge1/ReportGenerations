using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WindowsFormsApp1.Models;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.IO;

namespace WindowsFormsApp1.ExcelReader 
{
    public static class ExcelReader
    {
        public static List<Tree> ReadExcelFile(string fileName)
        {
            Dictionary<string, TreeSpecie> tressSpecies = TryReadTreeSpecies();

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

                // Read all rows and convert to Trees
                foreach (Row r in rows.Skip(1))
                {
                    if (r.Elements<Cell>().All(c => GetCellValue(c, sharedStringTable) == string.Empty))
                    {
                        break;
                    }

                    Tree tree = ConvertRowToTree(r, sharedStringTable, tressSpecies);
                    trees.Add(tree);
                }

                return trees;
            }
        }

        private static Dictionary<string, TreeSpecie> TryReadTreeSpecies(string path = @"ExcelReader\TreeSpeciesToScientificNameAndRate.csv")
        {
            Dictionary<string, TreeSpecie> treeNameToScientificNameAndValueDic = new Dictionary<string, TreeSpecie>();

            try
            {
                using (var reader = new StreamReader(path))
                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                {
                    IEnumerable<TreeSpecie> records = csv.GetRecords<TreeSpecie>();
                    foreach (TreeSpecie record in records)
                    {
                        treeNameToScientificNameAndValueDic[record.HebrewName] = record;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return treeNameToScientificNameAndValueDic;
        }

        private static Tree ConvertRowToTree(Row row, SharedStringTable sharedString, Dictionary<string, TreeSpecie> treesSpecies)
        {
        int index = -1;
        string species = string.Empty;
        double height = -1.0;
        double diameter = -1.0;
        double canopy = -1.0;
        int health = -1;
        int location = -1;
        int speciesRate = -1;
        double price = -1;
        string scientificName = "Unknown";
        int numberOfStems = 1;
        bool isTserifi = false;
        string comment = String.Empty;
            
        foreach (Cell cell in row.Elements<Cell>())
        {
            string cellValue = GetCellValue(cell, sharedString).Trim();

            if (string.IsNullOrWhiteSpace(cellValue))
            {
                continue;
            }

            var columneName = GetColumnName(cell.CellReference.Value);

            try
            {
                switch (columneName)
                {
                    case "A":
                        index = int.Parse(cellValue);
                        break;
                    case "B":
                        species = cellValue;
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
                        speciesRate = int.Parse(cellValue);
                        break;

                    case "I":
                        scientificName = cellValue;
                        break;

                    case "J":
                        comment = cellValue;
                        break;

                    case "K":
                        numberOfStems = !string.IsNullOrWhiteSpace(cellValue) && int.TryParse(cellValue, out int numOfStems) ? numOfStems : 1;
                        break;
                    case "L":
                        isTserifi = !string.IsNullOrWhiteSpace(cellValue) && int.TryParse(cellValue, out int tserifi) && tserifi == 1;
                        break;

                    default:
                        throw new ArgumentException($"Excel table shoudn't have a cell at column : {columneName}");
                    }
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Column :{columneName} in row {row.RowIndex} can not be: {cellValue}\n{ex}");
            }
        }

        // Infer species scientific name and value from internal map
        if (treesSpecies.TryGetValue(species, out TreeSpecie spicie))
        {
            scientificName = spicie.ScientificName;
            speciesRate = spicie.SpeciesRate;
        }

        return new Tree(index, species, height, diameter, health, canopy, location, speciesRate, price, scientificName, numberOfStems, isTserifi)
        {
            Comments = comment
        };
    }


        public static string GetColumnName(string cellReference)
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
