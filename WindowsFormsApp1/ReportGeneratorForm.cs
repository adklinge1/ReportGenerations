using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Path = System.IO.Path;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Linq;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1
{
    public partial class ReportGeneratorForm : Form
    {
        private readonly TreeCalculator _treeCalculator;

        public ReportGeneratorForm(TreeCalculator treeCalculator)
        {
            InitializeComponent();
            _treeCalculator = treeCalculator;

            // Initialize the controls
            statusStrip1 = new StatusStrip();
            toolStripStatusLabel1 = new ToolStripStatusLabel();

            // Add the controls to the form
            Controls.Add(statusStrip1);
            statusStrip1.Items.Add(toolStripStatusLabel1);

            // Configure ToolStripStatusLabel
            toolStripStatusLabel1.Text = @"© 2023 Created with ♥ for Yarden The Agronomist. All rights reserved."; ;
        }

        private void btnGenerateDocument_Click(object sender, EventArgs e)
        {
            string directoryPath = txtDirectoryPath.Text?.TrimEnd('"')?.TrimStart('"');
            string templatePath = templatePathTextBox.Text?.TrimEnd('"')?.TrimStart('"');
            string reportName = reportNameTextBox.Text;
            string excelFilePath = excelFilePahTextBox.Text?.TrimEnd('"')?.TrimStart('"'); 

            // Save the document inside the provided directoryPath with the name "GeneratedDocument.docx"
            string outputReportPath = Path.Combine(directoryPath, $"{reportName}.docx");

            // Check if the document exists/ open, and pop en error message
            if (File.Exists(outputReportPath))
            {
                MessageBox.Show(@"Document already exists.Please delete it, or choose a different report name", "Document already exists", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(directoryPath))
            {
                MessageBox.Show("Image folder does not exist", "Folder not exists", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                return;
            }

            // Check if the template document exists
            if (!File.Exists(templatePath))
            {
                MessageBox.Show(@"Template document not found. Generating report on new document", "Temlpate not exists", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                templatePath = null;

            }

            // check that the template is a word document
            else if (!templatePath.EndsWith(".docx"))
            {
                MessageBox.Show(@"Template document must be a *.docx file", "Incompatible template", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                templatePath = null;
            }

            // Check if excel file exists
            if (!File.Exists(excelFilePath))
            {
                MessageBox.Show(@"There is no such excel file", "Excel file doesn't already exists", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Get all image files in the directory and sort them by name
                string[] imagePaths = ValidateAndExtractImageFiles(directoryPath);

                if (imagePaths.Length == 0)
                {
                    return;
                }

                statusTxtLabel.Text = $@"Found: {imagePaths.Length} images to load to report";
                // Create a new Word application
                Application wordApp = new Application();

                // Create a new document
                Document doc = string.IsNullOrWhiteSpace(templatePath) ? wordApp.Documents.Add() : wordApp.Documents.Open(templatePath);

                // TODO: add table after images
                List<Tree> trees = ExcelReader.ExcelReader.ReadExcelFile(excelFilePath);

                if (trees.Count != imagePaths.Length)
                {
                    MessageBox.Show($@"The number of trees in the excel ({trees.Count}) does not match the numbers of images in folder ({imagePaths.Length})", @"Inconsistent number of trees", MessageBoxButtons.OKCancel | MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // Insert an empty paragraph for spacing
                doc.Paragraphs.Add().Range.InsertParagraphBefore();

                ExtractAllImagesIntoDoc(imagePaths, doc);

                // Insert an empty paragraph for spacing
                doc.Paragraphs.Add().Range.InsertParagraphBefore();

                AddTreesTables(doc, trees);


                doc.SaveAs2(outputReportPath);

                // Close Word application
                wordApp.Quit();

                // Release COM objects to avoid memory leaks
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

                // Display the output document
                System.Diagnostics.Process.Start(outputReportPath);

                MessageBox.Show($@"Report was successfully created in :{outputReportPath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExtractAllImagesIntoDoc(string[] imagePaths, Document doc)
        {
            int i = 1;
            // Iterate through all sorted image files in the directory
            foreach (string imagePath in imagePaths)
            {
                statusTxtLabel.Text = $@"Extracting image {i}";

                // Get the image name without extension
                string imageName = Path.GetFileNameWithoutExtension(imagePath);

                // Add content to the Word document
                Paragraph imageNameParagraph = doc.Paragraphs.Add();
                Range imageNameRange = imageNameParagraph.Range;
                imageNameRange.Text = imageName; // Set the image name

                // Add content to the Word document
                Paragraph paragraph = doc.Paragraphs.Add();
                Range range = paragraph.Range;

                // Load the image file and add it to the Word document
                InlineShape inlineShape = doc.InlineShapes.AddPicture(imagePath, Range: range);

                // Optionally, adjust the size of the image
                inlineShape.LockAspectRatio = MsoTriState.msoFalse;
                inlineShape.Width = 300; // Set the width as needed
                inlineShape.Height = 200; // Set the height as needed

                // Release COM objects to avoid memory leaks
                System.Runtime.InteropServices.Marshal.ReleaseComObject(inlineShape);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(paragraph);

                i++;
            }
        }

        private void AddTreesTables(Document doc, List<Tree> trees)
        {
            statusTxtLabel.Text = $@"Reads: {trees?.Count ?? 0} trees from excel file";

            foreach (var tree in trees)
            {
                if (tree.PriceInNis <= 0)
                {
                    tree.PriceInNis = _treeCalculator.TryToGetTreePrice(tree) ?? tree.PriceInNis;
                }
            }

            AddFirstTable(doc, trees);

            AddTitleParagraph(doc, "טבלת ערכיות העצים");

            
            AddSecondTable(doc, trees);
            AddThirdTable(doc, trees);
        }

        private void AddThirdTable(Document doc, List<Tree> trees)
        {
            statusTxtLabel.Text = @"Building third table";

            var tableColumns = new List<string>(15)
            {
                "מספר סידורי",
                "מין העץ ",
                "כמות עצים",
                "גובה העץ",
                "מספר גזעים",
                "קוטר גזע",
                "(סך ערכיות העץ (0-20",
                "ערכיות העץ",
                "סטטוס מוצע"
            };

            Table table = CreateTableWithHeaders(
                doc,
                rows: trees.Count + 1,
                cols: tableColumns.Count,
                tableColumns,
                title: "טבלת מידע מרכזת");

            int rowNumber = 2;

            foreach (Tree tree in trees)
            {
                Row newRow = table.Rows[rowNumber];

                // Set the text for each cell in the row.
                newRow.Cells[9].Range.Text = tree.Index.ToString();
                newRow.Cells[8].Range.Text = tree.Species;
                newRow.Cells[7].Range.Text = tree.Quatity.ToString();
                newRow.Cells[6].Range.Text = tree.Height.ToString();
                newRow.Cells[5].Range.Text = tree.NumberOfStems.ToString();
                newRow.Cells[4].Range.Text = tree.StemDiameter.ToString();
                newRow.Cells[3].Range.Text = tree.SumOfValues.ToString();
                newRow.Cells[2].Range.Text = tree.TreeEvaluation;
                newRow.Cells[1].Range.Text = tree.Status;

                // Color by evaluation
                newRow.Cells[2].Range.Shading.BackgroundPatternColor = tree.TreeEvaluation == TreeEvaluations.Low ? WdColor.wdColorYellow :
                    tree.TreeEvaluation == TreeEvaluations.Medium ? WdColor.wdColorGray30 :
                    tree.TreeEvaluation == TreeEvaluations.High ? WdColor.wdColorGreen :
                    WdColor.wdColorRed;

                rowNumber++;
            }
        }

        private void AddSecondTable(Document doc, List<Tree> trees)
        {
            statusTxtLabel.Text = @"Building second table";

            var tableColumns = new List<string>()
            {
                "מין העץ",
                "שם מדעי",
                "ערכיות גבוהה מאד",
                "ערכיות גבוהה",
                "ערכיות בינונית",
                "ערכיות נמוכה",
                "סכהכ"
            };

            IEnumerable<IGrouping<string, Tree>> groupsBySpecie = trees.GroupBy(t => t.Species);

            // Number of rows is number of distinct species + 2 summary rows + header
            int numberOfRows = groupsBySpecie.Count() + 2 + 1;

            Table table = CreateTableWithHeaders(doc, rows: numberOfRows, cols: tableColumns.Count, tableColumns, "טבלת סיכום סיווג");

            int rowNumber = 2;
            foreach (var specieGroup in groupsBySpecie)
            {
                Row newRow = table.Rows[rowNumber];

                // Set the text for each cell in the row.
                newRow.Cells[7].Range.Text = specieGroup.Key;
                newRow.Cells[6].Range.Text = specieGroup.FirstOrDefault()?.ScientificName ?? "< Unknown >";
                newRow.Cells[5].Range.Text = specieGroup.Count(t => t.TreeEvaluation == TreeEvaluations.VeryHigh).ToString();
                newRow.Cells[4].Range.Text = specieGroup.Count(t => t.TreeEvaluation == TreeEvaluations.High).ToString();
                newRow.Cells[3].Range.Text = specieGroup.Count(t => t.TreeEvaluation == TreeEvaluations.Medium).ToString();
                newRow.Cells[2].Range.Text = specieGroup.Count(t => t.TreeEvaluation == TreeEvaluations.Low).ToString();
                newRow.Cells[1].Range.Text = specieGroup.Count().ToString();

                rowNumber++;
            }

            string[] treeEvaluations = { TreeEvaluations.VeryHigh, TreeEvaluations.High, TreeEvaluations.Medium, TreeEvaluations.Low };
            IEnumerable<IGrouping<string, Tree>> groupsByEvaluation = trees.GroupBy(t => t.TreeEvaluation);
            Dictionary<string, int> countByEvaluationDic = treeEvaluations
                .ToDictionary(e => e, e => groupsByEvaluation.FirstOrDefault(g => g.Key == e)?.Count() ?? 0);

            // Write count summary row
            Row countSummaryRow = table.Rows[rowNumber];

            countSummaryRow.Cells[5].Range.Text = countByEvaluationDic[TreeEvaluations.VeryHigh].ToString();
            countSummaryRow.Cells[4].Range.Text = countByEvaluationDic[TreeEvaluations.High].ToString();
            countSummaryRow.Cells[3].Range.Text = countByEvaluationDic[TreeEvaluations.Medium].ToString();
            countSummaryRow.Cells[2].Range.Text = countByEvaluationDic[TreeEvaluations.Low].ToString();
            countSummaryRow.Cells[1].Range.Text = trees.Count.ToString();
            rowNumber++;

            // Write percentage summary row
            Row percentageSummaryRow = table.Rows[rowNumber];
            int totalNumberOfTress = trees.Count;
            percentageSummaryRow.Cells[5].Range.Text = $"{Math.Round(100.0 * countByEvaluationDic[TreeEvaluations.VeryHigh] / totalNumberOfTress, 2)} %";
            percentageSummaryRow.Cells[4].Range.Text = $"{Math.Round(100.0 * countByEvaluationDic[TreeEvaluations.High] / totalNumberOfTress, 2)} %";
            percentageSummaryRow.Cells[3].Range.Text = $"{Math.Round(100.0 * countByEvaluationDic[TreeEvaluations.Medium] / totalNumberOfTress, 2)} %";
            percentageSummaryRow.Cells[2].Range.Text = $"{Math.Round(100.0 * countByEvaluationDic[TreeEvaluations.Low] / totalNumberOfTress, 2)} %";
            percentageSummaryRow.Cells[1].Range.Text = "100 %";

            // color the header
            Row headerRow = table.Rows[1];
            headerRow.Cells[5].Range.Shading.BackgroundPatternColor = WdColor.wdColorRed;
            headerRow.Cells[4].Range.Shading.BackgroundPatternColor = WdColor.wdColorGreen;
            headerRow.Cells[3].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray30;
            headerRow.Cells[2].Range.Shading.BackgroundPatternColor = WdColor.wdColorYellow;
        }

        private void AddFirstTable(Document doc, List<Tree> trees)
        {
            statusTxtLabel.Text = @"Building first table";

            var tableColumns = new List<string>(15)
            {
                "מספר סידורי",
                "מין העץ ",
                "כמות עצים",
                "גובה העץ",
                "קוטר גזע",
                "(מצב בריאות (0-5",
                "(מיקום העץ (0-5",
                "(ערך מין העץ (0-5",
                "(חופת העץ (0-5",
                "(סך ערכיות העץ (0-20",
                "שווי העץ (₪)",
                "(אזור שורשים מוגן (רדיוס במטרים",
                "היתכנות העתקה ",
                "הערות",
                "סטטוס מוצע"
            };

            Table table = CreateTableWithHeaders(doc, rows: trees.Count + 1, cols: tableColumns.Count, tableColumns);

            // Add data rows
            for (int i = 0; i < trees.Count; i++)
            {
                AddTreeToFirstTable(table, trees[i], rowNumber: i + 2);
            }
        }

        private static void AddTitleParagraph(Document doc, string title, int size = 16, int bold = 1, WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphCenter)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return;
            }
            // Add a title to the document
            Paragraph titleParagraph = doc.Paragraphs.Add();
            titleParagraph.Range.Text = title;

            // Format the title (optional)
            titleParagraph.Range.Font.Bold = bold; // Make the text bold
            titleParagraph.Range.Font.Size = size; // Set the font size
            titleParagraph.Range.ParagraphFormat.Alignment = alignment;
        }

        private static Table CreateTableWithHeaders(Document doc, int rows, int cols, List<string> tableColumns, string title = "")
        {
             AddTitleParagraph(doc, title);

            // revers since this is hebrew
            tableColumns.Reverse();

            // Add a new paragraph before adding the new table
            doc.Paragraphs.Add();

            // Get the range after the last paragraph
            Range range = doc.Paragraphs[doc.Paragraphs.Count].Range;

            Table table = doc.Tables.Add(range, NumRows: rows, NumColumns: cols);

            // Make the table borders visible
            table.Borders.Enable = 1;
            table.Range.Font.Size = 8;

            // Add header row
            for (int i = 1; i <= tableColumns.Count; i++)
            {
                Cell headerCell = table.Cell(Row: 1, Column: i);
                headerCell.Range.Text = tableColumns[i - 1];

                // Make the text in the header cell bold
                headerCell.Range.Font.Bold = 1;
            }

            return table;
        }

        static void AddTreeToFirstTable(Table table, Tree tree, int rowNumber)
        {
            Row newRow = table.Rows[rowNumber];

            // Set the text for each cell in the row.
            newRow.Cells[15].Range.Text = tree.Index.ToString();
            newRow.Cells[14].Range.Text = tree.Species;
            newRow.Cells[13].Range.Text = tree.Quatity.ToString();
            newRow.Cells[12].Range.Text = tree.Height.ToString();
            newRow.Cells[11].Range.Text = tree.StemDiameter.ToString();
            newRow.Cells[10].Range.Text = tree.HealthRate.ToString();
            newRow.Cells[9].Range.Text = tree.LocationRate.ToString();
            newRow.Cells[8].Range.Text = tree.SpeciesRate.ToString();
            newRow.Cells[7].Range.Text = tree.CanopyRate.ToString();
            newRow.Cells[6].Range.Text = tree.SumOfValues.ToString();
            newRow.Cells[5].Range.Text = tree.PriceInNis.ToString();
            newRow.Cells[4].Range.Text = tree.RootsAreaRadiusInMeters.ToString();
            newRow.Cells[3].Range.Text = tree.Clonability;
            newRow.Cells[2].Range.Text = tree.Comments;
            newRow.Cells[1].Range.Text = tree.Status;
            
            // Set the font size for each cell.
            foreach (Cell cell in newRow.Cells)
            {
                cell.Range.Font.Size = 8; // Set the desired font size.
            }

            // Set color to SumOfValues: 0-6 (yellow), 7-13(gray), 14-16(green), 17-20 (red)
            int sumOfValues = tree.SumOfValues;
            
            newRow.Cells[6].Range.Shading.BackgroundPatternColor = sumOfValues <= 6 ? WdColor.wdColorYellow :
                sumOfValues <= 13 ? WdColor.wdColorGray30 :
                sumOfValues <= 16 ? WdColor.wdColorGreen :
                WdColor.wdColorRed;

        }

        private string[] ValidateAndExtractImageFiles(string directoryPath)
        {
            // Get all image files in the directory and sort them by name
            string[] imagePaths = Directory.GetFiles(directoryPath, "*.jpg", SearchOption.AllDirectories)
                .Concat(Directory.GetFiles(directoryPath, "*.jpeg", SearchOption.AllDirectories))
                .Concat(Directory.GetFiles(directoryPath, "*.png", SearchOption.AllDirectories))
                .ToArray();

            string[] invalidImageNames = imagePaths?.Where(img => int.TryParse(Path.GetFileNameWithoutExtension(img), out _) == false).ToArray();

            if (invalidImageNames.Any())
            {
                MessageBox.Show($"Image file must have a numeric name. Invalid image names : {string.Join(", ", invalidImageNames)}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new string[] { };
            }

            if (imagePaths.Length == 0)
            {
                MessageBox.Show("No images found in folder. Please select the folder containing the images of the report", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new string[] { };
            }

            // Order numerically
            return imagePaths.OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f))).ToArray();
        }

        private void browseImgFolder_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] imagePaths = Directory.GetFiles(fbd.SelectedPath, "*.jpg", SearchOption.AllDirectories)
                        .Concat(Directory.GetFiles(fbd.SelectedPath, "*.jpeg", SearchOption.AllDirectories))
                        .Concat(Directory.GetFiles(fbd.SelectedPath, "*.png", SearchOption.AllDirectories))
                        .ToArray();

                    txtDirectoryPath.Text = fbd.SelectedPath;
                    statusTxtLabel.Text = $@"Found {imagePaths.Length} images in folder";
                }
            }
        }

        private void selectExcelFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Select excel file",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "txt",
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = false,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                excelFilePahTextBox.Text = openFileDialog1.FileName;
            }
        }

        private void templateBrowser_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Select Template File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "docs",
                Filter = "Word File (.docx ,.doc)|*.docx;",

                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = false,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                templatePathTextBox.Text = openFileDialog1.FileName;
            }
        }

        private void ReportGeneratorForm_Load(object sender, EventArgs e)
        {

        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void reportNameLabel_Click(object sender, EventArgs e)
        {

        }

        private void reportNameTextBox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
