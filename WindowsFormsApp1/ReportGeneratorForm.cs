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
        public ReportGeneratorForm()
        {
            InitializeComponent();

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

                folderPathTxtLabel.Text = $@"Found: {imagePaths.Length} images to load to report";
                // Create a new Word application
                Application wordApp = new Application();

                // Create a new document
                Document doc = string.IsNullOrWhiteSpace(templatePath) ? wordApp.Documents.Add() : wordApp.Documents.Open(templatePath);

                // TODO: add table after images
                AddTreesTable(doc, excelFilePath);

                // Insert an empty paragraph for spacing
                Paragraph spacingParagraph = doc.Paragraphs.Add();
                spacingParagraph.Range.InsertParagraphBefore();

                // ExtractAllImagesIntoDoc(imagePaths, doc);
                
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

        private static void ExtractAllImagesIntoDoc(string[] imagePaths, Document doc)
        {
            // Iterate through all sorted image files in the directory
            foreach (string imagePath in imagePaths)
            {
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
            }
        }

        private void AddTreesTable(Document doc, string excelFile)
        {
            List<Tree> trees = ExcelReader.ReadExcelFile(excelFile);
            folderPathTxtLabel.Text = $@"Read: {trees?.Count ?? 0} trees from excel file";

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

            tableColumns.Reverse();

            Table table = doc.Tables.Add(doc.Range(), NumRows: trees.Count + 1, NumColumns: tableColumns.Count);

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

            // Add data rows
            for (int i = 0; i < trees.Count; i++)
            {
                AddTreeToTable(table, trees[i], rowNumber: i + 2);
            }
        }

        static void AddTreeToTable(Table table, Tree tree, int rowNumber)
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
                sumOfValues <= 13 ? WdColor.wdColorGray10 :
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


        private void txtDirectoryPath_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void reportNameLabel_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }
    }
}
