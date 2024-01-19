using System;
using System.IO;
using System.Windows.Forms;
using Path = System.IO.Path;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Linq;

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
            string directoryPath = txtDirectoryPath.Text;
            string templatePath = templatePathTextBox.Text;
            string reportName = reportNameTextBox.Text;

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

                    // Iterate through all sorted image files in the directory
                    foreach (string imagePath in imagePaths)
                    {
                        // Get the image name without extension
                        string imageName = Path.GetFileNameWithoutExtension(imagePath);

                        // Add content to the Word document
                        Paragraph imageNameParagraph = doc.Paragraphs.Add();
                        Range imageNameRange = imageNameParagraph.Range;
                        imageNameRange.Text = imageName;  // Set the image name

                        // Add content to the Word document
                        Paragraph paragraph = doc.Paragraphs.Add();
                        Range range = paragraph.Range;

                        // Load the image file and add it to the Word document
                        InlineShape inlineShape = doc.InlineShapes.AddPicture(imagePath, Range: range);
                    
                        // Optionally, adjust the size of the image
                        inlineShape.LockAspectRatio = MsoTriState.msoFalse;
                        inlineShape.Width = 300;  // Set the width as needed
                        inlineShape.Height = 200; // Set the height as needed

                        // Release COM objects to avoid memory leaks
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(inlineShape);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(paragraph);
                    }

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

        private string[] ValidateAndExtractImageFiles(string directoryPath)
        {
            // Get all image files in the directory and sort them by name
            string[] imagePaths = Directory.GetFiles(directoryPath, "*.jpg", SearchOption.TopDirectoryOnly)
                .Concat(Directory.GetFiles(directoryPath, "*.jpeg", SearchOption.TopDirectoryOnly))
                .Concat(Directory.GetFiles(directoryPath, "*.png", SearchOption.TopDirectoryOnly))
                .ToArray();

            string[] invalidImageNames = imagePaths?.Where(img => int.TryParse(Path.GetFileNameWithoutExtension(img), out _) == false).ToArray();

            if (invalidImageNames.Any())
            {
                MessageBox.Show($"Image file must have a numeric name. Invalid image names : {string.Join(", ", invalidImageNames)}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new string[]{};
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
    }
}
