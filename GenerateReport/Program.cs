using Microsoft.Office.Interop.Word;

namespace GenerateReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // if (args.Length != 1)
            // {
            //     Console.WriteLine("Usage: Program.exe <FolderPath>");
            //     return;
            // }

            // Create Word Application and Document
            Application wordApp = new Application();
            Document doc = wordApp.Documents.Add();

            string folderPath = @"C:\Users\adklinge\Downloads\potatos"; //args[0];

            // if (!Directory.Exists(folderPath))
            // {
            //     Console.WriteLine("Invalid folder path. Please provide a valid path.");
            //     return;
            // }
            //
            // string[] imageFiles = Directory.GetFiles(folderPath, "*.jpg"); // Change the file extension accordingly
            //
            // if (imageFiles.Length == 0)
            // {
            //     Console.WriteLine("No image files found in the specified folder.");
            //     return;
            // }
            //
            // // Create Word Application and Document
            // Application wordApp = new Application();
            // Document doc = wordApp.Documents.Add();
            //
            // // Insert pictures into the Word document
            // foreach (var imagePath in imageFiles)
            // {
            //     InlineShape inlineShape = doc.InlineShapes.AddPicture(imagePath, Type.Missing, Type.Missing, Type.Missing);
            //     // Adjust the picture size if needed
            //     inlineShape.Width = 300;
            //     inlineShape.Height = 200;
            // }
            //
            // // Save and close the Word document
            // string outputFilePath = Path.Combine(folderPath, "Report.docx");
            // doc.SaveAs2(outputFilePath);
            // doc.Close();
            // wordApp.Quit();
            //
            // Console.WriteLine($"Word document created at: {outputFilePath}");  // if (!Directory.Exists(folderPath))
            // {
            //     Console.WriteLine("Invalid folder path. Please provide a valid path.");
            //     return;
            // }
            //
            // string[] imageFiles = Directory.GetFiles(folderPath, "*.jpg"); // Change the file extension accordingly
            //
            // if (imageFiles.Length == 0)
            // {
            //     Console.WriteLine("No image files found in the specified folder.");
            //     return;
            // }
            //
            // // Create Word Application and Document
            // Application wordApp = new Application();
            // Document doc = wordApp.Documents.Add();
            //
            // // Insert pictures into the Word document
            // foreach (var imagePath in imageFiles)
            // {
            //     InlineShape inlineShape = doc.InlineShapes.AddPicture(imagePath, Type.Missing, Type.Missing, Type.Missing);
            //     // Adjust the picture size if needed
            //     inlineShape.Width = 300;
            //     inlineShape.Height = 200;
            // }
            //
            // // Save and close the Word document
            // string outputFilePath = Path.Combine(folderPath, "Report.docx");
            // doc.SaveAs2(outputFilePath);
            // doc.Close();
            // wordApp.Quit();
            //
            // Console.WriteLine($"Word document created at: {outputFilePath}");
        }
    }
}