using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the text file that contains CSV data
        string txtFilePath = "input.txt";

        // Read the CSV content from the text file
        string csvContent = File.ReadAllText(txtFilePath);

        // Load the workbook from CSV content
        byte[] csvBytes = Encoding.UTF8.GetBytes(csvContent);
        using (MemoryStream ms = new MemoryStream(csvBytes))
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(ms, loadOptions);

            // Access the VBA project associated with the workbook (if any)
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is protected with a password
            bool isProtected = vbaProject != null && vbaProject.IsProtected;

            // Output the result
            Console.WriteLine($"VBA project protected: {isProtected}");
        }
    }
}