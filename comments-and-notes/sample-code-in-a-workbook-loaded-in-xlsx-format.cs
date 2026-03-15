using System;
using Aspose.Cells;

namespace AsposeCellsSample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Write some data to cells
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue(DateTime.Now);
            sheet.Cells["C1"].PutValue(123.45);

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook loaded from '{inputPath}', modified, and saved as '{outputPath}'.");
        }
    }
}