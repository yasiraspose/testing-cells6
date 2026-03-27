using System;
using Aspose.Cells;

namespace AsposeCellsSxcExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source SXC file
            string sourcePath = "input.sxc";

            // Load the existing SXC workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example modification: write a header and some data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(200);

            // Save the modified workbook back to SXC format
            string outputPath = "output.sxc";
            workbook.Save(outputPath, SaveFormat.Sxc);

            Console.WriteLine($"Workbook successfully read from '{sourcePath}', modified, and saved to '{outputPath}'.");
        }
    }
}