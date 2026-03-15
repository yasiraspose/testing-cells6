using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook using the constructor (lifecycle rule)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Write basic workbook information to the console
        Console.WriteLine($"Workbook loaded from: {inputPath}");
        Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

        // Iterate through each worksheet and display its name and the value of cell A1
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            Console.WriteLine($"Worksheet {i} name: {sheet.Name}");

            // Retrieve the value of cell A1
            var cell = sheet.Cells["A1"];
            if (cell.Value != null)
            {
                Console.WriteLine($"  A1 = {cell.Value}");
            }
            else
            {
                Console.WriteLine("  A1 is empty");
            }
        }

        // Save a copy of the workbook (optional) using the Save method (lifecycle rule)
        string outputPath = "output_copy.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved as a copy to: {outputPath}");
    }
}