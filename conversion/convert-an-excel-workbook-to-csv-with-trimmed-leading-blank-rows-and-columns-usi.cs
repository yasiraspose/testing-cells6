using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel workbook (XLSX)
        string sourcePath = "input.xlsx";

        // Desired output CSV file path
        string destPath = "output.csv";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Configure CSV save options to trim leading blank rows and columns
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            TrimLeadingBlankRowAndColumn = true
        };

        // Save the workbook as CSV using the configured options
        workbook.Save(destPath, csvOptions);

        Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
    }
}