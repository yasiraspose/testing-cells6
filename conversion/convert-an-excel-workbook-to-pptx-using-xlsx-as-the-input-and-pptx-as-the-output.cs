using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel workbook (XLSX)
        string sourcePath = "input.xlsx";

        // Desired output PowerPoint file (PPTX)
        string destPath = "output.pptx";

        // Convert the Excel file to PPTX using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine($"Conversion completed successfully. Output file: {destPath}");
    }
}