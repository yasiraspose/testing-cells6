using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Desired output path for the XPS file
        string destPath = "output.xps";

        // Convert the Excel workbook to XPS using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion from XLSX to XPS completed successfully.");
    }
}