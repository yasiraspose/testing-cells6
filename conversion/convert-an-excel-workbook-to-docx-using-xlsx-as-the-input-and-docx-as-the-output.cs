using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ExcelToDocxConverter
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Desired output path for the DOCX file
        string destinationPath = "output.docx";

        // Convert the Excel workbook to DOCX using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destinationPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}