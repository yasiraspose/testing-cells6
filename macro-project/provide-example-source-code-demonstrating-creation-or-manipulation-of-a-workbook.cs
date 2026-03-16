using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class FodsDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(30);

        // Configure ODS save options for Flat OpenDocument Spreadsheet (FODS)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice; // Set generator type
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12; // Optional ODF version

        // Save the workbook as a .fods file using the configured options
        workbook.Save("Sample.fods", saveOptions);

        Console.WriteLine("FODS file created successfully.");
    }
}