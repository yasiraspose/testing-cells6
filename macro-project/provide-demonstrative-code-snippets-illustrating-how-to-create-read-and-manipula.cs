using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class XpsDemo
{
    static void Main()
    {
        // Create a workbook and save it as XPS
        CreateAndSaveXps();

        // Read the generated XPS file into memory
        byte[] xpsData = ReadXpsFile("DemoDocument.xps");
        Console.WriteLine($"Read XPS file, size: {xpsData.Length} bytes");

        // Modify the workbook and re‑save with different XPS options
        ManipulateAndResaveXps();
    }

    static void CreateAndSaveXps()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample data
        sheet.Cells["A1"].PutValue("Aspose.Cells XPS Demo");
        sheet.Cells["A2"].PutValue(DateTime.Now);
        sheet.Cells["B2"].PutValue(12345);
        sheet.Cells["A3"].PutValue("Another row");
        sheet.Cells["B3"].PutValue(67890);

        // Initialize XpsSaveOptions using the provided constructor
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        // Configure desired options
        saveOptions.OnePagePerSheet = true;               // each sheet on a single page
        saveOptions.DefaultFont = "Arial";                // fallback font
        saveOptions.PageIndex = 0;                        // start from first page
        saveOptions.PageCount = 1;                        // save only one page
        saveOptions.AllColumnsInOnePagePerSheet = true;   // fit all columns on one page

        // Save the workbook as XPS using the options (provided save rule)
        workbook.Save("DemoDocument.xps", saveOptions);
        Console.WriteLine("Workbook saved as XPS: DemoDocument.xps");
    }

    static byte[] ReadXpsFile(string path)
    {
        // Load the XPS file into a byte array (read operation)
        return File.ReadAllBytes(path);
    }

    static void ManipulateAndResaveXps()
    {
        // Load a new workbook (could also load an existing Excel file)
        Workbook workbook = new Workbook();

        // Add or modify data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Modified Content");
        sheet.Cells["B1"].PutValue(999);

        // Create XpsSaveOptions with different settings
        XpsSaveOptions options = new XpsSaveOptions
        {
            OnePagePerSheet = false,               // allow multiple pages per sheet
            DefaultFont = "Times New Roman",
            GridlineType = GridlineType.Dotted,
            TextCrossType = TextCrossType.Default,
            PrintingPageType = PrintingPageType.Default
        };

        // Save the modified workbook as a new XPS document
        workbook.Save("ModifiedDemoDocument.xps", options);
        Console.WriteLine("Modified workbook saved as XPS: ModifiedDemoDocument.xps");
    }
}