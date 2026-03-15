using System;
using Aspose.Cells;

namespace AsposeCellsXmlConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the resulting XML file
            string outputPath = "output.xml";

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(sourcePath);

            // Create XML save options (default settings are sufficient for a basic conversion)
            XmlSaveOptions xmlSaveOptions = new XmlSaveOptions();

            // Save the workbook as an XML file using the specified options
            workbook.Save(outputPath, xmlSaveOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been successfully converted to XML at '{outputPath}'.");
        }
    }
}