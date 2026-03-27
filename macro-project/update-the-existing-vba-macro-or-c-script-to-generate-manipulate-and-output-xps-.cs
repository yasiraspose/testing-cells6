using System;
using Aspose.Cells;

namespace AsposeCellsXpsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("XPS Save Options Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // Initialize XpsSaveOptions
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                OnePagePerSheet = true,
                DefaultFont = "Arial",
                CheckWorkbookDefaultFont = true,
                CheckFontCompatibility = true,
                IsFontSubstitutionCharGranularity = true,
                AllColumnsInOnePagePerSheet = true,
                IgnoreError = false,
                OutputBlankPageWhenNothingToPrint = false,
                PageIndex = 0,
                PageCount = 1,
                PrintingPageType = PrintingPageType.Default,
                GridlineType = GridlineType.Dotted,
                TextCrossType = TextCrossType.Default,
                DefaultEditLanguage = DefaultEditLanguage.English
                // SheetSet property removed as it is not available in this version
            };

            // Save the workbook as XPS
            string outputPath = "XpsSaveOptionsDemo.xps";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved as XPS to '{outputPath}'.");
        }
    }
}