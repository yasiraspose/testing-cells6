using System;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the generated HTML file
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options (optional: set HTML version to HTML5)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HtmlVersion = HtmlVersion.Html5;

            // Save the workbook as HTML using the options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been converted to HTML at '{outputPath}'.");
        }
    }
}