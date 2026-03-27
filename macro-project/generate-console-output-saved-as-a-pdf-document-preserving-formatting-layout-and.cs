using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Saving;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory Excel file)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["D1"].PutValue("Total");

            // Add a few rows of data
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Product {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 2);
                sheet.Cells[$"C{i}"].PutValue(10.5 * i);
                // Formula to calculate total = Quantity * Price
                sheet.Cells[$"D{i}"].Formula = $"=B{i}*C{i}";
            }

            // Apply simple formatting
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            AsposeRange headerRange = sheet.Cells.CreateRange("A1:D1");
            headerRange.ApplyStyle(headerStyle, new StyleFlag { All = true });

            // Auto‑fit columns for better layout
            sheet.AutoFitColumns();

            // Configure page setup (optional, affects PDF layout)
            PageSetup pageSetup = sheet.PageSetup;
            // Orientation property removed to avoid compatibility issues
            pageSetup.PaperSize = PaperSizeType.PaperA4;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 0; // let height adjust automatically

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each worksheet starts on a new page
                OnePagePerSheet = true
            };

            // Define output PDF path
            string pdfPath = "ReportOutput.pdf";

            // Save the workbook as PDF
            workbook.Save(pdfPath, pdfOptions);

            // Console output to indicate success and show some workbook info
            Console.WriteLine("Workbook has been created and saved as PDF.");
            Console.WriteLine($"PDF file location: {pdfPath}");
            Console.WriteLine($"Total worksheets: {workbook.Worksheets.Count}");
            Console.WriteLine($"First worksheet name: {workbook.Worksheets[0].Name}");
            Console.WriteLine($"Rows with data: {sheet.Cells.MaxDataRow + 1}");
        }
    }
}