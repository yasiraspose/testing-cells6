using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (create rule is used internally by the constructor)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Manipulate the worksheet: add a header and some data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            // Add a simple formula to calculate total price
            sheet.Cells["D1"].PutValue("Total");
            sheet.Cells["D2"].Formula = "B2*C2";
            sheet.Cells["D3"].Formula = "B3*C3";

            // Ensure formulas are calculated before saving (save rule)
            workbook.CalculateFormula();

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export document structure for better accessibility
                ExportDocumentStructure = true,

                // Set PDF/A-1b compliance
                Compliance = PdfCompliance.PdfA1b,

                // Use Flate compression for smaller file size
                PdfCompression = PdfCompressionCore.Flate,

                // Show the document title in the PDF viewer title bar
                DisplayDocTitle = true,

                // Set a custom producer string
                Producer = "Aspose.Cells for .NET Demo"
            };

            // Optional: add security settings (owner/user passwords and permissions)
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                OwnerPassword = "ownerPass123",
                UserPassword = "userPass123",
                PrintPermission = true,
                ModifyDocumentPermission = false,
                ExtractContentPermission = false,
                AnnotationsPermission = true,
                FillFormsPermission = true
            };
            pdfOptions.SecurityOptions = security;

            // Save the manipulated workbook as a PDF (save rule)
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("Excel file has been read, modified, and saved as PDF with specified options.");
        }
    }
}