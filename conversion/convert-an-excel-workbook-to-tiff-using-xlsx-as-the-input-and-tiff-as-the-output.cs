using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ConvertXlsxToTiff
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure image options for TIFF rendering
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,                     // Set output image type to TIFF
            HorizontalResolution = 300,                     // Optional: set DPI
            VerticalResolution = 300,
            TiffCompression = Aspose.Cells.Rendering.TiffCompression.CompressionLZW // Optional: set compression
        };

        // Create a SheetRender object with the worksheet and options
        SheetRender renderer = new SheetRender(worksheet, options);

        // Render the entire worksheet to a multi‑page TIFF file
        renderer.ToTiff("output.tiff");
    }
}