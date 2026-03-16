using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Utility;

class AdvancedXlsxHandling
{
    static void Main()
    {
        // 1. Create a new workbook (create rule)
        Workbook wb = new Workbook();

        // Access the first worksheet and set its name
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Data";

        // Populate some sample data
        ws.Cells["A1"].PutValue("ID");
        ws.Cells["B1"].PutValue("Name");
        ws.Cells["C1"].PutValue("Score");
        ws.Cells["A2"].PutValue(1);
        ws.Cells["B2"].PutValue("Alice");
        ws.Cells["C2"].PutValue(85);
        ws.Cells["A3"].PutValue(2);
        ws.Cells["B3"].PutValue("Bob");
        ws.Cells["C3"].PutValue(92);

        // Apply a simple header style
        Style headerStyle = wb.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;
        ws.Cells.CreateRange("A1:C1").ApplyStyle(headerStyle, new StyleFlag { FontBold = true, CellShading = true });

        // 2. Save the workbook to a MemoryStream using OoxmlSaveOptions (save rule)
        OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions
        {
            CompressionType = OoxmlCompressionType.Level6,
            ExportCellName = false, // improves performance for large files
            EnableZip64 = true
        };

        MemoryStream ms = new MemoryStream();
        wb.Save(ms, ooxmlOptions);
        ms.Position = 0; // Reset stream position for reading

        // 3. Load a workbook from the MemoryStream (load rule)
        Workbook loadedWb = new Workbook(ms);
        Console.WriteLine("Loaded worksheet name: " + loadedWb.Worksheets[0].Name);
        Console.WriteLine("Cell B2 value: " + loadedWb.Worksheets[0].Cells["B2"].StringValue);

        // 4. Convert the original workbook to PDF using ConversionUtility (conversion rule)
        string tempXlsxPath = Path.Combine(Path.GetTempPath(), "temp.xlsx");
        string pdfPath = Path.Combine(Path.GetTempPath(), "output.pdf");
        wb.Save(tempXlsxPath); // Save to a temporary file for conversion
        ConversionUtility.Convert(tempXlsxPath, pdfPath);
        Console.WriteLine("Converted to PDF: " + pdfPath);

        // 5. Detect file format of the generated PDF using FileFormatUtil (utility rule)
        var formatInfo = FileFormatUtil.DetectFileFormat(pdfPath);
        Console.WriteLine("Detected format of PDF file: " + formatInfo.FileFormatType);

        // 6. Convert a CSV file to XLSX using LoadOptions and OoxmlSaveOptions
        string csvPath = Path.Combine(Path.GetTempPath(), "sample.csv");
        File.WriteAllText(csvPath, "Id,Value\n1,100\n2,200");
        LoadOptions csvLoadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook csvWb = new Workbook(csvPath, csvLoadOptions);
        string xlsxFromCsv = Path.Combine(Path.GetTempPath(), "fromCsv.xlsx");
        OoxmlSaveOptions csvSaveOptions = new OoxmlSaveOptions
        {
            CompressionType = OoxmlCompressionType.Level1
        };
        csvWb.Save(xlsxFromCsv, csvSaveOptions);
        Console.WriteLine("CSV converted to XLSX: " + xlsxFromCsv);

        // 7. Save the workbook as an older XLS format using XlsSaveOptions
        string xlsPath = Path.Combine(Path.GetTempPath(), "legacy.xls");
        XlsSaveOptions xlsOptions = new XlsSaveOptions
        {
            MatchColor = true
        };
        wb.Save(xlsPath, xlsOptions);
        Console.WriteLine("Saved as legacy XLS: " + xlsPath);

        // Optional cleanup of temporary files
        // File.Delete(tempXlsxPath);
        // File.Delete(csvPath);
        // File.Delete(xlsxFromCsv);
        // File.Delete(xlsPath);
    }
}