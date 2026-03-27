using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Step 1: Create sample XLS files
        string file1 = "Sample1.xls";
        string file2 = "Sample2.xls";
        CreateSampleWorkbook(file1, "File1", "Alice", 100);
        CreateSampleWorkbook(file2, "File2", "Bob", 200);

        // Step 2: Merge files into a single XLS
        string mergedFile = "Merged.xls";
        string cacheFile = "Cache.tmp";
        CellsHelper.MergeFiles(new string[] { file1, file2 }, cacheFile, mergedFile);
        Console.WriteLine($"Merged file created: {mergedFile}");

        // Step 3: Process merged file with LightCellsDataHandler for performance
        LightCellsDataHandlerDemo handler = new LightCellsDataHandlerDemo();
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = handler;
        Workbook mergedWorkbook = new Workbook(mergedFile, loadOptions);
        Console.WriteLine("Processing completed via LightCellsDataHandler.");

        // Step 4: Save processed workbook with XlsSaveOptions (optimizations)
        XlsSaveOptions saveOptions = new XlsSaveOptions();
        saveOptions.MatchColor = true;
        saveOptions.ValidateMergedAreas = true;
        saveOptions.RefreshChartCache = true;
        saveOptions.ClearData = false;
        mergedWorkbook.Save("Processed.xls", saveOptions);
        Console.WriteLine("Processed workbook saved with XlsSaveOptions.");

        // Step 5: Convert final XLS to PDF using ConversionUtility
        string pdfFile = "Processed.pdf";
        ConversionUtility.Convert("Processed.xls", pdfFile);
        Console.WriteLine($"Converted to PDF: {pdfFile}");

        // Clean up temporary files
        File.Delete(file1);
        File.Delete(file2);
        File.Delete(cacheFile);
    }

    static void CreateSampleWorkbook(string path, string sheetName, string name, double value)
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = sheetName;
        ws.Cells["A1"].PutValue("Name");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue(name);
        ws.Cells["B2"].PutValue(value);
        wb.Save(path);
    }
}

// Implementation of LightCellsDataHandler to demonstrate streaming processing
class LightCellsDataHandlerDemo : LightCellsDataHandler
{
    public bool StartSheet(Worksheet sheet)
    {
        Console.WriteLine($"Start processing sheet: {sheet.Name}");
        return true;
    }

    public bool StartRow(int rowIndex)
    {
        // Process every row
        return true;
    }

    public bool ProcessRow(Row row)
    {
        // Example: output row index
        Console.WriteLine($"Processing row {row.Index}");
        return true;
    }

    public bool StartCell(int columnIndex)
    {
        // Process every cell
        return true;
    }

    public bool ProcessCell(Cell cell)
    {
        // Example: output cell address and value
        Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
        return true;
    }
}