using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path for the resulting JSON file
        string destPath = "output.json";

        // Load the Excel workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Configure options for saving as JSON
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Include empty cells as null in the JSON output
            ExportEmptyCells = true,
            // Treat the first row as header names
            HasHeaderRow = true,
            // Do not create a nested parent‑child hierarchy
            ExportNestedStructure = false
        };

        // Save the workbook to JSON using the configured options
        workbook.Save(destPath, jsonOptions);

        Console.WriteLine("Excel workbook has been successfully converted to JSON.");
    }
}