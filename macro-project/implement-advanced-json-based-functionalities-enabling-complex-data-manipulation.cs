using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and import JSON data with layout options
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook (lifecycle rule)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample JSON containing objects, arrays, numbers and dates
            string jsonInput = @"
            {
                ""Employees"": [
                    { ""Name"": ""John Doe"", ""Age"": 30, ""HireDate"": ""2022-01-15"", ""Salary"": 55000.75 },
                    { ""Name"": ""Jane Smith"", ""Age"": 28, ""HireDate"": ""2021-07-01"", ""Salary"": 62000.00 },
                    { ""Name"": ""Bob Johnson"", ""Age"": null, ""HireDate"": null, ""Salary"": null }
                ]
            }";

            // Configure how the JSON is mapped to the worksheet
            JsonLayoutOptions importOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,               // treat arrays as tables
                ConvertNumericOrDate = true,       // convert numeric strings and dates
                IgnoreNull = true,                 // skip null values
                NumberFormat = "$0.00",            // format numeric values as currency
                DateFormat = "yyyy-MM-dd",         // format dates
                TitleStyle = new Style()           // optional: style for header titles
            };
            // Apply a simple bold style to the title row
            importOptions.TitleStyle.Font.IsBold = true;

            // Import JSON data starting at cell A1
            JsonUtility.ImportData(jsonInput, cells, 0, 0, importOptions); // import (rule)

            // ------------------------------------------------------------
            // 2. Export a specific range (the imported table) to JSON with save options
            // ------------------------------------------------------------
            // Determine the used range of the worksheet
            Aspose.Cells.Range usedRange = cells.MaxDisplayRange; // create range (rule)

            // Configure JSON export options
            JsonSaveOptions exportOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,          // include empty cells as null
                HasHeaderRow = true,              // first row contains headers
                ExportAsString = true,            // export all values as strings
                Indent = "  ",                    // pretty‑print with two‑space indent
                ExportNestedStructure = false,    // flat array of objects
                AlwaysExportAsJsonObject = true   // ensure output is a JSON object even for single sheet
            };

            // Export the range to a JSON string
            string exportedJson = JsonUtility.ExportRangeToJson(usedRange, exportOptions); // export (rule)

            // Output the exported JSON to console
            Console.WriteLine("Exported JSON:");
            Console.WriteLine(exportedJson);

            // ------------------------------------------------------------
            // 3. Save the workbook to Excel and also to a JSON file
            // ------------------------------------------------------------
            // Save as Excel file using default format
            workbook.Save("ImportedData.xlsx"); // save (rule)

            // Save the entire workbook as JSON using the same export options
            workbook.Save("WorkbookExport.json", exportOptions); // save with options (rule)

            // ------------------------------------------------------------
            // 4. Demonstrate loading a JSON file with JsonLoadOptions and LayoutOptions
            // ------------------------------------------------------------
            // Assume the JSON file "sample.json" exists in the application directory
            string jsonFilePath = "sample.json";

            // Configure load options, including layout options for import
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                StartCell = "A1",
                MultipleWorksheets = true,
                KeptSchema = true,
                LayoutOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,
                    ConvertNumericOrDate = true,
                    NumberFormat = "0.00",
                    DateFormat = "MM/dd/yyyy",
                    IgnoreNull = false
                }
            };

            // Load the JSON file into a new workbook
            Workbook loadedWorkbook = new Workbook(jsonFilePath, loadOptions); // load (rule)

            // Save the loaded workbook as an Excel file to verify the import
            loadedWorkbook.Save("LoadedFromJson.xlsx"); // save (rule)

            Console.WriteLine("Processing completed.");
        }
    }
}