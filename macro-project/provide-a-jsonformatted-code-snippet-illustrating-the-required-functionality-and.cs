using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonUtilityDemo
{
    public static void Run()
    {
        // -------------------------------------------------
        // 1. Create a workbook and populate sample data
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // -------------------------------------------------
        // 2. Export a range to JSON using JsonSaveOptions
        // -------------------------------------------------
        Aspose.Cells.Range exportRange = worksheet.Cells.CreateRange("A1:B3");

        JsonSaveOptions exportOptions = new JsonSaveOptions
        {
            // 4 spaces indentation for pretty‑printed JSON
            Indent = "    ",
            HasHeaderRow = true,
            ExportEmptyCells = false,
            ExportNestedStructure = true
        };

        string json = JsonUtility.ExportRangeToJson(exportRange, exportOptions);
        Console.WriteLine("Exported JSON:");
        Console.WriteLine(json);

        // -------------------------------------------------
        // 3. Import the JSON back into a new workbook
        // -------------------------------------------------
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            // Treat JSON arrays as tables
            ArrayAsTable = true
        };

        Workbook importWorkbook = new Workbook();
        JsonUtility.ImportData(json, importWorkbook.Worksheets[0].Cells, 0, 0, importOptions);

        // Save the imported data as an Excel file
        importWorkbook.Save("ImportedData.xlsx");

        // -------------------------------------------------
        // 4. Save the imported workbook as JSON with a schema
        // -------------------------------------------------
        string schema = @"{
            ""$schema"": ""http://json-schema.org/draft-07/schema#"",
            ""type"": ""object"",
            ""properties"": {
                ""Name"": { ""type"": ""string"" },
                ""Age"": { ""type"": ""integer"" }
            },
            ""required"": [""Name"", ""Age""]
        }";

        JsonSaveOptions schemaOptions = new JsonSaveOptions
        {
            // Attach the schema for validation
            Schemas = new[] { schema },
            ExportNestedStructure = true,
            // Use a tab character for indentation
            Indent = "\t"
        };

        // Save the workbook as JSON, embedding the schema
        importWorkbook.Save("ImportedDataWithSchema.json", schemaOptions);
    }

    static void Main(string[] args)
    {
        Run();
    }
}