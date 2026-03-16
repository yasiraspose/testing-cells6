using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Define the range that includes the data
        Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B3");

        // Set JSON export options (indentation and header row handling)
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            Indent = "    ", // 4 spaces for pretty‑printed JSON
            HasHeaderRow = true
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(range, jsonOptions);

        // Write the JSON string to the console
        Console.WriteLine(json);
    }
}