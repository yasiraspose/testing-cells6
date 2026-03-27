using System;
using Aspose.Cells;
using System.Text.Json;

class Program
{
    static void Main()
    {
        // Load a workbook (macro-enabled file)
        Workbook workbook = new Workbook("input.xlsm");

        // Retrieve the protection status of the VBA project
        bool isProtected = workbook.VbaProject.IsProtected;

        // Create an anonymous object for JSON serialization
        var result = new { IsProtected = isProtected };

        // Serialize the result to JSON
        string json = JsonSerializer.Serialize(result);

        // Output the JSON string
        Console.WriteLine(json);
    }
}