using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main(string[] args)
    {
        // Expect the first argument to be the path of the Excel file to inspect
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as an argument.");
            return;
        }

        string filePath = args[0];

        // Load the workbook (macro-enabled or regular workbook)
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Prepare the result as a JSON object
        var result = new { IsProtected = isProtected };
        string json = JsonSerializer.Serialize(result);

        // Output the JSON string to the console
        Console.WriteLine(json);
    }
}