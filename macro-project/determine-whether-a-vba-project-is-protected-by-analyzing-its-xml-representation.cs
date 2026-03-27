using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    // Entry point of the application
    static void Main(string[] args)
    {
        // Path to the Excel file (macro-enabled) to be examined.
        // If a path is provided as a command‑line argument it will be used,
        // otherwise a default file name is assumed.
        string filePath = args.Length > 0 ? args[0] : "sample.xlsm";

        // Load the workbook. Aspose.Cells automatically parses the internal
        // XML parts, including the VBA project, when the file is opened.
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook.
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected.
        // The IsProtected property reflects the protection state extracted
        // from the VBA project's XML representation.
        bool isProtected = vbaProject.IsProtected;

        // Output the result.
        Console.WriteLine($"VBA Project Protected: {isProtected}");
    }
}