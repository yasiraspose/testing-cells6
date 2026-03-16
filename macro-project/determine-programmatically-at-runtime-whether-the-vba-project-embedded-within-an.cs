using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main(string[] args)
    {
        // Path to the macro-enabled workbook (XLSM). Use first argument if provided.
        string filePath = args.Length > 0 ? args[0] : "sample.xlsm";

        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(filePath);

        // Verify that the workbook actually contains a VBA project.
        if (!workbook.HasMacro)
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        // Access the VBA project associated with the workbook.
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected.
        bool isProtected = vbaProject.IsProtected;

        // Output the protection status.
        Console.WriteLine($"VBA Project Protected: {isProtected}");
    }
}