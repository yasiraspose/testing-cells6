using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the macro‑enabled workbook (XLSM) to be examined
        string filePath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if the VBA project is protected (cannot be edited)
        bool isProtected = vbaProject.IsProtected;

        // Output the result
        Console.WriteLine($"VBA Project Protected: {isProtected}");
    }
}