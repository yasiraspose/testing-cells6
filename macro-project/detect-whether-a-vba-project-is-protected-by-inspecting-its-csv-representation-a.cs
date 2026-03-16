using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the macro-enabled Excel file
        string filePath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if the VBA project is protected
        bool isProtected = vbaProject.IsProtected;
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the protection status
        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
    }
}