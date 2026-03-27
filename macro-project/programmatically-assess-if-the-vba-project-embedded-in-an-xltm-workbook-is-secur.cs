using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the XLTM workbook
        string workbookPath = "template.xltm";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Output the protection status
        Console.WriteLine($"VBA Project Protected: {isProtected}");

        // Additional info: whether it is locked for viewing
        Console.WriteLine($"VBA Project Locked for Viewing: {vbaProject.IslockedForViewing}");
    }
}