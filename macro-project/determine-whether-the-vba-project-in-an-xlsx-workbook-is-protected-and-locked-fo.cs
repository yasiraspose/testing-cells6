using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the macro‑enabled workbook (XLSM) that may contain a VBA project
        string filePath = "sample.xlsm";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(filePath);

        // Retrieve the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that a VBA project exists (null for non‑macro workbooks)
        if (vbaProject != null)
        {
            // Check whether the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Check whether the VBA project is locked for viewing
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            Console.WriteLine($"VBA Project Protected: {isProtected}");
            Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}