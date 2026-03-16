using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the Excel workbook (must be macro-enabled, e.g., .xlsm)
        string workbookPath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is password‑protected
        bool isProtected = vbaProject.IsProtected;

        // Output the result
        Console.WriteLine($"Is VBA Project Protected: {isProtected}");
    }
}