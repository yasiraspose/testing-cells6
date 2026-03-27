using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load a macro-enabled workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Display whether the VBA project is protected
        Console.WriteLine("VBA Project Protected: " + vbaProject.IsProtected);

        // Display whether the VBA project is locked for viewing
        Console.WriteLine("VBA Project Locked for Viewing: " + vbaProject.IslockedForViewing);
    }
}