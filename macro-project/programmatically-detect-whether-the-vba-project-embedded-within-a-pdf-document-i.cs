using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the Excel file that may contain an embedded VBA project
        string workbookPath = "sample.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Retrieve the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected (encrypted)
        bool isVbaProtected = vbaProject != null && vbaProject.IsProtected;

        Console.WriteLine("Is the VBA project protected? " + isVbaProtected);
    }
}