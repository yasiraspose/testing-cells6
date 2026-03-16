using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the ODS spreadsheet
        string odsPath = "sample.ods";

        // Load the ODS file
        Workbook workbook = new Workbook(odsPath);

        // Determine if a VBA project exists and check its protection status
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            bool isProtected = workbook.VbaProject.IsProtected;
            Console.WriteLine($"VBA project is protected: {isProtected}");
        }
        else
        {
            Console.WriteLine("The ODS file does not contain a VBA project.");
        }
    }
}