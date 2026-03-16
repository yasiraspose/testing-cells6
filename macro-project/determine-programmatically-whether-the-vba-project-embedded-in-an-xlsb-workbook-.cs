using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSB workbook (modify as needed)
            string filePath = "sample.xlsb";

            // Load the workbook (XLSB format is automatically detected)
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Output the result
            Console.WriteLine($"VBA Project Protected: {isProtected}");

            // Optional: also indicate if the project is locked for viewing
            Console.WriteLine($"VBA Project Locked for Viewing: {vbaProject.IslockedForViewing}");
        }
    }
}