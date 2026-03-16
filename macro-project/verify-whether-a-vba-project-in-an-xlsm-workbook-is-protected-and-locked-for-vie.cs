using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro-enabled workbook (XLSM) you want to inspect
            string workbookPath = "sample.xlsm";

            // Load the workbook (create/load lifecycle handled by Aspose.Cells)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is protected
            bool isProtected = vbaProject.IsProtected; // uses VbaProject.IsProtected property

            // Check if the VBA project is locked for viewing
            bool isLockedForViewing = vbaProject.IslockedForViewing; // uses VbaProject.IslockedForViewing property

            // Output the results
            Console.WriteLine($"VBA Project Protected: {isProtected}");
            Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");

            // Optional: demonstrate validation of protection password if needed
            // string passwordToTest = "yourPassword";
            // bool passwordValid = vbaProject.ValidatePassword(passwordToTest);
            // Console.WriteLine($"Password validation result: {passwordValid}");
        }
    }
}