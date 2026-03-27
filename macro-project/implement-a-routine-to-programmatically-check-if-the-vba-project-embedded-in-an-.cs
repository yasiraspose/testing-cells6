using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    public class VbaProtectionChecker
    {
        // Checks whether the VBA project in the specified XLTX workbook is protected.
        public static void CheckVbaProjectProtection(string workbookPath)
        {
            // Load the workbook (XLTX is a macro‑enabled template format)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // If the workbook does not contain a VBA project, IsProtected will be false.
            bool isProtected = vbaProject.IsProtected;

            // Output the result
            Console.WriteLine($"Workbook: {workbookPath}");
            Console.WriteLine($"VBA Project Protected: {isProtected}");
        }

        // Example usage
        public static void Main()
        {
            // Replace with the path to your XLTX file
            string path = "template.xltx";

            CheckVbaProjectProtection(path);
        }
    }
}