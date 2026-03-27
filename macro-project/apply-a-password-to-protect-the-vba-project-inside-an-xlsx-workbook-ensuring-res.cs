using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (will be saved as macro-enabled to hold VBA project)
        Workbook workbook = new Workbook();

        // Protect the VBA project and lock it for viewing with a password
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook as an XLSM file (macro-enabled) so the VBA project is retained
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}