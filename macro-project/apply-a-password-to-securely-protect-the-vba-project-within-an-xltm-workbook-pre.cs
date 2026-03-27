using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (macro‑enabled)
        Workbook workbook = new Workbook();

        // Protect the VBA project and lock it for viewing with a password
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook as a macro‑enabled template (XLTM)
        workbook.Save("ProtectedTemplate.xltm", SaveFormat.Xltm);
    }
}