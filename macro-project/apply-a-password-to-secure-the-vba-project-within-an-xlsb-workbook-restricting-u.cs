using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ProtectVbaProject
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Save as a temporary XLSB to initialize the VBA project, then reload
        string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsb");
        wb.Save(tempPath, SaveFormat.Xlsb);
        wb = new Workbook(tempPath);
        File.Delete(tempPath);

        // Protect the VBA project and lock it for viewing with a password
        wb.VbaProject.Protect(true, "MySecretPassword");

        // Save the final workbook with the protected VBA project
        string outputPath = "ProtectedVbaProject.xlsb";
        wb.Save(outputPath, SaveFormat.Xlsb);
    }
}