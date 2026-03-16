using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.IO;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Save as macro-enabled workbook to create a VBA project, then reload it
        string tempPath = "temp.xlsm";
        workbook.Save(tempPath, SaveFormat.Xlsm);
        workbook = new Workbook(tempPath);
        File.Delete(tempPath);

        // Protect the VBA project and lock it for viewing with a password
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook with the protected VBA project
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}