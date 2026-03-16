using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ApplyVbaProjectPassword
{
    static void Main()
    {
        // Create a new workbook (macro‑enabled format will be used on save)
        Workbook workbook = new Workbook();

        // Ensure a VBA project exists by saving as a macro‑enabled file and reloading
        string tempPath = "temp.xlsm";
        workbook.Save(tempPath, SaveFormat.Xlsm);
        workbook = new Workbook(tempPath);
        System.IO.File.Delete(tempPath);

        // Protect the VBA project and lock it for viewing with a password
        // islockedForViewing = true means the project cannot be opened without the password
        workbook.VbaProject.Protect(true, "MyVbaPassword");

        // Save the workbook with the protected VBA project
        workbook.Save("VbaProjectProtected.xlsm", SaveFormat.Xlsm);
    }
}