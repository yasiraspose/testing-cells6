using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class SetVbaProjectPassword
{
    static void Main()
    {
        // Create a new workbook (initially without a VBA project)
        Workbook workbook = new Workbook();

        // Add a worksheet to ensure the workbook can hold a VBA project
        workbook.Worksheets.Add();

        // Save as a macro‑enabled workbook to create the VBA project container
        string tempPath = "temp.xlsm";
        workbook.Save(tempPath, SaveFormat.Xlsm);

        // Reload the workbook so the VBA project is initialized
        workbook = new Workbook(tempPath);

        // Protect the VBA project and lock it for viewing with a password
        // islockedForViewing = true means the project cannot be opened without the password
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the final workbook with the protected VBA project
        string outputPath = "ProtectedVbaProject.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);

        // Optional: verify protection status
        Console.WriteLine("VBA Project IsProtected: " + workbook.VbaProject.IsProtected);
        Console.WriteLine("VBA Project IslockedForViewing: " + workbook.VbaProject.IslockedForViewing);
    }
}