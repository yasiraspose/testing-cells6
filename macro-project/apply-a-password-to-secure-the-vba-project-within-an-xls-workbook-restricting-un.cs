using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (initially without a VBA project)
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled file to instantiate the VBA project
        workbook.Save("temp.xlsm", SaveFormat.Xlsm);

        // Reload the workbook so the VBA project is available
        workbook = new Workbook("temp.xlsm");

        // OPTIONAL: add a simple VBA module (demonstrates that the project works)
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
        workbook.VbaProject.Modules[moduleIndex].Codes =
            "Sub Hello()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

        // Protect the VBA project and lock it for viewing with a password
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook with the protected VBA project
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}