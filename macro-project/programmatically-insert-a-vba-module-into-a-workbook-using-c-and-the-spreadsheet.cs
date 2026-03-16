using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural VBA module named "MyModule" (lifecycle: add module)
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");

        // Retrieve the newly added module and assign VBA code to it
        VbaModule module = vbaProject.Modules[moduleIndex];
        module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Save the workbook as a macro‑enabled file (lifecycle: save)
        // Using the OOXML macro‑enabled format (.xlsm) which preserves the VBA project
        workbook.Save("MyWorkbook.xlsm", SaveFormat.Xlsm);
    }
}