using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the source XLS workbook (may already contain macros)
        string inputPath = "MacroWorkbook.xls";

        Workbook workbook;

        // Load existing workbook if it exists; otherwise create a new one
        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook();
        }

        // Ensure a VBA project exists; if not, create one by saving as a macro‑enabled file and reloading
        if (workbook.VbaProject == null)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
            workbook.Save(tempPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempPath);
            File.Delete(tempPath);
        }

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new procedural module named "Automation"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "Automation");

        // Retrieve the newly added module
        VbaModule automationModule = vbaProject.Modules[moduleIndex];

        // Add VBA code to the module
        string vbaCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Cells!""
End Sub
";
        automationModule.Codes = vbaCode;

        // Save the workbook as a macro‑enabled file
        string outputPath = "MacroWorkbook_WithAutomation.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}