using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaModuleDemo
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new workbook and ensure it has a VBA project.
        // -----------------------------------------------------------------
        Workbook wb = new Workbook();

        // Aspose.Cells creates a VBA project only for macro-enabled files.
        // Save as a temporary macro-enabled workbook, reload, then delete the temp file.
        string tempXlsm = Path.Combine(Path.GetTempPath(), "temp.xlsm");
        wb.Save(tempXlsm, SaveFormat.Xlsm);
        wb = new Workbook(tempXlsm);
        File.Delete(tempXlsm);

        // -----------------------------------------------------------------
        // 2. Add VBA modules and set their code.
        // -----------------------------------------------------------------
        // Add a procedural module named "Utility"
        int utilIdx = wb.VbaProject.Modules.Add(VbaModuleType.Procedural, "Utility");
        VbaModule utilModule = wb.VbaProject.Modules[utilIdx];
        utilModule.Codes = "Public Sub ShowMessage()\r\n    MsgBox \"Hello from Utility module!\"\r\nEnd Sub";

        // Add a class module named "HelperClass"
        int classIdx = wb.VbaProject.Modules.Add(VbaModuleType.Class, "HelperClass");
        VbaModule classModule = wb.VbaProject.Modules[classIdx];
        classModule.Codes = "Public Function AddNumbers(a As Long, b As Long) As Long\r\n    AddNumbers = a + b\r\nEnd Function";

        // -----------------------------------------------------------------
        // 3. Save the workbook as a macro‑enabled template (.xltm).
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "VbaTemplateDemo.xltm");
        wb.Save(templatePath, SaveFormat.Xltm);

        // -----------------------------------------------------------------
        // 4. Load the saved template and extract VBA module code.
        // -----------------------------------------------------------------
        Workbook loadedWb = new Workbook(templatePath);
        VbaProject vbaProj = loadedWb.VbaProject;

        // Extract code from the "Utility" module
        VbaModule loadedUtil = vbaProj.Modules["Utility"];
        Console.WriteLine("Original Utility Module Code:");
        Console.WriteLine(loadedUtil.Codes);

        // -----------------------------------------------------------------
        // 5. Edit the code of the "Utility" module.
        // -----------------------------------------------------------------
        loadedUtil.Codes += "\r\nPublic Sub NewProcedure()\r\n    MsgBox \"New procedure added at runtime.\"\r\nEnd Sub";

        // -----------------------------------------------------------------
        // 6. Remove the class module "HelperClass".
        // -----------------------------------------------------------------
        vbaProj.Modules.Remove("HelperClass");

        // -----------------------------------------------------------------
        // 7. Save the modified workbook as a regular macro‑enabled file.
        // -----------------------------------------------------------------
        string modifiedPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "VbaTemplateModified.xlsm");
        loadedWb.Save(modifiedPath, SaveFormat.Xlsm);

        Console.WriteLine("Modified workbook saved to: " + modifiedPath);
    }
}