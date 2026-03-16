using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaModuleDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Ensure the workbook contains a VBA project by saving as a macro‑enabled XLSB and reloading
        string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsb");
        workbook.Save(tempPath, SaveFormat.Xlsb);
        workbook = new Workbook(tempPath);
        File.Delete(tempPath);

        // Access the VBA project (read‑only property)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural module named "UtilityModule"
        int procModuleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "UtilityModule");
        VbaModule procModule = vbaProject.Modules[procModuleIndex];
        procModule.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Add a module associated with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        int wsModuleIndex = vbaProject.Modules.Add(sheet);
        VbaModule wsModule = vbaProject.Modules[wsModuleIndex];
        wsModule.Codes = "Sub SheetMacro()\r\n    MsgBox \"This macro belongs to sheet \" & ActiveSheet.Name\r\nEnd Sub";

        // Modify the code of the procedural module
        procModule.Codes = "Sub HelloWorld()\r\n    MsgBox \"Updated message!\"\r\nEnd Sub";

        // Remove the worksheet‑specific module by its name (the worksheet's CodeName)
        vbaProject.Modules.Remove(sheet.CodeName);

        // Save the workbook with the remaining VBA code as a macro‑enabled XLSB file
        string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "VbaDemo.xlsb");
        workbook.Save(outputPath, SaveFormat.Xlsb);
    }
}