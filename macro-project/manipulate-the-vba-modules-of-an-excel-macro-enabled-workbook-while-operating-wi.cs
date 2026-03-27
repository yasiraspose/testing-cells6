using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the XLTX template file
        string templatePath = "TemplateFile.xltx";

        // Load the template workbook
        Workbook workbook = new Workbook(templatePath);

        // Ensure a VBA project exists; if not, create one by saving as .xlsm and reloading
        if (!workbook.HasMacro)
        {
            string tempMacroPath = Path.Combine(Path.GetTempPath(), "temp_macro.xlsm");
            workbook.Save(tempMacroPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempMacroPath);
            File.Delete(tempMacroPath);
        }

        // Add a new procedural VBA module named "Automation"
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "Automation");
        VbaModule automationModule = workbook.VbaProject.Modules[moduleIndex];
        automationModule.Codes = "Sub AutoRun()\r\n    MsgBox \"Automation macro executed.\"\r\nEnd Sub";

        // Example: remove an existing module named "OldModule" if it exists
        try
        {
            VbaModule oldModule = workbook.VbaProject.Modules["OldModule"];
            if (oldModule != null)
            {
                workbook.VbaProject.Modules.Remove("OldModule");
            }
        }
        catch
        {
            // Ignore if the module does not exist
        }

        // Save the workbook as a macro‑enabled file
        string outputPath = "ResultWorkbook.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}