using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file (can be an existing template or will be created if missing)
        string sourcePath = "Template.xlsx";

        // Load the workbook if the file exists; otherwise create a new workbook and add sample data
        Workbook workbook;
        if (File.Exists(sourcePath))
        {
            workbook = new Workbook(sourcePath);
        }
        else
        {
            workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Demo");
        }

        // Ensure the workbook is macro‑enabled.
        // If it does not contain a VBA project, save it as .xlsm and reload to create the project.
        if (!workbook.HasMacro)
        {
            string tempMacroPath = "temp.xlsm";
            workbook.Save(tempMacroPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempMacroPath);
            File.Delete(tempMacroPath);
        }

        // Add a new procedural VBA module named "MyMacro"
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyMacro");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];

        // Set the VBA code for the newly added module
        module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from Aspose.Cells VBA!\"\r\nEnd Sub";

        // Example: remove an existing module named "OldModule" if it exists
        try
        {
            workbook.VbaProject.Modules.Remove("OldModule");
        }
        catch
        {
            // Ignored – module may not be present
        }

        // Save the final workbook as a macro‑enabled file
        string outputPath = "Result.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}