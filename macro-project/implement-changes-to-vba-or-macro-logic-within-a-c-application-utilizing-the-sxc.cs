using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroSxcDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled workbook.
            string sourcePath = "MacroTemplate.xlsm";

            Workbook workbook;

            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                // Create a new workbook and save it as a macro‑enabled file to initialize a VBA project.
                workbook = new Workbook();
                workbook.Save(sourcePath, SaveFormat.Xlsm);
                workbook = new Workbook(sourcePath);
            }

            // -----------------------------------------------------------------
            // Example 1: Add a new VBA module with a simple macro
            // -----------------------------------------------------------------
            VbaProject vbaProject = workbook.VbaProject;
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "DemoModule");
            VbaModule module = vbaProject.Modules[moduleIndex];
            module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from Aspose.Cells!\"\r\nEnd Sub";

            // Verify that the workbook now contains a macro
            Console.WriteLine("HasMacro after adding module: " + workbook.HasMacro);

            // -----------------------------------------------------------------
            // Example 2: Remove all VBA/macros from the workbook
            // -----------------------------------------------------------------
            workbook.RemoveMacro();

            // Verify removal
            Console.WriteLine("HasMacro after removal: " + workbook.HasMacro);

            // -----------------------------------------------------------------
            // Save the resulting workbook in StarOffice Calc (SXC) format
            // -----------------------------------------------------------------
            string outputPath = "ResultWithoutMacro.sxc";
            workbook.Save(outputPath, SaveFormat.Sxc);

            Console.WriteLine("Workbook saved as SXC without macros at: " + outputPath);
        }
    }
}