using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Path to the source macro‑enabled workbook (XLSM)
        string sourceXlsmPath = "SourceWithMacro.xlsm";

        // Ensure the source file exists; create a minimal macro‑enabled workbook if it does not.
        if (!File.Exists(sourceXlsmPath))
        {
            var tempWb = new Workbook();
            // Save as XLSM to make it macro‑enabled (even though it contains no macros yet)
            tempWb.Save(sourceXlsmPath, SaveFormat.Xlsm);
        }

        // Load the workbook that contains VBA macros
        Workbook macroWorkbook = new Workbook(sourceXlsmPath);

        // Verify that the workbook indeed has a VBA project
        Console.WriteLine("Source workbook HasMacro: " + macroWorkbook.HasMacro);

        // Access the VBA project (read‑only property)
        VbaProject vbaProject = macroWorkbook.VbaProject;

        // Add a new procedural module (optional, demonstrates manipulation)
        if (vbaProject != null)
        {
            int newModuleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "NewModule");
            VbaModule newModule = vbaProject.Modules[newModuleIndex];
            newModule.Codes = "Sub NewMacro()\r\n    MsgBox \"Added by Aspose.Cells\"\r\nEnd Sub";
        }

        // Prepare ODS save options
        OdsSaveOptions odsOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice,
            IsStrictSchema11 = false,
            ClearData = false
        };

        // Save the workbook as ODS. VBA macros are not preserved in ODS format.
        string odsPath = "ExportedFromMacro.ods";
        macroWorkbook.Save(odsPath, odsOptions);
        Console.WriteLine("Workbook saved as ODS to: " + odsPath);

        // Load the ODS file back
        Workbook odsWorkbook = new Workbook(odsPath);
        Console.WriteLine("Loaded ODS workbook HasMacro: " + odsWorkbook.HasMacro);

        // Attempt to copy the original VBA project into a new macro‑enabled workbook
        if (vbaProject != null)
        {
            // Create a new macro‑enabled workbook to hold the copied VBA project
            Workbook tempMacroWorkbook = new Workbook();
            tempMacroWorkbook.VbaProject.Copy(vbaProject);

            // Verify copy
            Console.WriteLine("Temp workbook after copy HasMacro: " + tempMacroWorkbook.HasMacro);
            Console.WriteLine("Number of modules copied: " + tempMacroWorkbook.VbaProject.Modules.Count);

            // Save the temporary workbook as XLSM to retain macros
            string tempXlsmPath = "TempWithCopiedMacro.xlsm";
            tempMacroWorkbook.Save(tempXlsmPath, SaveFormat.Xlsm);
            Console.WriteLine("Temporary workbook with copied macros saved to: " + tempXlsmPath);
        }
    }
}