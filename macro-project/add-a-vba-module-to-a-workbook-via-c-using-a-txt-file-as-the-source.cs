using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaModuleFromTxt
{
    static void Main()
    {
        string vbaFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MacroCode.txt");
        string vbaCode = File.Exists(vbaFilePath)
            ? File.ReadAllText(vbaFilePath)
            : "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

        Workbook workbook = new Workbook();

        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "ImportedModule");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];
        module.Codes = vbaCode;

        workbook.Save("WorkbookWithVba.xlsm", SaveFormat.Xlsm);
    }
}