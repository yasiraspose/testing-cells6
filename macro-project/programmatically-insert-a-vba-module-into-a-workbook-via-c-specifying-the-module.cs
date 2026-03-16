using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class InsertVbaModuleFromMht
{
    static void Main()
    {
        // Create a new workbook (macro‑enabled format will be used when saving)
        Workbook workbook = new Workbook();

        // Access the VBA project of the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new procedural VBA module with a custom name
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MhtModule");

        // Retrieve the added module
        VbaModule module = vbaProject.Modules[moduleIndex];

        // Example MHT content (normally you would read this from an .mht file)
        string mhtContent = @"<html>
<head>
<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
<title>Sample MHT VBA</title>
</head>
<body>
<!-- VBA code embedded in MHT -->
Sub SampleMacro()
    MsgBox ""Hello from MHT VBA!""
End Sub
</body>
</html>";

        // Assign the MHT content to the module's code.
        // Aspose.Cells treats the string as VBA source; the MHT wrapper is ignored by the VBA engine,
        // but this demonstrates how to set the source from an MHT‑formatted string.
        module.Codes = mhtContent;

        // Save the workbook as a macro‑enabled file
        string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WorkbookWithMhtModule.xlsm");
        workbook.Save(outputPath, SaveFormat.Xlsm);

        Console.WriteLine($"Workbook saved to: {outputPath}");
    }
}