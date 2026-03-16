using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace VbaHtmlIntegrationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro‑enabled)
            Workbook workbook = new Workbook();

            // Add a VBA class module to the workbook
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "HtmlModule");
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];

            // Embed HTML‑formatted code inside the VBA module
            // The VBA macro stores an HTML string and displays it in a message box
            vbaModule.Codes =
                "Sub ShowHtml()\n" +
                "    Dim html As String\n" +
                "    html = \"<html><body><h1>Hello from VBA</h1><p>This is an <b>HTML</b> snippet.</p></body></html>\"\n" +
                "    MsgBox html\n" +
                "End Sub";

            // Save the workbook as a macro‑enabled file
            workbook.Save("VbaHtmlModuleDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}