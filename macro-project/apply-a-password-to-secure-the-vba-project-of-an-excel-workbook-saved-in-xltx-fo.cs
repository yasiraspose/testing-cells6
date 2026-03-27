using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a VBA module so that a VBA project exists
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];
        module.Codes = "Sub Demo()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Protect the VBA project and lock it for viewing with a password
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook as an Excel template (XLTX)
        workbook.Save("ProtectedVbaProjectTemplate.xltx", SaveFormat.Xltx);
    }
}