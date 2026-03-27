using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new class module to the VBA project
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "TestModule");

        // Retrieve the added module
        VbaModule vbaModule = vbaProject.Modules[moduleIndex];

        // Set VBA code for the module
        vbaModule.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Build an XML representation of the VBA module definition
        XDocument xmlDoc = new XDocument(
            new XElement("VbaModule",
                new XElement("Name", vbaModule.Name),
                new XElement("Type", vbaModule.Type.ToString()),
                new XElement("Codes", vbaModule.Codes)
            )
        );

        // Save the XML definition to a file
        string xmlPath = Path.Combine(Environment.CurrentDirectory, "VbaModuleDefinition.xml");
        xmlDoc.Save(xmlPath);

        // Save the workbook as a macro‑enabled file (lifecycle: save)
        string workbookPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithVba.xlsm");
        workbook.Save(workbookPath, SaveFormat.Xlsm);

        Console.WriteLine($"Workbook saved to: {workbookPath}");
        Console.WriteLine($"VBA module definition XML saved to: {xmlPath}");
    }
}