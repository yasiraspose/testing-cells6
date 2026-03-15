using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook and add a VBA module ----------
            Workbook workbook = new Workbook(); // create new workbook
            // Add a procedural VBA module named "MyModule"
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");
            // Retrieve the added module and set its VBA code
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];
            module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from Aspose.Cells VBA!\"\r\nEnd Sub";
            // Save as a macro‑enabled workbook
            string pathCreated = "CreatedWorkbook.xlsm";
            workbook.Save(pathCreated, SaveFormat.Xlsm);

            // ---------- Load the workbook, edit the existing module ----------
            Workbook loadedWorkbook = new Workbook(pathCreated); // load existing workbook
            // Access the module by name
            VbaModule loadedModule = loadedWorkbook.VbaProject.Modules["MyModule"];
            // Append additional code to the existing module
            loadedModule.Codes += "\r\nSub GoodbyeWorld()\r\n    MsgBox \"Goodbye!\"\r\nEnd Sub";
            // Save the edited workbook
            string pathEdited = "EditedWorkbook.xlsm";
            loadedWorkbook.Save(pathEdited, SaveFormat.Xlsm);

            // ---------- Remove the VBA module by name ----------
            Workbook toRemoveWorkbook = new Workbook(pathEdited); // load edited workbook
            // Remove the module using its name
            toRemoveWorkbook.VbaProject.Modules.Remove("MyModule");
            // Save the final workbook without the module
            string pathFinal = "FinalWorkbook.xlsm";
            toRemoveWorkbook.Save(pathFinal, SaveFormat.Xlsm);

            Console.WriteLine("Workbook creation, editing, and removal of VBA module completed.");
        }
    }
}