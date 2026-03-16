using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "sample.xlsx";

        // Create a sample workbook and save it if the file does not exist
        if (!System.IO.File.Exists(filePath))
        {
            Workbook tempWb = new Workbook();
            tempWb.Worksheets[0].Cells["A1"].PutValue("Sample");
            tempWb.Save(filePath, SaveFormat.Xlsx);
        }

        // Load the workbook from the file
        Workbook workbook = new Workbook(filePath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if a VBA project exists and whether it is protected
        bool hasVbaProject = vbaProject != null;
        bool isProtected = hasVbaProject && vbaProject.IsProtected;

        // Output the detection results
        Console.WriteLine("VBA project present: " + hasVbaProject);
        Console.WriteLine("VBA project protected: " + isProtected);
    }
}