using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Workbook with protected VBA project");

        // Save to a memory stream as a macro‑enabled workbook to ensure a VBA project exists (save rule)
        using (MemoryStream tempStream = new MemoryStream())
        {
            workbook.Save(tempStream, SaveFormat.Xlsm);
            tempStream.Position = 0;

            // Load the workbook from the stream (load rule)
            Workbook wbWithVba = new Workbook(tempStream);

            // Protect the VBA project with a password and lock it for viewing
            wbWithVba.VbaProject.Protect(true, "MyVbaPassword");

            // Save the result as PDF (save rule)
            wbWithVba.Save("ProtectedVbaProject.pdf", SaveFormat.Pdf);
        }
    }
}