using System;
using Aspose.Cells;

namespace AsposeCellsOxpsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells XPS Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(67.89);

            // Save the workbook as XPS (XML Paper Specification)
            workbook.Save("DemoOutput.xps", SaveFormat.Xps);

            Console.WriteLine("Workbook successfully saved as XPS file: DemoOutput.xps");
        }
    }
}