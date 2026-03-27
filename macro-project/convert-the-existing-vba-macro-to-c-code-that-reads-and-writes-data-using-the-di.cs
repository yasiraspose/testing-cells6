using System;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Saving;

namespace AsposeCellsDifExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and populate it with sample data.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];               // access first worksheet

            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(60);

            // -----------------------------------------------------------------
            // 2. Save the workbook to DIF format using DifSaveOptions.
            // -----------------------------------------------------------------
            DifSaveOptions saveOptions = new DifSaveOptions
            {
                ClearData = true,          // clear data after saving
                CreateDirectory = true,    // create target folder if missing
                RefreshChartCache = true   // refresh chart cache (if any)
            };

            string difPath = "SampleData.dif";
            workbook.Save(difPath, saveOptions);   // save as DIF

            // -----------------------------------------------------------------
            // 3. Load the previously saved DIF file using DifLoadOptions.
            // -----------------------------------------------------------------
            DifLoadOptions loadOptions = new DifLoadOptions(); // create load options
            Workbook loadedWorkbook = new Workbook(difPath, loadOptions); // load DIF

            // -----------------------------------------------------------------
            // 4. Read data from the loaded workbook and display it.
            // -----------------------------------------------------------------
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            int maxRow = loadedSheet.Cells.MaxDataRow;
            int maxCol = loadedSheet.Cells.MaxDataColumn;

            Console.WriteLine("Data read from DIF file:");
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Console.Write(loadedSheet.Cells[row, col].StringValue + "\t");
                }
                Console.WriteLine();
            }

            // -----------------------------------------------------------------
            // 5. Optionally, modify data and save back to a new DIF file.
            // -----------------------------------------------------------------
            loadedSheet.Cells["B2"].PutValue(130); // update quantity for Apple

            string difPathModified = "SampleData_Modified.dif";
            loadedWorkbook.Save(difPathModified, saveOptions); // reuse same save options

            Console.WriteLine($"Modified DIF file saved as '{difPathModified}'.");
        }
    }
}