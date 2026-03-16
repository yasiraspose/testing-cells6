using System;
using Aspose.Cells;
using Aspose.Cells.Loading;

namespace AsposeCellsDifDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and populate it with sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook (rule)
            Worksheet sheet = workbook.Worksheets[0];               // access first worksheet

            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(75);

            // -----------------------------------------------------------------
            // 2. Save the workbook as a DIF file using DifSaveOptions
            // -----------------------------------------------------------------
            DifSaveOptions saveOptions = new DifSaveOptions
            {
                ClearData = true,          // optional: clear data after saving
                CreateDirectory = true,    // create target folder if missing
                RefreshChartCache = true   // optional: refresh chart cache
            };
            string difPath = "SampleData.dif";
            workbook.Save(difPath, saveOptions);                 // save (rule)

            Console.WriteLine($"Workbook saved to DIF format at '{difPath}'.");

            // -----------------------------------------------------------------
            // 3. Load the DIF file back using DifLoadOptions
            // -----------------------------------------------------------------
            DifLoadOptions loadOptions = new DifLoadOptions();   // create load options (rule)
            Workbook loadedWorkbook = new Workbook(difPath, loadOptions); // load (rule)

            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 4. Read data from the loaded workbook and display it
            // -----------------------------------------------------------------
            Console.WriteLine("Data read from the loaded DIF file:");
            for (int row = 0; row <= 4; row++)
            {
                string product = loadedSheet.Cells[row, 0].StringValue;
                string quantity = loadedSheet.Cells[row, 1].StringValue;
                Console.WriteLine($"{product}\t{quantity}");
            }

            // -----------------------------------------------------------------
            // 5. Optionally, save the loaded data to another format (e.g., XLSX) for verification
            // -----------------------------------------------------------------
            string xlsxPath = "LoadedFromDif.xlsx";
            loadedWorkbook.Save(xlsxPath, SaveFormat.Xlsx);
            Console.WriteLine($"Loaded data also saved to '{xlsxPath}'.");
        }
    }
}