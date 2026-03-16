using System;
using Aspose.Cells;

namespace AsposeCellsMemoryOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the large Excel file
            string inputPath = "largeDataset.xlsx";

            // Create LoadOptions and set memory mode to MemoryPreference for lower memory consumption
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;

            // Open the workbook with the specified load options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example: read a value from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook (optional, can be saved to a different file)
            workbook.Save("OptimizedCopy.xlsx", SaveFormat.Xlsx);

            // Release resources
            workbook.Dispose();
        }
    }
}