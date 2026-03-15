using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear all comments (including threaded comments) from the worksheet
                // This uses the Worksheet.ClearComments method (lifecycle: operation)
                sheet.ClearComments();
            }

            // Path to save the workbook after removing threaded comments
            string outputPath = "OutputWorkbook_NoThreadedComments.xlsx";

            // Save the modified workbook (lifecycle: save)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Threaded comments removed and workbook saved to: {outputPath}");
        }
    }
}