using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Iterate through all worksheets and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clears all comments from the worksheet
                sheet.ClearComments();
            }

            // Save the workbook after removing threaded comments
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}