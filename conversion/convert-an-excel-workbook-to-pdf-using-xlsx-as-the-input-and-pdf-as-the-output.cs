using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertXlsxToPdfDemo
    {
        public static void Main()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.pdf";

            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}