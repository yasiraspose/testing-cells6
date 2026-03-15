using System;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path for the resulting CSV file
        string destPath = "output.csv";

        // Load the Excel workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure CSV save options to keep separators for blank rows
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            KeepSeparatorsForBlankRow = true, // Preserve empty fields for blank rows
            Encoding = Encoding.UTF8           // Use UTF-8 encoding
        };

        // Save the workbook as CSV with the specified options
        workbook.Save(destPath, csvOptions);
    }
}