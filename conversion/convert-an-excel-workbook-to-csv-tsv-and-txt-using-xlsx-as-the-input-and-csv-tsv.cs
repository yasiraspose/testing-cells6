using System;
using Aspose.Cells;
using Aspose.Cells.Utility; // For TxtSaveOptions

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Load the workbook (create & load lifecycle)
        Workbook workbook = new Workbook(sourcePath);

        // Save as CSV
        workbook.Save("output.csv", SaveFormat.Csv);

        // Save as TSV
        workbook.Save("output.tsv", SaveFormat.Tsv);

        // Save as TXT using TxtSaveOptions with a tab separator
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
        txtOptions.SeparatorString = "\t"; // Use tab as delimiter for TXT
        workbook.Save("output.txt", txtOptions);
    }
}