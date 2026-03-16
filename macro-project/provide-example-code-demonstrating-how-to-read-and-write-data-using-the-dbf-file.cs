using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Loading;

namespace AsposeCellsDbfDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and write data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Save as DBF using DbfSaveOptions
            DbfSaveOptions saveOptions = new DbfSaveOptions
            {
                ExportAsString = true // export all values as strings (optional)
            };
            string dbfPath = "SampleData.dbf";
            workbook.Save(dbfPath, saveOptions);
            Console.WriteLine($"Workbook saved to DBF file: {dbfPath}");

            // ---------- Load the DBF file ----------
            DbfLoadOptions loadOptions = new DbfLoadOptions();
            Workbook dbfWorkbook = new Workbook(dbfPath, loadOptions);
            Worksheet dbfSheet = dbfWorkbook.Worksheets[0];

            // Read and display some data from the loaded DBF
            Console.WriteLine("Data read from DBF file:");
            Console.WriteLine($"Row 2, Column A: {dbfSheet.Cells["A2"].StringValue}");
            Console.WriteLine($"Row 2, Column B: {dbfSheet.Cells["B2"].StringValue}");
            Console.WriteLine($"Row 3, Column A: {dbfSheet.Cells["A3"].StringValue}");
            Console.WriteLine($"Row 3, Column B: {dbfSheet.Cells["B3"].StringValue}");

            // Save the loaded data to an Excel file for verification
            string xlsxPath = "DbfDataConverted.xlsx";
            dbfWorkbook.Save(xlsxPath, SaveFormat.Xlsx);
            Console.WriteLine($"Loaded DBF data saved to Excel file: {xlsxPath}");
        }
    }
}