using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Saving;
using Aspose.Cells.Vba;

class DbfMacroIntegration
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Load a macro‑enabled workbook (if you need to work with VBA)
        // ------------------------------------------------------------
        string macroPath = "template.xlsm";
        if (File.Exists(macroPath))
        {
            Workbook macroWorkbook = new Workbook(macroPath);

            if (macroWorkbook.HasMacro)
            {
                foreach (VbaModule module in macroWorkbook.VbaProject.Modules)
                {
                    string originalCode = module.Codes;
                    module.Codes = "' Modified by C# integration\n" + originalCode;
                }

                macroWorkbook.RemoveMacro();
            }
        }
        else
        {
            Console.WriteLine($"Macro workbook not found: {macroPath}. Skipping macro processing.");
        }

        // ------------------------------------------------------------
        // 2. Load an existing DBF file or create a new one if missing
        // ------------------------------------------------------------
        string dbfPath = "data.dbf";
        Workbook dbfWorkbook;
        Worksheet sheet;

        if (File.Exists(dbfPath))
        {
            DbfLoadOptions loadOptions = new DbfLoadOptions();
            dbfWorkbook = new Workbook(dbfPath, loadOptions);
            sheet = dbfWorkbook.Worksheets[0];
        }
        else
        {
            Console.WriteLine($"DBF file not found: {dbfPath}. Creating a new workbook with sample data.");
            dbfWorkbook = new Workbook();
            sheet = dbfWorkbook.Worksheets[0];

            // Sample headers
            sheet.Cells[0, 0].PutValue("Col1");
            sheet.Cells[0, 1].PutValue("Col2");

            // Sample data rows
            for (int i = 1; i <= 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i);
                sheet.Cells[i, 1].PutValue(i * 10);
            }
        }

        // ------------------------------------------------------------
        // 3. Manipulate the DBF data (e.g., add a computed column)
        // ------------------------------------------------------------
        int newColumnIndex = sheet.Cells.MaxColumn + 1;
        sheet.Cells[0, newColumnIndex].PutValue("Total");

        for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
        {
            double firstValue = sheet.Cells[row, 0].DoubleValue;
            double secondValue = sheet.Cells[row, 1].DoubleValue;
            sheet.Cells[row, newColumnIndex].PutValue(firstValue + secondValue);
        }

        // ------------------------------------------------------------
        // 4. Save the workbook back to DBF format with ExportAsString = true
        // ------------------------------------------------------------
        DbfSaveOptions saveOptions = new DbfSaveOptions
        {
            ExportAsString = true
        };

        dbfWorkbook.Save(dbfPath, saveOptions);
        Console.WriteLine($"DBF file saved successfully to {dbfPath}");
    }
}