using System;
using Aspose.Cells;

namespace AsposeCellsNumberFormatDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some numeric values
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(0.75);
            sheet.Cells["A3"].PutValue(0.1234);
            sheet.Cells["A4"].PutValue(1234567);

            // ---------- Built‑in Currency format (Number = 5) ----------
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 5;                     // $#,##0_);($#,##0)
            StyleFlag currencyFlag = new StyleFlag();
            currencyFlag.NumberFormat = true;
            Aspose.Cells.Range currencyRange = sheet.Cells.CreateRange(0, 0, 1, 1); // A1
            currencyRange.ApplyStyle(currencyStyle, currencyFlag);

            // ---------- Built‑in Percentage format (Number = 9) ----------
            Style percentStyle = workbook.CreateStyle();
            percentStyle.Number = 9;                      // 0%
            StyleFlag percentFlag = new StyleFlag();
            percentFlag.NumberFormat = true;
            Aspose.Cells.Range percentRange = sheet.Cells.CreateRange(1, 0, 1, 1); // A2
            percentRange.ApplyStyle(percentStyle, percentFlag);

            // ---------- Custom number format using the Custom property ----------
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0.00;[Red]-#,##0.00"; // Positive;Negative in red
            StyleFlag customFlag = new StyleFlag();
            customFlag.NumberFormat = true;
            Aspose.Cells.Range customRange = sheet.Cells.CreateRange(2, 0, 1, 1); // A3
            customRange.ApplyStyle(customStyle, customFlag);

            // ---------- Built‑in Accounting format (Number = 37) ----------
            Style accountingStyle = workbook.CreateStyle();
            accountingStyle.Number = 37;                  // #,##0_);(#,##0)
            StyleFlag accountingFlag = new StyleFlag();
            accountingFlag.NumberFormat = true;
            Aspose.Cells.Range accountingRange = sheet.Cells.CreateRange(3, 0, 1, 1); // A4
            accountingRange.ApplyStyle(accountingStyle, accountingFlag);

            // Save the workbook to an XLSX file
            workbook.Save("NumbersFormatDemo.xlsx");

            Console.WriteLine("Workbook with various number formats has been created.");
        }
    }
}