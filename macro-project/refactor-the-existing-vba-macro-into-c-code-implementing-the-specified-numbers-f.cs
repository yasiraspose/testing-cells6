using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsNumberFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Apply built‑in number format using Style.Number (e.g., 0.00)
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue(1234.567);
            Style builtinStyle = workbook.CreateStyle();
            builtinStyle.Number = 2; // 2 corresponds to "0.00"
            sheet.Cells["A1"].SetStyle(builtinStyle);

            // ------------------------------------------------------------
            // 2. Apply a custom number format using Style.Custom and StyleFlag
            // ------------------------------------------------------------
            sheet.Cells["B1"].PutValue(9876.543);
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "_-€ * #,##0.00_ ;_-€ * -#,##0.00_ ;_-€ * \"-\"??_ ;_-@_";
            // Use StyleFlag to apply only the number format part
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            Aspose.Cells.Range rangeB1 = sheet.Cells.CreateRange(0, 1, 1, 1); // Fully qualified Range
            rangeB1.ApplyStyle(customStyle, flag);

            // ------------------------------------------------------------
            // 3. Create a chart and set DataLabels number format
            // ------------------------------------------------------------
            // Add sample data for the chart
            sheet.Cells["D1"].PutValue("Category");
            sheet.Cells["E1"].PutValue("Value");
            sheet.Cells["D2"].PutValue("A");
            sheet.Cells["E2"].PutValue(1500);
            sheet.Cells["D3"].PutValue("B");
            sheet.Cells["E3"].PutValue(2500);
            sheet.Cells["D4"].PutValue("C");
            sheet.Cells["E4"].PutValue(3500);

            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("E2:E4", true);
            chart.NSeries.CategoryData = "D2:D4";

            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            // Built‑in number format (e.g., 0.00) and custom format for data labels
            series.DataLabels.Number = 2; // Built‑in "0.00"
            series.DataLabels.NumberFormat = "\"$\"#,##0.00";

            // ------------------------------------------------------------
            // 4. Create a pivot table and set number format for a data field
            // ------------------------------------------------------------
            // Add source data for pivot
            sheet.Cells["G1"].PutValue("Product");
            sheet.Cells["H1"].PutValue("Sales");
            sheet.Cells["G2"].PutValue("Apple");
            sheet.Cells["H2"].PutValue(1200);
            sheet.Cells["G3"].PutValue("Banana");
            sheet.Cells["H3"].PutValue(800);
            sheet.Cells["G4"].PutValue("Cherry");
            sheet.Cells["H4"].PutValue(1500);

            int pivotIdx = sheet.PivotTables.Add("G1:H4", "J3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            PivotField dataField = pivot.DataFields[dataFieldIdx];
            dataField.Function = ConsolidationFunction.Sum;
            dataField.NumberFormat = "$#,##0.00";

            // ------------------------------------------------------------
            // 5. Import a DataTable with column‑specific number formats
            // ------------------------------------------------------------
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("TimeValue", typeof(DateTime));
            dt.Rows.Add(1, "Item 1", DateTime.Parse("1:30 PM"));
            dt.Rows.Add(2, "Item 2", DateTime.Parse("3:45 PM"));

            ImportTableOptions importOpts = new ImportTableOptions
            {
                IsFieldNameShown = true,
                NumberFormats = new string[] { null, null, "h:mm AM/PM" } // Apply time format to third column
            };
            sheet.Cells.ImportData(dt, 0, 10, importOpts); // Import starting at column K (index 10)

            // ------------------------------------------------------------
            // 6. Adjust regional separators via WorkbookSettings
            // ------------------------------------------------------------
            workbook.Settings.NumberDecimalSeparator = ',';
            workbook.Settings.NumberGroupSeparator = '.';
            // Apply a style that uses group separator
            Style regionalStyle = workbook.CreateStyle();
            regionalStyle.Custom = "#,##0.00";
            sheet.Cells["L1"].PutValue(12345.678);
            sheet.Cells["L1"].SetStyle(regionalStyle);

            // ------------------------------------------------------------
            // 7. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("NumberFormattingDemo.xlsx");
        }
    }
}