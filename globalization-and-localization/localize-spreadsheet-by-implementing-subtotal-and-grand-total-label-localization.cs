using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (XLSX format)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Create and configure globalization settings for totals
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Set custom grand total names for different consolidation functions (e.g., Chinese)
            globalSettings.SetGrandTotalName(ConsolidationFunction.Sum, "合计");          // "Grand Total" for Sum
            globalSettings.SetGrandTotalName(ConsolidationFunction.Count, "计数合计");   // "Grand Total" for Count

            // Set custom total names (used in Subtotal operation)
            globalSettings.SetTotalName(ConsolidationFunction.Sum, "总计");               // "Total" for Sum
            globalSettings.SetTotalName(ConsolidationFunction.Count, "计数总计");        // "Total" for Count

            // ------------------------------------------------------------
            // Create and configure pivot globalization settings for subtotals
            // ------------------------------------------------------------
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Set custom text for subtotal types (e.g., Chinese)
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "小计");          // "Subtotal" for Sum
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "计数小计");   // "Subtotal" for Count

            // Attach pivot settings to the global settings
            globalSettings.PivotSettings = pivotSettings;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // ------------------------------------------------------------
            // Demonstrate Subtotal operation which will use the custom labels
            // ------------------------------------------------------------
            // Define the range to subtotal (e.g., A1:B5)
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0-4, columns 0-1

            // Apply subtotal: group by column 0 (first column), sum column 1, show subtotals and grand total
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // ------------------------------------------------------------
            // Save the modified workbook
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}