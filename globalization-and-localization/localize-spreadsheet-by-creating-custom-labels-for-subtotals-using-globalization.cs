using System;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for ConsolidationFunction enum

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create customizable globalization settings
        SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();

        // Define a custom label for the subtotal of the SUM function
        gSettings.SetTotalName(ConsolidationFunction.Sum, "Custom Subtotal");

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = gSettings;

        // Reference the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the range that will be subtotaled (e.g., A1:B5)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, columns 0‑1

        // Apply subtotal:
        // - Group by column 0 (e.g., "Region")
        // - Use SUM as the consolidation function
        // - Subtotal column 1 (e.g., "Sales")
        // - Replace existing subtotals, no page breaks, place summary below data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}