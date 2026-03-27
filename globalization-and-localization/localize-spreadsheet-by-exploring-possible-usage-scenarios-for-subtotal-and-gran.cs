using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing workbook (XLSX format)
        Workbook workbook = new Workbook("input.xlsx");

        // ------------------------------------------------------------
        // 1. Globalize subtotal and grand total labels for regular subtotals
        // ------------------------------------------------------------
        // Create a settable globalization settings instance
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

        // Customize the total label for the SUM function
        globalization.SetTotalName(ConsolidationFunction.Sum, "My Sum Total");

        // Customize the grand total label for the SUM function
        globalization.SetGrandTotalName(ConsolidationFunction.Sum, "My Sum Grand Total");

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Apply Subtotal on a sample data range (A1:B5)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the area that contains the data (rows 0‑4, columns 0‑1)
        CellArea dataArea = CellArea.CreateCellArea(0, 0, 4, 1);

        // Group by column 0 (e.g., Region) and calculate SUM on column 1 (e.g., Sales)
        // Parameters: area, columnIndexToGroupBy, function, columnsToSubtotal, 
        //             replace, useGrandTotal, useSubTotal
        cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, true, true);

        // ------------------------------------------------------------
        // 2. Globalize labels inside a PivotTable (Grand Total & Subtotal)
        // ------------------------------------------------------------
        // Create a settable pivot globalization settings instance
        SettablePivotGlobalizationSettings pivotGlobalization = new SettablePivotGlobalizationSettings();

        // Change the default "Grand Total" text
        pivotGlobalization.SetTextOfGrandTotal("My Pivot Grand Total");

        // Change the text for the SUM subtotal type
        pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "My Pivot Sum Subtotal");

        // Attach the pivot globalization settings to the workbook's globalization settings
        globalization.PivotSettings = pivotGlobalization;

        // Create a PivotTable to demonstrate the localized labels
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "MyPivotTable");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the PivotTable fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field (Region)
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Data field (Sales)

        // Set the aggregation function for the data field
        pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

        // Refresh and calculate the PivotTable so that labels appear
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // ------------------------------------------------------------
        // Save the modified workbook
        // ------------------------------------------------------------
        workbook.Save("output.xlsx");
    }
}