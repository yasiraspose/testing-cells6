using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize subtotal and grand total labels
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize the total name used by Subtotal operation (e.g., "Sum Total")
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            switch (functionType)
            {
                case ConsolidationFunction.Sum:
                    return "合計 (Sum)";          // Japanese example
                case ConsolidationFunction.Count:
                    return "件数 (Count)";
                case ConsolidationFunction.Average:
                    return "平均 (Average)";
                default:
                    return base.GetTotalName(functionType);
            }
        }

        // Localize the grand total name for a given function (used in pivot tables)
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            if (functionType == ConsolidationFunction.Sum)
                return "総計 (Grand Total)";
            return base.GetGrandTotalName(functionType);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the range to which Subtotal will be applied (A1:B5)
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

            // Apply Subtotal: group by column A, sum column B
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // Assign the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Demonstrate the localized names
            string localizedTotal = workbook.Settings.GlobalizationSettings.GetTotalName(ConsolidationFunction.Sum);
            string localizedGrandTotal = workbook.Settings.GlobalizationSettings.GetGrandTotalName(ConsolidationFunction.Sum);
            Console.WriteLine($"Localized Total Name: {localizedTotal}");
            Console.WriteLine($"Localized Grand Total Name: {localizedGrandTotal}");

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}