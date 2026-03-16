using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Ensure there is sample data for subtotal and chart
        if (cells["A1"].Value == null)
        {
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            string[] categories = { "A", "B", "C", "D", "E", "F" };
            int[] values = { 10, 20, 5, 8, 12, 3 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);   // Column A
                cells[i + 1, 1].PutValue(values[i]);      // Column B
            }
        }

        // Apply subtotal on the data range; the total row will use the custom total name
        CellArea dataArea = CellArea.CreateCellArea(0, 0, cells.MaxDataRow, 1);
        // Subtotal by Category (column 0) using Sum on the Value column (column 1)
        cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

        // Assign custom globalization settings (covers subtotal total name and chart "Other" label)
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Create a pie chart to demonstrate the customized "Other" label
        int chartIdx = sheet.Charts.Add(ChartType.Pie, 10, 0, 25, 10);
        Chart pieChart = sheet.Charts[chartIdx];

        // Use the same data range for the chart
        pieChart.NSeries.Add("B2:B7", true);
        pieChart.NSeries.CategoryData = "A2:A7";
        pieChart.Title.Text = "Sample Pie Chart";

        // Show data labels (category name and percentage) so the "Other" label can appear if needed
        pieChart.NSeries[0].DataLabels.ShowCategoryName = true;
        pieChart.NSeries[0].DataLabels.ShowPercentage = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings class
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        public CustomGlobalizationSettings()
        {
            // Attach custom chart globalization settings
            this.ChartSettings = new CustomChartGlobalizationSettings();
        }

        // Override the total name used by Subtotal operation
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Sum Total";
            return base.GetTotalName(functionType);
        }
    }

    // Custom chart globalization to change the "Other" label text
    class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetOtherName()
        {
            return "Other (Custom)";
        }
    }
}