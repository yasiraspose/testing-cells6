using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // -------------------------------------------------
            // 1. Create globalization settings for subtotals
            // -------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize the total label for the SUM function (used by Subtotal)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // -------------------------------------------------
            // 2. Create chart globalization settings
            // -------------------------------------------------
            SettableChartGlobalizationSettings chartGlobals = new SettableChartGlobalizationSettings();

            // Customize various chart labels
            chartGlobals.SetSeriesName("Custom Series");
            chartGlobals.SetChartTitleName("Custom Pie Chart Title");
            chartGlobals.SetLegendTotalName("Custom Legend Total");
            chartGlobals.SetOtherName("Other Category");

            // Attach the chart globalization to the main settings
            globalization.ChartSettings = chartGlobals;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // -------------------------------------------------
            // 3. Locate a pie chart and optionally set its title
            // -------------------------------------------------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = 0; i < sheet.Charts.Count; i++)
                {
                    Chart chart = sheet.Charts[i];
                    if (chart.Type == ChartType.Pie)
                    {
                        // The chart will now use the custom globalization labels
                        chart.Title.Text = "Demo Pie Chart";
                    }
                }
            }

            // -------------------------------------------------
            // 4. Save the modified workbook
            // -------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}