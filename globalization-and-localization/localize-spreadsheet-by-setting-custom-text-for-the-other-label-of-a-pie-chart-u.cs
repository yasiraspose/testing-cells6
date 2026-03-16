using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook that contains a pie chart
        Workbook workbook = new Workbook("input.xlsx");

        // Create chart globalization settings and set custom text for the "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Miscellaneous Items");

        // Create a settable globalization settings instance and assign the chart settings to it
        SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings();
        globalizationSettings.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalizationSettings;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}