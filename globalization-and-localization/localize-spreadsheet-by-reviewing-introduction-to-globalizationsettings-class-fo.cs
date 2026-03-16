using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (lifecycle rule: load)
            Workbook workbook = new Workbook("input.xlsx");

            // Create an instance of SettableGlobalizationSettings (lifecycle rule: create)
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize list separator (e.g., use semicolon instead of comma)
            globalization.SetListSeparator(';');

            // Customize boolean display strings
            globalization.SetBooleanValueString(true, "TRUE_LOCAL");
            globalization.SetBooleanValueString(false, "FALSE_LOCAL");

            // Map standard function names to localized names (e.g., SUM -> SOMME, AVERAGE -> MOYENNE)
            globalization.SetLocalFunctionName("SUM", "SOMME", true);
            globalization.SetLocalFunctionName("AVERAGE", "MOYENNE", true);

            // Map a built‑in name (e.g., Total) to a localized version
            globalization.SetLocalBuiltInName("Total", "TOTAL_LOCAL", true);

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Demonstrate usage of localized boolean values
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(true);   // Will display "TRUE_LOCAL"
            sheet.Cells["A2"].PutValue(false);  // Will display "FALSE_LOCAL"

            // Demonstrate usage of localized function names in formulas
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);

            // Standard function name (still works because of bidirectional mapping)
            sheet.Cells["C1"].Formula = "=SUM(B1:B3)";

            // Localized function name
            sheet.Cells["C2"].Formula = "=SOMME(B1:B3)";

            // Localized function name for AVERAGE
            sheet.Cells["C3"].Formula = "=MOYENNE(B1:B3)";

            // Calculate all formulas
            workbook.CalculateFormula();

            // Output some results to console for verification
            Console.WriteLine($"A1 (boolean true) displayed as: {sheet.Cells["A1"].StringValue}");
            Console.WriteLine($"A2 (boolean false) displayed as: {sheet.Cells["A2"].StringValue}");
            Console.WriteLine($"C1 (SUM) result: {sheet.Cells["C1"].Value}");
            Console.WriteLine($"C2 (SOMME) result: {sheet.Cells["C2"].Value}");
            Console.WriteLine($"C3 (MOYENNE) result: {sheet.Cells["C3"].Value}");

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save("output.xlsx");
        }
    }
}