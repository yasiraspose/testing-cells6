using System;
using Aspose.Cells;

namespace FormulaLocalLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Scenario 1: Default locale (en-US) – set standard formula and read FormulaLocal
            // -------------------------------------------------
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.Formula = "=SUM(B1:C1)"; // standard English formula
            Console.WriteLine("Scenario 1");
            Console.WriteLine($"Standard Formula (A1): {cellA1.Formula}");
            Console.WriteLine($"Localized Formula (A1): {cellA1.FormulaLocal}");
            Console.WriteLine();

            // -------------------------------------------------
            // Scenario 2: Change workbook region to German and observe FormulaLocal
            // -------------------------------------------------
            workbook.Settings.Region = CountryCode.Germany; // set locale to German
            Cell cellA2 = sheet.Cells["A2"];
            cellA2.Formula = "=SUM(B2:C2)"; // still using English syntax
            Console.WriteLine("Scenario 2 (German region)");
            Console.WriteLine($"Standard Formula (A2): {cellA2.Formula}");
            Console.WriteLine($"Localized Formula (A2): {cellA2.FormulaLocal}"); // should show German function name "SUMME"
            Console.WriteLine();

            // -------------------------------------------------
            // Scenario 3: Set formula using localized (German) syntax via FormulaLocal
            // -------------------------------------------------
            Cell cellA3 = sheet.Cells["A3"];
            cellA3.FormulaLocal = "=SUMME(B3:C3)"; // German function name
            Console.WriteLine("Scenario 3 (set FormulaLocal with German syntax)");
            Console.WriteLine($"Standard Formula (A3): {cellA3.Formula}");
            Console.WriteLine($"Localized Formula (A3): {cellA3.FormulaLocal}");
            Console.WriteLine();

            // -------------------------------------------------
            // Scenario 4: Retrieve formulas with GetFormula (localized vs standard)
            // -------------------------------------------------
            Console.WriteLine("Scenario 4 (GetFormula)");
            Console.WriteLine($"Standard GetFormula (A3): {cellA3.GetFormula(false, false)}");
            Console.WriteLine($"Localized GetFormula (A3): {cellA3.GetFormula(false, true)}");
            Console.WriteLine();

            // -------------------------------------------------
            // Scenario 5: Locale‑dependent formula parsing (using standard API)
            // -------------------------------------------------
            Cell cellA4 = sheet.Cells["A4"];
            // Using standard English syntax; the locale identifier inside the format string is preserved.
            cellA4.Formula = "TEXT(TODAY(),\"[$-fr-FR]dddd, dd mmmm yyyy\")";
            Console.WriteLine("Scenario 5 (Locale‑dependent formula parsing)");
            Console.WriteLine($"Formula (A4): {cellA4.Formula}");
            Console.WriteLine($"Localized Formula (A4): {cellA4.FormulaLocal}");
            Console.WriteLine();
        }
    }
}