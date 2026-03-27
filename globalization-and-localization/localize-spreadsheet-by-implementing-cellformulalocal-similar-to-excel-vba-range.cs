using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (XLSX format)
            // Replace "input.xlsx" with the actual path to your file
            Workbook workbook = new Workbook("input.xlsx");

            // Set the workbook's locale (region) to German for demonstration
            // This influences how FormulaLocal is interpreted and displayed
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell (A1)
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // -----------------------------------------------------------------
            // 1. Set a formula using the standard (English) syntax
            // -----------------------------------------------------------------
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("After setting standard formula:");
            Console.WriteLine("Standard Formula   : " + cell.Formula);
            Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 2. Set a formula using the localized (German) syntax via FormulaLocal
            // -----------------------------------------------------------------
            // In German Excel, the SUM function is "SUMME"
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the conversion
            Console.WriteLine("After setting localized formula (FormulaLocal):");
            Console.WriteLine("Standard Formula   : " + cell.Formula);
            Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 3. Demonstrate GetFormula with localization flags
            // -----------------------------------------------------------------
            // GetFormula(isR1C1, isLocal)
            string englishFormula = cell.GetFormula(false, false); // English, A1 style
            string germanFormula = cell.GetFormula(false, true);   // Localized, A1 style

            Console.WriteLine("Using GetFormula:");
            Console.WriteLine("English (isLocal=false) : " + englishFormula);
            Console.WriteLine("German  (isLocal=true)  : " + germanFormula);
            Console.WriteLine();

            // Optionally calculate the workbook to ensure formula results are up‑to‑date
            workbook.CalculateFormula();

            // Save the modified workbook
            // Replace "output.xlsx" with the desired output path
            workbook.Save("output.xlsx");
        }
    }
}