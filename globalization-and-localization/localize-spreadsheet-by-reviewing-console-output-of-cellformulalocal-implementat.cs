using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX workbook to be examined.
            // Replace with your actual file path or pass as a command‑line argument.
            string workbookPath = args.Length > 0 ? args[0] : "Sample.xlsx";

            // Load the workbook from the specified XLSX file.
            Workbook workbook = new Workbook(workbookPath);

            // Set the workbook's locale to German (de-DE) to see localized function names.
            // You can change this to any culture you need to test.
            workbook.Settings.CultureInfo = new System.Globalization.CultureInfo("de-DE");

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"--- Worksheet: {sheet.Name} ---");

                // Get the range of used cells to limit the iteration.
                var maxRow = sheet.Cells.MaxDataRow;
                var maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only cells that contain a formula.
                        if (cell.IsFormula)
                        {
                            // Standard (English) formula.
                            string standardFormula = cell.Formula;

                            // Locale‑formatted formula using the FormulaLocal property.
                            string localizedFormula = cell.FormulaLocal;

                            // The same using GetFormula with the isLocal flag.
                            string getFormulaLocal = cell.GetFormula(false, true);

                            Console.WriteLine($"Cell {cell.Name}:");
                            Console.WriteLine($"  Standard Formula : {standardFormula}");
                            Console.WriteLine($"  FormulaLocal     : {localizedFormula}");
                            Console.WriteLine($"  GetFormula(true) : {getFormulaLocal}");
                        }
                    }
                }
            }

            // Optionally, save the workbook after any modifications (none in this demo).
            // workbook.Save("LocalizedOutput.xlsx");
        }
    }
}