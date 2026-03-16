using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize Boolean and error values
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values (e.g., Russian)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error strings (e.g., Russian equivalents)
        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!":  return "#ДЕЛ/0!";
                case "#REF!":    return "#ССЫЛКА!";
                case "#VALUE!":  return "#ЗНАЧ!";
                case "#N/A":     return "#Н/Д";
                case "#NUM!":    return "#ЧИСЛО!";
                case "#NULL!":   return "#ПУСТО!";
                default:         return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (must exist)
            string inputPath = "input.xlsx";

            // Load the workbook using the provided constructor (load rule)
            Workbook wb = new Workbook(inputPath);

            // Apply the custom globalization settings (error & Boolean localization)
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Demonstrate Boolean localization
            cells[0, 0].PutValue(true);   // A1
            cells[0, 1].PutValue(false);  // B1

            // Demonstrate error value localization
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            for (int i = 0; i < errors.Length; i++)
            {
                // Place each error string starting from C1
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Recalculate formulas (if any) to ensure error values are generated
            wb.CalculateFormula();

            // Output the localized string values to the console
            Console.WriteLine("Localized cell values (first row):");
            for (int col = 0; col < errors.Length + 2; col++)
            {
                Console.WriteLine($"Cell[0,{col}] : {cells[0, col].StringValue}");
            }

            // Save the localized workbook using the provided Save method (save rule)
            string outputPath = "localized_output.xlsx";
            wb.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}