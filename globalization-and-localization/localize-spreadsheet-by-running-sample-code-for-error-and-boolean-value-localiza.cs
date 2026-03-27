using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings that provide Russian translations for Boolean and error values
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            // Return localized strings for TRUE/FALSE
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            // Map standard Excel error strings to Russian equivalents
            return err switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (must exist)
            string inputPath = "sample.xlsx";

            // Load the workbook using the standard constructor (create/load rule)
            Workbook wb = new Workbook(inputPath);

            // Apply the custom globalization settings to the workbook
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Prepare data: Boolean values and a set of error strings
            Cells cells = wb.Worksheets[0].Cells;
            cells[0, 0].PutValue(true);   // Boolean TRUE
            cells[0, 1].PutValue(false);  // Boolean FALSE

            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };

            for (int i = 0; i < errors.Length; i++)
            {
                // Place each error string in the first row, starting from column 2
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Display the localized string values in the console
            for (int col = 0; col < errors.Length + 2; col++)
            {
                Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
            }

            // Save the localized workbook (save rule)
            string outputPath = "localized_output.xlsx";
            wb.Save(outputPath);
        }
    }
}