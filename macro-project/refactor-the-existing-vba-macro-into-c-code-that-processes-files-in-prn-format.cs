using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsPrnProcessor
{
    public class PrnToXlsxConverter
    {
        /// <summary>
        /// Converts a PRN (printer) file to an XLSX workbook, removes any macros if present, and saves the result.
        /// </summary>
        /// <param name="prnFilePath">Full path to the source PRN file.</param>
        /// <param name="outputXlsxPath">Full path where the resulting XLSX file will be saved.</param>
        public static void Convert(string prnFilePath, string outputXlsxPath)
        {
            // Validate input file existence
            if (!File.Exists(prnFilePath))
                throw new FileNotFoundException($"Source file not found: {prnFilePath}");

            // Load the PRN file. Aspose.Cells does not have a dedicated PRN format,
            // but PRN files are typically delimited text. We treat it as CSV.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(prnFilePath, loadOptions);

            // If the workbook somehow contains VBA macros, remove them.
            if (workbook.HasMacro)
            {
                workbook.RemoveMacro();
            }

            // Save the workbook as a macro‑free XLSX file.
            workbook.Save(outputXlsxPath, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Main()
        {
            try
            {
                string sourcePrn = @"C:\Data\sample.prn";
                string destinationXlsx = @"C:\Data\sample_converted.xlsx";

                Convert(sourcePrn, destinationXlsx);

                Console.WriteLine($"Conversion completed successfully. Output file: {destinationXlsx}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}