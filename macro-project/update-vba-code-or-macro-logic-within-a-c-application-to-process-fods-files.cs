using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ProcessFodsFile
{
    static void Main()
    {
        // Paths for the source FODS file and the target XLSX file
        string sourcePath = "sample.fods";
        string intermediatePath = "converted.xlsx";

        // Convert the FODS file to XLSX using the utility method (handles load and save internally)
        ConversionUtility.Convert(sourcePath, intermediatePath);

        // Load the converted workbook to perform additional processing (e.g., remove macros, replace placeholders)
        Workbook workbook = new Workbook(intermediatePath);

        // If the workbook contains VBA macros, remove them
        if (workbook.HasMacro)
        {
            workbook.RemoveMacro();
        }

        // Example placeholder replacement: replace all occurrences of "{{Name}}" with "John Doe"
        workbook.Replace("{{Name}}", "John Doe");

        // Save the final workbook (overwrites the intermediate file)
        workbook.Save(intermediatePath, SaveFormat.Xlsx);

        Console.WriteLine("FODS file processed and saved as XLSX: " + intermediatePath);
    }
}