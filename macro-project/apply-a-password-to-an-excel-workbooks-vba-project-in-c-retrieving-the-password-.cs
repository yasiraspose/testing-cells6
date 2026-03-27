using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ApplyVbaProjectPassword
{
    static void Main()
    {
        // Path to the TSV file containing the password (first column)
        string tsvPath = "passwords.tsv";

        // Read the first line and extract the password (assumes password is the first field)
        string password = "";
        if (File.Exists(tsvPath))
        {
            string firstLine = File.ReadLines(tsvPath).FirstOrDefault();
            if (!string.IsNullOrEmpty(firstLine))
            {
                // Split by tab and take the first column as password
                string[] parts = firstLine.Split('\t');
                if (parts.Length > 0)
                {
                    password = parts[0];
                }
            }
        }

        // Fallback if password not found
        if (string.IsNullOrEmpty(password))
        {
            Console.WriteLine("Password not found in TSV file.");
            return;
        }

        // Create a new workbook (macro-enabled)
        Workbook workbook = new Workbook();

        // Ensure there is at least one worksheet (VBA project is created automatically)
        workbook.Worksheets.Add();

        // Protect the VBA project with the retrieved password (not locked for viewing)
        workbook.VbaProject.Protect(false, password);

        // Save the workbook as a macro-enabled file
        string outputPath = "ProtectedVbaProject.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);

        Console.WriteLine($"Workbook saved with VBA project protected. Password: {password}");
    }
}