using System;
using System.IO;
using Aspose.Cells;

class VbaProjectPasswordProtection
{
    static void Main()
    {
        string csvPath = "passwords.csv";
        string vbaPassword = "defaultPassword";

        if (File.Exists(csvPath))
        {
            using (var reader = new StreamReader(csvPath))
            {
                string line = reader.ReadLine();
                if (!string.IsNullOrEmpty(line))
                {
                    vbaPassword = line.Split(',')[0];
                }
            }
        }

        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Name = "Data";
        sheet.Cells["A1"].PutValue("Sample data");

        string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsm");
        wb.Save(tempPath, SaveFormat.Xlsm);
        wb.Dispose();

        Workbook macroWb = new Workbook(tempPath);
        // Protect the VBA project with the specified password (read‑only flag set to false)
        macroWb.VbaProject.Protect(false, vbaPassword);

        string outputPath = "ProtectedVbaProject.xlsm";
        macroWb.Save(outputPath, SaveFormat.Xlsm);
        macroWb.Dispose();

        if (File.Exists(tempPath))
            File.Delete(tempPath);

        Console.WriteLine($"VBA project protected and saved to '{outputPath}'.");
    }
}