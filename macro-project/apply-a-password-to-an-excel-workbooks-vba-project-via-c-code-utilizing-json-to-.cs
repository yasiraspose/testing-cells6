using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // JSON containing the VBA project password
        string json = @"{ ""password"": ""myVbaPwd123"" }";

        // Parse JSON to extract the password
        using JsonDocument doc = JsonDocument.Parse(json);
        string vbaPassword = doc.RootElement.GetProperty("password").GetString();

        // Create a new workbook (macro‑enabled format will be used when saving)
        Workbook workbook = new Workbook();

        // Protect the VBA project with the password (not locked for viewing)
        workbook.VbaProject.Protect(false, vbaPassword);

        // Save the workbook as a macro‑enabled file
        workbook.Save("VbaProtectedWorkbook.xlsm", SaveFormat.Xlsm);
    }
}