using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the Excel file that may contain a VBA project
        string inputPath = "sample.xlsm";

        // Path where the HTML report will be saved
        string reportPath = "VbaProtectionReport.html";

        // Load the workbook (create rule)
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Retrieve protection status
        bool isProtected = vbaProject.IsProtected;
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Create a new workbook to hold the report (create rule)
        Workbook reportWorkbook = new Workbook();
        Worksheet sheet = reportWorkbook.Worksheets[0];
        sheet.Name = "Report";

        // Build simple HTML content
        string htmlContent = "<html><head><title>VBA Protection Report</title></head><body>";
        htmlContent += $"<h1>VBA Protection Report for {System.IO.Path.GetFileName(inputPath)}</h1>";
        htmlContent += $"<p><strong>IsProtected:</strong> {isProtected}</p>";
        htmlContent += $"<p><strong>IsLockedForViewing:</strong> {isLockedForViewing}</p>";
        htmlContent += "</body></html>";

        // Place the HTML into a cell
        sheet.Cells["A1"].PutValue(htmlContent);

        // Save the report as an HTML file (save rule)
        reportWorkbook.Save(reportPath, SaveFormat.Html);
    }
}