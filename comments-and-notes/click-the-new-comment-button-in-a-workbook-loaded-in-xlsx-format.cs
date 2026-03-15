using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (you can choose any worksheet as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a new comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "New comment added via code.";
        comment.Author = "AutomatedProcess";

        // Save the workbook with the new comment
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}