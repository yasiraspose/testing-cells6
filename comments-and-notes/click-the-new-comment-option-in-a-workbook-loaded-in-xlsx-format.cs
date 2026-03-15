using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // Workbook(string) ctor

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a new comment to cell A1 (row 0, column 0)
        int commentIndex = worksheet.Comments.Add(0, 0); // CommentCollection.Add(int, int)
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "New comment added programmatically.";
        comment.Author = "Automation";

        // Save the workbook with the new comment
        string outputPath = "output.xlsx";
        workbook.Save(outputPath); // Save the workbook
    }
}