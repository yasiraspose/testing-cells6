using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one comment; add one for demonstration if none exist
        if (sheet.Comments.Count == 0)
        {
            // Add a comment at cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment newComment = sheet.Comments[commentIndex];
            newComment.Author = "Author";
            newComment.Note = "Original comment";
        }

        // Retrieve the first comment
        Comment comment = sheet.Comments[0];

        // Edit the comment text
        comment.Note = "Edited comment text";

        // Save the workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        // Clean up
        workbook.Dispose();
    }
}