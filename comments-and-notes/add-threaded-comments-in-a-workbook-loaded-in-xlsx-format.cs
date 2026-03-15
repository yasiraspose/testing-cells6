using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a threaded comment author
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "jdoe", "Provider1");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell B2 using row/column indices (row 1, column 1)
        worksheet.Comments.AddThreadedComment(1, 1, "Initial threaded comment.", author);

        // Add another threaded comment to the same cell using the cell name
        worksheet.Comments.AddThreadedComment("B2", "Follow‑up comment.", author);

        // Retrieve and display all threaded comments for cell B2
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(1, 1);
        foreach (ThreadedComment comment in threadedComments)
        {
            Console.WriteLine($"Author: {comment.Author.Name}, Text: {comment.Notes}");
        }

        // Save the workbook with the new threaded comments
        workbook.Save("Output.xlsx");
    }
}