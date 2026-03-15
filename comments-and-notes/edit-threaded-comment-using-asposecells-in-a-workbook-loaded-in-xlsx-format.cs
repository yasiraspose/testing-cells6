using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the comment collection of the worksheet
        CommentCollection comments = worksheet.Comments;

        // Ensure a threaded comment author exists (add if necessary)
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Editor", "editor@example.com", "provider");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Target cell (row 2, column B) – zero‑based indices
        int row = 1;   // Row index for B2
        int column = 1; // Column index for B2

        // Add a threaded comment if the cell does not already have one
        // (this step can be omitted if the comment already exists)
        comments.AddThreadedComment(row, column, "Original comment", author);

        // Retrieve all threaded comments for the specified cell
        ThreadedCommentCollection threadedComments = comments.GetThreadedComments(row, column);

        // Edit the first threaded comment's text (Notes property)
        if (threadedComments.Count > 0)
        {
            ThreadedComment firstComment = threadedComments[0];
            firstComment.Notes = "Edited comment text";
        }

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}