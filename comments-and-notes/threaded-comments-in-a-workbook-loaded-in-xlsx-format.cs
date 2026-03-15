using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the worksheet's comment collection
        CommentCollection comments = worksheet.Comments;

        // Specify the cell for which to retrieve threaded comments
        string cellName = "B2";

        // Get all threaded comments attached to the cell
        ThreadedCommentCollection threadedComments = comments.GetThreadedComments(cellName);

        // Display each threaded comment's author and text
        if (threadedComments != null && threadedComments.Count > 0)
        {
            Console.WriteLine($"Threaded comments for cell {cellName}:");
            foreach (ThreadedComment tc in threadedComments)
            {
                Console.WriteLine($"- Author: {tc.Author.Name}, Notes: {tc.Notes}");
            }
        }
        else
        {
            Console.WriteLine($"No threaded comments found for cell {cellName}.");
        }

        // Add a new threaded comment to the same cell (optional demonstration)
        int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("Demo User", "demo@example.com", "DemoProvider");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];
        comments.AddThreadedComment(cellName, "Added via code", author);

        // Save the workbook with the modifications
        workbook.Save("output.xlsx");
    }
}