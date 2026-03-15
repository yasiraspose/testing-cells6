using System;
using Aspose.Cells;

class ReadThreadedComments
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the collection of comments for the worksheet
        CommentCollection comments = worksheet.Comments;

        // Iterate through each comment in the collection
        foreach (Comment comment in comments)
        {
            // Check if the comment is a threaded comment
            if (comment.IsThreadedComment)
            {
                // Retrieve the threaded comments for the cell using its row and column
                ThreadedCommentCollection threadedComments = comments.GetThreadedComments(comment.Row, comment.Column);

                // Output information about each threaded comment
                Console.WriteLine($"Threaded comments for cell {CellsHelper.CellIndexToName(comment.Row, comment.Column)}:");
                foreach (ThreadedComment tc in threadedComments)
                {
                    string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                    Console.WriteLine($"- Author: {authorName}, Text: {tc.Notes}, Created: {tc.CreatedTime}");
                }
                Console.WriteLine();
            }
        }

        // (Optional) Save the workbook if any modifications were made
        workbook.Save("output.xlsx");
    }
}