using System;
using Aspose.Cells;

class ReadThreadedComments
{
    static void Main()
    {
        // Load the workbook (XLSX format) using the provided load rule
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
        {
            Worksheet worksheet = workbook.Worksheets[wsIndex];
            Console.WriteLine($"Worksheet: {worksheet.Name}");

            // Access the comment collection of the current worksheet
            CommentCollection comments = worksheet.Comments;

            // Loop through all comments present in the worksheet
            for (int c = 0; c < comments.Count; c++)
            {
                Comment comment = comments[c];

                // Retrieve threaded comments for the cell that holds this comment
                ThreadedCommentCollection threadedComments = comments.GetThreadedComments(comment.Row, comment.Column);

                // If there are no threaded comments, continue to the next comment
                if (threadedComments == null || threadedComments.Count == 0)
                    continue;

                // Display cell address
                string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                Console.WriteLine($"  Cell: {cellName}");

                // Iterate through each threaded comment and output its details
                foreach (ThreadedComment tc in threadedComments)
                {
                    string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                    Console.WriteLine($"    Author: {authorName}");
                    Console.WriteLine($"    Note  : {tc.Notes}");
                    Console.WriteLine($"    Row   : {tc.Row}, Column: {tc.Column}");
                }
            }
        }

        // No saving required for read‑only operation; if needed, use the provided save rule:
        // workbook.Save("output.xlsx");
    }
}