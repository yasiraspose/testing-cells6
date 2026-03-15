using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // load rule

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the comments collection of the current worksheet
            CommentCollection comments = sheet.Comments;

            // Loop through each comment in the collection
            foreach (Comment comment in comments)
            {
                // Row and column where the comment is placed
                int row = comment.Row;
                int col = comment.Column;

                // Retrieve threaded comments for this cell (by row/column)
                ThreadedCommentCollection threadedComments = comments.GetThreadedComments(row, col);

                // Output the CreatedTime of each threaded comment, if any
                for (int i = 0; i < threadedComments.Count; i++)
                {
                    ThreadedComment tc = threadedComments[i];
                    Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {CellsHelper.CellIndexToName(row, col)}, ThreadedComment #{i + 1}, CreatedTime: {tc.CreatedTime}");
                }
            }
        }

        // Save the workbook (no modifications made, but save rule applied)
        workbook.Save("output.xlsx"); // save rule
    }
}