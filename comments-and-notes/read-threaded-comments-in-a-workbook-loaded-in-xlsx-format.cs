using System;
using Aspose.Cells;

namespace ThreadedCommentsReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input XLSX file
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (XLSX format)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Access the comment collection of the current worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through all comments in the collection
                for (int i = 0; i < comments.Count; i++)
                {
                    // Retrieve the comment object
                    Comment comment = comments[i];

                    // Determine the cell address of the comment
                    string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    Console.WriteLine($"  Comment at cell {cellName}:");

                    // Retrieve threaded comments for this cell using row/column indices
                    ThreadedCommentCollection threadedComments = comments.GetThreadedComments(comment.Row, comment.Column);

                    // If there are no threaded comments, continue to the next comment
                    if (threadedComments == null || threadedComments.Count == 0)
                    {
                        Console.WriteLine("    (No threaded comments)");
                        continue;
                    }

                    // Iterate through each threaded comment and display its details
                    for (int j = 0; j < threadedComments.Count; j++)
                    {
                        ThreadedComment tc = threadedComments[j];
                        string authorName = tc.Author != null ? tc.Author.Name : "Unknown";
                        Console.WriteLine($"    Thread {j + 1}:");
                        Console.WriteLine($"      Author : {authorName}");
                        Console.WriteLine($"      Notes  : {tc.Notes}");
                        Console.WriteLine($"      Row    : {tc.Row}, Column: {tc.Column}");
                        Console.WriteLine($"      Created: {tc.CreatedTime}");
                    }
                }
            }

            // Optionally, save the workbook after processing (e.g., to a new file)
            // workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}