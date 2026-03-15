using System;
using Aspose.Cells;

class ThreadedCommentExample
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Get the first worksheet (you can change the index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author (name, userId, providerId)
        int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("John Doe", "john.doe@example.com", "PROVIDER_1");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

        // Add a threaded comment to cell B2 (row 1, column 1) using the author
        worksheet.Comments.AddThreadedComment(1, 1, "This is a threaded comment added via Aspose.Cells.", author);

        // Optionally retrieve and display the added threaded comment(s)
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(1, 1);
        foreach (ThreadedComment tc in threadedComments)
        {
            Console.WriteLine($"Threaded comment by {tc.Author.Name}: {tc.Notes}");
        }

        // Save the workbook with the new threaded comment
        workbook.Save("OutputWorkbook.xlsx");
    }
}