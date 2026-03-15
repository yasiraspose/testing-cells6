using System;
using Aspose.Cells;

class ThreadedCommentExample
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author (name, userId, providerId)
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "Demo Author",          // Author name
            "demo_user",            // User ID
            "demo_provider");       // Provider ID

        // Retrieve the author object
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell B2 (row 1, column 1) using the author
        worksheet.Comments.AddThreadedComment(1, 1, "This is a threaded comment added to B2.", author);

        // Optionally, retrieve and display the added threaded comment
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(1, 1);
        foreach (ThreadedComment tc in threadedComments)
        {
            Console.WriteLine($"Comment at B2 by {tc.Author.Name}: {tc.Notes}");
        }

        // Save the workbook with the new threaded comment
        workbook.Save("OutputWorkbook.xlsx");
    }
}