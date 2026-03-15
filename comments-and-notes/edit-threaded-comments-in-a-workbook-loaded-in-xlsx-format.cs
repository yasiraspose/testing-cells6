using System;
using Aspose.Cells;

namespace ThreadedCommentEditor
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (lifecycle rule: use constructor with file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the comment collection of the worksheet
            CommentCollection comments = sheet.Comments;

            // Ensure there is at least one threaded comment author; create one if needed
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("Editor", "editorUser", "provider");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Example 1: Add a new threaded comment to cell B2 (row 1, column 1)
            comments.AddThreadedComment(1, 1, "Initial threaded comment.", author);

            // Example 2: Retrieve all threaded comments for cell B2 and update their text
            var threadedComments = comments.GetThreadedComments(1, 1);
            foreach (ThreadedComment tc in threadedComments)
            {
                // Append a suffix to the existing comment text
                tc.Notes = tc.Notes + " (edited)";
            }

            // Example 3: Add a reply to the existing threaded comment in B2
            comments.AddThreadedComment(1, 1, "Reply to the comment.", author);

            // Save the modified workbook (lifecycle rule: use Save method)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Threaded comments edited and workbook saved to '{outputPath}'.");
        }
    }
}