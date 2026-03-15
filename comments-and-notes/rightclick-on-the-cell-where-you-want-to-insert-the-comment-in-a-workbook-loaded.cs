using System;
using Aspose.Cells;

class AddCommentExample
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx"); // load rule

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the target cell has some content (optional)
        worksheet.Cells["B2"].PutValue("Target Cell");

        // Add a comment to cell B2 using the cell name overload
        int commentIndex = worksheet.Comments.Add("B2");
        Comment comment = worksheet.Comments[commentIndex];

        // Set comment properties
        comment.Note = "This comment was added programmatically.";
        comment.Author = "AsposeUser";
        comment.Font.Name = "Calibri";
        comment.Font.Size = 11;
        comment.IsVisible = true;

        // Save the modified workbook
        workbook.Save("output.xlsx"); // save rule
    }
}