using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell A1 (or retrieve existing one)
        int commentIdx = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIdx];
        comment.Note = "Comment with custom background";

        // Set the comment shape's fill format to change its background color
        comment.CommentShape.FillFormat.IsVisible = true;
        comment.CommentShape.FillFormat.ForeColor = Color.Yellow;   // primary fill color
        comment.CommentShape.FillFormat.BackColor = Color.Orange;   // secondary fill color (optional)

        // Make the comment visible in the worksheet
        comment.IsVisible = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}