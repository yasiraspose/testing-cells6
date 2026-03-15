using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1 (or retrieve existing one)
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];

        // Set the comment text
        comment.Note = "This comment has its text direction changed.";

        // Change the text direction (orientation) of the comment.
        // Options: ClockWise, CounterClockWise, NoRotation, TopToBottom
        comment.TextOrientationType = TextOrientationType.ClockWise;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}