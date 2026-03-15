using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Try to get the comment at cell A1 (row 0, column 0)
        Comment comment = sheet.Comments[0, 0];

        // If the comment does not exist, create one
        if (comment == null)
        {
            int commentIndex = sheet.Comments.Add(0, 0); // Add comment to A1
            comment = sheet.Comments[commentIndex];
            comment.Note = "This is a new comment.";
        }

        // Change the font color of the entire comment to red
        comment.Font.Color = Color.Red;

        // Make the comment visible (optional)
        comment.IsVisible = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}