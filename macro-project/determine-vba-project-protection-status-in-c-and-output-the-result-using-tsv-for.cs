using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main(string[] args)
    {
        // Header for TSV output
        Console.WriteLine("FilePath\tIsProtected");

        // If no file path is provided, indicate usage and exit
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program <excel_file_path>");
            return;
        }

        string filePath = args[0];

        // Load the workbook (lifecycle: load)
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project; it may be null if the workbook has no VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected
        bool isProtected = vbaProject != null && vbaProject.IsProtected;

        // Output the result in TSV format
        Console.WriteLine($"{filePath}\t{isProtected}");
    }
}