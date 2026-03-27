using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        string fileName = "sample.mht";
        string mhtFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

        if (!File.Exists(mhtFilePath))
        {
            Console.WriteLine($"File not found: {mhtFilePath}");
            return;
        }

        Workbook workbook = new Workbook(mhtFilePath);

        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            bool isVbaProtected = workbook.VbaProject.IsProtected;
            Console.WriteLine($"VBA project is protected: {isVbaProtected}");
        }
        else
        {
            Console.WriteLine("The loaded file does not contain a VBA project.");
        }
    }
}