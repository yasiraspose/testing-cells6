using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Convert XLSX to ODS
        string odsPath = "output.ods";
        ConversionUtility.Convert(sourcePath, odsPath);

        // Convert XLSX to SXC
        string sxcPath = "output.sxc";
        ConversionUtility.Convert(sourcePath, sxcPath);

        // Convert XLSX to FODS
        string fodsPath = "output.fods";
        ConversionUtility.Convert(sourcePath, fodsPath);

        Console.WriteLine("Conversion to ODS, SXC, and FODS completed successfully.");
    }
}