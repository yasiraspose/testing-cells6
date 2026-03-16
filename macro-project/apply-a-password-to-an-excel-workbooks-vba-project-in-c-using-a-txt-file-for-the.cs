using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the text file that contains the VBA project password
        string passwordFilePath = "vba_password.txt";

        // Read the password from the file (trim to remove any trailing newline characters)
        string vbaPassword = File.ReadAllText(passwordFilePath).Trim();

        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled workbook to ensure a VBA project is created,
        // then reload it so the VbaProject object becomes available.
        string tempFile = "temp.xlsm";
        workbook.Save(tempFile, SaveFormat.Xlsm);
        workbook = new Workbook(tempFile);
        File.Delete(tempFile);

        // Protect the VBA project with the password read from the TXT file.
        // The first argument (false) means the project is not locked for viewing.
        workbook.VbaProject.Protect(false, vbaPassword);

        // Save the final workbook with the protected VBA project.
        workbook.Save("WorkbookWithVbaPassword.xlsm", SaveFormat.Xlsm);
    }
}