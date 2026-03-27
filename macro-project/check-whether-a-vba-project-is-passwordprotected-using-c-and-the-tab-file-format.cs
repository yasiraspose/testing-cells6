using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProjectProtection
{
    static void Main()
    {
        // Path to the macro‑enabled workbook (XLSM) that may contain a VBA project.
        // The file can be any format that supports VBA (e.g., .xlsm, .xlsb).
        string inputPath = "sample.xlsm";

        // Load the workbook. No password is required for opening the file itself.
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project associated with the workbook.
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that a VBA project actually exists.
        if (vbaProject == null)
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        // Check whether the VBA project is protected.
        bool isProtected = vbaProject.IsProtected;
        Console.WriteLine($"VBA project protected: {isProtected}");

        // If the project is protected, you can validate a password (example password shown).
        if (isProtected)
        {
            string passwordToTest = "testPassword";
            bool passwordValid = vbaProject.ValidatePassword(passwordToTest);
            Console.WriteLine($"Password '{passwordToTest}' validation result: {passwordValid}");
        }

        // OPTIONAL: Save the workbook in TAB (tab‑delimited) format.
        // Note: Saving to TAB format discards VBA content, but demonstrates the required format.
        string outputPath = "output.tab";
        workbook.Save(outputPath, SaveFormat.TabDelimited);
    }
}