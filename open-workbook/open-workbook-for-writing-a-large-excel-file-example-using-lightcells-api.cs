using System;
using Aspose.Cells;

class LargeExcelLightCellsExample
{
    static void Main()
    {
        // Create an empty workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Instantiate the custom LightCellsDataProvider
        var provider = new LargeDataProvider();

        // Configure save options to use LightCells mode (lifecycle rule)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = provider
        };

        // Save the workbook using the LightCells provider (lifecycle rule)
        workbook.Save("LargeFile.xlsx", saveOptions);
    }

    // Custom implementation of LightCellsDataProvider that streams a large dataset
    class LargeDataProvider : LightCellsDataProvider
    {
        private const int TotalRows = 100_000;   // Example large row count
        private const int TotalCols = 50;        // Example column count per row

        private int currentRow = -1;
        private int currentCol = -1;
        private bool processCurrentSheet = false;

        // Called once per worksheet during save
        public bool StartSheet(int sheetIndex)
        {
            // Process only the first worksheet (index 0)
            processCurrentSheet = sheetIndex == 0;
            return processCurrentSheet;
        }

        // Provides the next row index to be saved
        public int NextRow()
        {
            if (!processCurrentSheet) return -1;

            currentRow++;
            currentCol = -1; // reset column for new row
            return currentRow < TotalRows ? currentRow : -1; // -1 signals no more rows
        }

        // Allows optional row-level configuration
        public void StartRow(Row row)
        {
            // Example: set a uniform row height
            row.Height = 15;
        }

        // Provides the next column index within the current row
        public int NextCell()
        {
            currentCol++;
            return currentCol < TotalCols ? currentCol : -1; // -1 signals end of cells in this row
        }

        // Populates the cell with data
        public void StartCell(Cell cell)
        {
            // Simple example: write a string indicating its position
            cell.PutValue($"R{currentRow + 1}C{currentCol + 1}");
        }

        // Indicates whether string values should be gathered into a global pool for efficiency
        public bool IsGatherString()
        {
            return true;
        }
    }
}