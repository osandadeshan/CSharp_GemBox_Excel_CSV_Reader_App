using System.Collections.Generic;
using GemBox.Spreadsheet;

namespace ExcelCsvReaderApp.Util
{
    public class CsvReader
    {
        private readonly string _csvFilePath;

        public CsvReader(string csvFilePath)
        {
            _csvFilePath = csvFilePath;
        }

        public string GetCellValue(int rowNumber, int columnNumber)
        {
            // If using Professional version, put your serial key below
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Load CSV workbook from file path
            var workbook = ExcelFile.Load(_csvFilePath);

            // Select the worksheet by index
            var worksheet = workbook.Worksheets[0];

            // Select the row by row number
            var row = worksheet.Rows[rowNumber - 1];

            // Select the cell by row and column number
            var cell = row.Cells[columnNumber - 1];

            return cell.Value.ToString();
        }

        public int GetRowNumberByCellValue(string expectedCellValue, int columnNumber)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = ExcelFile.Load(_csvFilePath);
            var worksheet = workbook.Worksheets[0];
            var numberOfRows = worksheet.Rows.Count;

            var rowNumber = 0;

            for (var i = 1; i <= numberOfRows; i++)
            {
                var cellValue = worksheet.Cells[(i - 1), (columnNumber - 1)].Value;
                if (cellValue != null && cellValue.ToString().Equals(expectedCellValue))
                {
                    rowNumber = i;
                    break;
                }
            }

            if (rowNumber == 0)
            {
                throw new KeyNotFoundException("Failed to find '" + expectedCellValue + "' in CSV file");
            }

            return rowNumber;
        }

        public int GetNumberOfDuplicatesByCellValue(string expectedCellValue, int columnNumber)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = ExcelFile.Load(_csvFilePath);
            var worksheet = workbook.Worksheets[0];
            var numberOfRows = worksheet.Rows.Count;

            var count = 0;

            for (var i = 1; i <= numberOfRows; i++)
            {
                var cellValue = worksheet.Cells[(i - 1), (columnNumber - 1)].Value;
                if (cellValue != null && cellValue.ToString().Equals(expectedCellValue))
                {
                    count++;
                }
            }

            return count;
        }
    }
}