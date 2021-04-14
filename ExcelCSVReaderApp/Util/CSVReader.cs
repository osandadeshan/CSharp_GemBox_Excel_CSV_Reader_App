using GemBox.Spreadsheet;

namespace ExcelCSVReaderApp.Util
{
    public class CsvReader
    {
        
        private readonly string _csvFilePath;

        public CsvReader(string csvFilePath)
        {
            _csvFilePath = csvFilePath;
        }

        public string Read(int rowIndex, int columnIndex)
        {
            // If using Professional version, put your serial key below
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Load Excel workbook from file path
            var workbook = ExcelFile.Load(_csvFilePath);

            // Select the worksheet by name
            var worksheet = workbook.Worksheets["Sheet1"];

            // Select the row by index
            var row = worksheet.Rows[rowIndex];

            // Select the cell by row and column number
            var cell = row.Cells[columnIndex];

            return cell.Value as string;
        }
    }
}