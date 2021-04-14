using GemBox.Spreadsheet;

namespace ExcelCSVReaderApp.Util
{
    public class ExcelReader
    {
        private readonly string _excelFilePath;

        public ExcelReader(string excelFilePath)
        {
            _excelFilePath = excelFilePath;
        }

        public string Read(string sheetName, int rowIndex, int columnIndex)
        {
            // If using Professional version, put your serial key below
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Load Excel workbook from file path
            var workbook = ExcelFile.Load(_excelFilePath);

            // Select the worksheet by name
            var worksheet = workbook.Worksheets[sheetName];

            // Display sheet's name
            // Console.WriteLine("SheetName is: " + worksheet.Name);

            // Select the row by index
            var row = worksheet.Rows[rowIndex];

            // Select the cell by row and column number
            var cell = row.Cells[columnIndex];

            // Display cell value
            // Console.WriteLine("Cell value is: " + cell.Value);

            return cell.Value as string;
        }
    }
}