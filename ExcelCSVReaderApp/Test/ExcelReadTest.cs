using System;
using System.Reflection;
using ExcelCsvReaderApp.Util;
using NUnit.Framework;

namespace ExcelCsvReaderApp.Test
{
    [TestFixture]
    public class ExcelReadTest
    {
        private string _excelFilePath;
        private ExcelReader _excelReader;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            var path = Assembly.GetCallingAssembly().CodeBase;
            if (path == null) return;
            var projectPath = new Uri(path.Substring(0, path.LastIndexOf("bin", StringComparison.Ordinal))).LocalPath;

            _excelFilePath = new Uri(projectPath).LocalPath + @"TestData.xlsx";
            Console.WriteLine("Excel file path is: " + _excelFilePath);

            _excelReader = new ExcelReader(_excelFilePath);
        }

        [Test]
        public void TestExcel()
        {
            var c2CellValue = _excelReader.GetCellValue("TestDataSheet", 2, 3);
            Console.WriteLine("C2 (2nd Row, 3rd Column) Cell Value: " + c2CellValue);
            Assert.AreEqual("JS Community", c2CellValue);

            var b3CellValue = _excelReader.GetCellValue("TestDataSheet", 3, 2);
            Console.WriteLine("B3 (3rd Row, 2nd Column) Cell Value: " + b3CellValue);
            Assert.AreEqual("password", b3CellValue);
        }

        [Test]
        public void TestExcelCellRowNumber()
        {
            var rowNumber1 = _excelReader.GetRowNumberByCellValue("TestDataSheet", "Value", 3);
            Console.WriteLine("'Value' is located in " + rowNumber1 + " row");
            Assert.AreEqual(1, rowNumber1);

            var rowNumber2 = _excelReader.GetRowNumberByCellValue("TestDataSheet", "JS Community", 3);
            Console.WriteLine("'JS Community' is located in " + rowNumber2 + " row");
            Assert.AreEqual(2, rowNumber2);

            var rowNumber3 = _excelReader.GetRowNumberByCellValue("TestDataSheet", "firstName", 2);
            Console.WriteLine("'firstName' is located in " + rowNumber3 + " row");
            Assert.AreEqual(4, rowNumber3);
        }

        [Test]
        public void TestExcelDuplicateValueCount()
        {
            var duplicateCount = _excelReader.GetNumberOfDuplicatesByCellValue("TestDataSheet", "LoginTest", 1);
            Console.WriteLine("'LoginTest' is duplicated " + duplicateCount + " time(s)");
            Assert.AreEqual(3, duplicateCount);
        }
    }
}