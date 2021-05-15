using System;
using System.Reflection;
using ExcelCsvReaderApp.Util;
using NUnit.Framework;

namespace ExcelCsvReaderApp.Test
{
    public class CsvReadTest
    {
        private string _csvFilePath;
        private CsvReader _csvReader;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            var path = Assembly.GetCallingAssembly().CodeBase;
            if (path == null) return;
            var projectPath = new Uri(path.Substring(0, path.LastIndexOf("bin", StringComparison.Ordinal))).LocalPath;

            _csvFilePath = new Uri(projectPath).LocalPath + @"TestData.csv";
            Console.WriteLine("CSV file path is: " + _csvFilePath);

            _csvReader = new CsvReader(_csvFilePath);
        }

        [Test]
        public void TestCsv()
        {
            var b2CellValue = _csvReader.GetCellValue(2, 2);
            Console.WriteLine("B2 (2nd Row, 2nd Column) Cell Value: " + b2CellValue);
            Assert.AreEqual("username", b2CellValue);

            var c3CellValue = _csvReader.GetCellValue(3, 3);
            Console.WriteLine("C3 (3rd Row, 3rd Column) Cell Value: " + c3CellValue);
            Assert.AreEqual("Wiley", c3CellValue);
        }

        [Test]
        public void TestCsvCellRowNumber()
        {
            var rowNumber1 = _csvReader.GetRowNumberByCellValue("Value", 3);
            Console.WriteLine("'Value' is located in " + rowNumber1 + " row");
            Assert.AreEqual(1, rowNumber1);

            var rowNumber2 = _csvReader.GetRowNumberByCellValue("JS Community", 3);
            Console.WriteLine("'JS Community' is located in " + rowNumber2 + " row");
            Assert.AreEqual(2, rowNumber2);

            var rowNumber3 = _csvReader.GetRowNumberByCellValue("firstName", 2);
            Console.WriteLine("'firstName' is located in " + rowNumber3 + " row");
            Assert.AreEqual(4, rowNumber3);
        }

        [Test]
        public void TestCsvDuplicateValueCount()
        {
            var duplicateCount = _csvReader.GetNumberOfDuplicatesByCellValue("LoginTest", 1);
            Console.WriteLine("'LoginTest' is duplicated " + duplicateCount + " time(s)");
            Assert.AreEqual(3, duplicateCount);
        }
    }
}