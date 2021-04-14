using System;
using System.Reflection;
using ExcelCSVReaderApp.Util;

namespace ExcelCSVReaderApp
{
    internal static class Program
    {
        private static void Main()
        {
            // Set Excel and CSV file paths
            var path = Assembly.GetCallingAssembly().CodeBase;
            if (path == null) return;
            var projectPath = new Uri(path.Substring(0, path.LastIndexOf("bin", StringComparison.Ordinal))).LocalPath;
            
            var excelFilePath = new Uri(projectPath).LocalPath + @"TestData.xlsx";
            var csvFilePath = new Uri(projectPath).LocalPath + @"TestData.csv";

            Console.WriteLine("Excel file path is: " + excelFilePath);
            Console.WriteLine("CSV file path is: " + csvFilePath);

            // Read Excel file
            Console.WriteLine("Excel cell value is: " + new ExcelReader(excelFilePath).Read("TestDataSheet", 1, 2));
            
            // Read CSV file
            Console.WriteLine("CSV cell value is: " + new CsvReader(csvFilePath).Read(1, 2));
        }
    }
}