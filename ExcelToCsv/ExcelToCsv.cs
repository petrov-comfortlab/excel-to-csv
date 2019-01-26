using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelToCsv
{
    public static class ExcelToCsv
    {
        private static string _csvDirectory;
        private static string _excelFile;

        public static bool Create_SheetName_CsvFile(string excelFile, List<string> sheetNames = null, string csvDirectory = null)
        {
            try
            {
                if (string.IsNullOrEmpty(csvDirectory))
                {
                    var directory = Path.GetDirectoryName(excelFile);
                    var folder = Path.GetFileNameWithoutExtension(excelFile);

                    if (!string.IsNullOrEmpty(directory) && !string.IsNullOrEmpty(folder))
                    {
                        csvDirectory = Path.Combine(directory, folder);

                        if (!Directory.Exists(csvDirectory))
                            Directory.CreateDirectory(csvDirectory);
                    }
                    else if (!string.IsNullOrEmpty(directory))
                        csvDirectory = directory;
                    else
                        return false;
                }

                _csvDirectory = csvDirectory;

                var path = Environment.GetFolderPath(Environment.SpecialFolder.Templates);
                var fileName = Path.GetFileName(excelFile);
                if (fileName == null)
                    return false;
                
                var tempExcelFile = Path.Combine(path, fileName);
                File.Copy(excelFile, tempExcelFile, overwrite: true);

                XSSFWorkbook xssfWorkbook;
                using (var fileStream = new FileStream(tempExcelFile, FileMode.Open, FileAccess.Read))
                    xssfWorkbook = new XSSFWorkbook(fileStream);

                if (sheetNames == null)
                {
                    sheetNames = new List<string>();
                    var numberOfSheets = xssfWorkbook.NumberOfSheets;

                    for (var i = 0; i < numberOfSheets; i++)
                    {
                        var sheet = xssfWorkbook.GetSheetAt(i);

                        if (sheet.SheetName.ToLower().Contains("data") ||
                            sheet.SheetName.StartsWith("#"))
                            continue;

                        sheetNames.Add(sheet.SheetName);
                    }
                }

                sheetNames.ForEach(n => CreateCsvFile(excelFile, xssfWorkbook, n));

                File.Delete(tempExcelFile);

                return true;
            }
            catch (Exception exception)
            {
                MessageBox.Show($"{excelFile}\n{exception.Message}\n{exception.StackTrace}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        public static string CreateCsvFile(string excelFile, IWorkbook xssfWorkbook, string sheetName)
        {
            _excelFile = excelFile;

            var sheet = xssfWorkbook.GetSheet(sheetName);

            if (sheet == null)
                return null;

            var rowsCount = sheet.LastRowNum;

            if (rowsCount < 2)
                return null;

            var columnsCount = sheet.GetRow(0).LastCellNum;
            var content = GetHeader(sheet);

            for (var rowIndex = 1; rowIndex <= rowsCount; rowIndex++)
                content += GetRowData(sheet, columnsCount, rowIndex);

            var csvFile = CreateCsvFile(sheetName, content);

            return csvFile;
        }

        private static string GetHeader(ISheet sheet)
        {
            var cellIndex = 1;
            var rowString = string.Empty;
            var rowHeader = sheet.GetRow(0);
                
            for (cellIndex = 1; cellIndex < rowHeader.LastCellNum; cellIndex++)
                rowString += $";{rowHeader.GetCell(cellIndex).StringCellValue}";
            
            return rowString;
        }

        private static string GetRowData(ISheet sheet, short columnsCount, int rowIndex)
        {
            var cellIndex = 1;
            var rowString = string.Empty;
            var row = sheet.GetRow(rowIndex);
            var firstCell = row?.GetCell(0);
            var isFirstCellHasValue = !string.IsNullOrEmpty(firstCell?.StringCellValue);

            if (!isFirstCellHasValue)
                return rowString;

            var firstCellValue = $@"""{firstCell.StringCellValue.Trim()}""";

            rowString += $"\n{firstCellValue}";
            for (cellIndex = 1; cellIndex < columnsCount; cellIndex++)
                rowString += $";{(row.GetCell(cellIndex).NumericCellValue).ToString(CultureInfo.InvariantCulture).Replace(",", ".")}";

            return rowString;
        }

        private static string CreateCsvFile(string sheetName, string content)
        {
            var csvFileName = $"{Path.GetFileNameWithoutExtension(_excelFile)}_{sheetName}.csv";
            var csvFile = Path.Combine(_csvDirectory, csvFileName);

            using (var streamWriter = new StreamWriter(csvFile, false, Encoding.GetEncoding(1251)))
            {
                streamWriter.Write(content);
            }

            return csvFileName;
        }
    }
}