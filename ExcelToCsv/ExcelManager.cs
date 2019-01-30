using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelToCsv
{
    public class ExcelManager
    {
        private readonly string _csvDirectory;
        private readonly string _excelFile;
        private readonly IWorkbook _xssfWorkbook;

        public ExcelManager(string excelFile)
        {
            if (string.IsNullOrEmpty(excelFile))
                return;

            using (var fileStream = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                _xssfWorkbook = new XSSFWorkbook(fileStream);

            _excelFile = excelFile;
            _csvDirectory = GetCsvDirectory(excelFile);
        }

        private static string GetCsvDirectory(string excelFile)
        {
            var directory = Path.GetDirectoryName(excelFile);
            var folder = Path.GetFileNameWithoutExtension(excelFile);

            if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(folder))
                throw new DirectoryNotFoundException();

            var csvDirectory = Path.Combine(directory, folder);

            if (Directory.Exists(csvDirectory))
                return csvDirectory;

            var directoryInfo = Directory.CreateDirectory(csvDirectory);

            return directoryInfo.Exists ? directoryInfo.FullName : throw new DirectoryNotFoundException();
        }

        public void CreateCsvFile(string sheetName)
        {
            try
            {
                var sheet = _xssfWorkbook.GetSheet(sheetName);

                if (sheet == null)
                    return;

                var rowsCount = sheet.LastRowNum;

                if (rowsCount < 2)
                    return;

                var columnsCount = sheet.GetRow(0).LastCellNum;
                var content = string.Empty;

                for (var rowIndex = 0; rowIndex <= rowsCount; rowIndex++)
                {
                    var rowData = GetRowData(sheet, columnsCount, rowIndex);

                    if (rowData.Any(n => n != ';' && n != '"'))
                        content += rowData + "\n";
                }

                CreateCsvFile(sheetName, content);
            }
            catch (Exception exception)
            {
                MessageBox.Show($"{exception.Message}\n\n{exception.StackTrace}");
                return;
            }
        }

        private static string GetRowData(ISheet sheet, short columnsCount, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            var cellValues = new List<string>();

            for (var cellIndex = 0; cellIndex < columnsCount; cellIndex++)
            {
                var cell = row.GetCell(cellIndex);

                if (cell == null)
                {
                    cellValues.Add(string.Empty);
                    continue;
                }

                var cellValue = GetCellValue(cell);

                cellValues.Add(cellValue);
            }

            return string.Join(";", cellValues);
        }

        private static string GetCellValue(ICell cell, CellType cellType = CellType.Unknown)
        {
            switch (cellType != CellType.Unknown ? cellType : cell.CellType)
            {
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture).Replace(",", ".");
                case CellType.String:
                    return $@"""{cell.StringCellValue.Trim()}""";
                case CellType.Boolean:
                    return cell.BooleanCellValue ? "1" : "0";
                case CellType.Formula:
                    return GetCellValue(cell, cell.CachedFormulaResultType);
                default:
                    return string.Empty;
            }
        }

        private void CreateCsvFile(string sheetName, string content)
        {
            var csvFileName = $"{Path.GetFileNameWithoutExtension(_excelFile)}_{sheetName}.csv";
            var csvFile = Path.Combine(_csvDirectory, csvFileName);

            try
            {
                using (var streamWriter = new StreamWriter(csvFile, false, Encoding.GetEncoding(1251)))
                {
                    streamWriter.Write(content);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public static List<string> GetExcelPages(string excelFile)
        {
            if (string.IsNullOrEmpty(excelFile))
                return new List<string>();

            IWorkbook xssfWorkbook;
            using (var fileStream = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                xssfWorkbook = new XSSFWorkbook(fileStream);

            var sheetNames = new List<string>();
            var numberOfSheets = xssfWorkbook.NumberOfSheets;

            for (var i = 0; i < numberOfSheets; i++)
            {
                var sheet = xssfWorkbook.GetSheetAt(i);
                sheetNames.Add(sheet.SheetName);
            }

            return sheetNames;
        }
    }
}