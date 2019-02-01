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
    public class ExcelManager : IDisposable
    {
        private readonly string _csvDirectory;
        private readonly string _excelFile;
        private readonly string _copyExcelFile;
        private readonly IWorkbook _workbook;

        public ExcelManager(string excelFile)
        {
            if (string.IsNullOrEmpty(excelFile))
                return;

            _copyExcelFile = CreateTemplateFile(excelFile);
            _excelFile = excelFile;
            _csvDirectory = GetCsvDirectory(excelFile);

            using (var fileStream = new FileStream(_copyExcelFile, FileMode.Open, FileAccess.Read))
                _workbook = new XSSFWorkbook(fileStream);
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
                var sheet = _workbook.GetSheet(sheetName);

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

        public List<string> GetExcelSheets(string excelFile)
        {
            if (string.IsNullOrEmpty(excelFile))
                return new List<string>();

            var sheetNames = new List<string>();
            var numberOfSheets = _workbook.NumberOfSheets;

            for (var i = 0; i < numberOfSheets; i++)
            {
                var sheet = _workbook.GetSheetAt(i);
                sheetNames.Add(sheet.SheetName);
            }

            return sheetNames;
        }

        private static string CreateTemplateFile(string excelFile)
        {
            var templatePath = Environment.GetFolderPath(Environment.SpecialFolder.Templates);
            var fileName = Path.GetFileName(excelFile);
            var tempExcelFile = Path.Combine(templatePath, fileName);

            File.Copy(excelFile, tempExcelFile, overwrite: true);
            return tempExcelFile;
        }

        public void Dispose()
        {
            File.Delete(_copyExcelFile);
        }

        public void CommentSheet(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName) ||
                sheetName.StartsWith("#"))
                return;

            using (var fileStream = new FileStream(_excelFile, FileMode.Open, FileAccess.ReadWrite))
            {
                var xssfWorkbook = new XSSFWorkbook(fileStream);
                var sheetIndex = xssfWorkbook.GetSheetIndex(sheetName);
                xssfWorkbook.SetSheetName(sheetIndex, $"#{sheetName}");
            }
        }

        public void UncommentSheet(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName) ||
                !sheetName.StartsWith("#"))
                return;

            //using (var fileStream = new FileStream(_excelFile, FileMode.Open, FileAccess.ReadWrite))
            //{
            //    var xssfWorkbook = new XSSFWorkbook(fileStream);
            //    var sheetIndex = xssfWorkbook.GetSheetIndex(sheetName);
            //    xssfWorkbook.SetSheetName(sheetIndex, sheetName.Substring(1));
            //    xssfWorkbook.Write(fileStream);
            //    fileStream.Close();
            //}

            //XSSFWorkbook xssfWorkbook;
            //using (var file = new FileStream(_copyExcelFile, FileMode.Open, FileAccess.ReadWrite))
            //{
            //    xssfWorkbook = new XSSFWorkbook(file);
            //}

            //var mstream = new MemoryStream();
            //xssfWorkbook.Write(mstream);

            //var fileStream = new FileStream(_excelFile, FileMode.Create, System.IO.FileAccess.Write);
            //var bytes = new byte[mstream.Length];
            //mstream.Read(bytes, 0, (int)mstream.Length);
            //fileStream.Write(bytes, 0, bytes.Length);
            //fileStream.Close();
            //mstream.Close();
        }
    }
}