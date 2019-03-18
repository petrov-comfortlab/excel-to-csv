using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelToCsv.Helpers
{
    public class ExcelManager
    {
        private static ExcelManager _instance;

        private string _csvDirectory;
        private string _excelFile;
        private string _copyExcelFile;
        private IWorkbook _workbook;

        private ExcelManager() { }

        private static ExcelManager Instance => _instance ?? (_instance = new ExcelManager());

        public static string ExcelFile
        {
            get => Instance._excelFile;
            set
            {
                if (string.IsNullOrEmpty(value))
                    return;

                Dispose();

                Instance._copyExcelFile = FileManager.CreateTemplateFile(value);
                Instance._excelFile = value;
                Instance._csvDirectory = GetCsvDirectory(value);

                using (var fileStream = new FileStream(Instance._copyExcelFile, FileMode.Open, FileAccess.Read))
                {
                    Instance.CreateWorkbook(fileStream, Path.GetExtension(Instance._copyExcelFile));
                }
            }
        }

        private static string GetCsvDirectory(string excelFile)
        {
            var directory = Path.GetDirectoryName(excelFile);
            var folder = Path.GetFileNameWithoutExtension(excelFile);

            if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(folder))
                throw new DirectoryNotFoundException();

            return Path.Combine(directory, folder);
        }

        private void CreateWorkbook(FileStream fileStream, string extension)
        {
            switch (extension)
            {
                case ".xls":
                    _workbook = new HSSFWorkbook(fileStream);
                    break;

                case ".xlsx":
                case ".xlsm":
                    _workbook = new XSSFWorkbook(fileStream);
                    break;

                default:
                    throw new FileFormatException();
            }
        }

        public static DataTable GetSheetDataTable(string sheetName)
        {
            if (Instance._workbook == null) throw new NullReferenceException(nameof(IWorkbook));

            try
            {
                var sheet = Instance._workbook.GetSheet(sheetName);

                if (sheet == null)
                    return new DataTable();

                var dataTable = new DataTable();

                for (var rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    AddDataTableRow(sheet, dataTable, rowIndex);

                return dataTable;
            }
            catch (Exception exception)
            {
                MessageBox.Show($"{exception.Message}\n\n{exception.StackTrace}");
                return new DataTable();
            }
        }

        #region GetSheetDataTable

        private static void AddDataTableRow(ISheet sheet, DataTable dataTable, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            var cellsNumber = row?.LastCellNum;
            var cellsValues = new List<object>();

            for (var cellIndex = 0; cellIndex < cellsNumber; cellIndex++)
            {
                if (cellIndex >= dataTable.Columns.Count)
                    dataTable.Columns.Add();

                var cell = row.GetCell(cellIndex);

                cellsValues.Add(cell != null ? GetCellValue(cell) : string.Empty);
            }

            var dataTableRow = dataTable.NewRow();
            dataTable.Rows.Add(dataTableRow);
            dataTableRow.ItemArray = cellsValues.ToArray();
        }

        #endregion

        public static void CreateCsvFile(string sheetName)
        {
            if (Instance._workbook == null) throw new NullReferenceException(nameof(IWorkbook));

            try
            {
                var sheet = Instance._workbook.GetSheet(sheetName);

                if (sheet == null)
                    return;

                var content = GetSheetContent(sheet);

                if (string.IsNullOrEmpty(content))
                    return;

                var csvFilePath = Instance.GetCsvFilePath(sheetName);

                CsvManager.CreateCsvFile(csvFilePath, content);
            }
            catch (Exception exception)
            {
                MessageBox.Show($"{exception.Message}\n\n{exception.StackTrace}");
                return;
            }
        }

        #region CreateCsvFile

        private static string GetSheetContent(ISheet sheet)
        {
            var rowsCount = sheet.LastRowNum;

            if (rowsCount == 0)
                return null;

            var columnsCount = sheet.GetRow(0).LastCellNum;
            var content = string.Empty;

            for (var rowIndex = 0; rowIndex <= rowsCount; rowIndex++)
            {
                var rowData = GetRowAsString(sheet, columnsCount, rowIndex);

                if (rowData.Any(n => n != ';' && n != '"'))
                    content += rowData + "\n";
            }

            return content;
        }

        private static string GetRowAsString(ISheet sheet, short columnsCount, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            var cellsValues = new List<string>();

            for (var cellIndex = 0; cellIndex < columnsCount; cellIndex++)
            {
                var cell = row.GetCell(cellIndex);

                cellsValues.Add(cell != null ? GetCellValue(cell) : string.Empty);
            }

            return string.Join(";", cellsValues);
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

        private string GetCsvFilePath(string sheetName)
        {
            if (!Directory.Exists(_csvDirectory))
                Directory.CreateDirectory(_csvDirectory);

            var csvFileName = $"{Path.GetFileNameWithoutExtension(_excelFile)}_{sheetName}.csv";
            return Path.Combine(_csvDirectory, csvFileName);
        }

        #endregion

        public static IEnumerable<string> GetCsvFiles()
        {
            try
            {
                if (string.IsNullOrEmpty(Instance._csvDirectory) ||
                    !Directory.Exists(Instance._csvDirectory))
                    return new List<string>();

                return Directory.GetFiles(Instance._csvDirectory)
                    .Where(n => string.Equals(Path.GetExtension(n), ".csv", StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception)
            {
                return new List<string>();
            }
        }

        public static List<(int Index, string Name)> GetExcelSheets()
        {
            if (Instance._workbook == null) throw new NullReferenceException(nameof(IWorkbook));

            if (string.IsNullOrEmpty(ExcelFile))
                return new List<(int, string)>();

            var sheetNames = new List<(int, string)>();
            var numberOfSheets = Instance._workbook.NumberOfSheets;

            for (var i = 0; i < numberOfSheets; i++)
            {
                var sheet = Instance._workbook.GetSheetAt(i);
                sheetNames.Add((i, sheet.SheetName));
            }

            return sheetNames;
        }

        public static void RenameSheet(int index, string newName)
        {
            if (Instance._workbook == null) throw new NullReferenceException(nameof(IWorkbook));

            if (!FileManager.GetAccessFile(ExcelFile) ||
                string.IsNullOrEmpty(newName))
                return;

            Instance._workbook.SetSheetName(index, newName);
            Instance.SaveExcelFile();
        }

        public static void CommentSheet(string sheetName)
        {
            if (Instance._workbook == null) throw new NullReferenceException(nameof(IWorkbook));

            if (!FileManager.GetAccessFile(ExcelFile) ||
                string.IsNullOrEmpty(sheetName) ||
                sheetName.StartsWith("#"))
                return;

            var sheetIndex = Instance._workbook.GetSheetIndex(sheetName);
            Instance._workbook.SetSheetName(sheetIndex, $"#{sheetName}");

            Instance.SaveExcelFile();
        }

        public static void UncommentSheet(string sheetName)
        {
            if (Instance._workbook == null) throw new NullReferenceException(nameof(IWorkbook));

            if (!FileManager.GetAccessFile(ExcelFile) ||
                string.IsNullOrEmpty(sheetName) ||
                !sheetName.StartsWith("#"))
                return;

            var sheetIndex = Instance._workbook.GetSheetIndex(sheetName);
            Instance._workbook.SetSheetName(sheetIndex, sheetName.TrimStart('#'));

            Instance.SaveExcelFile();
        }

        private void SaveExcelFile()
        {
            try
            {
                using (var fileStream = new FileStream(_excelFile, FileMode.Create, FileAccess.Write))
                {
                    _workbook.Write(fileStream);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, Resources.Dictionary.FileNotEditable, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        public static void Dispose()
        {
            if (File.Exists(Instance._copyExcelFile) &&
                FileManager.GetAccessFile(Instance._copyExcelFile, showMessage: false))
                File.Delete(Instance._copyExcelFile);

            Instance._workbook?.Close();

            _instance = null;
        }
    }
}