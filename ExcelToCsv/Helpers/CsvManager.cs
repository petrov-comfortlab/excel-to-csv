using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;

namespace ExcelToCsv.Helpers
{
    public class CsvManager
    {
        private static readonly Regex CsvRegex = new Regex("(?<=^|;)(\"(?:[^\"]|\"\")*\"|[^;]*)");

        public static void CreateCsvFile(string csvFile, string content)
        {
            if (File.Exists(csvFile) &&
                !FileManager.GetAccessFile(csvFile))
                return;

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

        public static DataTable GetDataTable(string file)
        {
            if (string.IsNullOrEmpty(file) ||
                !File.Exists(file) ||
                Path.GetExtension(file).ToLower() != ".csv")
                return new DataTable();

            var lines = GetLines(file);
            var allCells = new List<List<object>>();
            var columnNumber = 0;

            foreach (var line in lines)
            {
                var cells = GetCells(line);

                if (columnNumber < cells.Count)
                    columnNumber = cells.Count;

                allCells.Add(cells);
            }

            return GetDataTable(allCells, columnNumber);
        }

        private static IEnumerable<string> GetLines(string file)
        {
            var templateCsvFile = FileManager.CreateTemplateFile(file);
            var lines = File.ReadLines(templateCsvFile, Encoding.Default).ToList();

            if (File.Exists(templateCsvFile))
                File.Delete(templateCsvFile);

            return lines;
        }

        private static List<object> GetCells(string line)
        {
            var cells = new List<object>();

            foreach (Match m in CsvRegex.Matches(line))
                cells.Add(m.Groups[1].ToString().Trim('"'));

            return cells;
        }

        private static DataTable GetDataTable(IEnumerable<List<object>> allCells, int columnNumber)
        {
            var dataTable = new DataTable();

            for (var i = 0; i < columnNumber; i++)
                dataTable.Columns.Add();

            foreach (var rowCells in allCells)
                dataTable.Rows.Add(rowCells.ToArray());

            return dataTable;
        }
    }
}