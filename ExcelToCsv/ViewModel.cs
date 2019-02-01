using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using ExcelToCsv.Annotations;
using ExcelToCsv.Elements;
using Microsoft.Win32;
using File = ExcelToCsv.Elements.File;

namespace ExcelToCsv
{
    public class ViewModel : INotifyPropertyChanged, IDisposable
    {
        private string _csvFileName;
        private string _workDirectory;
        private DataTable _csvDataGrid;
        private ExcelManager _excelManager;
        private File _selectedCsvFile;
        private File _selectedExcelFile;
        private ObservableCollection<File> _excelFiles;
        private ObservableCollection<Sheet> _excelFileSheets;
        private ObservableCollection<File> _csvFiles;
        private Visibility _csvTextVisibility;
        private Visibility _excelFilesVisibility;
        private Sheet _selectedSheet;
        private Visibility _commentMenuItemVisibility;
        private Visibility _uncommentMenuItemVisibility;

        private static readonly List<string> ExcelExtensions = new List<string>
        {
            ".xls",
            ".xlsx",
            ".xlsm",
        };


        public ViewModel()
        {
            WorkDirectory = $@"E:\YanDisk\#ostec\Setup\Source Files\Common\!_Excel\Детали";
            //SetDefaultExcelFiles();

            ShowExcelFiles();
        }

        private void SetDefaultExcelFiles()
        {
            WorkDirectory = Directory.GetCurrentDirectory();
        }

        public string WorkDirectory
        {
            get => _workDirectory;
            set
            {
                _workDirectory = value;
                OnPropertyChanged();
                SetExcelFiles();
            }
        }

        #region SetExcelFiles

        private void SetExcelFiles()
        {
            var files = Directory.GetFiles(WorkDirectory)
                .Where(n => !Path.GetFileName(n).StartsWith("~$") &&
                            ExcelExtensions.Contains(Path.GetExtension(n)))
                .Select(n => new File(n));

            ExcelFiles = new ObservableCollection<File>(files);
        }

        #endregion

        public ObservableCollection<File> ExcelFiles
        {
            get => _excelFiles;
            set
            {
                _excelFiles = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<Sheet> ExcelFileSheets
        {
            get => _excelFileSheets;
            set
            {
                _excelFileSheets = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<File> CsvFiles
        {
            get => _csvFiles;
            set
            {
                _csvFiles = value;
                OnPropertyChanged();
            }
        }

        public File SelectedExcelFile
        {
            get => _selectedExcelFile;
            set
            {
                _selectedExcelFile = value;
                _excelManager = new ExcelManager(value?.FullPath);

                OnPropertyChanged();

                SetExcelFilePages(value);
                SetCsvFiles(value);
            }
        }

        #region SetSelectedExcelFile

        private void SetExcelFilePages(File excelFile)
        {
            var pages = _excelManager.GetExcelSheets(excelFile?.FullPath).Select(n => new Sheet(n));
            ExcelFileSheets = new ObservableCollection<Sheet>(pages);
        }

        private void SetCsvFiles(File excelFile)
        {
            if (excelFile == null)
            {
                CsvFiles = new ObservableCollection<File>();
                return;
            }

            var csvFolder = Path.GetFileNameWithoutExtension(excelFile.FullPath);
            var csvDirectory = Path.Combine(excelFile.Directory, csvFolder ?? throw new InvalidOperationException());

            if (!Directory.Exists(csvDirectory))
            {
                CsvFiles = new ObservableCollection<File>();
                return;
            }

            var files = Directory.GetFiles(csvDirectory)
                .Where(n => string.Equals(Path.GetExtension(n), ".csv", StringComparison.OrdinalIgnoreCase))
                .Select(n => new File(n));

            CsvFiles = new ObservableCollection<File>(files);
        }

        #endregion

        public Sheet SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                if (Equals(value, _selectedSheet)) return;
                _selectedSheet = value;
                OnPropertyChanged();

                SetMenuItemsVisibility(value);
            }
        }

        private void SetMenuItemsVisibility(Sheet value)
        {
            if (string.IsNullOrEmpty(value?.SheetName))
            {
                CommentMenuItemVisibility = Visibility.Collapsed;
                UncommentMenuItemVisibility = Visibility.Collapsed;
                return;
            }

            if (value.SheetName.StartsWith("#"))
            {
                CommentMenuItemVisibility = Visibility.Collapsed;
                UncommentMenuItemVisibility = Visibility.Visible;
            }
            else
            {
                CommentMenuItemVisibility = Visibility.Visible;
                UncommentMenuItemVisibility = Visibility.Collapsed;
            }
        }

        public File SelectedCsvFile
        {
            get => _selectedCsvFile;
            set
            {
                if (value == null || Equals(value, _selectedCsvFile)) return;
                _selectedCsvFile = value;
                OnPropertyChanged();
                ShowCsvFiles();
                CsvFileName = SelectedCsvFile?.FileName;
                SetCsvDataGrid(SelectedCsvFile?.FullPath);
            }
        }

        private void SetCsvDataGrid(string file)
        {
            if (string.IsNullOrEmpty(file) ||
                !System.IO.File.Exists(file) ||
                Path.GetExtension(file).ToLower() != ".csv")
                return;

            var lines = System.IO.File.ReadLines(file, Encoding.GetEncoding(1251)).ToList();
            var allCells = new List<List<object>>();

            var regex = new Regex("(?<=^|;)(\"(?:[^\"]|\"\")*\"|[^;]*)");
            var columnNumber = 0;

            foreach (var line in lines)
            {
                var cells = new List<object>();

                foreach (Match m in regex.Matches(line))
                {
                    cells.Add(m.Groups[1].ToString().Trim('"'));
                }
                
                if (columnNumber < cells.Count)
                    columnNumber = cells.Count;

                allCells.Add(cells);
            }

            var dataTable = new DataTable();

            for (var i = 0; i < columnNumber; i++)
                dataTable.Columns.Add();

            foreach (var rowCells in allCells)
                dataTable.Rows.Add(rowCells.ToArray());

            CsvDataGrid = dataTable;
        }

        public string CsvFileName
        {
            get => _csvFileName;
            set
            {
                if (value == _csvFileName) return;
                _csvFileName = value;
                OnPropertyChanged();
            }
        }

        public DataTable CsvDataGrid
        {
            get => _csvDataGrid;
            set
            {
                _csvDataGrid = value;
                OnPropertyChanged();
            }
        }

        public Visibility CsvTextVisibility
        {
            get => _csvTextVisibility;
            set
            {
                if (value == _csvTextVisibility) return;
                _csvTextVisibility = value;
                OnPropertyChanged();
            }
        }

        public Visibility ExcelFilesVisibility
        {
            get => _excelFilesVisibility;
            set
            {
                if (value == _excelFilesVisibility) return;
                _excelFilesVisibility = value;
                OnPropertyChanged();
            }
        }

        public Visibility CommentMenuItemVisibility
        {
            get => _commentMenuItemVisibility;
            set
            {
                if (value == _commentMenuItemVisibility) return;
                _commentMenuItemVisibility = value;
                OnPropertyChanged();
            }
        }

        public Visibility UncommentMenuItemVisibility
        {
            get => _uncommentMenuItemVisibility;
            set
            {
                if (value == _uncommentMenuItemVisibility) return;
                _uncommentMenuItemVisibility = value;
                OnPropertyChanged();
            }
        }

        private void ShowCsvFiles()
        {
            CsvTextVisibility = Visibility.Visible;
            ExcelFilesVisibility = Visibility.Collapsed;
        }

        private void ShowExcelFiles()
        {
            CsvTextVisibility = Visibility.Collapsed;
            ExcelFilesVisibility = Visibility.Visible;
        }

        #region Commands

        public ICommand GoToFolderCommand => new CommandHandler
        {
            CanExecuteFunc = () => Directory.Exists(WorkDirectory),
            CommandAction = () => Process.Start(WorkDirectory)
        };

        public ICommand OpenFolderCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = OpenFolder
        };

        #region OpenFolder

        private void OpenFolder()
        {
            var excelExtensionsFilter = string.Join(";", ExcelExtensions.Select(n => $"*{n}"));
            var openFileDialog = new OpenFileDialog
            {
                Filter = $"Excel files ({excelExtensionsFilter})|{excelExtensionsFilter}|All files (*.*)|*.*",
                InitialDirectory = WorkDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() != true)
                return;

            WorkDirectory = Path.GetDirectoryName(openFileDialog.FileName);
        }

        #endregion

        public ICommand OpenDefaultFolderCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = SetDefaultExcelFiles
        };

        public ICommand CreateCsvFilesCommand => new CommandHandler
        {
            CanExecuteFunc = () => SelectedExcelFile != null && (ExcelFileSheets?.Any() ?? false),
            CommandAction = CreateCsvFiles
        };

        #region CreateCsvFiles

        private void CreateCsvFiles()
        {
            var selectedFiles = ExcelFiles.Where(n => n.IsSelected).Select(n => n.FullPath).ToList();
            selectedFiles.ForEach(CreateCsvFile);
            SetCsvFiles(SelectedExcelFile);
        }

        private void CreateCsvFile(string excelFile)
        {
            _excelManager = new ExcelManager(excelFile);

            var pages = _excelManager.GetExcelSheets(excelFile);

            foreach (var page in pages)
            {
                if (page.StartsWith("#"))
                    continue;

                _excelManager.CreateCsvFile(page);
            }
        }

        #endregion

        public ICommand DeleteCsvFilesCommand => new CommandHandler
        {
            CanExecuteFunc = () => SelectedExcelFile != null && (CsvFiles?.Any() ?? false),
            CommandAction = DeleteCsvFiles
        };

        #region DeleteCsvFiles

        private void DeleteCsvFiles()
        {
            var selectedFiles = ExcelFiles.Where(n => n.IsSelected).Select(n => n.FullPath).ToList();
            selectedFiles.ForEach(DeleteCavFile);
            SetCsvFiles(SelectedExcelFile);
        }

        private void DeleteCavFile(string excelFile)
        {
            var csvFolder = Path.GetFileNameWithoutExtension(excelFile);
            var csvDirectory = Path.Combine(WorkDirectory, csvFolder ?? throw new InvalidOperationException());

            if (Directory.Exists(csvDirectory))
                Directory.Delete(csvDirectory, recursive: true);
        }

        #endregion

        public ICommand OpenExcelFileCommand => new CommandHandler
        {
            CanExecuteFunc = () => SelectedExcelFile != null,
            CommandAction = () => ExcelFiles.Where(n => n.IsSelected).ToList().ForEach(n => Process.Start(n.FullPath))
        };

        public ICommand CloseCsvFileCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = ShowExcelFiles
        };

        public ICommand CommentSheetCommand => new CommandHandler
        {
            CanExecuteFunc = () => !SelectedSheet?.SheetName?.StartsWith("#") ?? false,
            CommandAction = () => _excelManager.CommentSheet(SelectedSheet?.SheetName)
        };

        public ICommand UncommentSheetCommand => new CommandHandler
        {
            CanExecuteFunc = () => SelectedSheet?.SheetName?.StartsWith("#") ?? false,
            CommandAction = () => _excelManager.UncommentSheet(SelectedSheet?.SheetName)
        };

        #endregion

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void Dispose()
        {
            _csvDataGrid?.Dispose();
            _excelManager?.Dispose();
        }
    }
}