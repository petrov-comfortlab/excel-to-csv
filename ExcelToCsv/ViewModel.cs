using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using ExcelToCsv.Annotations;
using ExcelToCsv.Elements;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using File = ExcelToCsv.Elements.File;

namespace ExcelToCsv
{
    public class ViewModel : INotifyPropertyChanged
    {
        private string _workDirectory;
        private ObservableCollection<File> _excelFiles;
        private ObservableCollection<Page> _excelFilePages;
        private ObservableCollection<File> _csvFiles;
        private File _selectedExcelFile;

        private static readonly List<string> ExcelExtensions = new List<string>
        {
            ".xls",
            ".xlsx",
            ".xlsm",
        };


        public ViewModel()
        {
            WorkDirectory = $@"D:\Серёжа\#ostec\Детали";
            //SetDefaultExcelFiles();
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
                .Where(n => ExcelExtensions.Contains(Path.GetExtension(n)))
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

        public ObservableCollection<Page> ExcelFilePages
        {
            get => _excelFilePages;
            set
            {
                _excelFilePages = value;
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
                OnPropertyChanged();
                SetExcelFilePages(value);
                SetCsvFiles(value);
            }
        }

        #region SetSelectedExcelFile

        private void SetExcelFilePages(File excelFile)
        {
            var pages = ExcelManager.GetExcelPages(excelFile?.FullPath).Select(n => new Page(n));
            ExcelFilePages = new ObservableCollection<Page>(pages);
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
            CanExecuteFunc = () => true,
            CommandAction = CreateCsvFiles
        };

        #region CreateCsvFiles

        private void CreateCsvFiles()
        {
            var selectedFiles = ExcelFiles.Where(n => n.IsSelected).Select(n => n.FullPath).ToList();
            selectedFiles.ForEach(CreateCsvFile);
            SetCsvFiles(SelectedExcelFile);
        }

        private static void CreateCsvFile(string excelFile)
        {
            var excelManager = new ExcelManager(excelFile);
            var pages = ExcelManager.GetExcelPages(excelFile);

            foreach (var page in pages)
            {
                if (page.StartsWith("#"))
                    continue;

                excelManager.CreateCsvFile(page);
            }
        }

        #endregion

        public ICommand DeleteCsvFilesCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
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
            CanExecuteFunc = () => true,
            CommandAction = () => ExcelFiles.Where(n => n.IsSelected).ToList().ForEach(n => Process.Start(n.FullPath))
        };

        #endregion

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}