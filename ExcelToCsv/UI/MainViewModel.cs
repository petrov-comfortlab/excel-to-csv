using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelToCsv.CommandHandlers;
using ExcelToCsv.Elements;
using ExcelToCsv.Helpers;
using ExcelToCsv.Properties;
using Microsoft.Win32;
using File = ExcelToCsv.Elements.File;

namespace ExcelToCsv.UI
{
    public class MainViewModel : INotifyPropertyChanged, IDisposable
    {
        private string _fileName;
        private string _workDirectory;
        private DataTable _dataTable;
        private File _selectedCsvFile;
        private File _selectedExcelFile;
        private Sheet _selectedSheet;
        private ObservableCollection<File> _quickDirectories;
        private ObservableCollection<File> _csvFiles;
        private ObservableCollection<File> _excelFiles;
        private ObservableCollection<Sheet> _excelFileSheets;
        private Visibility _fileContentVisibility;
        private Visibility _excelFilesVisibility;
        private Visibility _commentMenuItemVisibility;
        private Visibility _uncommentMenuItemVisibility;

        private static readonly List<string> ExcelExtensions = new List<string>
        {
            ".xls",
            ".xlsx",
            ".xlsm",
        };


        public MainViewModel()
        {
            SetQuickDirectories();
            SetDefaultDirectory();
            ShowExcelFiles();
        }

        private void SetQuickDirectories()
        {
            QuickDirectories = new ObservableCollection<File>(
                RegistryManager.LoadQuickDirectories()
                    .Where(Directory.Exists)
                    .Select(n => new File(n)));
        }

        private void SetDefaultDirectory()
        {
            WorkDirectory = QuickDirectories.Any()
                ? QuickDirectories.First().FullPath
                : Directory.GetCurrentDirectory();
        }

        public string WorkDirectory
        {
            get => _workDirectory;
            set
            {
                try
                {
                    _workDirectory = value;
                    OnPropertyChanged();
                    SetExcelFiles();
                    AddQuickDirectory(value);
                }
                catch (Exception exception)
                {
                    Logger.Log.Error(exception);
                }
            }
        }

        #region SetWorkDirectory

        private void SetExcelFiles()
        {
            var files = Directory.GetFiles(WorkDirectory)
                .Where(n => !Path.GetFileName(n).StartsWith("~$") &&
                            ExcelExtensions.Contains(Path.GetExtension(n)))
                .Select(n => new File(n));

            ExcelFiles = new ObservableCollection<File>(files);
        }

        private void AddQuickDirectory(string value)
        {
            if (QuickDirectories.Select(n => n.FullPath).Contains(value))
            {
                var existDirectory = QuickDirectories.FirstOrDefault(n => n.FullPath == value);
                var oldIndex = QuickDirectories.IndexOf(existDirectory);
                QuickDirectories.Move(oldIndex, 0);
            }
            else
            {
                QuickDirectories.Insert(0, new File(value));

                if (QuickDirectories.Count > 30)
                    QuickDirectories.Remove(QuickDirectories.Last());
            }
        }

        #endregion

        public ObservableCollection<File> QuickDirectories
        {
            get => _quickDirectories;
            set
            {
                _quickDirectories = value;
                OnPropertyChanged();
            }
        }

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
                try
                {
                    _selectedExcelFile = value;
                    ExcelManager.ExcelFile = value?.FullPath;

                    OnPropertyChanged();

                    SetExcelFileSheets();
                    SetCsvFiles(value);
                }
                catch (Exception exception)
                {
                    Logger.Log.Error(exception);
                }
            }
        }

        #region SetSelectedExcelFile

        private void SetExcelFileSheets()
        {
            var sheets = SelectedExcelFile != null
                ? ExcelManager.GetExcelSheets().Select(n => new Sheet(SelectedExcelFile.FullPath, n.Index, n.Name))
                : new List<Sheet>();

            ExcelFileSheets = new ObservableCollection<Sheet>(sheets);
        }

        private void SetCsvFiles(File excelFile)
        {
            CsvFiles = excelFile != null
                ? new ObservableCollection<File>(ExcelManager.GetCsvFiles().Select(n => new File(n)))
                : new ObservableCollection<File>();
        }

        #endregion

        public Sheet SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                try
                {
                    if (_selectedSheet != null)
                        _selectedSheet.PropertyChanged -= SelectedSheetOnPropertyChanged;

                    _selectedSheet = value;
                    OnPropertyChanged();

                    SetSelectedSheet(value);

                    value.PropertyChanged += SelectedSheetOnPropertyChanged;
                }
                catch (Exception exception)
                {
                    Logger.Log.Error(exception);
                }
            }
        }

        #region SetSelectedSheet

        private void SetSelectedSheet(Sheet sheet)
        {
            SetMenuItemsVisibility(sheet);

            if (sheet == null)
                return;

            ShowFileContent(sheet);
            FileName = $"{SelectedExcelFile?.FileName}: {sheet.SheetName}";
            SetExcelSheetDataGrid(SelectedExcelFile?.FullPath, sheet.SheetName);
        }

        private void SetExcelSheetDataGrid(string file, string sheetName)
        {
            if (string.IsNullOrEmpty(file) ||
                !System.IO.File.Exists(file) ||
                !ExcelExtensions.Contains(Path.GetExtension(file).ToLower()))
                return;

            DataTable = ExcelManager.GetSheetDataTable(sheetName);
        }

        private void SelectedSheetOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Sheet.SheetName))
                ExcelManager.RenameSheet(SelectedSheet.Index, SelectedSheet.SheetName);
        }

        #endregion

        public File SelectedCsvFile
        {
            get => _selectedCsvFile;
            set
            {
                try
                {
                    if (value == null) return;
                    _selectedCsvFile = value;
                    OnPropertyChanged();

                    SetSelectedCsvFile(value);
                }
                catch (Exception exception)
                {
                    Logger.Log.Error(exception);
                }
            }
        }

        #region SetSelectedCsvFile

        private void SetSelectedCsvFile(File value)
        {
            ShowFileContent(value);
            FileName = value.FileName;
            DataTable = CsvManager.GetDataTable(value.FullPath);
        }

        #endregion

        public string FileName
        {
            get => _fileName;
            set
            {
                if (value == _fileName) return;
                _fileName = value;
                OnPropertyChanged();
            }
        }

        public DataTable DataTable
        {
            get => _dataTable;
            set
            {
                _dataTable = value;
                OnPropertyChanged();
            }
        }

        public Visibility FileContentVisibility
        {
            get => _fileContentVisibility;
            set
            {
                if (value == _fileContentVisibility) return;
                _fileContentVisibility = value;
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

        private void ShowFileContent(object value)
        {
            if (value == null)
                return;

            FileContentVisibility = Visibility.Visible;
            ExcelFilesVisibility = Visibility.Collapsed;
        }

        private void ShowExcelFiles()
        {
            FileContentVisibility = Visibility.Collapsed;
            ExcelFilesVisibility = Visibility.Visible;
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

        #region Commands

        #region Folders Commands

        public ICommand OpenFolderInNewWindowCommand => new CommandHandler
        {
            CanExecuteFunc = () => Directory.Exists(WorkDirectory),
            CommandAction = OpenFolderInNewWindow
        };

        #region OpenFolderInNewWindow

        private void OpenFolderInNewWindow()
        {
            try
            {
                Process.Start(WorkDirectory);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        #endregion

        public ICommand OpenFolderCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = OpenFolder
        };

        #region OpenFolder

        private void OpenFolder()
        {
            try
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
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        #endregion

        public ICommand OpenFastFolderCommand => new CommandParameterHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = OpenFastFolder
        };

        #region OpenFastFolder

        private void OpenFastFolder(object obj)
        {
            try
            {
                if (!(obj is File file))
                    return;

                ShowExcelFiles();
                WorkDirectory = file.FullPath;
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        #endregion

        #endregion

        #region CSV Files Commands

        public ICommand CreateCsvFilesCommand => new CommandHandler
        {
            CanExecuteFunc = () => SelectedExcelFile != null && (ExcelFileSheets?.Any() ?? false),
            CommandAction = CreateCsvFiles
        };

        #region CreateCsvFiles

        private void CreateCsvFiles()
        {
            try
            {
                var selectedFiles = ExcelFiles.Where(n => n.IsSelected).Select(n => n.FullPath).ToList();
                selectedFiles.ForEach(CreateCsvFile);
                SetCsvFiles(SelectedExcelFile);
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        private static void CreateCsvFile(string excelFile)
        {
            ExcelManager.ExcelFile = excelFile;

            var sheets = ExcelManager.GetExcelSheets();

            foreach (var sheet in sheets)
            {
                if (sheet.Name.StartsWith("#"))
                    continue;

                ExcelManager.CreateCsvFile(sheet.Name);
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
            try
            {
                var selectedFiles = ExcelFiles.Where(n => n.IsSelected).Select(n => n.FullPath).ToList();
                selectedFiles.ForEach(DeleteCsvFile);
                SetCsvFiles(SelectedExcelFile);
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        private void DeleteCsvFile(string excelFile)
        {
            var csvFolder = Path.GetFileNameWithoutExtension(excelFile);
            var csvDirectory = Path.Combine(WorkDirectory, csvFolder ?? throw new InvalidOperationException());

            if (Directory.Exists(csvDirectory) &&
                FileManager.GetAccessDirectory(csvDirectory))
                Directory.Delete(csvDirectory, recursive: true);
        }

        #endregion

        #endregion

        #region Excel Files Commands

        public ICommand OpenExcelFileCommand => new CommandHandler
        {
            CanExecuteFunc = () => SelectedExcelFile != null,
            CommandAction = OpenExcelFile
        };

        #region OpenExcelFile

        private void OpenExcelFile()
        {
            try
            {
                ExcelFiles.Where(n => System.IO.File.Exists(n.FullPath) && n.IsSelected).ToList().ForEach(n => Process.Start(n.FullPath));
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        #endregion

        public ICommand SelectAllExcelFilesCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = () => ExcelFiles.ToList().ForEach(n => n.IsSelected = true)
        };

        #endregion

        #region Sheets Commands

        public ICommand RenameSelectedSheetCommand => new CommandParameterHandler
        {
            CanExecuteFunc = () => SelectedSheet != null,
            CommandAction = RenameSelectedSheet
        };

        #region RenameSelectedSheet

        private static void RenameSelectedSheet(object obj)
        {
            try
            {
                if (!(obj is ListBox listBox))
                    return;

                var listBoxItem = (ListBoxItem)listBox.ItemContainerGenerator.ContainerFromItem(listBox.SelectedItem);

                if (CheckFileAccess(listBoxItem))
                    ListBoxItemManager.RenameItem(listBoxItem);
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        private static bool CheckFileAccess(ListBoxItem listBoxItem)
        {
            return listBoxItem?.DataContext is Sheet sheet &&
                   FileManager.GetAccessFile(sheet.File);
        }

        #endregion

        public ICommand CommentSheetCommand => new CommandHandler
        {
            CanExecuteFunc = () => !SelectedSheet?.SheetName?.StartsWith("#") ?? false,
            CommandAction = CommentSheet
        };

        #region CommentSheet

        private void CommentSheet()
        {
            try
            {
                ExcelManager.CommentSheet(SelectedSheet?.SheetName);
                SetExcelFileSheets();
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        #endregion

        public ICommand UncommentSheetCommand => new CommandHandler
        {
            CanExecuteFunc = () => SelectedSheet?.SheetName?.StartsWith("#") ?? false,
            CommandAction = UncommentSheet
        };

        #region UncommentSheet

        private void UncommentSheet()
        {
            try
            {
                ExcelManager.UncommentSheet(SelectedSheet?.SheetName);
                SetExcelFileSheets();
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        #endregion

        #endregion

        public ICommand CloseFileCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = ShowExcelFiles
        };

        public ICommand CsvFilesGotFocusCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = () => SelectedCsvFile = SelectedCsvFile
        };

        public ICommand ExcelSheetsGotFocusCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = () => SelectedSheet = SelectedSheet
        };

        public ICommand AboutCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = () => new AboutWindow().ShowDialog()
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
            RegistryManager.SaveQuickDirectories(QuickDirectories.Select(n => n.FullPath));

            _dataTable?.Dispose();
            ExcelManager.Dispose();
        }
    }
}