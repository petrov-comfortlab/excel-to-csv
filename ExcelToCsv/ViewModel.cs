using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using ExcelToCsv.Annotations;
using Microsoft.Win32;

namespace ExcelToCsv
{
    public class ViewModel : INotifyPropertyChanged
    {
        private string _workDirectory;
        private ObservableCollection<File> _excelFiles;
        private ObservableCollection<Page> _excelFilePages;
        private ObservableCollection<File> _csvFiles;

        private static readonly List<string> ExcelExtensions = new List<string>
        {
            ".xls",
            ".xlsx",
        };


        public ViewModel()
        {
            SetDefaultExcelFiles();
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

        #region Commands

        public ICommand OpenFolderCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = OpenFolder
        };

        private void OpenFolder()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls|All files (*.*)|*.*",
                InitialDirectory = Directory.GetCurrentDirectory(),
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() != true)
                return;

            WorkDirectory = Path.GetDirectoryName(openFileDialog.FileName);
        }

        public ICommand OpenDefaultFolderCommand => new CommandHandler
        {
            CanExecuteFunc = () => true,
            CommandAction = () => WorkDirectory = Directory.GetCurrentDirectory()
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