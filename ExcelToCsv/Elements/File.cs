using System;
using System.CodeDom;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using ExcelToCsv.Properties;

namespace ExcelToCsv.Elements
{
    public class File : INotifyPropertyChanged
    {
        private bool _isSelected;

        public File(string fullPath)
        {
            if (!System.IO.File.Exists(fullPath) &&
                !System.IO.Directory.Exists(fullPath))
                throw new FileNotFoundException();

            FullPath = fullPath;
        }

        public string FullPath { get; private set; }

        public string FileName
        {
            get => Path.GetFileName(FullPath);
            set => FullPath = Path.Combine(Path.GetDirectoryName(FullPath) ?? throw new InvalidOperationException(), value);
        }

        public string Directory => Path.GetDirectoryName(FullPath);

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}