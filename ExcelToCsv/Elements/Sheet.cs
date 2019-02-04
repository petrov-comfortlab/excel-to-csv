using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using ExcelToCsv.Properties;

namespace ExcelToCsv.Elements
{
    public class Sheet : INotifyPropertyChanged
    {
        private string _sheetName;

        public Sheet(string file, int index, string sheetName)
        {
            if (!System.IO.File.Exists(file)) throw new FileNotFoundException();

            File = file;
            Index = index;
            SheetName = sheetName;
        }

        public int Index { get; }
        public string File { get; }

        public string SheetName
        {
            get => _sheetName;
            set
            {
                _sheetName = value;
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