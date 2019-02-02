using System.ComponentModel;
using System.Runtime.CompilerServices;
using ExcelToCsv.Properties;

namespace ExcelToCsv.Elements
{
    public class Sheet : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _sheetName;

        public Sheet(string pageName) => SheetName = pageName;

        public string SheetName
        {
            get => _sheetName;
            set
            {
                _sheetName = value;
                OnPropertyChanged();
            }
        }

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