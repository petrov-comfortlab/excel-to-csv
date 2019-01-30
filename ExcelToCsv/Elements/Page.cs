using System.ComponentModel;
using System.Runtime.CompilerServices;
using ExcelToCsv.Annotations;

namespace ExcelToCsv.Elements
{
    public class Page : INotifyPropertyChanged
    {
        private bool _isSelected;

        public Page(string pageName) => PageName = pageName;

        public string PageName { get; }

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