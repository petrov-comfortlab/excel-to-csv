using System.ComponentModel;
using System.Runtime.CompilerServices;
using ExcelToCsv.Annotations;

namespace ExcelToCsv
{
    public class ViewModel : INotifyPropertyChanged
    {
        public ViewModel()
        {

        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}