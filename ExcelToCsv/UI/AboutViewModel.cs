using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using ExcelToCsv.Annotations;
using ExcelToCsv.Resources;

namespace ExcelToCsv.UI
{
    public class AboutViewModel : INotifyPropertyChanged
    {
        public AboutViewModel()
        {
            Company = $"© {AppData.Company}, {DateTime.Today.Year}";

            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
            Version = $"{Dictionary.Version} {fileVersionInfo.FileVersion}";
        }

        public string Company { get; set; }
        public string Version { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}