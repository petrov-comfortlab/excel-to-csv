using System.Diagnostics;
using System.Windows;
using System.Windows.Documents;

namespace ExcelToCsv.UI
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();

            this.ShowInTaskbar = false;
            this.Topmost = true;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            DataContext = new AboutViewModel();
        }

        private void Hyperlink_OnClick(object sender, RoutedEventArgs e)
        {
            var hyperlink = (Hyperlink)e.Source;
            Process.Start(hyperlink.NavigateUri.AbsoluteUri);
        }
    }
}
