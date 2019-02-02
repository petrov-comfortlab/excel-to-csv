using System.Windows;
using System.Windows.Controls;

namespace ExcelToCsv
{
    public partial class MainWindow : Window
    {
        private readonly ViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new ViewModel();
            DataContext = _viewModel;

            Closed += (sender, args) => { _viewModel.Dispose(); };
        }

        private void Selector_OnSelected(object sender, RoutedEventArgs e)
        {
            if (!(sender is ListBox listBox))
                return;

            listBox.SelectedItem = null;
        }
    }
}
