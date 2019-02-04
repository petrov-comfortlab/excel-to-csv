using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelToCsv.Elements;
using ExcelToCsv.Helpers;

namespace ExcelToCsv.UI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            try
            {
                InitializeComponent();

                var viewModel = new ViewModel();
                DataContext = viewModel;

                Closed += (sender, args) => { viewModel.Dispose(); };
            }
            catch (Exception exception)
            {
                Logger.Log.Error(exception);
            }
        }

        private void Selector_OnSelected(object sender, RoutedEventArgs e)
        {
            if (!(sender is ListBox listBox))
                return;

            listBox.SelectedItem = null;
        }

        private void ListBoxItem_OnHandler(object sender, KeyEventArgs e)
        {
            if (!(sender is ListBoxItem listBoxItem &&
                  CheckFileAccess(listBoxItem)))
                return;

            if (e.Key == Key.F2)
                ListBoxItemManager.RenameItem(listBoxItem);
        }

        private void ListBoxItem_OnHandler(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is ListBoxItem listBoxItem &&
                  CheckFileAccess(listBoxItem)))
                return;

            if (e.ChangedButton == MouseButton.Left)
                ListBoxItemManager.RenameItem(listBoxItem);
        }

        private static bool CheckFileAccess(ListBoxItem listBoxItem)
        {
            return listBoxItem.DataContext is Sheet sheet &&
                   FileManager.GetAccessFile(sheet.File);
        }
    }
}
