using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ExcelToCsv
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new ViewModel();
        }

        private void EventSetter_OnHandler(object sender, MouseButtonEventArgs e)
        {
            var listBoxItem = (ListBoxItem)sender;
            listBoxItem.IsSelected = true;
        }
    }

    public class MyListView : ListView
    {
        protected override void OnPreviewMouseDown(System.Windows.Input.MouseButtonEventArgs e)
        {
            DependencyObject listViewItem = (DependencyObject)e.OriginalSource;
            while (listViewItem != null && !(listViewItem is ListViewItem))
                listViewItem = VisualTreeHelper.GetParent(listViewItem);

            SelectedItem = ((ListViewItem)listViewItem).Content;

            base.OnPreviewMouseDown(e);
        }
    }
}
