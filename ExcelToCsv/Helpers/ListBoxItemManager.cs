using System.Windows;
using System.Windows.Controls;
using ExcelToCsv.Controls;

namespace ExcelToCsv.Helpers
{
    public class ListBoxItemManager
    {
        public static void RenameItem(ListBoxItem listBoxItem)
        {
            if (listBoxItem == null)
                return;

            var dependencyObject = (DependencyObject)listBoxItem;
            var editableTextBlockDependencyObject = DependencyObjectManager.GetFirstChildByType(dependencyObject, typeof(EditableTextBlock));

            if (editableTextBlockDependencyObject is EditableTextBlock editableTextBlock)
                editableTextBlock.IsInEditMode = true;
        }
    }
}