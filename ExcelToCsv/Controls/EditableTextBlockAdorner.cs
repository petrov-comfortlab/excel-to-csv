using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace ExcelToCsv.Controls
{
    /// <inheritdoc />
    /// <summary>
    /// Adorner class which shows textbox over the text block when the Edit mode is on.
    /// </summary>
    public class EditableTextBlockAdorner : Adorner
    {
        private readonly VisualCollection _collection;
        private readonly TextBox _textBox;
        private readonly TextBlock _textBlock;

        public EditableTextBlockAdorner(EditableTextBlock adornedElement) : base(adornedElement)
        {
            var binding = new Binding("Text") {Source = adornedElement};
            _textBlock = adornedElement;

            _textBox = new TextBox
            {
                TextWrapping = TextWrapping.NoWrap,
            };

            _textBox.SetBinding(TextBox.TextProperty, binding);
            _textBox.MaxLength = adornedElement.MaxLength;

            _collection = new VisualCollection(this) {_textBox};
        }

        protected override Visual GetVisualChild(int index)
        {
            return _collection[index];
        }

        protected override int VisualChildrenCount => _collection.Count;

        protected override Size ArrangeOverride(Size finalSize)
        {
            _textBox.Arrange(new Rect(-3, -1, _textBlock.DesiredSize.Width + 50, _textBlock.DesiredSize.Height));
            _textBox.Focus();
            _textBox.CaretIndex = _textBox.Text.Length;
            return finalSize;
        }

        public event RoutedEventHandler TextBoxLostFocus
        {
            add => _textBox.LostFocus += value;
            remove => _textBox.LostFocus -= value;
        }

        public event KeyEventHandler TextBoxKeyUp
        {
            add => _textBox.KeyUp += value;
            remove => _textBox.KeyUp -= value;
        }
    }
}
