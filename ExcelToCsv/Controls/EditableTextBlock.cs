using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;

namespace ExcelToCsv.Controls
{
    public class EditableTextBlock : TextBlock
    {
        private bool _isCancel;
        private string _oldText;

        public bool IsInEditMode
        {
            get => (bool)GetValue(IsInEditModeProperty);
            set
            {
                SetValue(IsInEditModeProperty, value);

                if (value)
                {
                    _isCancel = false;
                    _oldText = Text;
                }
                else if (_isCancel)
                {
                    Text = _oldText;
                }
            }
        }

        private EditableTextBlockAdorner _adorner;

        public static readonly DependencyProperty IsInEditModeProperty =
            DependencyProperty.Register("IsInEditMode", typeof(bool), typeof(EditableTextBlock), new UIPropertyMetadata(false, IsInEditModeUpdate));

        private static void IsInEditModeUpdate(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            if (!(obj is EditableTextBlock textBlock))
                return;

            var layer = AdornerLayer.GetAdornerLayer(textBlock);

            if (textBlock.IsInEditMode)
            {
                if (null == textBlock._adorner)
                {
                    textBlock._adorner = new EditableTextBlockAdorner(textBlock);
                    textBlock._adorner.TextBoxKeyUp += textBlock.TextBoxKeyUp;
                    textBlock._adorner.TextBoxLostFocus += textBlock.TextBoxLostFocus;
                }
                layer.Add(textBlock._adorner);
            }
            else
            {
                var adorners = layer.GetAdorners(textBlock);
                if (adorners != null)
                {
                    foreach (var adorner in adorners)
                    {
                        if (adorner is EditableTextBlockAdorner)
                        {
                            layer.Remove(adorner);
                        }
                    }
                }

                var expression = textBlock.GetBindingExpression(TextProperty);
                expression?.UpdateTarget();
            }
        }

        public int MaxLength
        {
            get => (int)GetValue(MaxLengthProperty);
            set => SetValue(MaxLengthProperty, value);
        }

        public static readonly DependencyProperty MaxLengthProperty =
            DependencyProperty.Register("MaxLength", typeof(int), typeof(EditableTextBlock), new UIPropertyMetadata(0));

        private void TextBoxLostFocus(object sender, RoutedEventArgs e)
        {
            IsInEditMode = false;
        }

        private void TextBoxKeyUp(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                    IsInEditMode = false;
                    break;

                case Key.Escape:
                    _isCancel = true;
                    IsInEditMode = false;
                    break;
            }
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
                IsInEditMode = true;
        }
    }
}