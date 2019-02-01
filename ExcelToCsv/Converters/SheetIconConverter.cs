using System;
using System.Drawing;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ExcelToCsv.Resources;

namespace ExcelToCsv.Converters
{
    public class SheetIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is string text))
                return null;

            var bitmap = text.StartsWith("#")
                ? Icons.UnactiveSheet_16x16
                : Icons.Sheet_16x16;

            var imageSource = Bitmap2BitmapSource(bitmap);
            var image = new System.Windows.Controls.Image
            {
                Width = imageSource.Width,
                Height = imageSource.Height,
                Stretch = Stretch.None,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Source = imageSource
            };

            Logger.Log.Debug($"image: {image.Source.Height}");

            return image;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public static BitmapSource Bitmap2BitmapSource(Bitmap bitmap)
        {
            try
            {
                if (bitmap == null)
                    return null;

                return Imaging.CreateBitmapSourceFromHBitmap(
                    bitmap.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex);
                return null;
            }
        }
    }
}