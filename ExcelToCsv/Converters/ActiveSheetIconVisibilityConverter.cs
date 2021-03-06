﻿using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ExcelToCsv.Converters
{
    public class ActiveSheetIconVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is string text))
                return Visibility.Collapsed;

            return text.StartsWith("#") ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}