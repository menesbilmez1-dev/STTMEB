using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace sttmerve_cc.Helpers
{
    public class BooleanToColorConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            bool isRunning = value is bool b && b;
            return isRunning ? Brushes.Green : Brushes.Red;
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class BooleanToStatusConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            bool isRunning = value is bool b && b;
            return isRunning ? "Çalışıyor" : "Durduruldu";
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
