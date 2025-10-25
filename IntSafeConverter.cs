using System;
using System.Globalization;
using System.Windows.Data;

namespace PointAC
{
    public class IntSafeConverter : IValueConverter
    {
        // This will remember the last valid value per binding.
        private string _lastValidValue = "0";

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            _lastValidValue = value?.ToString() ?? "0";
            return _lastValidValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string input = value?.ToString() ?? string.Empty;

            if (int.TryParse(input, out int result))
            {
                _lastValidValue = input;
                return result;
            }

            return new RevertTextBindingResult(_lastValidValue);
        }

        private class RevertTextBindingResult
        {
            public string Value { get; }
            public RevertTextBindingResult(string value) => Value = value;
        }
    }
}
