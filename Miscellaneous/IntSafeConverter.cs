using System.Windows.Data;
using System.Globalization;

namespace PointAC.Miscellaneous
{
    public class IntSafeConverter : IValueConverter
    {
        private string lastValidInput = "0";

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            lastValidInput = value?.ToString() ?? "0";
            return lastValidInput;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string input = value?.ToString() ?? string.Empty;

            if (int.TryParse(input, out int result))
            {
                lastValidInput = input;
                return result;
            }

            return new RevertTextBindingResult(lastValidInput);
        }

        private class RevertTextBindingResult
        {
            public string Value { get; }
            public RevertTextBindingResult(string value) => Value = value;
        }
    }
}
