using PointAC.Management;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;

namespace PointAC
{
    public static class UIManager
    {
        public static void ApplyTheme(Window window, string theme, string backdrop)
        {
            window.ThemeMode = theme switch
            {
                "Light" => ThemeMode.Light,
                "Dark" => ThemeMode.Dark,
                _ => ThemeMode.System
            };
            BackdropManager.ApplyBackdrop(window, backdrop, theme);
        }

        public static void ApplyLanguage(Window window, string language)
        {
            if (window == null)
                return;

            string cultureName = language.Equals("System", StringComparison.OrdinalIgnoreCase)
                ? CultureInfo.CurrentUICulture.Name
                : language;

            var culture = new CultureInfo(cultureName);

            string[] candidates =
            {
                culture.Name,
                culture.TwoLetterISOLanguageName,
                "en-US"
            };

            ResourceDictionary? dict = null;

            foreach (var candidate in candidates)
            {
                try
                {
                    dict = TryLoadDictionary(candidate);
                    if (dict != null)
                        break;
                }
                catch { }
            }

            if (dict == null && !culture.Name.Equals("en-US", StringComparison.OrdinalIgnoreCase))
            {
                string langPrefix = culture.TwoLetterISOLanguageName;
                dict = TryFindSameLanguageVariant(langPrefix);
            }

            if (dict == null)
                return;

            var previous = Application.Current.Resources.MergedDictionaries
                .FirstOrDefault(d => d.Source?.OriginalString.Contains("/Localization/") == true);

            if (previous != null)
                Application.Current.Resources.MergedDictionaries.Remove(previous);

            Application.Current.Resources.MergedDictionaries.Add(dict);

            window.FlowDirection = culture.TextInfo.IsRightToLeft
                ? FlowDirection.RightToLeft
                : FlowDirection.LeftToRight;
        }

        private static ResourceDictionary? TryLoadDictionary(string cultureCode)
        {
            try
            {
                var uri = new Uri($"pack://application:,,,/Localization/{cultureCode}.xaml", UriKind.Absolute);
                var dict = new ResourceDictionary { Source = uri };
                return dict;
            }
            catch
            {
                return null;
            }
        }

        private static ResourceDictionary? TryFindSameLanguageVariant(string langPrefix)
        {
            try
            {
                var assembly = typeof(UIManager).Assembly;
                string[] resourceNames = assembly
                    .GetManifestResourceNames()
                    .Where(n => n.Contains(".Localization.") && n.EndsWith(".xaml"))
                    .ToArray();

                string? match = resourceNames
                    .Select(r => r.Split(new[] { ".Localization." }, StringSplitOptions.None).Last())
                    .FirstOrDefault(name => name.StartsWith(langPrefix + "-", StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    string cultureFile = match.Replace(".xaml", "");
                    return TryLoadDictionary(cultureFile);
                }
            }
            catch { }

            return null;
        }
    }
}