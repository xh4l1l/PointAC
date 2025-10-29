using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace PointAC.Management
{
    public static class BackdropManager
    {
        private enum DWMSBT
        {
            Auto = 0,
            None = 1,
            Mica = 2,
            Acrylic = 3,
            Tabbed = 4
        }

        private enum DWMWA : uint
        {
            USE_IMMERSIVE_DARK_MODE = 20,
            SYSTEMBACKDROP_TYPE = 38
        }

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, DWMWA attr, ref int pv, int cb);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, DWMWA attr, out int pv, int cb);

        public static void ApplyBackdrop(Window window, string backdrop, string theme)
        {
            var hwnd = new WindowInteropHelper(window).EnsureHandle();

            int requestedType = backdrop switch
            {
                "None" => (int)DWMSBT.None,
                "Mica" => (int)DWMSBT.Mica,
                "Tabbed" => (int)DWMSBT.Tabbed,
                "MicaAlt" => (int)DWMSBT.Tabbed,
                "Acrylic" => (int)DWMSBT.Acrylic,
                _ => (int)DWMSBT.Mica
            };

            int darkMode = theme.Equals("Dark", StringComparison.OrdinalIgnoreCase) ? 1 :
                           theme.Equals("Light", StringComparison.OrdinalIgnoreCase) ? 0 :
                           GetSystemDarkMode() ? 1 : 0;

            DwmSetWindowAttribute(hwnd, DWMWA.USE_IMMERSIVE_DARK_MODE, ref darkMode, sizeof(int));

            if (!TrySetBackdrop(hwnd, requestedType))
            {
                if (requestedType == (int)DWMSBT.Acrylic)
                {
                    if (!TrySetBackdrop(hwnd, (int)DWMSBT.Mica))
                    {
                        TrySetBackdrop(hwnd, (int)DWMSBT.None);
                    }
                }
                else if (requestedType == (int)DWMSBT.Mica)
                {
                    TrySetBackdrop(hwnd, (int)DWMSBT.None);
                }
            }
        }

        private static bool TrySetBackdrop(IntPtr hwnd, int type)
        {
            try
            {
                int hr = DwmSetWindowAttribute(hwnd, DWMWA.SYSTEMBACKDROP_TYPE, ref type, sizeof(int));
                if (hr == 0)
                {
                    if (DwmGetWindowAttribute(hwnd, DWMWA.SYSTEMBACKDROP_TYPE, out int applied, sizeof(int)) == 0)
                        return applied == type;
                }
            }
            catch { }
            return false;
        }

        private static bool GetSystemDarkMode()
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                return (int?)key?.GetValue("AppsUseLightTheme", 1) == 0;
            }
            catch { return false; }
        }
    }
}