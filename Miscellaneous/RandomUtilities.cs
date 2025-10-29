using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Point = System.Drawing.Point;

namespace PointAC.Miscellaneous
{
    public static class RandomUtilities
    {
        public static bool IsClickInsideAppWindow(Window window, Point screenPoint)
        {
            var hwnd = new WindowInteropHelper(window).Handle;

            if (GetWindowRect(hwnd, out RECT rect))
            {
                return screenPoint.X >= rect.Left && screenPoint.X <= rect.Right &&
                              screenPoint.Y >= rect.Top && screenPoint.Y <= rect.Bottom;
            }

            return false;
        }

        public static T? FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && parent is not T)
                parent = VisualTreeHelper.GetParent(parent);
            return parent as T;
        }

        public static T? FindDescendant<T>(DependencyObject root) where T : DependencyObject
        {
            for (int i = 0, n = VisualTreeHelper.GetChildrenCount(root); i < n; i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T wanted) return wanted;
                var deeper = FindDescendant<T>(child);
                if (deeper != null) return deeper;
            }
            return null;
        }

        public static System.Drawing.Rectangle GetElementScreenBounds(Window window, FrameworkElement element)
        {
            if (!element.IsLoaded)
                return System.Drawing.Rectangle.Empty;

            var transform = element.TransformToAncestor(window);
            var topLeft = transform.Transform(new System.Windows.Point(0, 0));
            var bottomRight = transform.Transform(new System.Windows.Point(element.ActualWidth, element.ActualHeight));

            var screenTopLeft = window.PointToScreen(topLeft);
            var screenBottomRight = window.PointToScreen(bottomRight);

            return System.Drawing.Rectangle.FromLTRB(
                (int)screenTopLeft.X,
                (int)screenTopLeft.Y,
                (int)screenBottomRight.X,
                (int)screenBottomRight.Y
            );
        }

        public static double Distance(Point a, Point b)
        {
            int dx = a.X - b.X;
            int dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        public static bool IsValidBitmap(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            try
            {
                Uri uri = Uri.TryCreate(path, UriKind.Absolute, out var absoluteUri)
                    ? absoluteUri
                    : new Uri(Path.GetFullPath(path));

                using var stream = Application.GetResourceStream(uri)?.Stream
                                   ?? (File.Exists(path) ? File.OpenRead(path) : null);

                if (stream == null)
                    return false;

                var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.IgnoreColorProfile, BitmapCacheOption.None);
                var frame = decoder.Frames.FirstOrDefault();
                return frame != null && frame.PixelWidth > 0 && frame.PixelHeight > 0;
            }
            catch
            {
                return false;
            }
        }

        public static void ScrollIntoViewIfNotVisible(FrameworkElement element, ScrollViewer scroller)
        {
            element.Dispatcher.InvokeAsync(() =>
            {
                var presenter = FindDescendant<ScrollContentPresenter>(scroller);
                if (presenter == null)
                {
                    element.BringIntoView();
                    return;
                }

                var relativePos = element.TranslatePoint(new System.Windows.Point(0, 0), presenter);

                double elementTop = relativePos.Y;
                double elementBottom = elementTop + element.ActualHeight;

                double viewTop = 0;
                double viewBottom = scroller.ViewportHeight;

                bool isOutOfView = elementTop < viewTop + 10 || elementBottom > viewBottom - 10;

                if (isOutOfView)
                {
                    double targetOffset = scroller.VerticalOffset
                                         + (elementTop - (scroller.ViewportHeight / 2))
                                         + (element.ActualHeight / 2);

                    targetOffset = Math.Max(0, Math.Min(targetOffset, scroller.ScrollableHeight));
                    scroller.ScrollToVerticalOffset(targetOffset);
                }
            }, System.Windows.Threading.DispatcherPriority.Background);
        }

        public static bool IsShortcutPressed(KeyEventArgs e, HashSet<Key> shortcut)
        {
            bool ctrl = shortcut.Contains(Key.LeftCtrl) || shortcut.Contains(Key.RightCtrl);
            bool shift = shortcut.Contains(Key.LeftShift) || shortcut.Contains(Key.RightShift);
            bool alt = shortcut.Contains(Key.LeftAlt) || shortcut.Contains(Key.RightAlt);

            bool ctrlPressed = (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl));
            bool shiftPressed = (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift));
            bool altPressed = (Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt));

            var mainKeys = shortcut.Except(new[] {
                Key.LeftCtrl, Key.RightCtrl,
                Key.LeftShift, Key.RightShift,
                Key.LeftAlt, Key.RightAlt
            });

            return mainKeys.Contains(e.Key)
                && (!ctrl || ctrlPressed)
                && (!shift || shiftPressed)
                && (!alt || altPressed);
        }

        public static string GetFormattedShortcut(HashSet<Key> shortcut)
        {
            if (shortcut == null || !shortcut.Any())
                return string.Empty;

            var ordered = shortcut.OrderBy(k => k switch
            {
                Key.LeftCtrl or Key.RightCtrl => 0,
                Key.LeftShift or Key.RightShift => 1,
                Key.LeftAlt or Key.RightAlt => 2,
                Key.LWin or Key.RWin => 3,
                _ => 4
            });

            var formattedKeys = ordered.Select(key =>
            {
                return FormatKeyName(key);
            });

            return string.Join(" + ", formattedKeys);
        }

        public static string FormatKeyName(Key key)
        {
            if (key >= Key.D0 && key <= Key.D9)
                return ((int)key - (int)Key.D0).ToString();

            if (key >= Key.NumPad0 && key <= Key.NumPad9)
                return "Num" + ((int)key - (int)Key.NumPad0);

            return key switch
            {
                Key.LeftCtrl or Key.RightCtrl => "Ctrl",
                Key.LeftShift or Key.RightShift => "Shift",
                Key.LeftAlt or Key.RightAlt => "Alt",
                Key.LWin or Key.RWin => "Win",

                >= Key.F1 and <= Key.F24 => key.ToString().ToUpper(),

                Key.Return => "Enter",
                Key.Back => "Backspace",
                Key.Delete => "Del",
                Key.Insert => "Ins",
                Key.Tab => "Tab",
                Key.Space => "Space",
                Key.Escape => "Esc",
                Key.Capital => "CapsLock",

                Key.Home => "Home",
                Key.End => "End",
                Key.PageUp => "PgUp",
                Key.PageDown => "PgDn",
                Key.Up => "↑",
                Key.Down => "↓",
                Key.Left => "←",
                Key.Right => "→",

                Key.OemPlus => "Plus",
                Key.OemMinus => "Minus",
                Key.OemComma => ",",
                Key.OemPeriod => ".",
                Key.Oem1 => ";",
                Key.Oem2 => "/",
                Key.Oem3 => "`",
                Key.Oem4 => "[",
                Key.Oem5 => "\\",
                Key.Oem6 => "]",
                Key.Oem7 => "'",

                Key.Multiply => "Multiply",
                Key.Divide => "Divide",
                Key.Subtract => "Subtract",
                Key.Add => "Add",
                Key.Decimal => "Decimal",

                _ => key.ToString()
            };
        }

        #region Win32 Interop
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion
    }
}
